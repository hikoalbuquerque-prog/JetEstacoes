/**
 * campo_v3_backend_gs.gs
 * Adicionar como arquivo novo no projeto GAS.
 * Contem: getRankingCampoHoje, editarEstacaoFromMapa, geocodeEndereco
 */

/* ----------------------------------------------------------------
 * getRankingCampoHoje
 * Retorna array { email, nome, total } ordenado desc por cadastros do dia
 * ---------------------------------------------------------------- */
function getRankingCampoHoje() {
  try {
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('Estacoes') || ss.getSheetByName('ESTACOES');
    if (!aba) return [];

    var data    = aba.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });

    // Encontrar colunas
    var iOp   = headers.indexOf('CriadoPor');
    if (iOp < 0) iOp = headers.indexOf('Operador');
    var iData = headers.indexOf('DataCriacao');

    if (iOp < 0) return [];

    var hoje = new Date(); hoje.setHours(0,0,0,0);
    var ranking = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var op  = String(row[iOp] || '').trim();
      if (!op) continue;
      if (iData >= 0) {
        var d = new Date(row[iData]);
        if (isNaN(d.getTime()) || d < hoje) continue;
      }
      ranking[op] = (ranking[op] || 0) + 1;
    }

    // Buscar nomes na aba USUARIOS
    var nomes = {};
    try {
      var uAba  = ss.getSheetByName('USUARIOS');
      if (uAba) {
        var uData = uAba.getDataRange().getValues();
        for (var j = 1; j < uData.length; j++) {
          var email = String(uData[j][0] || '').trim().toLowerCase();
          var nome  = String(uData[j][1] || '').trim();
          if (email) nomes[email] = nome || email;
        }
      }
    } catch(e) {}

    return Object.keys(ranking).map(function(email) {
      return { email: email, nome: nomes[email.toLowerCase()] || email, total: ranking[email] };
    }).sort(function(a,b){ return b.total - a.total; });

  } catch(e) {
    Logger.log('getRankingCampoHoje erro: ' + e);
    return [];
  }
}

/* ----------------------------------------------------------------
 * editarEstacaoFromMapa
 * Edita uma estacao existente na aba Estacoes.
 * Payload: { codigo, tipo, endereco, observacoes, lat, lng, foto }
 * ---------------------------------------------------------------- */
function editarEstacaoFromMapa(payload) {
  try {
    var codigo = String((payload && payload.codigo) || '').trim();
    if (!codigo) return { ok: false, error: 'Codigo obrigatorio.' };

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('Estacoes') || ss.getSheetByName('ESTACOES');
    if (!aba) return { ok: false, error: 'Aba Estacoes nao encontrada.' };

    var data    = aba.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });

    // Mapear colunas pelos nomes reais da planilha
    function col(name) { return headers.indexOf(name); }

    var iCod  = col('CodigoEstacao');
    var iTipo = col('TipoEstacao');
    var iEnd  = col('Endereço completo da estação');
    var iObs  = col('ObservacaoPrivado');
    var iMod  = col('Modalidade');
    var iDim  = col('Dimensões da Estação');
    var iLarg = col('Largura da Faixa Livre (m)');
    var iLat  = col('Latitude');
    var iLng  = col('Longitude');
    var iLoc  = col('Localização');
    var iFoto = col('Foto da Estação');

    if (iCod < 0) {
      // Fallback: tentar nomes alternativos
      iCod = col('Codigo') >= 0 ? col('Codigo') : col('codigo');
    }
    if (iCod < 0) return { ok: false, error: 'Coluna CodigoEstacao nao encontrada. Verifique os headers da aba Estacoes.' };

    // Encontrar a linha com o codigo
    for (var i = 1; i < data.length; i++) {
      var cellCod = String(data[i][iCod] || '').trim();
      if (cellCod !== codigo) continue;

      var row = i + 1; // linha 1-indexed na planilha

      if (iTipo >= 0 && payload.tipo)
        aba.getRange(row, iTipo + 1).setValue(payload.tipo);

      if (iEnd >= 0 && payload.endereco)
        aba.getRange(row, iEnd + 1).setValue(payload.endereco);

      // Modalidade -- coluna propria
      if (iMod >= 0 && payload.modalidade) {
        var modVal = String(payload.modalidade).toUpperCase().trim() || 'PATINETE';
        aba.getRange(row, iMod + 1).setValue(modVal);
      }
      // Dimensoes e Largura
      if (iDim >= 0 && payload.dimensoes !== undefined)
        aba.getRange(row, iDim + 1).setValue(payload.dimensoes);
      if (iLarg >= 0 && payload.larguraFaixa !== undefined)
        aba.getRange(row, iLarg + 1).setValue(payload.larguraFaixa);
      // Observacoes
      if (iObs >= 0 && payload.observacoes !== undefined) {
        aba.getRange(row, iObs + 1).setValue(payload.observacoes);
      }

      if (iLat >= 0 && payload.lat)
        aba.getRange(row, iLat + 1).setValue(Number(payload.lat));

      if (iLng >= 0 && payload.lng)
        aba.getRange(row, iLng + 1).setValue(Number(payload.lng));

      if (iLoc >= 0 && payload.lat && payload.lng)
        aba.getRange(row, iLoc + 1).setValue(payload.lat + ',' + payload.lng);

      // Foto nova
      if (payload.foto && payload.foto.base64) {
        try {
          var bytes = Utilities.base64Decode(payload.foto.base64);
          if (bytes.length > 8 * 1024 * 1024)
            return { ok: false, error: 'Foto muito grande.' };

          var blob  = Utilities.newBlob(bytes, payload.foto.mime || 'image/jpeg', 'foto_' + codigo + '.jpg');
          var pasta = DriveApp.getFolderById(TARGET_FOTOS_FOLDER_ID);
          var arq   = pasta.createFile(blob);
          arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

          var fotoUrl = 'https://drive.google.com/uc?id=' + arq.getId();
          if (iFoto >= 0) aba.getRange(row, iFoto + 1).setValue(fotoUrl);

        } catch(fe) {
          Logger.log('editarEstacaoFromMapa foto erro: ' + fe);
          // nao bloquear o resto da edicao por erro na foto
        }
      }

      return { ok: true, codigo: codigo };
    }

    return { ok: false, error: 'Estacao com codigo "' + codigo + '" nao encontrada.' };

  } catch(e) {
    Logger.log('editarEstacaoFromMapa erro: ' + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/* ----------------------------------------------------------------
 * geocodeEndereco
 * Forward geocode chamado via google.script.run no campo.
 * Retorna { ok, lat, lng, endereco, bairro, cidade, pais }
 * ---------------------------------------------------------------- */
function geocodeEndereco(enderecoStr) {
  try {
    var key = getMapsApiKey_();
    var url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?address=' + encodeURIComponent(enderecoStr)
      + '&key='     + key;

    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    if (!data.results || !data.results.length) return { ok: false, error: 'Nao encontrado.' };

    var r     = data.results[0];
    var loc   = r.geometry.location;
    var comps = r.address_components || [];

    function get(type) {
      var f = comps.filter(function(c){ return c.types.indexOf(type) !== -1; })[0];
      return f ? f.long_name : '';
    }

    return {
      ok:       true,
      lat:      loc.lat,
      lng:      loc.lng,
      endereco: r.formatted_address || '',
      bairro:   get('sublocality_level_1') || get('neighborhood') || '',
      cidade:   get('locality') || get('administrative_area_level_2') || '',
      estado:   get('administrative_area_level_1') || '',
      pais:     get('country') === 'Mexico' ? 'MX' : 'BR'
    };
  } catch(e) {
    return { ok: false, error: String(e) };
  }
}

/* ----------------------------------------------------------------
 * excluirEstacaoFromMapa
 * Remove uma estacao da aba Estacoes pelo CodigoEstacao.
 * Payload: { codigo, email }
 * ---------------------------------------------------------------- */
function excluirEstacaoFromMapa(payload) {
  try {
    var codigo = String((payload && payload.codigo) || '').trim();
    if (!codigo) return { ok: false, error: 'Codigo obrigatorio.' };

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('Estacoes') || ss.getSheetByName('ESTACOES');
    if (!aba) return { ok: false, error: 'Aba Estacoes nao encontrada.' };

    var data    = aba.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var iCod    = headers.indexOf('CodigoEstacao');
    if (iCod < 0) iCod = headers.indexOf('Codigo');
    if (iCod < 0) return { ok: false, error: 'Coluna CodigoEstacao nao encontrada.' };

    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][iCod]).trim() === codigo) {
        aba.deleteRow(i + 1);
        Logger.log('Estacao excluida: ' + codigo);
        return { ok: true, codigo: codigo };
      }
    }

    return { ok: false, error: 'Estacao "' + codigo + '" nao encontrada.' };

  } catch(e) {
    Logger.log('excluirEstacaoFromMapa erro: ' + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/**
 * PATCH MANUAL NO Código.gs — addEstacaoFromMapa
 *
 * Localizar a linha:
 *   setCell(newRow, COL.ObservacaoPrivado, obsFinal);
 *
 * ADICIONAR logo após:
 *   // Dimensoes e Largura da Faixa Livre
 *   if (payload.dimensoes) setCell(newRow, 'Dimensões da Estação', payload.dimensoes);
 *   if (payload.larguraFaixa) setCell(newRow, 'Largura da Faixa Livre (m)', Number(payload.larguraFaixa) || payload.larguraFaixa);
 */
