/**
 * reorganizar_curitiba_fotos.gs
 * Reorganiza APENAS as fotos de Curitiba para:
 *   Fotos_Estacoes / Curitiba / <Bairro> / Fotos / arquivo
 *
 * Filtra subpastas pelo bairros da aba Estacoes (Cidade = "Curitiba").
 * Subpastas de outras cidades permanecem intactas.
 *
 * Retomavel:
 *   iniciarReorgFotosCuritiba()    -- inicia do zero
 *   continuarReorgFotosCuritiba()  -- retoma de onde parou
 *   statusReorgFotosCuritiba()     -- progresso no Logger
 *   resetarReorgFotosCuritiba()    -- limpa estado salvo
 */

var REORG_FOTOS_CWB = {
  FOTOS_RAIZ:  '1hc5-whvQdYNqHbmkzW96MMnOjm4nk964',

  SHEET_NAME:  'Estacoes',
  COL_CIDADE:  'Cidade',
  COL_BAIRRO:  'Bairro',
  CIDADE_ALVO: 'Curitiba',

  KEY_BAIRROS: 'REORG_FOTOS_CWB_BAIRROS',
  KEY_IDX:     'REORG_FOTOS_CWB_IDX',
  KEY_STATUS:  'REORG_FOTOS_CWB_STATUS',

  LIMITE_MS: 5 * 60 * 1000
};

// ── Helpers ────────────────────────────────────────────────────────

function rfcNorm_(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function rfcGetOuCriar_(pasta, nome) {
  var it = pasta.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pasta.createFolder(nome);
}

function rfcMover_(arquivo, destino) {
  var pais = arquivo.getParents();
  destino.addFile(arquivo);
  while (pais.hasNext()) {
    var pai = pais.next();
    if (pai.getId() !== destino.getId()) pai.removeFile(arquivo);
  }
}

function rfcBairrosCuritiba_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(REORG_FOTOS_CWB.SHEET_NAME);
  if (!sh) throw new Error('Aba "' + REORG_FOTOS_CWB.SHEET_NAME + '" nao encontrada.');

  var data   = sh.getDataRange().getValues();
  var header = data[0].map(function(h) { return String(h).trim(); });
  var iCidade = header.indexOf(REORG_FOTOS_CWB.COL_CIDADE);
  var iBairro = header.indexOf(REORG_FOTOS_CWB.COL_BAIRRO);

  if (iCidade < 0 || iBairro < 0) {
    throw new Error('Colunas "Cidade" e/ou "Bairro" nao encontradas no cabecalho.');
  }

  var set = {};
  for (var i = 1; i < data.length; i++) {
    var cidade = String(data[i][iCidade] || '').trim();
    var bairro = String(data[i][iBairro] || '').trim();
    if (cidade === REORG_FOTOS_CWB.CIDADE_ALVO && bairro) {
      set[rfcNorm_(bairro)] = bairro;
    }
  }
  return set;
}

function rfcFiltrarSubpastas_(pastaRaiz, bairrosSet) {
  var result = [];
  var subs = pastaRaiz.getFolders();
  while (subs.hasNext()) {
    var sub  = subs.next();
    var nome = sub.getName();
    if (nome === REORG_FOTOS_CWB.CIDADE_ALVO) continue; // pula pasta Curitiba ja criada
    if (bairrosSet[rfcNorm_(nome)] !== undefined) {
      result.push({ id: sub.getId(), nome: nome });
    } else {
      Logger.log('[SKIP] ' + nome + ' (nao e bairro de Curitiba)');
    }
  }
  return result;
}

// ── Iniciar ────────────────────────────────────────────────────────

function iniciarReorgFotosCuritiba() {
  var props = PropertiesService.getScriptProperties();

  Logger.log('Lendo bairros de Curitiba na planilha...');
  var bairrosSet = rfcBairrosCuritiba_();
  Logger.log('Bairros unicos encontrados: ' + Object.keys(bairrosSet).length);

  var pastaRaiz = DriveApp.getFolderById(REORG_FOTOS_CWB.FOTOS_RAIZ);
  var bairros   = rfcFiltrarSubpastas_(pastaRaiz, bairrosSet);

  Logger.log('Subpastas que serao reorganizadas: ' + bairros.length);

  props.setProperty(REORG_FOTOS_CWB.KEY_BAIRROS, JSON.stringify(bairros));
  props.setProperty(REORG_FOTOS_CWB.KEY_IDX,     '0');
  props.setProperty(REORG_FOTOS_CWB.KEY_STATUS,  'RODANDO');

  _executarReorgFotosCuritiba_();
}

// ── Continuar ──────────────────────────────────────────────────────

function continuarReorgFotosCuritiba() {
  var props  = PropertiesService.getScriptProperties();
  var status = props.getProperty(REORG_FOTOS_CWB.KEY_STATUS);

  if (status === 'CONCLUIDO') {
    Logger.log('Ja concluido. Use resetarReorgFotosCuritiba() para reiniciar.');
    return;
  }
  if (!status) {
    Logger.log('Sem estado salvo. Use iniciarReorgFotosCuritiba() primeiro.');
    return;
  }
  Logger.log('Retomando...');
  _executarReorgFotosCuritiba_();
}

// ── Status ─────────────────────────────────────────────────────────

function statusReorgFotosCuritiba() {
  var props   = PropertiesService.getScriptProperties();
  var status  = props.getProperty(REORG_FOTOS_CWB.KEY_STATUS) || 'nao iniciado';
  var idx     = props.getProperty(REORG_FOTOS_CWB.KEY_IDX)    || '0';
  var raw     = props.getProperty(REORG_FOTOS_CWB.KEY_BAIRROS);
  var total   = raw ? JSON.parse(raw).length : 0;

  Logger.log('=== STATUS REORG FOTOS CURITIBA ===');
  Logger.log('Status : ' + status);
  Logger.log('Progresso: bairro ' + idx + ' / ' + total);
}

// ── Reset ──────────────────────────────────────────────────────────

function resetarReorgFotosCuritiba() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(REORG_FOTOS_CWB.KEY_BAIRROS);
  props.deleteProperty(REORG_FOTOS_CWB.KEY_IDX);
  props.deleteProperty(REORG_FOTOS_CWB.KEY_STATUS);
  Logger.log('Estado resetado.');
}

// ── Motor ──────────────────────────────────────────────────────────

function _executarReorgFotosCuritiba_() {
  var props   = PropertiesService.getScriptProperties();
  var inicio  = Date.now();

  var bairros = JSON.parse(props.getProperty(REORG_FOTOS_CWB.KEY_BAIRROS) || '[]');
  var idx     = parseInt(props.getProperty(REORG_FOTOS_CWB.KEY_IDX) || '0', 10);

  var pastaRaiz = DriveApp.getFolderById(REORG_FOTOS_CWB.FOTOS_RAIZ);
  var pastaCwb  = rfcGetOuCriar_(pastaRaiz, REORG_FOTOS_CWB.CIDADE_ALVO);

  for (var i = idx; i < bairros.length; i++) {

    if (Date.now() - inicio > REORG_FOTOS_CWB.LIMITE_MS) {
      props.setProperty(REORG_FOTOS_CWB.KEY_IDX, String(i));
      Logger.log('Tempo limite atingido no bairro ' + i + '/' + bairros.length);
      Logger.log('Execute continuarReorgFotosCuritiba() para retomar.');
      return;
    }

    var bairro       = bairros[i];
    var pastaBairro  = DriveApp.getFolderById(bairro.id);
    var pastaCwbB    = rfcGetOuCriar_(pastaCwb, bairro.nome);
    var pastaDestino = rfcGetOuCriar_(pastaCwbB, 'Fotos');

    var arqs = pastaBairro.getFiles();
    var n = 0;
    while (arqs.hasNext()) {
      rfcMover_(arqs.next(), pastaDestino);
      n++;
    }
    Logger.log('[OK] ' + bairro.nome + ': ' + n + ' foto(s) movida(s)');
  }

  props.setProperty(REORG_FOTOS_CWB.KEY_IDX,    String(bairros.length));
  props.setProperty(REORG_FOTOS_CWB.KEY_STATUS, 'CONCLUIDO');
  Logger.log('=== REORGANIZACAO FOTOS CURITIBA CONCLUIDA ===');
}

function verCidadesCaraguatatuba() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var count = {};
  data.slice(1).forEach(function(row) {
    var c = String(row[iCidade] || '').trim();
    if (c.toLowerCase().indexOf('caragu') !== -1 ||
        c.toLowerCase().indexOf('santo andr') !== -1) {
      count[c] = (count[c] || 0) + 1;
    }
  });
  Logger.log(JSON.stringify(count, null, 2));
}

function diagnosticarCidades() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var count = {};
  data.slice(1).forEach(function(row) {
    var c = String(row[iCidade] || '').trim();
    if (!c) return;
    count[c] = (count[c] || 0) + 1;
  });
  // Mostrar cidades com menos de 5 estacoes
  var pequenas = Object.keys(count)
    .filter(function(c){ return count[c] < 5; })
    .sort();
  Logger.log('Cidades com < 5 estacoes:');
  pequenas.forEach(function(c){ Logger.log(count[c] + 'x "' + c + '"'); });
  Logger.log('Total cidades: ' + Object.keys(count).length);
}

function normalizarNomesCidades() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  
  var mapa = {
    'Del Valle Nte':        'Del Valle Norte',
    'Mexico City':          'Ciudad de México',
    'cmdx':                 'Ciudad de México',
    'Naucalpan de Juárez':  'Naucalpan',
    'Del. Cuauhtemoc':      'Cuauhtémoc'
  };
  
  var corrigidos = 0;
  for (var i = 1; i < data.length; i++) {
    var c = String(data[i][iCidade] || '').trim();
    if (mapa[c]) {
      sh.getRange(i + 1, iCidade + 1).setValue(mapa[c]);
      corrigidos++;
    }
  }
  Logger.log('Corrigidos: ' + corrigidos);
}

function contarSaoPaulo() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var n = 0;
  data.slice(1).forEach(function(row) {
    if (String(row[iCidade]).trim() === 'São Paulo') n++;
  });
  Logger.log('São Paulo: ' + n + ' estacoes');
}

function diagnosticarCoordsInteiras() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var iLat = header.indexOf('Latitude');
  var iLng = header.indexOf('Longitude');
  var iEnd = header.indexOf('Endereço completo da estação');
  
  var inteiras = 0, semEnd = 0;
  for (var i = 1; i < data.length; i++) {
    var cidade = String(data[i][iCidade] || '').trim();
    if (cidade !== 'São Paulo') continue;
    var lat = Number(data[i][iLat]);
    var lng = Number(data[i][iLng]);
    if (lat === Math.round(lat) && lng === Math.round(lng)) {
      inteiras++;
      if (!data[i][iEnd]) semEnd++;
    }
  }
  Logger.log('SP coords inteiras: ' + inteiras);
  Logger.log('SP sem endereço: ' + semEnd);
}

function preencherPaisPorCoords() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  
  var iPais = header.indexOf('Pais');
  var iLat  = header.indexOf('Latitude');
  var iLng  = header.indexOf('Longitude');
  
  if (iPais < 0) { Logger.log('Coluna Pais nao encontrada'); return; }
  
  var atualizados = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][iPais]) continue; // ja tem pais, pula
    
    var lat = parseFloat(String(data[i][iLat]).replace(',', '.'));
    var lng = parseFloat(String(data[i][iLng]).replace(',', '.'));
    if (!isFinite(lat) || !isFinite(lng)) continue;
    
    var pais = 'BR';
    if (lat >= 14.5 && lat <= 32.7 && lng >= -117.1 && lng <= -86.7) pais = 'MX';
    
    sh.getRange(i + 1, iPais + 1).setValue(pais);
    atualizados++;
  }
  
  Logger.log('Pais preenchido em ' + atualizados + ' linhas');
}

function corrigirCoordsInteirasSP() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];

  var iCidade = header.indexOf('Cidade');
  var iLat    = header.indexOf('Latitude');
  var iLng    = header.indexOf('Longitude');
  var iLoc    = header.indexOf('Localização');
  var iEnd    = header.indexOf('Endereço completo da estação');

  var key = PropertiesService.getScriptProperties().getProperty('GMAPS_API_KEY');
  var corrigidos = 0;
  var erros = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iCidade]).trim() !== 'São Paulo') continue;

    var lat = Number(data[i][iLat]);
    var lng = Number(data[i][iLng]);

    // Pular se já tem decimal
    if (lat !== Math.round(lat) || lng !== Math.round(lng)) continue;

    var endereco = String(data[i][iEnd] || '').trim();
    if (!endereco) { erros++; continue; }

    // Geocode pelo endereço
    var url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?address=' + encodeURIComponent(endereco + ', São Paulo, Brasil')
      + '&key=' + key;

    try {
      var resp = JSON.parse(UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getContentText());
      if (resp.status === 'OK' && resp.results.length) {
        var loc = resp.results[0].geometry.location;
        sh.getRange(i + 1, iLat + 1).setValue(loc.lat);
        sh.getRange(i + 1, iLng + 1).setValue(loc.lng);
        if (iLoc >= 0) sh.getRange(i + 1, iLoc + 1).setValue(loc.lat + ',' + loc.lng);
        corrigidos++;
      } else {
        erros++;
      }
    } catch(e) {
      erros++;
    }

    // Evitar rate limit
    if (corrigidos % 10 === 0) Utilities.sleep(500);

    // Salvar progresso a cada 50
    if (corrigidos % 50 === 0) {
      PropertiesService.getScriptProperties().setProperty('CORRIGIR_SP_ROW', String(i));
      Logger.log('Progresso: ' + corrigidos + ' corrigidos, linha ' + i);
    }
  }

  Logger.log('Concluido: ' + corrigidos + ' corrigidos, ' + erros + ' erros');
}

function corrigirCoordsInteirasSP() {
  var props = PropertiesService.getScriptProperties();
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];

  var iCidade = header.indexOf('Cidade');
  var iLat    = header.indexOf('Latitude');
  var iLng    = header.indexOf('Longitude');
  var iLoc    = header.indexOf('Localização');
  var iEnd    = header.indexOf('Endereço completo da estação');

  var key = PropertiesService.getScriptProperties().getProperty('GMAPS_API_KEY');
  var startRow = Number(props.getProperty('CORRIGIR_SP_ROW') || 1);
  var HARD_LIMIT = 4 * 60 * 1000;
  var startTime = Date.now();
  var corrigidos = 0, erros = 0, pulados = 0;

  for (var i = startRow; i < data.length; i++) {

    if (Date.now() - startTime > HARD_LIMIT) {
      props.setProperty('CORRIGIR_SP_ROW', String(i));
      Logger.log('PAUSADO na linha ' + i + '. Corrigidos: ' + corrigidos + '. Rode continuarCorrigirSP()');
      return;
    }

    if (String(data[i][iCidade]).trim() !== 'São Paulo') continue;

    var latRaw = data[i][iLat];
    var latStr = String(latRaw || '');

    // Pular se já tem decimal
    var temDecimal = latStr.indexOf('.') !== -1 || latStr.indexOf(',') !== -1;
    if (temDecimal) { pulados++; continue; }

    var endereco = String(data[i][iEnd] || '').trim();
    // Remover emojis e prefixos do endereço
    endereco = endereco.replace(/[\u{1F300}-\u{1FFFF}]/gu, '').replace(/^[^a-zA-Z0-9\u00C0-\u017E]+/, '').trim();

    if (!endereco) { erros++; continue; }

    var url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?address=' + encodeURIComponent(endereco + ', São Paulo, SP, Brasil')
      + '&key=' + key;

    try {
      var resp = JSON.parse(UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getContentText());
      if (resp.status === 'OK' && resp.results && resp.results.length) {
        var loc = resp.results[0].geometry.location;
        // Validar que é realmente SP (bbox)
        if (loc.lat < -24.0 || loc.lat > -23.3 || loc.lng < -47.0 || loc.lng > -46.3) {
          erros++; continue;
        }
        sh.getRange(i + 1, iLat + 1).setValue(loc.lat);
        sh.getRange(i + 1, iLng + 1).setValue(loc.lng);
        if (iLoc >= 0) sh.getRange(i + 1, iLoc + 1).setValue(loc.lat + ',' + loc.lng);
        corrigidos++;
      } else {
        erros++;
      }
    } catch(e) {
      erros++;
    }

    Utilities.sleep(100);
  }

  props.deleteProperty('CORRIGIR_SP_ROW');
  Logger.log('CONCLUIDO. Corrigidos: ' + corrigidos + ', Erros: ' + erros + ', Pulados: ' + pulados);
}

function continuarCorrigirSP() {
  corrigirCoordsInteirasSP();
}

function statusCorrigirSP() {
  var row = PropertiesService.getScriptProperties().getProperty('CORRIGIR_SP_ROW');
  Logger.log(row ? 'Vai retomar da linha ' + row : 'Não iniciado ou concluído');
}

function resetarCorrigirSP() {
  PropertiesService.getScriptProperties().deleteProperty('CORRIGIR_SP_ROW');
  Logger.log('Reset feito');
}

function diagnosticarLatsSP() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var iLat = header.indexOf('Latitude');
  
  var inteiras = 0, decimais = 0, exemplos = [];
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iCidade]).trim() !== 'São Paulo') continue;
    var lat = data[i][iLat];
    var latStr = String(lat);
    var temDecimal = latStr.indexOf('.') !== -1 || latStr.indexOf(',') !== -1;
    if (temDecimal) {
      decimais++;
    } else {
      inteiras++;
      if (exemplos.length < 3) exemplos.push('tipo:' + typeof lat + ' valor:' + lat + ' string:"' + latStr + '"');
    }
  }
  
  Logger.log('Decimais: ' + decimais + ', Inteiras: ' + inteiras);
  Logger.log('Exemplos inteiras: ' + JSON.stringify(exemplos));
}

function verValoresReais() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var iLat = header.indexOf('Latitude');
  var iEnd = header.indexOf('Endereço completo da estação');
  
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iCidade]).trim() !== 'São Paulo') continue;
    var lat = data[i][iLat];
    if (lat === -23 || lat === -23.0) {
      count++;
      if (count <= 3) Logger.log('L' + (i+1) + ' lat=' + lat + ' end=' + String(data[i][iEnd]).substring(0,50));
    }
  }
  Logger.log('Total com lat==-23: ' + count);
}

function debugLatsSP() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var iCidade = header.indexOf('Cidade');
  var iLat = header.indexOf('Latitude');
  var iLng = header.indexOf('Longitude');
  var iLoc = header.indexOf('Localização');
  
  var inteiros = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iCidade]).trim() !== 'São Paulo') continue;
    var lat = Number(data[i][iLat]);
    var lng = Number(data[i][iLng]);
    if (lat === Math.round(lat)) {
      inteiros++;
      if (inteiros <= 3) {
        Logger.log('L'+(i+1)+' lat='+data[i][iLat]+' lng='+data[i][iLng]+' loc="'+data[i][iLoc]+'"');
      }
    }
  }
  Logger.log('Inteiros: ' + inteiros);
}

function normalizarLatLngSP() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  
  var iCidade = header.indexOf('Cidade');
  var iLat    = header.indexOf('Latitude');
  var iLng    = header.indexOf('Longitude');
  var iLoc    = header.indexOf('Localização');
  
  var corrigidos = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iCidade]).trim() !== 'São Paulo') continue;
    
    var latRaw = data[i][iLat];
    var lngRaw = data[i][iLng];
    
    // Converter string com vírgula para número
    var lat = parseFloat(String(latRaw).replace(',', '.'));
    var lng = parseFloat(String(lngRaw).replace(',', '.'));
    
    if (isNaN(lat) || isNaN(lng)) continue;
    if (lat === Number(latRaw) && lng === Number(lngRaw)) continue; // já é número correto
    
    sh.getRange(i + 1, iLat + 1).setValue(lat);
    sh.getRange(i + 1, iLng + 1).setValue(lng);
    if (iLoc >= 0) sh.getRange(i + 1, iLoc + 1).setValue(lat + ',' + lng);
    corrigidos++;
    
    if (corrigidos % 100 === 0) Logger.log('Progresso: ' + corrigidos);
  }
  
  Logger.log('Concluido: ' + corrigidos + ' corrigidos');
}

function normalizarLatLngTodos() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  
  var iLat = header.indexOf('Latitude');
  var iLng = header.indexOf('Longitude');
  var iLoc = header.indexOf('Localização');
  
  var corrigidos = 0;
  var HARD_LIMIT = 4 * 60 * 1000;
  var startTime = Date.now();
  var props = PropertiesService.getScriptProperties();
  var startRow = Number(props.getProperty('NORM_LAT_ROW') || 1);
  
  for (var i = startRow; i < data.length; i++) {
    if (Date.now() - startTime > HARD_LIMIT) {
      props.setProperty('NORM_LAT_ROW', String(i));
      Logger.log('PAUSADO linha ' + i + ', corrigidos: ' + corrigidos + '. Rode continuarNormLatLng()');
      return;
    }
    
    var latRaw = data[i][iLat];
    var lngRaw = data[i][iLng];
    if (latRaw === '' || latRaw === null || latRaw === undefined) continue;
    
    var lat = parseFloat(String(latRaw).replace(',', '.'));
    var lng = parseFloat(String(lngRaw).replace(',', '.'));
    
    if (isNaN(lat) || isNaN(lng)) continue;
    if (lat === Number(latRaw) && lng === Number(lngRaw)) continue;
    
    sh.getRange(i + 1, iLat + 1).setValue(lat);
    sh.getRange(i + 1, iLng + 1).setValue(lng);
    if (iLoc >= 0) sh.getRange(i + 1, iLoc + 1).setValue(lat + ',' + lng);
    corrigidos++;
  }
  
  props.deleteProperty('NORM_LAT_ROW');
  Logger.log('CONCLUIDO. Corrigidos: ' + corrigidos);
}

function continuarNormLatLng() {
  normalizarLatLngTodos();
}
