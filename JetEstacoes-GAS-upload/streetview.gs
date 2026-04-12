/**
 * streetview_gs.gs
 * Endpoint para gerar Street View de uma estação individual a partir do InfoWindow.
 * Adicionar como novo arquivo no projeto GAS.
 *
 * Como usar:
 *   - Clique em "Gerar Street View" no InfoWindow de qualquer estação sem foto de rua
 *   - O backend chama a Street View Static API, salva no Drive e grava na coluna
 *   - O InfoWindow atualiza na hora com o link
 *
 * Dependências no projeto:
 *   - fetchStreetView_CORE()  → croqui_core_gs.gs
 *   - IMAGENS_FOLDER_ID       → Código.gs (constante já declarada)
 *   - SPREADSHEET_ID          → Código.gs
 *   - SHEET_NAME              → Código.gs ('Estacoes')
 *   - COL.ImgStreet           → Código.gs ('Street View')
 */

function gerarStreetViewEstacao(payload) {
  try {
    var codigo = String((payload && payload.codigo) || '').trim();
    var lat    = Number((payload && payload.lat)    || 0);
    var lng    = Number((payload && payload.lng)    || 0);

    if (!codigo) return { ok: false, error: 'Codigo obrigatorio.' };
    if (!isFinite(lat) || !isFinite(lng) || lat === 0)
      return { ok: false, error: 'Coordenadas invalidas.' };

    // 1. Gerar imagem via Street View Static API
    var blob;
    try {
      blob = fetchStreetView_CORE(lat, lng, {
        size:    '640x640',
        fov:     90,
        heading: 0,
        pitch:   0
      });
    } catch(apiErr) {
      return { ok: false, error: 'Street View API: ' + String(apiErr.message || apiErr) };
    }

    // 2. Salvar no Drive
    var nomeArq  = String(codigo).replace(/[^a-zA-Z0-9_-]/g, '_') + '_STREET.png';
    var pasta    = DriveApp.getFolderById(IMAGENS_FOLDER_ID);
    var file     = pasta.createFile(blob.setName(nomeArq));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileUrl  = file.getUrl();

    // 3. Gravar na planilha -- coluna 'Street View'
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba     = ss.getSheetByName(SHEET_NAME);
    if (!aba) return { ok: false, error: 'Aba Estacoes nao encontrada.' };

    var dados   = aba.getDataRange().getValues();
    var headers = dados[0].map(function(h){ return String(h).trim(); });

    var iCod    = headers.indexOf('CodigoEstacao');
    var iStreet = headers.indexOf(COL.ImgStreet);   // 'Street View'

    if (iCod < 0)    return { ok: false, error: 'Coluna CodigoEstacao nao encontrada.' };
    if (iStreet < 0) return { ok: false, error: 'Coluna "Street View" nao encontrada na planilha.' };

    var atualizado = false;
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][iCod] || '').trim() === codigo) {
        aba.getRange(i + 1, iStreet + 1).setValue(fileUrl);
        atualizado = true;
        break;
      }
    }

    if (!atualizado) {
      // Estacao nao encontrada -- pode ser uma nova (SOLICITADO sem codigo definitivo)
      // Gravar assim mesmo e retornar a URL para o frontend usar
      Logger.log('gerarStreetViewEstacao: estacao ' + codigo + ' nao encontrada na planilha.');
    }

    Logger.log('Street View gerado para ' + codigo + ': ' + fileUrl);
    return { ok: true, url: fileUrl, codigo: codigo };

  } catch(e) {
    Logger.log('gerarStreetViewEstacao erro: ' + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

function analisarCalcadaComGemini(params) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, error: 'GEMINI_API_KEY nao configurada' };

  var imagens = params.imagens || [];
  if (!imagens.length) return { ok: false, error: 'Sem imagens' };

  var parts = [{
    text: 'Analise estas imagens de calcada urbana. Avalie se tem faixa livre minima'
      + ' de 2.8 metros para patinetes. Retorne APENAS JSON:'
      + ' {"aprovado":true,"larguraEstimada":"X metros","observacoes":"texto","confianca":"alta","score":35}'
  }];

  imagens.forEach(function(b64) {
    parts.push({
      inline_data: { mime_type: 'image/jpeg', data: b64 }
    });
  });

  var payload = {
    contents: [{ parts: parts }],
    generationConfig: { maxOutputTokens: 300, temperature: 0.1 }
  };

  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + key;
    var resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code !== 200) return { ok: false, error: 'HTTP ' + code + ': ' + resp.getContentText().substring(0, 100) };

    var data = JSON.parse(resp.getContentText());
    var text = data.candidates[0].content.parts[0].text;
    var resultado = JSON.parse(text.replace(/```json|```/g, '').trim());
    return { ok: true, resultado: resultado };

  } catch(e) {
    return { ok: false, error: String(e) };
  }
}
