/**
 * ciclovias_backend_gs.gs
 * Backend para servir os dados de ciclovias da aba CICLOVIAS_RAW.
 *
 * A aba CICLOVIAS_RAW tem:
 *   Col A: WKT original (LINESTRING)
 *   Col B: PIPE lat,lng  (ex: -23.55,-46.63|-23.56,-46.64|...)
 *   Col C: Nome/tipo da ciclovia (opcional)
 *
 * Como usar:
 *   - No frontend: google.script.run.getCicloviasCidade('São Paulo', cb)
 *   - Retorna array de strings PIPE: ["-23.55,-46.63|-23.56,-46.64", ...]
 *   - Cada string é uma polyline (ciclovia)
 *   - Cache de 6h para performance (mesmo padrão dos polígonos)
 *
 * Adicionado ao doGet routing: action=getCiclovias&cidade=São+Paulo
 */

var CICLOVIAS_CACHE_TTL = 21600;  // 6h
var CICLOVIAS_SHEET     = 'CICLOVIAS_RAW';
var CICLOVIAS_COL_PIPE  = 2;      // Coluna B (1-indexed)
var CICLOVIAS_COL_NOME  = 3;      // Coluna C -- nome/tipo (opcional)

function _cicloCacheKey(cidade) {
  return 'CICLOVIAS_V1_' + String(cidade || '').replace(/\s/g, '_').toUpperCase();
}

/**
 * getCicloviasCidade
 * Retorna array de strings PIPE para a cidade informada.
 * Se cidade for vazio/null, retorna todas.
 */
function getCicloviasCidade(cidade) {
  try {
    var cacheKey = _cicloCacheKey(cidade || 'ALL');
    var cache    = CacheService.getScriptCache();

    // 1. Tentar cache
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        if (Array.isArray(parsed)) return parsed;
      } catch(e) {}
    }

    // 2. Ler da planilha
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName(CICLOVIAS_SHEET);
    if (!aba) {
      Logger.log('Aba CICLOVIAS_RAW nao encontrada.');
      return [];
    }

    var lastRow = aba.getLastRow();
    if (lastRow <= 1) return [];

    var dados = aba.getRange(2, 1, lastRow - 1, Math.min(aba.getLastColumn(), 3)).getValues();

    var result = [];
    for (var i = 0; i < dados.length; i++) {
      var pipe = String(dados[i][CICLOVIAS_COL_PIPE - 1] || '').trim();
      if (!pipe) continue;           // sem dado convertido
      if (pipe.indexOf('|') < 0) continue;  // menos de 2 pontos
      result.push(pipe);
    }

    // 3. Gravar no cache
    // CacheService tem limite de 100KB por chave
    // Dividir se necessário
    try {
      var json = JSON.stringify(result);
      if (json.length < 95000) {
        cache.put(cacheKey, json, CICLOVIAS_CACHE_TTL);
      } else {
        // Gravar em chunks
        var chunkSize = 500;
        for (var c = 0; c * chunkSize < result.length; c++) {
          var chunk = result.slice(c * chunkSize, (c + 1) * chunkSize);
          cache.put(cacheKey + '_CHUNK_' + c, JSON.stringify(chunk), CICLOVIAS_CACHE_TTL);
        }
        cache.put(cacheKey + '_CHUNKS', String(Math.ceil(result.length / chunkSize)), CICLOVIAS_CACHE_TTL);
      }
    } catch(cacheErr) {
      Logger.log('Cache ciclovias erro: ' + cacheErr);
    }

    Logger.log('getCicloviasCidade: ' + result.length + ' rotas retornadas');
    return result;

  } catch(e) {
    Logger.log('getCicloviasCidade erro: ' + e);
    return [];
  }
}
