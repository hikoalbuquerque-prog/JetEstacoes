// * JET CROSS vFinal - Retorna APENAS NOVO e CANCELADO - Cache por cidade - Sem efeitos colaterais

function getJetCrossMapa(cidade) {
 
  // ========== CACHE ==========
  var cache      = CacheService.getScriptCache();
  var cidadeKey  = cidade ? normalizarTexto(cidade) : 'ALL';
  var cacheKey   = 'JET_CROSS::CIDADE::' + cidadeKey;
 
  var cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }
 
  // ========== PLANILHAS ==========
  var ss    = SpreadsheetApp.getActive();
  var shJet = ss.getSheetByName('JET_IMPORT');
  var shEst = ss.getSheetByName('Estacoes');
 
  if (!shJet || !shEst) {
    throw new Error('Abas JET_IMPORT ou Estacoes nao encontradas');
  }
 
  var cidadeNorm = cidade ? normalizarTexto(cidade) : null;
 
  // ========== LEITURA ==========
  var jet  = shJet.getDataRange().getValues();
  var est  = shEst.getDataRange().getValues();
  var hJet = jet.shift();
  var hEst = est.shift();
 
  var jLat = hJet.indexOf('JetLat');
  var jLng = hJet.indexOf('JetLng');
  var jCid = hJet.indexOf('JetCidade');
 
  var eLat = hEst.indexOf('Latitude');
  var eLng = hEst.indexOf('Longitude');
  var eCod = hEst.indexOf('CodigoEstacao');
  var eCid = hEst.indexOf('Cidade');
 
  // ========== ESTACOES ==========
  // Usa parseNumSafe() e normalizarTexto() de utils_math.gs
  var estacoes = est
    .filter(function(r) {
      if (!cidadeNorm) return true;
      return normalizarTexto(r[eCid]) === cidadeNorm;
    })
    .map(function(r) {
      return {
        codigo: r[eCod] || '',
        lat:    parseNumSafe(r[eLat]),
        lng:    parseNumSafe(r[eLng])
      };
    })
    .filter(function(e) { return isFinite(e.lat) && isFinite(e.lng); });
 
  // ========== CRUZAMENTO ==========
  var resultado = [];
 
  jet.forEach(function(r) {
    if (cidadeNorm && normalizarTexto(r[jCid]) !== cidadeNorm) return;
 
    var lat = parseNumSafe(r[jLat]);
    var lng = parseNumSafe(r[jLng]);
    if (!isFinite(lat) || !isFinite(lng)) return;
 
    var melhor = null;
 
    estacoes.forEach(function(e) {
      // Usa distanciaMetros() de utils_math.gs
      var d = distanciaMetros(lat, lng, e.lat, e.lng);
      if (d <= 30 && (!melhor || d < melhor.distancia)) {
        melhor = { codigo: e.codigo, distancia: Math.round(d) };
      }
    });
 
    // JET CROSS vFinal — retorna apenas NOVO e CANCELADO
    if (melhor) {
      if (melhor.distancia > 15) {
        resultado.push({
          lat:       lat,
          lng:       lng,
          status:    'NOVO',
          codigo:    melhor.codigo,
          distancia: melhor.distancia
        });
      }
      // distancia <= 15 = MATCH_OK — nao retornado intencionalmente
    } else {
      resultado.push({
        lat:       lat,
        lng:       lng,
        status:    'CANCELADO',
        codigo:    '',
        distancia: null
      });
    }
  });
 
  // ========== CACHE SAVE ==========
  cache.put(cacheKey, JSON.stringify(resultado), 900); // 15 min
 
  return resultado;
}
 
/**
 * DEBUG — JET CROSS vFinal
 * Validacao isolada (rodar no Editor do Apps Script)
 */
function debugJetCross() {
  var testes = [null, '', 'Sao Paulo'];
 
  testes.forEach(function(cidade) {
    var r = getJetCrossMapa(cidade);
    Logger.log('---------------------------');
    Logger.log('Cidade: ' + (cidade || '(todas)'));
    Logger.log('Total: ' + r.length);
    Logger.log(JSON.stringify(r.slice(0, 5), null, 2));
  });
}
 

































