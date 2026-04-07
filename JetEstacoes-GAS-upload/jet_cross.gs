function cruzarJetComEstacoes(cidade) {
  var ss = SpreadsheetApp.getActive();
 
  var shJet = ss.getSheetByName('JET_IMPORT');
  var shEst = ss.getSheetByName('Estacoes');
  var shOut = ss.getSheetByName('JET_X_ESTACOES');
 
  if (!shJet || !shEst || !shOut) {
    throw new Error('Abas JET_IMPORT, Estacoes ou JET_X_ESTACOES nao encontradas');
  }
 
  // Usa normalizarTexto() de utils_math.gs
  var cidadeNorm = cidade ? normalizarTexto(cidade) : null;
 
  var jet = shJet.getDataRange().getValues();
  var est = shEst.getDataRange().getValues();
 
  var hJet = jet[0];
  var hEst = est[0];
 
  var jLat = hJet.indexOf('JetLat');
  var jLng = hJet.indexOf('JetLng');
  var jCid = hJet.indexOf('JetCidade');
 
  var eLat = hEst.indexOf('Latitude');
  var eLng = hEst.indexOf('Longitude');
  var eCod = hEst.indexOf('CodigoEstacao');
  var eCid = hEst.indexOf('Cidade');
 
  shOut.clearContents();
  shOut.appendRow([
    'Cidade', 'Origem',
    'JetLat', 'JetLng',
    'EstacaoCodigo', 'EstacaoLat', 'EstacaoLng',
    'DistanciaMetros', 'MatchNivel', 'StatusComparacao'
  ]);
 
  // Usa parseNumSafe() de utils_math.gs
  var estacoes = est.slice(1)
    .filter(function(r) {
      if (!cidadeNorm) return true;
      return normalizarTexto(r[eCid]) === cidadeNorm;
    })
    .map(function(r) {
      return {
        codigo:  r[eCod],
        lat:     parseNumSafe(r[eLat]),
        lng:     parseNumSafe(r[eLng]),
        matched: false
      };
    })
    .filter(function(e) { return isFinite(e.lat) && isFinite(e.lng); });
 
  jet.slice(1).forEach(function(r) {
    if (cidadeNorm && normalizarTexto(r[jCid]) !== cidadeNorm) return;
 
    var lat = parseNumSafe(r[jLat]);
    var lng = parseNumSafe(r[jLng]);
    if (!isFinite(lat) || !isFinite(lng)) return;
 
    var best = null;
 
    estacoes.forEach(function(e) {
      // Usa distanciaMetros() de utils_math.gs
      var d = distanciaMetros(lat, lng, e.lat, e.lng);
      if (d <= 30 && (!best || d < best.d)) {
        best = { d: d, e: e };
      }
    });
 
    if (best) {
      best.e.matched = true;
      shOut.appendRow([
        cidade || 'TODAS', 'JET',
        lat, lng,
        best.e.codigo, best.e.lat, best.e.lng,
        Math.round(best.d),
        best.d <= 15 ? 'FORTE' : 'MEDIO',
        best.d <= 15 ? 'MATCH_OK' : 'DIVERGENTE'
      ]);
    } else {
      shOut.appendRow([
        cidade || 'TODAS', 'JET',
        lat, lng,
        '', '', '',
        '',
        'NENHUM', 'JET_ONLY'
      ]);
    }
  });
 
  estacoes.forEach(function(e) {
    if (!e.matched) {
      shOut.appendRow([
        cidade || 'TODAS', 'ESTACAO',
        '', '',
        e.codigo, e.lat, e.lng,
        '',
        'NENHUM', 'ESTACAO_ONLY'
      ]);
    }
  });
 
  return {
    cidade:        cidade || 'TODAS',
    jetTotal:      jet.length - 1,
    estacoesTotal: estacoes.length,
    gerados:       shOut.getLastRow() - 1
  };
}
 

































