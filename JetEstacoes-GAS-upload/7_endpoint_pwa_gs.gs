/**
 * webapp_pwa_endpoints.gs -- ARQUIVO NOVO
 *
 * Endpoints extras para o PWA de campo.
 *
 * COMO INTEGRAR AO doPost():
 * Use o arquivo 8_doGet_routing_gs.gs que ja contem
 * o doPost() e dispatchAction_() completos.
 * Este arquivo contem apenas as funcoes de negocio.
 *
 * DEPENDENCIAS: utils_auth.gs, Codigo.gs (getMapsApiKey_)
 */

/**
 * Reverse geocode para o PWA.
 * Nao requer auth -- apenas valida lat/lng.
 *
 * @param {{ lat:number, lng:number }} params
 * @returns {{ ok:boolean, geo?:object }}
 */
function reverseGeocodePWA(params) {
  try {
    params = params || {};
    var lat = Number(params.lat);
    var lng = Number(params.lng);

    if (!isFinite(lat) || !isFinite(lng)) {
      return { ok: false, error: 'Lat/Lng invalidos.' };
    }

    var key = getMapsApiKey_();
    var url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?latlng=' + lat + ',' + lng
      + '&key=' + key;

    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    if (!data.results || !data.results.length) {
      return { ok: false, error: 'Sem resultados para estas coordenadas.' };
    }

    var comps = data.results[0].address_components || [];
    var get = function(type) {
      var f = comps.filter(function(c) { return c.types.indexOf(type) !== -1; })[0];
      return f ? f.long_name : '';
    };

    var paisCod = get('country');
    var pais    = paisCod === 'MX' ? 'MX' : 'BR';

    return {
      ok: true,
      geo: {
        endereco: data.results[0].formatted_address || '',
        bairro:   get('sublocality_level_1') || get('neighborhood') || get('sublocality') || '',
        cidade:   get('locality') || get('administrative_area_level_2') || '',
        estado:   get('administrative_area_level_1') || '',
        pais:     pais,
        alcaldia: pais === 'MX' ? (get('administrative_area_level_2') || '') : ''
      }
    };

  } catch (e) {
    Logger.log('reverseGeocodePWA erro: ' + e);
    return { ok: false, error: String(e) };
  }
}

/**
 * doPost -- ponto de entrada unico do WebApp.
 * SUBSTITUIR o doPost() atual em Codigo.gs por este bloco.
 */
function doPost(e) {
  var result;
  try {
    var body   = e.postData ? e.postData.contents : '{}';
    var params = JSON.parse(body || '{}');
    var action = String((e.parameter && e.parameter.action) || '').trim();
    result = dispatchAction_(action, params);
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatchAction_(action, params) {
  params = params || {};

  switch (action) {
    // Auth
    case 'validateUser':
      return validateUser(params.tokenId);

    // Estações
    case 'addEstacaoFromMapa':
      return addEstacaoFromMapa(params);

    // PWA
    case 'reverseGeocodePWA':
      return reverseGeocodePWA(params);
    case 'getMinhasEstacoes':
      return getMinhasEstacoes(params);
    case 'getEstacoesEquipe':
      return getEstacoesEquipe(params);
    case 'updateStatusEstacao':
      return updateStatusEstacao(params);
    case 'getDashboard':
      return getDashboard(params);
    case 'salvarUsuario':
      return salvarUsuarioApi(params);
    case 'listarUsuarios':
      return listarUsuariosApi(params);

    // Mapa desktop
    case 'getEstacoesWebApp':
      return getEstacoesWebApp();
    case 'getCitiesIndex':
      return getCitiesIndex();
    case 'getEstacoesByCity':
      return getEstacoesByCity(params.cityKey);
    case 'getJetCrossMapa':
      return getJetCrossMapa(params.cidade);
    case 'validateAddPass':
      return { ok: validateAddPass(params.pass) };

    // Solicitações de acesso
    case 'solicitarAcesso':
      return solicitarAcesso(params);
    case 'aprovarSolicitacao':
      return aprovarSolicitacao(params);
    case 'rejeitarSolicitacao':
      return rejeitarSolicitacao(params);
    case 'listarSolicitacoesPendentes':
      return listarSolicitacoesPendentes(params);

    // Auth Campo
    case 'loginCampo':
      return loginCampo(params.email, params.senha);
    case 'recuperarSenha':
      return recuperarSenha(params.email || params);

    // CRUD estações
    case 'editarEstacaoFromMapa':
      return editarEstacaoFromMapa(params);
    case 'excluirEstacaoFromMapa':
      return excluirEstacaoFromMapa(params);

    // Monitor
    case 'listarMonitor':
      return listarMonitor();
    case 'toggleMonitor':
      return toggleMonitor(params);

    // Street View / IA
    case 'gerarStreetViewEstacao':
      return gerarStreetViewEstacao(params);

    case 'analisarCalcadaIA':
      return (typeof analisarCalcadaIA === 'function' && analisarCalcadaIA.length >= 2)
        ? analisarCalcadaIA(params.lat, params.lng)
        : analisarCalcadaIA(params);

    case 'getMapsKey':
      return { ok: true, key: getMapsApiKey_() };

    case 'analisarCalcadaComGemini':
      return analisarCalcadaComGemini(params);

    // Geocode
    case 'geocodeEndereco':
      return geocodeEnderecoPWA_({ endereco: params.q || params.endereco || params });

    // Polígonos
    case 'getPoligonosCidade': {
      var cidade = (typeof params === 'string')
        ? params
        : (params.cidade || params.city || '');
      return getPoligonosCidade(cidade);
    }

    case 'salvarPoligono':
      return salvarPoligono(params);

    case 'atualizarPoligono':
      return atualizarPoligono(params);

    case 'excluirPoligono':
      return excluirPoligono(params);

    default:
      return { ok: false, error: 'Acao desconhecida: ' + action };
  }
}
