function parseLatLngStringSafe(str) {
  if (!str) return null;
  const m = String(str).match(/(-?\d+(\.\d+)?),\s*(-?\d+(\.\d+)?)/);
  if (!m) return null;
  return { lat:Number(m[1]), lng:Number(m[3]) };
}

function pickNeighborhood_(geocodeResult) {
  if (!geocodeResult?.results?.length) return '';
  const comps = geocodeResult.results[0].address_components || [];
  const pick = t => (comps.find(c => c.types.includes(t)) || {}).long_name || '';
  return pick('neighborhood') || pick('sublocality') || pick('sublocality_level_1') || '';
}

function bairroEhValido_(bairro, cidade) {
  if (!bairro) return false;
  if (bairro.toLowerCase() === String(cidade||'').toLowerCase()) return false;
  return bairro.length >= 3;
}

/**
 * debug_doGet_gs.gs -- FUNCAO TEMPORARIA DE DEBUG
 *
 * Rodar no Editor do Apps Script para verificar:
 * 1. Se GMAPS_API_KEY esta nas Script Properties
 * 2. Se o replace esta funcionando no conteudo do pwaCampo
 *
 * Como usar:
 *   Editor > Selecionar funcao "debugCampoReplace" > Executar
 *   Ver o log em Execucoes (ou Ctrl+Enter)
 */
function debugCampoReplace() {
  var props    = PropertiesService.getScriptProperties();
  var gmapsKey = props.getProperty('GMAPS_API_KEY') || '';

  Logger.log('=== DEBUG CAMPO REPLACE ===');
  Logger.log('GMAPS_API_KEY definida: ' + (gmapsKey ? 'SIM (' + gmapsKey.slice(0,8) + '...)' : 'NAO -- ESSA E A CAUSA'));

  if (!gmapsKey) {
    Logger.log('ACAO: va em Configuracoes > Salvar Google Maps API Key e salve a chave');
    return;
  }

  var content = HtmlService.createHtmlOutputFromFile('pwaCampo').getContent();
  Logger.log('Tamanho do conteudo: ' + content.length + ' chars');
  Logger.log('Tem MAPS_KEY_VALUE: '   + content.includes('MAPS_KEY_VALUE'));
  Logger.log('Tem __GMAPS_API_KEY__: ' + content.includes('__GMAPS_API_KEY__'));
  Logger.log('Tem <?= GMAPS: '         + content.includes('<?= GMAPS_API_KEY'));

  var replaced = content.replace('MAPS_KEY_VALUE', gmapsKey);
  Logger.log('Apos replace, ainda tem MAPS_KEY_VALUE: ' + replaced.includes('MAPS_KEY_VALUE'));

  // Procurar o trecho relevante no conteudo
  var idx = content.indexOf('MAPS_KEY_VALUE');
  if (idx >= 0) {
    Logger.log('Contexto do placeholder (+-60 chars):');
    Logger.log(content.substring(Math.max(0, idx-60), idx+80));
  } else {
    Logger.log('MAPS_KEY_VALUE NAO ENCONTRADO no conteudo -- arquivo antigo no Editor?');
    Logger.log('Verifique se o pwaCampo.html foi atualizado com o arquivo campo.html novo');

    // Mostrar o trecho do Maps no arquivo atual
    var mapsIdx = content.indexOf('googleapis.com/maps');
    if (mapsIdx >= 0) {
      Logger.log('Trecho do Maps no arquivo atual:');
      Logger.log(content.substring(Math.max(0, mapsIdx-80), mapsIdx+120));
    }
  }
}