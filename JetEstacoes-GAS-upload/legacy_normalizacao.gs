/**
 * LEGACY — NORMALIZAÇÃO / REPARO DE ENDEREÇOS
 * -------------------------------------------------
 * ⚠️ Arquivo LEGACY
 * - Não participa do fluxo de croqui
 * - Não chama croqui_core
 * - Usado apenas para manutenção / saneamento de dados
 *
 * Seguro manter separado.
 */

/***************** BAIRROS / GEO *****************/

function normalizarBairroBR_(lat, lng, endereco) {
  // 1) Reverse geocode (fonte oficial)
  try {
    const res = Maps.newGeocoder().reverseGeocode(lat, lng);
    const bairro = pickNeighborhood_(res);
    if (bairro) return bairro;
  } catch (e) {
    Logger.log('Reverse geocode falhou: ' + e);
  }

  // 2) Fallback por endereço textual
  const viaEndereco = guessBairroFromAddress_(endereco);
  if (viaEndereco) return viaEndereco;

  return '';
}

function bairroEhValido_(bairro, cidade) {
  if (!bairro) return false;

  const b = String(bairro).trim().toLowerCase();
  const c = String(cidade || '').trim().toLowerCase();

  if (b === c) return false;
  if (b.length < 3) return false;

  return true;
}

function pickNeighborhood_(geocodeResult) {
  if (!geocodeResult || !geocodeResult.results || !geocodeResult.results.length) return '';
  const components = geocodeResult.results[0].address_components || [];
  const pick = (type) => {
    const f = components.find(c => (c.types || []).indexOf(type) !== -1);
    return f ? f.long_name : '';
  };
  return (
    pick('neighborhood') ||
    pick('sublocality') ||
    pick('sublocality_level_1') ||
    ''
  );
}

/***************** NORMALIZAÇÃO BR *****************/

function normalizarBairrosBrasil() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  if (!sh) throw new Error('Aba "Estacoes" não encontrada.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const cBairro = headers.indexOf('Bairro') + 1;
  const cCidade = headers.indexOf('Cidade') + 1;
  const cEnd    = headers.indexOf('Endereço completo da estação') + 1;
  const cLocal  = headers.indexOf('Localização') + 1;

  if (!cBairro || !cLocal) {
    throw new Error('Colunas obrigatórias não encontradas.');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('NORMALIZAR_BR_NEXT', '2');

  _normalizarBairrosBrasilWorker_(sh, lastRow, headers, cBairro, cCidade, cEnd, cLocal, props);
}

function continuarNormalizarBairrosBrasil() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const cBairro = headers.indexOf('Bairro') + 1;
  const cCidade = headers.indexOf('Cidade') + 1;
  const cEnd    = headers.indexOf('Endereço completo da estação') + 1;
  const cLocal  = headers.indexOf('Localização') + 1;

  const props = PropertiesService.getScriptProperties();
  _normalizarBairrosBrasilWorker_(sh, lastRow, headers, cBairro, cCidade, cEnd, cLocal, props);
}

function _normalizarBairrosBrasilWorker_(sh, lastRow, headers, cBairro, cCidade, cEnd, cLocal, props) {
  const HARD_LIMIT_MS = 5 * 60 * 1000;
  const startTime = Date.now();

  let row = Number(props.getProperty('NORMALIZAR_BR_NEXT') || 2);
  let atualizados = 0;

  for (; row <= lastRow; row++) {

    if (Date.now() - startTime > HARD_LIMIT_MS) {
      props.setProperty('NORMALIZAR_BR_NEXT', String(row));
      SpreadsheetApp.getUi().alert(
        'Normalização parcial concluída.\n' +
        'Bairros atualizados: ' + atualizados + '\n' +
        'Retomar a partir da linha ' + row
      );
      return;
    }

    const bairroAtual = sh.getRange(row, cBairro).getValue();
    const cidade = cCidade ? sh.getRange(row, cCidade).getValue() : '';
    const endereco = cEnd ? sh.getRange(row, cEnd).getValue() : '';
    const locStr = sh.getRange(row, cLocal).getValue();

    if (bairroEhValido_(bairroAtual, cidade)) continue;

    const coords = parseLatLngStringSafe(locStr);
    if (!coords) continue;

    const bairroNovo = normalizarBairroBR_(coords.lat, coords.lng, endereco);
    if (!bairroNovo) continue;

    sh.getRange(row, cBairro).setValue(bairroNovo);
    atualizados++;

    Utilities.sleep(1100); // respeita rate limit do Maps
  }

  props.deleteProperty('NORMALIZAR_BR_NEXT');

  SpreadsheetApp.getUi().alert(
    'Normalização nacional concluída!\n\n' +
    'Bairros atualizados: ' + atualizados
  );
}

/***************** REPARO DE ENDEREÇOS *****************/

function repairProblemAddressesSelection() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  const sel = sh.getActiveRange();
  if (!sel) {
    SpreadsheetApp.getUi().alert('Selecione um intervalo.');
    return;
  }
  _repairAddressesCore_(sh, sel);
  SpreadsheetApp.getUi().alert('Reparo de endereços (seleção) concluído.');
}

function repairProblemAddressesAll() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  const last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2,1,last-1, sh.getLastColumn());
  _repairAddressesCore_(sh, rng);
  SpreadsheetApp.getUi().alert('Reparo de endereços (todos) concluído.');
}

function _repairAddressesCore_(sh, rowBlock) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const iEnd = headers.indexOf('Endereço completo da estação');
  const iLoc = headers.indexOf('Localização');
  const iBai = headers.indexOf('Bairro');
  if (iEnd === -1 || iLoc === -1) throw new Error('Colunas obrigatórias não encontradas.');

  const values = rowBlock.getValues();
  const out = values.map(r => r.slice());

  for (let r=0; r<values.length; r++) {
    const endereco = String(values[r][iEnd] || '').trim();
    const locStr   = String(values[r][iLoc] || '').trim();
    const bairro   = String(values[r][iBai] || '').trim();

    if (!_isProblematicAddress_(endereco)) continue;

    if (_hasLatLng_(locStr)) {
      const ll  = _parseLatLng_(locStr);
      const rev = _reverseGeocode_(ll.lat, ll.lng);
      if (rev.formatted) out[r][iEnd] = rev.formatted;
      if (!bairro) out[r][iBai] = _pickBairro_(rev.components) || bairro;
      _throttle_();
    }
  }

  for (let r=0; r<values.length; r++) {
    sh.getRange(rowBlock.getRow() + r, 1, 1, out[r].length)
      .setValues([out[r]]);
  }
}

/***************** HELPERS INTERNOS *****************/

function _isProblematicAddress_(addr) {
  if (!addr) return true;
  if (_isPlusCode_(addr)) return true;
  if (/^-?\d+\.\d+,\s*-?\d+\.\d+$/.test(addr)) return true;
  if (addr.length < 8) return true;
  return false;
}

function _isPlusCode_(addr) {
  return /[23456789CFGHJMPQRVWX]{4,}\+[2-9CFGHJMPQRVWX]{2}/i.test(addr);
}

function _hasLatLng_(locStr) {
  return /^-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?$/.test(String(locStr||''));
}

function _parseLatLng_(locStr) {
  const parts = String(locStr||'').split(',');
  return { lat:Number(parts[0]), lng:Number(parts[1]) };
}

function _reverseGeocode_(lat,lng) {
  try {
    const res = Maps.newGeocoder().reverseGeocode(lat,lng);
    if (res?.results?.length) {
      const best = res.results[0];
      return { formatted: best.formatted_address||'', components: best.address_components||[] };
    }
  } catch(e){}
  return {formatted:'', components:[]};
}

function _pickBairro_(components) {
  if(!components) return '';
  const byType = t => (components.find(c => (c.types||[]).includes(t)) || {}).long_name || '';
  return byType('neighborhood') || byType('sublocality_level_1') || byType('sublocality') || '';
}

function _throttle_(ms) {
  Utilities.sleep(ms || 200);
}
