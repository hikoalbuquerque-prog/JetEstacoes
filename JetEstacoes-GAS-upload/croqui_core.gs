/**
 * CROQUI CORE — ENGINE CENTRAL
 * --------------------------------
 * Engine pura, determinística, sem UI
 * Código.gs apenas ORQUESTRA
 */

/* ================= CONFIG ================= */

const CROQUI_CORE_CFG = {
  MAP: {
    SIZE: '640x640',
    SCALE: 2,
    ZOOM: 19
  }
};

/* ================= ENTRY POINT ================= */

/**
 * ÚNICO ponto de geração de croqui
 * Chamado por Código.gs (wrappers)
 */
function gerarCroqui_CORE(sheet, row) {

  if (!sheet || typeof sheet.getName !== 'function') {
    throw new Error('Sheet inválida');
  }
  if (!row || row < 2) {
    throw new Error('Linha inválida: ' + row);
  }

  const data = lerLinhaEstacao_CORE_(sheet, row);

  if (data.tipoEstacao === 'PUBLICA') {
    return gerarCroquiPublico_CORE_(data);
  }

  if (data.tipoEstacao === 'PRIVADA') {
    return gerarCroquiPrivado_CORE_(data);
  }

  throw new Error('TipoEstacao não suportado: ' + data.tipoEstacao);
}

/* ================= LEITURA ================= */

function lerLinhaEstacao_CORE_(sheet, row) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const values  = sheet.getRange(row,1,1,sheet.getLastColumn()).getValues()[0];

  const get = (name) => {
    const i = headers.indexOf(name);
    return i === -1 ? '' : values[i];
  };

  const loc = String(get('Localização')).trim();
  const coords = parseLatLng_CORE_(loc);

  return {
    row,
    codigo: get('CodigoEstacao'),
    tipoEstacao: String(get('TipoEstacao') || '').toUpperCase(),
    endereco: get('Endereço completo da estação'),
    cidade: get('Cidade'),
    bairro: get('Bairro'),
    subprefeitura: get('Subprefeitura'),
    lat: coords.lat,
    lng: coords.lng
  };
}

/* ================= CROQUI PÚBLICO ================= */

function gerarCroquiPublico_CORE_(d) {

  validarCampos_CORE_(d, ['codigo','endereco','lat','lng']);

  const imagens = gerarImagensCroqui_CORE(d.lat, d.lng);

  // Aqui você pluga Slides / PDF
  // (mantive mínimo porque o foco é estabilizar engine)

  return {
    ok: true,
    imagens
  };
}

/* ================= CROQUI PRIVADO ================= */

function gerarCroquiPrivado_CORE_(d) {

  validarCampos_CORE_(d, ['codigo','endereco','lat','lng']);

  const imagens = gerarImagensCroqui_CORE(d.lat, d.lng);

  return {
    ok: true,
    imagens
  };
}

/* ================= MAP ENGINE ================= */

function gerarImagensCroqui_CORE(lat, lng) {
  return {
    sat: fetchStaticMap_CORE(lat, lng, { maptype: 'satellite' }),
    map: fetchStaticMap_CORE(lat, lng, { maptype: 'roadmap' }),
    street: fetchStreetView_CORE(lat, lng, {})
  };
}

function fetchStaticMap_CORE(lat, lng, cfg) {

  validarLatLng_CORE_(lat, lng);

  const key = PropertiesService.getScriptProperties().getProperty('GMAPS_API_KEY');
  if (!key) throw new Error('GMAPS_API_KEY não definida');

  const config = Object.assign({
    zoom: CROQUI_CORE_CFG.MAP.ZOOM,
    size: CROQUI_CORE_CFG.MAP.SIZE,
    scale: CROQUI_CORE_CFG.MAP.SCALE,
    maptype: 'roadmap'
  }, cfg || {});

  const center = lat + ',' + lng;
  const markers = encodeURIComponent('color:red|' + center);

  const url =
    'https://maps.googleapis.com/maps/api/staticmap?' +
    'center=' + encodeURIComponent(center) +
    '&zoom=' + Number(config.zoom) +
    '&size=' + encodeURIComponent(String(config.size)) +
    '&scale=' + Number(config.scale) +
    '&maptype=' + encodeURIComponent(String(config.maptype)) +
    '&markers=' + markers +
    '&key=' + key;

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('StaticMap erro ' + resp.getResponseCode());
  }

  return resp.getBlob();
}

function fetchStreetView_CORE(lat, lng, cfg) {

  validarLatLng_CORE_(lat, lng);

  const key = PropertiesService.getScriptProperties().getProperty('GMAPS_API_KEY');
  if (!key) throw new Error('GMAPS_API_KEY não definida');

  const config = Object.assign({
    size: '640x640',
    fov: 90,
    heading: 0,
    pitch: 0
  }, cfg || {});

  const url =
    'https://maps.googleapis.com/maps/api/streetview?' +
    'size=' + encodeURIComponent(config.size) +
    '&location=' + encodeURIComponent(lat + ',' + lng) +
    '&fov=' + Number(config.fov) +
    '&heading=' + Number(config.heading) +
    '&pitch=' + Number(config.pitch) +
    '&key=' + key;

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('StreetView erro ' + resp.getResponseCode());
  }

  return resp.getBlob();
}

/* ================= HELPERS CORE ================= */

function parseLatLng_CORE_(str) {
  const m = String(str).match(/(-?\d+(\.\d+)?),\s*(-?\d+(\.\d+)?)/);
  if (!m) throw new Error('Localização inválida: ' + str);
  return { lat:Number(m[1]), lng:Number(m[3]) };
}

function validarLatLng_CORE_(lat, lng) {
  if (!isFinite(lat) || !isFinite(lng)) {
    throw new Error('Lat/Lng inválidos');
  }
}

function validarCampos_CORE_(obj, campos) {
  campos.forEach(c => {
    if (!obj[c]) throw new Error('Campo obrigatório ausente: ' + c);
  });
}

/* ================= DEBUG ================= */

function DEBUG_CORE_STATIC_MAP() {
  const blob = fetchStaticMap_CORE(-22.954771, -43.190604, { maptype:'roadmap' });
  DriveApp.getRootFolder().createFile(blob.setName('CORE_STATIC_OK.png'));
}

var REORG_ROW_KEY = 'REORG_CROQUI_ROW';

// ── Entrada principal ─────────────────────────────────────────
function reorganizarCroquisRJ() {
  _executarReorganizacao_();
}

function continuarReorganizarCroquis() {
  _executarReorganizacao_();
}

function statusReorganizarCroquis() {
  var saved = PropertiesService.getScriptProperties().getProperty(REORG_ROW_KEY);
  if (!saved) {
    Logger.log('[STATUS] Nenhum progresso salvo. Nao iniciado ou ja concluido.');
  } else {
    Logger.log('[STATUS] Vai retomar a partir da linha ' + saved + ' da planilha.');
  }
}

function resetarReorganizarCroquis() {
  PropertiesService.getScriptProperties().deleteProperty(REORG_ROW_KEY);
  Logger.log('[RESET] Progresso apagado. Rode reorganizarCroquisRJ() para comecar do zero.');
}

// ── Implementacao retomavel ───────────────────────────────────
function _executarReorganizacao_() {

  var props = PropertiesService.getScriptProperties();
  var sh    = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  if (!sh) {
    Logger.log('[ERRO] Aba "' + SHEET_NAME + '" nao encontrada.');
    return;
  }

  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });

  var iCidade  = headers.indexOf('Cidade');
  var iBairro  = headers.indexOf('Bairro');
  var iSubpref = headers.indexOf('Subprefeitura');
  var iTipo    = headers.indexOf('TipoEstacao');
  var iCroqui  = headers.indexOf('Croqui');

  if (iCroqui === -1) {
    Logger.log('[ERRO] Coluna "Croqui" nao encontrada.');
    return;
  }

  var pastaPublica = DriveApp.getFolderById(CROQUIS_PUBLICOS_FOLDER_ID);
  var pastaPrivada = DriveApp.getFolderById(CROQUIS_PRIVADOS_FOLDER_ID);

  // Retomar de onde parou
  var savedRow = Number(props.getProperty(REORG_ROW_KEY) || 1);
  var startIdx = Math.max(1, savedRow - 1); // data[] e 0-based, linha 1 = header

  var HARD_LIMIT_MS = 4 * 60 * 1000; // 4 min
  var startTime     = Date.now();
  var movidos       = 0;
  var jaOrg         = 0;
  var erros         = 0;

  Logger.log('[INICIO] Comecando da linha ' + (startIdx + 1) + ' de ' + data.length);

  for (var i = startIdx; i < data.length; i++) {
    var linhaReal = i + 1;

    // Salvar progresso e pausar se proximo do limite de tempo
    if (i % 10 === 0 && Date.now() - startTime > HARD_LIMIT_MS) {
      props.setProperty(REORG_ROW_KEY, String(linhaReal));
      Logger.log(
        '[PAUSADO] Tempo limite atingido.\n' +
        '  Movidos: ' + movidos + '\n' +
        '  Ja organizados/sem croqui: ' + jaOrg + '\n' +
        '  Erros: ' + erros + '\n' +
        '  Progresso salvo na linha ' + linhaReal + '.\n' +
        '  >> Rode reorganizarCroquisRJ() novamente para continuar.'
      );
      return;
    }

    var row       = data[i];
    var cidade    = String(row[iCidade]  || '').trim();
    var bairro    = String(row[iBairro]  || '').trim();
    var subpref   = String(row[iSubpref] || '').trim();
    var tipo      = String(row[iTipo]    || '').trim().toUpperCase();
    var croquiUrl = String(row[iCroqui]  || '').trim();

    // Pular linhas sem croqui gerado
    if (!croquiUrl) { jaOrg++; continue; }

    // Pular SP (ja organizado por subprefeitura)
    var cidNorm = cidade.toUpperCase()
      .replace(/[ÁÀÃÂ]/g, 'A')
      .replace(/[ÉÈÊ]/g, 'E')
      .replace(/[ÍÌÎ]/g, 'I')
      .replace(/[ÓÒÕÔ]/g, 'O')
      .replace(/[ÚÙÛ]/g, 'U');

    if (cidNorm === 'SAO PAULO') { jaOrg++; continue; }

    // Extrair file ID do link
    var fileId = null;
    var m = croquiUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!m) m = croquiUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (m) fileId = m[1];
    if (!fileId) { jaOrg++; continue; }

    try {
      var file = DriveApp.getFileById(fileId);

      // Pasta raiz conforme tipo
      var pastaRaiz = (tipo === 'PUBLICA' || tipo === 'PUBLICO')
        ? pastaPublica
        : pastaPrivada;

      // Nivel 1: Cidade
      var pastaCidade = getOuCriarPasta_(pastaRaiz, cleanName_(cidade || 'Sem_Cidade'));

      // Nivel 2: Bairro (ou Subprefeitura para SP -- ja filtrado acima)
      var nomeSub = cleanName_(bairro || 'Sem_Bairro');
      var pastaDestino = getOuCriarPasta_(pastaCidade, nomeSub);

      // Verificar se ja esta na pasta correta
      var jaEstaLa = false;
      var parentsIt = file.getParents();
      while (parentsIt.hasNext()) {
        if (parentsIt.next().getId() === pastaDestino.getId()) {
          jaEstaLa = true;
          break;
        }
      }

      if (jaEstaLa) { jaOrg++; continue; }

      // Mover: adicionar na destino e remover das outras
      pastaDestino.addFile(file);

      var parents2 = file.getParents();
      while (parents2.hasNext()) {
        var p = parents2.next();
        if (p.getId() !== pastaDestino.getId()) {
          p.removeFile(file);
        }
      }

      movidos++;
      Logger.log('[MOVE] L' + linhaReal + ' ' + cidade + '/' + (bairro || 'Sem_Bairro') + ' → ' + file.getName());
      Utilities.sleep(80);

    } catch (e) {
      erros++;
      Logger.log('[ERRO] L' + linhaReal + ': ' + e);
    }
  }

  // Concluido
  props.deleteProperty(REORG_ROW_KEY);
  Logger.log(
    '[CONCLUIDO]\n' +
    '  Movidos: ' + movidos + '\n' +
    '  Ja organizados/sem croqui: ' + jaOrg + '\n' +
    '  Erros: ' + erros
  );
}

// ── Helper: buscar ou criar subpasta ─────────────────────────
function getOuCriarPasta_(pastaBase, nomePasta) {
  nomePasta = String(nomePasta || 'Sem_Nome').trim();
  var it = pastaBase.getFoldersByName(nomePasta);
  return it.hasNext() ? it.next() : pastaBase.createFolder(nomePasta);
}
