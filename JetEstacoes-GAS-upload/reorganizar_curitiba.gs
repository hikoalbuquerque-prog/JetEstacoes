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