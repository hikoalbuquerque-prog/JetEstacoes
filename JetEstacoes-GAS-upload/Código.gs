/**
 * APP ESTAÇÕES — CORE SCRIPT
 * Última revisão: 2026-01
 * Tech Lead: Wesley
 *
 * ⚠️ Não chamar funções *_ diretamente
 * Use sempre os wrappers de fila e retry
 */

/*******************  CONFIG  *******************/

/*******************  HELPERS (WEBAPP SAFE)  *******************/
function getSS_() {
  const props = PropertiesService.getScriptProperties();
  const id =
    props.getProperty('SPREADSHEET_ID') ||
    props.getProperty('PLANILHA_ID') ||
    props.getProperty('SHEET_ID') ||
    props.getProperty('ID_PLANILHA');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetAny_(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}


// Ajuste apenas se quiser trocar pastas/organização
const TARGET_CROQUIS_FOLDER_ID = '1Xvp6fGAOEo-CYtSvIGsTwL7fxQwLznOe';  // Croquis (PDFs + imagens de mapa)
const TARGET_FOTOS_FOLDER_ID   = '1hc5-whvQdYNqHbmkzW96MMnOjm4nk964';  // Fotos das estações
const IMAGENS_FOLDER_ID        = '1hc5-whvQdYNqHbmkzW96MMnOjm4nk964';  // Imagens (Satélite / Mapa / Street)
const SLIDES_TEMPLATE_PUBLICO_ID = '1BFRPrjFO2h39vW1xnJYoy4VSBekHURBZZZEpf3J4i8g';
const SLIDES_TEMPLATE_PRIVADO_ID = '1bCLY8Gn2hpO6P1B5Tb4zxizahJkNjOGq9Nhw5pK69u4';
const CROQUIS_PUBLICOS_FOLDER_ID = '1Ww46_E6eK4Tv8_DNoQ7yNA7FL4IpexZj';
const CROQUIS_PRIVADOS_FOLDER_ID = '1-IBYPkHV32SbcFx4vrfwuRtcwEQdMH_j';
const SPREADSHEET_ID = '1w6GrXtwu2cJP0jP8I75kzkWzMjkOoCNGQ7vYUJglrdA'


// ================== MÉXICO (MX) ==================

// 📄 Templates Slides
const SLIDES_TEMPLATE_PUBLICO_MX_ID =
  '1nWlpr5kbwVb1uOop_lhsMm7Y-7kmRMw5lBOQ-j9TgXs';

const SLIDES_TEMPLATE_PRIVADO_MX_ID =
  '1yRHOfsH7dgEQimfjdwAvqlYv7dBFiLjqlGS9c60j9iE';

// 📁 Pastas CDMX (raiz)
const CROQUIS_PUBLICOS_MX_FOLDER_ID =
  '172jDUHBUXPhEckvAWsOxziScAlyglB-V';

const CROQUIS_PRIVADOS_MX_FOLDER_ID =
  '18_7LP9aogPGN_Y1Y7OKvDo5MDeXKWGzN';

// 📁 Pastas MX fora de CDMX (raiz)
const CROQUIS_PUBLICOS_MX_FORA_CDMX_FOLDER_ID =
  '1GWo7l-WJmldPW60_e74JD5CTKOZA_6Cj';

const CROQUIS_PRIVADOS_MX_FORA_CDMX_FOLDER_ID =
  '1QQwK2vax_q1EHkNFkEULXv3caWFA9iOX';


// Organização em subpastas: 'none' | 'bairro' | 'codigo'
const ORGANIZE_BY = 'bairro';

// Planilha / colunas
const SHEET_NAME         = 'Estacoes';
const CROQUIS_FOLDER_ID  = TARGET_CROQUIS_FOLDER_ID;
const FOTOS_FOLDER_ID    = TARGET_FOTOS_FOLDER_ID;

// API key – já deixei a sua aqui também (mas o ideal é salvar nas Script Properties)
const STATIC_MAPS_API_KEY = 'AIzaSyD-yITmwagjpKPhAlTX1ecAJA7SNYfae5E';
//////////////////////////////AIzaSyCFEVy4YPCq_lv1vP2RVfRUHIrjKU1q28E///////////////////////////////

PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'AIzaSyAWq-l8qSIoURs7Z24rYpjwcKcV3EzWVkY');
// Nomes exatos das colunas
const COL = {
  RowKey      : 'RowKey',
  Codigo      : 'CodigoEstacao',
  NomeEstacao : 'Nome da Estação',

  Cidade      : 'Cidade',
  Bairro      : 'Bairro',
  Subprefeitura : 'Subprefeitura',

  Endereco    : 'Endereço completo da estação',
  Localizacao : 'Localização',

  TipoEstacao : 'TipoEstacao',      // PUBLICA | PRIVADA | CONCORRENTE
  TipoPublica : 'TipoPublica',       // CALÇADA | RUA

  Dimensoes   : 'Dimensões da Estação',
  Largura     : 'Largura da Faixa Livre (m)',
  FaixaMinima : 'FaixaLivreMinima',
  Capacidade  : 'Capacidade',
  AreaTotal   : 'AreaTotal',
  Condicao    : 'CondicaoImplantacao',

  Foto        : 'Foto da Estação',
  ImgSat      : 'Imagem Satélite',
  ImgMapa     : 'Imagem Mapa',
  ImgStreet   : 'Street View',
  Croqui      : 'Croqui',

  // ===== PRIVADO =====
  NomeLocalPrivado      : 'NomeLocalPrivado',
  NomeAutorizante      : 'NomeAutorizante',
  CargoAutorizante     : 'CargoAutorizante',
  TelefoneAutorizante  : 'TelefoneAutorizante',
  EmailAutorizante     : 'EmailAutorizante',
  DocumentoAutorizacao : 'DocumentoAutorizacao',
  DataAutorizacao      : 'DataAutorizacao',
  ObservacaoPrivado    : 'ObservacaoPrivado'
};


function validarInfraTemplates_() {
  const erros = [];

  function validar(id, nome) {
    if (!id || typeof id !== 'string') {
      erros.push(`${nome}: ID vazio ou inválido`);
      return;
    }

    try {
      const file = DriveApp.getFileById(id);
      const mime = file.getMimeType();

      if (mime !== MimeType.GOOGLE_SLIDES) {
        erros.push(`${nome}: arquivo não é Google Slides (${mime})`);
      }
    } catch (e) {
      erros.push(`${nome}: sem acesso ou ID inexistente`);
    }
  }

  validar(SLIDES_TEMPLATE_PUBLICO_ID, 'Template Público');
  validar(SLIDES_TEMPLATE_PRIVADO_ID, 'Template Privado');

  if (erros.length) {
    throw new Error(
      '❌ ERRO DE INFRA — Templates inválidos:\n\n' +
      erros.join('\n')
    );
  }

  Logger.log('✅ Infra de templates validada com sucesso');
}


/******************** ON OPEN (menus) ********************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Croquis')
    .addItem('Gerar Croqui (linha atual)', 'gerarCroquisLinhaAtual')
    .addItem('Gerar Croqui (seleção – com fila)', 'gerarCroquisSelecaoFila')
    .addSeparator()
    .addItem('Continuar fila automática', 'processarFilaCroquis_')
    .addSeparator()
    .addItem('Criar gatilho automático (fila)', 'criarGatilhoFilaCroquis')
    .addItem('Remover gatilho automático', 'removerGatilhoFilaCroquis')
    .addToUi();

  ui.createMenu('Mapas')
    .addItem('Gerar Satélite (seleção)', 'gerarImagemSateliteSelecao')
    .addItem('Gerar Mapa (seleção)', 'gerarImagemMapaSelecao')
    .addItem('Gerar Street View (seleção)', 'gerarImagemStreetViewSelecao')
    .addToUi();

  ui.createMenu('⚙️ Estações')
    .addItem('🏙️ Preencher subprefeituras (SP)', 'preencherSubprefeiturasSP')
    .addItem('🔁 Reprocessar fotos por subprefeitura', 'reprocessarFotosPorSubprefeitura')
    .addToUi();

  ui.createMenu('Reorganizar')
    .addItem('Normalizar bairros (Brasil)', 'normalizarBairrosBrasil')
    .addItem('Continuar normalização BR', 'continuarNormalizarBairrosBrasil')
    .addItem('Normalizar lista (início)', 'normalizarListaAtual')
    .addItem('Continuar normalização', 'continuarNormalizacaoLista')
    .addItem('Atualizar bairros (auto)', 'updateBairrosAuto')
    .addToUi();

  ui.createMenu('Exportar')
    .addItem('Exportar Croquis por Cidade (fila)', 'exportarCroquisPorCidadeFila')
    .addToUi();

  // ✅ único ponto de criação do menu de polígonos
  addMenuPoligonos(ui);
}



/******************** UTIL: API KEY ********************/
function saveMapsApiKey() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Salvar GMAPS API Key',
    'Cole sua API key do Google Maps (não compartilhe):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const key = resp.getResponseText().trim();
  if (!key) {
    ui.alert('Chave vazia.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GMAPS_API_KEY', key);
  ui.alert('GMAPS_API_KEY salva nas Script Properties.');
}

function getMapsApiKey_() {
  const key = PropertiesService
    .getScriptProperties()
    .getProperty('GMAPS_API_KEY');

  if (!key) {
    throw new Error(
      'GMAPS_API_KEY não configurada. Use o menu para salvar a chave.'
    );
  }
  return key;
}

function getMapsApiKey() {
  return getMapsApiKey_();
}

function salvarApiKey() {
  PropertiesService
    .getScriptProperties()
    .setProperty('GMAPS_API_KEY', 'AIzaSyD-yITmwagjpKPhAlTX1ecAJA7SNYfae5E');
}


/******************* GATILHO ON CHANGE ********************/
function setupTriggers() {
  const all = ScriptApp.getProjectTriggers();
  for (const t of all) {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'onChangeAssets') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('onChangeAssets')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
  SpreadsheetApp.getUi().alert('Trigger onChangeAssets criado. Lembre-se de autorizar o script.');
}

function onChangeAssets(e) {
  try {
    // Se quiser automatizar algo aqui, descomente:
    // updateBairrosAuto();
  } catch (err) {
    Logger.log('onChangeAssets erro: ' + err);
  }
}


/******************* HELPERS GERAIS (Drive / planilha) ********************/
function gerarCroquiComRetry_(sheet, row) {

  const { headers } = makeRowAccessor_(sheet, row);

  const colStatus = headers.indexOf('CroquiStatus') + 1;
  const colTent   = headers.indexOf('CroquiTentativas') + 1;
  const colErro   = headers.indexOf('CroquiUltimoErro') + 1;
  const colExec   = headers.indexOf('CroquiUltimaExec') + 1;

  if (!colStatus || !colTent || !colErro || !colExec) {
    throw new Error('Colunas de controle de croqui ausentes.');
  }

  // 🔒 Validação de estado
  const statusAtual = sheet.getRange(row, colStatus).getValue();
  const validos = ['','PENDENTE','PROCESSANDO','OK','ERRO','IGNORADO'];

  if (validos.indexOf(statusAtual) === -1) {
    throw new Error('CroquiStatus inválido: ' + statusAtual);
  }

  let tentativas = Number(sheet.getRange(row, colTent).getValue() || 0);

  try {
    sheet.getRange(row, colStatus).setValue('PROCESSANDO');
    sheet.getRange(row, colExec).setValue(new Date());

    gerarCroquiParaLinha(sheet, row, true);

    sheet.getRange(row, colStatus).setValue('OK');
    sheet.getRange(row, colErro).setValue('');
  } catch (e) {
    tentativas++;
    sheet.getRange(row, colTent).setValue(tentativas);
    sheet.getRange(row, colErro).setValue(String(e.message || e));

    if (tentativas >= 3) {
      sheet.getRange(row, colStatus).setValue('ERRO');
    } else {
      sheet.getRange(row, colStatus).setValue('PENDENTE');
    }

    throw e;
  }
}



function toCentimeters_(val) {
  if (val === '' || val === null || val === undefined) return '';
  const n = Number(String(val).replace(',', '.'));
  if (isNaN(n)) return '';
  if (n > 10) return Math.round(n);
  return Math.round(n * 100);
}


function linhaSeValor_(label, valor, sufixo) {
  if (valor === null || valor === undefined) return '';
  const v = String(valor).trim();
  if (!v) return '';
  return label + ': ' + v + (sufixo || '');
}


/******************* EXPORTAÇÕES *******************/
function exportarCroquisPorCidadeFila() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Exportar Croquis (Fila)',
    'Digite o nome da cidade exatamente como na planilha:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const cidade = cleanName_(resp.getResponseText());
  if (!cidade) return;

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const cCidade = headers.indexOf(COL.Cidade);
  const cCroqui = headers.indexOf(COL.Croqui);
  const cStatus = headers.indexOf('CroquiStatus');

  if (cCidade === -1 || cCroqui === -1 || cStatus === -1) {
    ui.alert('Colunas Cidade, Croqui ou CroquiStatus não encontradas.');
    return;
  }

  const rows = [];
  const lastRow = sh.getLastRow();

  for (let r = 2; r <= lastRow; r++) {
    const cidadeRow = cleanName_(sh.getRange(r, cCidade + 1).getValue());
    const status    = String(sh.getRange(r, cStatus + 1).getValue()).toUpperCase();
    const croquiUrl = sh.getRange(r, cCroqui + 1).getValue();

    if (cidadeRow === cidade && status === 'OK' && croquiUrl) rows.push(r);
  }

  if (!rows.length) {
    ui.alert('Nenhum croqui válido (status OK) encontrado para esta cidade.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('EXPORT_FILA_CIDADE', cidade);
  props.setProperty('EXPORT_FILA_ROWS', JSON.stringify(rows));
  props.setProperty('EXPORT_FILA_INDEX', '0');
  props.setProperty(
    'EXPORT_FILA_DATA',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  );

  criarTriggerExportacao_();

  ui.alert(
    'Fila de exportação criada!\n\n' +
    'Cidade: ' + cidade + '\n' +
    'Croquis na fila: ' + rows.length + '\n\n' +
    'A exportação será feita automaticamente.'
  );
}

// mantém fora
function criarTriggerExportacao_() {
  removerTriggerExportacao_();
  ScriptApp.newTrigger('processarFilaExportacao_')
    .timeBased()
    .everyMinutes(1)
    .create();
}


function removerTriggerExportacao_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processarFilaExportacao_') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function finalizarFilaExportacao_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('EXPORT_FILA_CIDADE');
  props.deleteProperty('EXPORT_FILA_ROWS');
  props.deleteProperty('EXPORT_FILA_INDEX');
  props.deleteProperty('EXPORT_FILA_DATA');

  removerTriggerExportacao_();
}


function processarFilaExportacao_() {
  const props = PropertiesService.getScriptProperties();

  const cidade = props.getProperty('EXPORT_FILA_CIDADE');
  const rows   = JSON.parse(props.getProperty('EXPORT_FILA_ROWS') || '[]');
  let index    = Number(props.getProperty('EXPORT_FILA_INDEX') || 0);
  const data   = props.getProperty('EXPORT_FILA_DATA');

  if (!cidade || !rows.length) {
    finalizarFilaExportacao_();
    return;
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const cCroqui = headers.indexOf(COL.Croqui);

  const exportRoot = ensureSubFolder_(DriveApp.getRootFolder(), 'Exportacoes');
  const pastaCidade = ensureSubFolder_(exportRoot, cidade);
  const pastaData   = ensureSubFolder_(pastaCidade, data);
  const pastaCroq   = ensureSubFolder_(pastaData, 'Croquis');

  const BATCH_SIZE = 25;
  let processed = 0;

  while (index < rows.length && processed < BATCH_SIZE) {
    const row = rows[index];
    index++;

    try {
      const croquiUrl = sh.getRange(row, cCroqui + 1).getValue();
      const file = resolveFileFromCell_(croquiUrl, DriveApp.getRootFolder());
      if (!file) continue;

      file.makeCopy(file.getName(), pastaCroq);
      processed++;
    } catch (e) {
      Logger.log('Erro exportando linha ' + row + ': ' + e);
    }
  }

  props.setProperty('EXPORT_FILA_INDEX', String(index));

  if (index >= rows.length) {
    finalizarFilaExportacao_();
    SpreadsheetApp.getUi().alert(
      'Exportação concluída!\n\n' +
      'Cidade: ' + cidade + '\n' +
      'Croquis exportados: ' + rows.length + '\n\n' +
      'Pasta:\n' + pastaData.getUrl()
    );
  }
}


/******************* (2) FOTO – MOVE & RENAME *********************/
// Agora: renomeia apenas com o "Endereço completo da estação" e salva link público
function renomearFotoParaLinha_(sheet, row) {
  const { headers, get } = makeRowAccessor_(sheet, row);

  const bairro   = get(COL.Bairro);
  const endereco = get(COL.Endereco);
  const fotoVal  = get(COL.Foto);

  // Precisa ter endereço e alguma referência de foto
  if (!fotoVal || !endereco) return;

  const fotosRoot = DriveApp.getFolderById(FOTOS_FOLDER_ID);
  const dirBairro = ensureSubFolder_(fotosRoot, cleanName_(bairro || 'Outros'));

  const file = resolveFileFromCell_(fotoVal, fotosRoot);
  if (!file) {
    Logger.log('Foto não encontrada na linha ' + row + ': ' + fotoVal);
    return;
  }

  const baseName = cleanName_(endereco);

  // Mantém extensão original
  const oldName = file.getName();
  const mExt = oldName.match(/(\.[^\.]+)$/);
  const ext = mExt ? mExt[1] : '';
  const newName  = baseName + ext;

  moveFileToFolder_(file, dirBairro);
  file.setName(newName);

  const colFoto = headers.indexOf(COL.Foto) + 1;
  setCellToPublicUrl_(sheet, row, colFoto, file);
}


/******************* (3) CROQUI – CORE (FOTO + SAT + MAP + STREET) ********************/


function salvarImagensNoDrive2_(sheet, row, imagens) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  var colSat    = headers.indexOf(COL.ImgSat) + 1;
  var colMap    = headers.indexOf(COL.ImgMapa) + 1;
  var colStreet = headers.indexOf(COL.ImgStreet) + 1;

  var pastaImagens = DriveApp.getFolderById(IMAGENS_FOLDER_ID);

  function salvar(blob, sufixo, col) {
    if (!blob || !col) return;

    var novoBlob = Utilities.newBlob(
      blob.getBytes(),
      'image/png',
      'ESTACAO_' + row + '_' + sufixo + '.png'
    );

    var file = pastaImagens.createFile(novoBlob);

    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    } catch (e) {
      Logger.log(
        'Aviso: sem permissão para setSharing em ' +
        file.getName() + ' — ' + e
      );
    }

    sheet.getRange(row, col).setValue(file.getUrl());
  }

  salvar(imagens.sat, 'SAT', colSat);
  salvar(imagens.map, 'MAP', colMap);
  salvar(imagens.street, 'STREET', colStreet);
}

/**
 * Wrapper de compatibilidade
 * Código.gs NÃO gera croqui — apenas delega ao CORE
 */

function gerarCroquiPublico_(sheet, row, silencioso) {
  return gerarCroquiFinal_(sheet, row, 'PUBLICO');
}

function gerarCroquiPrivado_(sheet, row, silencioso) {
  return gerarCroquiFinal_(sheet, row, 'PRIVADO');
}


function gerarCroquisSelecaoFila() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const sel = sh.getActiveRange();

  if (!sel) {
    SpreadsheetApp.getUi().alert('Selecione linhas.');
    return;
  }

  const start = Math.max(2, sel.getRow());
  const end   = Math.min(sh.getLastRow(), start + sel.getNumRows() - 1);

  const props = PropertiesService.getScriptProperties();
  props.setProperty('CROQUI_FILA_ATUAL', start);
  props.setProperty('CROQUI_FILA_FIM', end);

  processarFilaCroquis_();
}


function gerarCroquisPendentesSelecao() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) {
    ui.alert('Aba "' + SHEET_NAME + '" não encontrada.');
    return;
  }

  const sel = sh.getActiveRange();
  if (!sel || sel.getRow() < 2) {
    ui.alert('Selecione linhas válidas.');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const colCroqui = headers.indexOf(COL.Croqui);
  if (colCroqui === -1) {
    ui.alert('Coluna "Croqui" não encontrada.');
    return;
  }

  const start = Math.max(2, sel.getRow());
  const end   = Math.min(sh.getLastRow(), start + sel.getNumRows() - 1);

  // 🔹 Monta lista apenas com linhas SEM croqui
  const filas = [];
  for (let row = start; row <= end; row++) {
    if (sh.isRowHiddenByFilter(row)) continue;

    const croquiVal = String(sh.getRange(row, colCroqui + 1).getValue() || '').trim();
    if (!croquiVal) filas.push(row);
  }

  if (!filas.length) {
    ui.alert('Nenhuma linha pendente de croqui na seleção.');
    return;
  }

  // 🔹 Salva fila explícita (array serializado)
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CROQUI_QUEUE_LIST', JSON.stringify(filas));
  props.setProperty('CROQUI_QUEUE_INDEX', '0');
  props.setProperty('CROQUI_QUEUE_RUNNING', 'true');

  criarTriggerFilaCroquis_();

  ui.alert(
    'Fila de croquis pendentes criada!\n\n' +
    'Total na fila: ' + filas.length + '\n' +
    'Processamento automático iniciado.'
  );
}



/********** helpers usados no croqui (Slides) **********/
function resolvePastaTerritorial_(cidade, bairro, subprefeitura) {
  const cid = String(cidade || '').toUpperCase();

  if (cid === 'SAO PAULO' || cid === 'SÃO PAULO') {
    return cleanName_('Subprefeitura_' + (subprefeitura || 'Sem_Subprefeitura'));
  }

  return cleanName_(bairro || 'Sem_Bairro');
}

function replaceTokensInSlidesNoClose_(presentation, map) {
  const slides = presentation.getSlides();
  for (const s of slides) {
    s.getPageElements().forEach(pe => {
      if (pe.getPageElementType && pe.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = pe.asShape();
        let txt = shape.getText().asString();
        Object.keys(map).forEach(token => {
          txt = txt.split(token).join(map[token] ?? '');
        });
        shape.getText().setText(txt);
      }
    });
  }
}

// encontra placeholder shape (que contenha o token)
function findPlaceholderShape_(target, token) {

  // Caso 1: veio a apresentação inteira
  if (target.getSlides) {
    const slides = target.getSlides();
    for (const slide of slides) {
      const shape = findPlaceholderShape_(slide, token);
      if (shape) return shape;
    }
    return null;
  }

  // Caso 2: veio um slide
  if (!target.getPageElements) return null;

  const elems = target.getPageElements();

  for (const pe of elems) {
    if (
      pe.getPageElementType &&
      pe.getPageElementType() === SlidesApp.PageElementType.SHAPE
    ) {
      const sh = pe.asShape();
      const txt = sh.getText && sh.getText().asString
        ? sh.getText().asString()
        : '';
      if (txt && txt.indexOf(token) !== -1) {
        return sh;
      }
    }
  }

  return null;
}


/**
 * FOTO_ESTACAO – caber 100% dentro do box (contain), sem vazar, mantendo proporção.
 */
function insertImageFitPlaceholder_(presentationOrSlide, token, blob) {

  const ph = findPlaceholderShape_(presentationOrSlide, token);
  if (!ph) {
    throw new Error('Placeholder de imagem não encontrado: ' + token);
  }

  const slide = ph.getParentPage();

  const boxLeft = ph.getLeft();
  const boxTop  = ph.getTop();
  const boxW    = ph.getWidth();
  const boxH    = ph.getHeight();

  try { ph.getText().setText(''); } catch (e) {}

  const img = slide.insertImage(blob);

  const natW = img.getWidth();
  const natH = img.getHeight();

  if (!natW || !natH) {
    img.setLeft(boxLeft).setTop(boxTop).setWidth(boxW).setHeight(boxH);
    return;
  }

  const scale = Math.min(boxW / natW, boxH / natH);
  const newW  = natW * scale;
  const newH  = natH * scale;

  const x = boxLeft + (boxW - newW) / 2;
  const y = boxTop  + (boxH - newH) / 2;

  img.setLeft(x).setTop(y).setWidth(newW).setHeight(newH);
}


/**
 * SAT_IMG, MAP_IMG, STREET_IMG – escalar para percent% do box (largura/altura), centralizado.
 * percent = 1.0 → 100%
 */
function insertImageScaledPlaceholder_(presentation, token, blob, percent) {
  const slides = presentation.getSlides();
  for (const s of slides) {
    const ph = findPlaceholderShape_(s, token);
    if (!ph) continue;

    const boxLeft = ph.getLeft();
    const boxTop  = ph.getTop();
    const boxW    = ph.getWidth();
    const boxH    = ph.getHeight();

    try { ph.getText().setText(''); } catch(e) {}

    const img = s.insertImage(blob);
    const natW = img.getWidth();
    const natH = img.getHeight();
    if (!natW || !natH) {
      img.setLeft(boxLeft).setTop(boxTop).setWidth(boxW * percent);
      return true;
    }

    // Escala para percent * min(boxW/natW, boxH/natH) → 100% do máximo que caberia
    const scaleMax = Math.min(boxW / natW, boxH / natH);
    const scale    = scaleMax * (percent || 1.0);
    const newW     = natW * scale;
    const newH     = natH * scale;

    const x = boxLeft + (boxW - newW) / 2;
    const y = boxTop  + (boxH - newH) / 2;

    img.setLeft(x).setTop(y).setWidth(newW).setHeight(newH);
    return true;
  }
  return false;
}




function validarEGerarPdfSeguro_(tmpFileId, pdfName, pastaDestino) {

  // Drive e eventual -- espera curta obrigatoria
  Utilities.sleep(800);

  var tmpFile = DriveApp.getFileById(tmpFileId);
  if (!tmpFile) {
    throw new Error('Arquivo temporario do croqui nao encontrado: ' + tmpFileId);
  }

  // Tentar exportar via URL primeiro (contorna cota de conversao)
  var blob = _exportarSlidesComoPdf_(tmpFileId, pdfName);

  // Fallback: getAs tradicional se a exportacao via URL falhar
  if (!blob) {
    Logger.log('validarEGerarPdfSeguro_: fallback para getAs (arquivo ' + tmpFileId + ')');
    blob = tmpFile.getAs('application/pdf');
  }

  if (!blob || !blob.getBytes || blob.getBytes().length < 1000) {
    throw new Error('Falha ao gerar PDF do croqui (blob invalido ou vazio).');
  }

  blob.setName(pdfName);
  var pdfFile = pastaDestino.createFile(blob);
  return pdfFile;
}

function _exportarSlidesComoPdf_(fileId, pdfName) {
  try {
    var token = ScriptApp.getOAuthToken();

    var url = 'https://www.googleapis.com/drive/v3/files/'
      + encodeURIComponent(fileId)
      + '/export?mimeType=application%2Fpdf';

    var opcoes = {
      method:      'GET',
      headers:     { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch(url, opcoes);
    var code = resp.getResponseCode();

    if (code !== 200) {
      Logger.log('_exportarSlidesComoPdf_ HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
      return null;
    }

    var blob = resp.getBlob();
    if (!blob || blob.getBytes().length < 1000) {
      Logger.log('_exportarSlidesComoPdf_: blob vazio ou muito pequeno');
      return null;
    }

    blob.setName(pdfName || 'croqui.pdf');
    blob.setContentType('application/pdf');
    return blob;

  } catch (e) {
    Logger.log('_exportarSlidesComoPdf_ erro: ' + e);
    return null;
  }
}


/******************* (4) CROQUIS EM BLOCO / RETOMADA ********************/
const CROQUI_STATE_KEY = 'croqui_next_row';
const NORMALIZAR_STATE_KEY  = 'normalizar_next_row';

function gerarCroquiParaLinha(sheet, row, silencioso) {

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const idxTipo = headers.indexOf(COL.TipoEstacao);
  const idxCodigo = headers.indexOf(COL.Codigo);

  if (idxTipo === -1) {
    throw new Error('Coluna TipoEstacao não encontrada.');
  }

  const tipoBruto = sheet.getRange(row, idxTipo + 1).getDisplayValue();
  const tipoEstacao = String(tipoBruto || '').trim().toUpperCase();

  Logger.log('DEBUG — linha: ' + row);
  Logger.log('DEBUG — TipoEstacao bruto: ' + tipoBruto);
  Logger.log('DEBUG — TipoEstacao normalizado: ' + tipoEstacao);

  if (tipoEstacao === 'PUBLICA' || tipoEstacao === 'PÚBLICA') {
    gerarCroquiPublico_(sheet, row, silencioso);
    return;
  }

  if (tipoEstacao === 'PRIVADA') {
    gerarCroquiPrivado_(sheet, row, silencioso);
    return;
  }

  // fallback explícito (antes era silencioso)
  Logger.log(
    '⚠️ Croqui ignorado. TipoEstacao inválido: "' +
    tipoEstacao + '" na linha ' + row
  );

  const idxStatus = headers.indexOf('CroquiStatus');
  if (idxStatus !== -1) {
    sheet.getRange(row, idxStatus + 1).setValue('IGNORADO');
  }
}


/**
 * resolverTemplatePorPaisETipo_
 * ----------------------------------------------------
 * Resolve template e pasta de destino do croqui
 * com base em País + Tipo de Estação
 */
function resolverTemplatePorPaisETipo_(pais, tipo, ctx) {
  ctx  = ctx  || {};
  pais = String(pais || 'BR').toUpperCase();
  tipo = String(tipo || '').toUpperCase();
  if (pais !== 'MX') pais = 'BR';

  var cidade   = normalizar(ctx.cidade   || '');
  var alcaldia = normalizar(ctx.alcaldia || '');

  // ── BRASIL ──────────────────────────────────────────────────
  if (pais === 'BR') {
    var templateId = tipo === 'PUBLICO'
      ? SLIDES_TEMPLATE_PUBLICO_ID
      : SLIDES_TEMPLATE_PRIVADO_ID;

    var pastaRaizId = tipo === 'PUBLICO'
      ? CROQUIS_PUBLICOS_FOLDER_ID
      : CROQUIS_PRIVADOS_FOLDER_ID;

    var pastaRaiz   = DriveApp.getFolderById(pastaRaizId);
    var pastaCidade = getOuCriarPasta_(pastaRaiz, cleanName_(ctx.cidade || 'Sem_Cidade'));

    var nomeSubpasta;
    var cidUp = cidade.toUpperCase();
    if (cidUp === 'SAO PAULO' || cidUp === 'SÃO PAULO') {
      nomeSubpasta = cleanName_('Subprefeitura_' + (ctx.subpref || ctx.subprefeitura || 'Sem_Subprefeitura'));
    } else {
      nomeSubpasta = cleanName_(ctx.bairro || 'Sem_Bairro');
    }

    return {
      templateId:   templateId,
      pastaDestino: getOuCriarPasta_(pastaCidade, nomeSubpasta)
    };
  }

  // ── MÉXICO ───────────────────────────────────────────────────
  var isCDMX = cidade.includes('CIUDAD DE MEXICO') || cidade.includes('CIUDAD DE MÉXICO');

  var templateId = tipo === 'PUBLICO'
    ? SLIDES_TEMPLATE_PUBLICO_MX_ID
    : SLIDES_TEMPLATE_PRIVADO_MX_ID;

  var pastaRaizMxId = isCDMX
    ? (tipo === 'PUBLICO' ? CROQUIS_PUBLICOS_MX_FOLDER_ID            : CROQUIS_PRIVADOS_MX_FOLDER_ID)
    : (tipo === 'PUBLICO' ? CROQUIS_PUBLICOS_MX_FORA_CDMX_FOLDER_ID  : CROQUIS_PRIVADOS_MX_FORA_CDMX_FOLDER_ID);

  var nomeSubpasta = isCDMX
    ? (ctx.bairro || alcaldia || 'VALIDAR_BAIRRO')
    : (ctx.cidade || 'VALIDAR_MUNICIPIO');

  var pastaBaseMx  = DriveApp.getFolderById(pastaRaizMxId);
  var pastaDestino = getOuCriarPasta_(pastaBaseMx, nomeSubpasta);

  return { templateId: templateId, pastaDestino: pastaDestino };
}

function getOuCriarPasta_(pastaBase, nomePasta) {
  nomePasta = String(nomePasta || 'Sem_Nome').trim();
  var it = pastaBase.getFoldersByName(nomePasta);
  return it.hasNext() ? it.next() : pastaBase.createFolder(nomePasta);
}




function gerarCroquiFinal_(sheet, row, tipo) {

  const { headers, get } = makeRowAccessor_(sheet, row);

  const colCroqui = headers.indexOf(COL.Croqui) + 1;
  if (!colCroqui) {
    throw new Error('Coluna "Croqui" não encontrada.');
  }

  const pais = (() => {
    const cidade = String(get('Cidade') || '').toUpperCase();
    const endereco = String(get('Endereço completo da estação') || '').toUpperCase();
    const alcaldia = String(get('Alcaldia') || '').toUpperCase();

    if (
      cidade.includes('CIUDAD') ||
      cidade.includes('MEXICO') ||
      endereco.includes('CDMX') ||
      endereco.includes('MÉXICO') ||
      alcaldia
    ) {
      return 'MX';
    }

    return 'BR';
  })();

  const colPais = headers.indexOf('Pais') + 1;
  if (colPais) {
    sheet.getRange(row, colPais).setValue(pais);
  }

  // ================= CORE =================
  const coreResult = gerarCroqui_CORE(sheet, row);
  if (!coreResult || !coreResult.imagens) {
    throw new Error('CORE não retornou imagens válidas.');
  }

  // Converte blobs para garantir compatibilidade
  const imagens = {
    sat: coreResult.imagens.sat
      ? Utilities.newBlob(coreResult.imagens.sat.getBytes(), 'image/png', 'sat.png')
      : null,
    map: coreResult.imagens.map
      ? Utilities.newBlob(coreResult.imagens.map.getBytes(), 'image/png', 'map.png')
      : null,
    street: coreResult.imagens.street
      ? Utilities.newBlob(coreResult.imagens.street.getBytes(), 'image/png', 'street.png')
      : null
  };

  // ================= SALVAR IMAGENS AUXILIARES =================
  try {
    salvarImagensNoDrive2_(sheet, row, imagens);
  } catch (e) {
    Logger.log('Aviso ao salvar imagens auxiliares do croqui na linha ' + row + ': ' + e);
  }

  // ================= TEMPLATE / PASTA =================
  const resolved = resolverTemplatePorPaisETipo_(pais, tipo, {
    cidade: get('Cidade'),
    alcaldia: get('Alcaldia'),
    bairro: get('Bairro'),
    subprefeitura: get('Subprefeitura')
  });

  const templateId = resolved.templateId;
  const pastaDestino = resolved.pastaDestino;

  // ================= SLIDES TEMP =================
  const tmpFile = DriveApp
    .getFileById(templateId)
    .makeCopy('TMP_CROQUI_' + tipo + '_' + row, pastaDestino);

  const presentation = SlidesApp.openById(tmpFile.getId());

  // ================= TOKENS =================
  const tokenMap = buildTokenMap_(sheet, row, {
    tipoEstacao: tipo,
    pais: pais,
    v2Juridico: true
  });

  replaceTokensInSlidesNoClose_(presentation, tokenMap);

  // ================= FOTO ESTAÇÃO (COM FALLBACK) =================
  let fotoBlob = null;

  try {
    const fotoVal = get(COL.Foto);
    if (fotoVal) {
      const fileFoto = resolveFileFromCell_(
        fotoVal,
        DriveApp.getFolderById(FOTOS_FOLDER_ID)
      );
      if (fileFoto) {
        fotoBlob = fileFoto.getBlob();
      }
    }
  } catch (e) {
    Logger.log('Aviso ao buscar foto da estação na linha ' + row + ': ' + e);
  }

  if (!fotoBlob && imagens.street) {
    fotoBlob = imagens.street;
  }

  if (fotoBlob) {
    insertImageFitPlaceholder_(
      presentation,
      '{{FOTO_ESTACAO}}',
      fotoBlob
    );
  }

  // ================= MAPAS =================
  if (imagens.sat) {
    insertImageScaledPlaceholder_(presentation, '{{SAT_IMG}}', imagens.sat, 1.0);
  }
  if (imagens.map) {
    insertImageScaledPlaceholder_(presentation, '{{MAP_IMG}}', imagens.map, 1.0);
  }
  if (imagens.street) {
    insertImageScaledPlaceholder_(presentation, '{{STREET_IMG}}', imagens.street, 1.0);
  }

  presentation.saveAndClose();

  // ================= PDF =================
  const nomeEstacao = cleanName_(get(COL.NomeEstacao) || ('ESTACAO_' + row));
  const pdfFile = validarEGerarPdfSeguro_(
    tmpFile.getId(),
    'Croqui_' + nomeEstacao + '.pdf',
    pastaDestino
  );

  // ================= PLANILHA =================
  sheet.getRange(row, colCroqui).setValue(pdfFile.getUrl());

  // ================= CLEANUP =================
  try {
    tmpFile.setTrashed(true);
  } catch (e) {
    Logger.log('Aviso no cleanup do arquivo temporário do croqui: ' + e);
  }

  return pdfFile.getUrl();
}


function gerarCroquisLinhaAtual() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Aba "Estacoes" não encontrada.');

  const range = sh.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Nenhuma célula selecionada.');
    return;
  }

  const row = range.getRow();

  // 🚫 BLOQUEIO DEFINITIVO
  if (row <= 1) {
    SpreadsheetApp.getUi().alert(
      'Selecione uma linha de estação válida (abaixo do cabeçalho).'
    );
    return;
  }

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idxRowKey = headers.indexOf(COL.RowKey);

  if (idxRowKey === -1) {
    throw new Error('Coluna RowKey não encontrada.');
  }

  const rowKey = sh.getRange(row, idxRowKey + 1).getValue();
  if (!rowKey) {
    SpreadsheetApp.getUi().alert(
      'Linha inválida.\n\nEsta linha não possui RowKey.'
    );
    return;
  }

  renomearFotoParaLinha_(sh, row);
  gerarCroquiComRetry_(sh, row);
}


function gerarCroquisSelecao() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const sel = sh.getActiveRange();
  if (!sel) {
    SpreadsheetApp.getUi().alert('Selecione um bloco de linhas.');
    return;
  }
  const start = Math.max(2, sel.getRow());
  const end   = Math.min(sh.getLastRow(), start + sel.getNumRows() - 1);
  _processarCroquisPorIntervalo_(sh, start, end);
}

function gerarCroquisIntervaloPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Gerar Croquis (intervalo)', 'Linha inicial (>= 2):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Gerar Croquis (intervalo)', 'Linha final:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  const start = Math.max(2, parseInt(r1.getResponseText(), 10));
  const end   = parseInt(r2.getResponseText(), 10);
  if (isNaN(start) || isNaN(end) || end < start) {
    ui.alert('Intervalo inválido.');
    return;
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  _processarCroquisPorIntervalo_(sh, start, Math.min(end, sh.getLastRow()));
}

function reprocessarCroquisComErro() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const cStatus = headers.indexOf('CroquiStatus') + 1;
  const cTent   = headers.indexOf('CroquiTentativas') + 1;

  if (!cStatus || !cTent) {
    SpreadsheetApp.getUi().alert('Colunas de controle não encontradas.');
    return;
  }

  const rows = [];
  const lastRow = sh.getLastRow();

  for (let r = 2; r <= lastRow; r++) {
    const status = String(sh.getRange(r, cStatus).getValue());
    if (status === 'ERRO') {
      sh.getRange(r, cTent).setValue(0);
      sh.getRange(r, cStatus).setValue('PENDENTE');
      rows.push(r);
    }
  }

  if (!rows.length) {
    SpreadsheetApp.getUi().alert('Nenhum croqui em erro para reprocessar.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('CROQUI_FILA_ROWS', JSON.stringify(rows));
  props.setProperty('CROQUI_FILA_INDEX', '0');

  criarTriggerFilaCroquis_();

  SpreadsheetApp.getUi().alert(
    'Reprocessamento iniciado.\n\n' +
    'Croquis recolocados na fila: ' + rows.length
  );
}

function gerarCroquisTodos() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  _processarCroquisPorIntervalo_(sh, 2, sh.getLastRow());
}

function continuarCroquisPendentes() {
  criarTriggerFilaCroquis_();
  SpreadsheetApp.getUi().alert('Fila retomada.');
}

function finalizarFilaCroquis_() {
  const props = PropertiesService.getScriptProperties();

  // ✅ Apaga somente chaves relacionadas à fila de croquis
  const keys = props.getKeys();
  keys.forEach(k => {
    if (k && k.startsWith('CROQUI_')) {
      props.deleteProperty(k);
    }
  });

  // remove triggers da fila
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'processarFilaCroquis_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  SpreadsheetApp.getUi().alert('Fila de croquis concluída com sucesso.');
}
function criarGatilhoFilaCroquis() {
  removerGatilhoFilaCroquis();

  ScriptApp.newTrigger('processarFilaCroquis_')
    .timeBased()
    .everyMinutes(2)
    .create();

  SpreadsheetApp.getUi().alert('Gatilho criado: fila roda a cada 2 minutos.');
}

function removerGatilhoFilaCroquis() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processarFilaCroquis_') {
      ScriptApp.deleteTrigger(t);
    }
  });
}


function criarTriggerFilaCroquis_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processarFilaCroquis_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('processarFilaCroquis_')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function _processarCroquisPorIntervalo_(sh, startRow, endRow) {
  const HARD_LIMIT_MS = 5 * 60 * 1000; // 5 min
  const startTime = Date.now();
  let processed = 0;
  const props = PropertiesService.getScriptProperties();

  for (let r = startRow; r <= endRow; r++) {
    if (Date.now() - startTime > HARD_LIMIT_MS) {
      props.setProperty(CROQUI_STATE_KEY, String(r));
      SpreadsheetApp.getUi().alert(
        'Tempo quase no limite. Processadas ' + processed +
        ' linha(s). Use "Continuar croquis pendentes". Próxima linha: ' + r
      );
      return;
    }
    try {
      renomearFotoParaLinha_(sh, r);
      gerarCroquiParaLinha(sh, r, true);
      processed++;
      Utilities.sleep(150);
    } catch (e) {
      Logger.log('Falha na linha ' + r + ': ' + e);
    }
  }

  props.deleteProperty(CROQUI_STATE_KEY);
  SpreadsheetApp.getUi().alert(
    'Croquis gerados para ' + processed + ' linha(s) (linhas ' + startRow + '–' + endRow + ').'
  );
}

function processarFilaCroquis_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) return;

  var props = PropertiesService.getScriptProperties();
  var start = Number(props.getProperty('CROQUI_FILA_ATUAL') || 0);
  var end   = Number(props.getProperty('CROQUI_FILA_FIM')   || 0);

  if (!start || !end) {
    SpreadsheetApp.getUi().alert('Nenhuma fila ativa.');
    return;
  }

  var HARD_LIMIT_MS = 5 * 60 * 1000;
  var startTime     = Date.now();
  var processadas   = 0;

  var headers   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colStatus = headers.indexOf('CroquiStatus')    + 1;
  var colErro   = headers.indexOf('CroquiUltimoErro') + 1;

  for (var row = start; row <= end; row++) {

    if (Date.now() - startTime > HARD_LIMIT_MS) {
      props.setProperty('CROQUI_FILA_ATUAL', String(row));
      SpreadsheetApp.getUi().alert(
        'Tempo limite atingido.\n' +
        'Processados: ' + processadas + '\n' +
        'Use "Continuar fila" para retomar da linha ' + row + '.'
      );
      return;
    }

    try {
      gerarCroquiComRetry_(sh, row, true);
      processadas++;
      Utilities.sleep(120);
    } catch (e) {
      var msg = String(e.message || e).toLowerCase();

      var isQuota =
        msg.indexOf('muitas vezes')            !== -1 ||
        msg.indexOf('too many calls')           !== -1 ||
        msg.indexOf('service invoked too many') !== -1 ||
        msg.indexOf('conversion')               !== -1 ||
        msg.indexOf('rate limit')               !== -1 ||
        msg.indexOf('exceeded')                 !== -1;

      if (isQuota) {
        if (colStatus) sh.getRange(row, colStatus).setValue('PAUSADO_QUOTA');
        if (colErro)   sh.getRange(row, colErro).setValue('Quota diaria: ' + String(e.message || e));
        props.setProperty('CROQUI_FILA_ATUAL', String(row));

        SpreadsheetApp.getUi().alert(
          'LIMITE DIARIO DE CONVERSAO ATINGIDO\n\n' +
          'Processados hoje: ' + processadas + '\n' +
          'Parado na linha: ' + row + '\n\n' +
          'Aguarde ate amanha e use "Continuar fila" para retomar.\n' +
          'A linha ' + row + ' esta marcada como PAUSADO_QUOTA na planilha.'
        );
        return;
      }

      Logger.log('Fila erro L' + row + ': ' + e);
    }
  }

  props.deleteProperty('CROQUI_FILA_ATUAL');
  props.deleteProperty('CROQUI_FILA_FIM');
  SpreadsheetApp.getUi().alert('Fila concluida! Croquis gerados: ' + processadas);
}



/***************** BAIRROS / REPARO DE ENDEREÇOS *****************/
const COL_LOC = 'Localização';

function updateBairrosAuto() {
  fillBairrosAuto();
}





function _normalizarListaCore_(sh, startRow, endRow, props) {
  const HARD_LIMIT_MS = 5 * 60 * 1000; // ~5 minutos
  const startTime = Date.now();

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];

  const cFoto  = headers.indexOf(COL.Foto)      + 1;
  const cSat   = headers.indexOf(COL.ImgSat)    + 1;
  const cMap   = headers.indexOf(COL.ImgMapa)   + 1;
  const cStr   = headers.indexOf(COL.ImgStreet) + 1;
  const cCroq  = headers.indexOf(COL.Croqui)    + 1;

  const fotosRoot   = DriveApp.getFolderById(FOTOS_FOLDER_ID);
  const imagensRoot = DriveApp.getFolderById(IMAGENS_FOLDER_ID);
  const croquisRoot = DriveApp.getFolderById(CROQUIS_FOLDER_ID);

  let countFoto  = 0;
  let countSat   = 0;
  let countMap   = 0;
  let countStr   = 0;
  let countCroqi = 0;

  for (let row = startRow; row <= endRow; row++) {

    // se estiver perto de estourar o tempo, salva o próximo row e sai
    if (Date.now() - startTime > HARD_LIMIT_MS) {
      props.setProperty(NORMALIZAR_STATE_KEY, String(row));
      SpreadsheetApp.getUi().alert(
        'Tempo quase no limite.\n' +
        'Normalização parcial concluída até a linha ' + (row - 1) + '.\n' +
        'Use "Continuar normalização pendente" para seguir a partir da linha ' + row + '.'
      );
      return;
    }

    try {
      // 1) Foto da Estação – move/renomeia pelo endereço e grava URL
      try {
        renomearFotoParaLinha_(sh, row);
        countFoto++;
      } catch (e) {
        Logger.log('Erro ao normalizar foto na linha ' + row + ': ' + e);
      }

      // 2) Imagem Satélite
      if (cSat) {
        const vSat = String(sh.getRange(row, cSat).getValue() || '').trim();
        if (vSat && !vSat.startsWith('http')) {
          const fileSat = resolveFileFromCell_(vSat, imagensRoot);
          if (fileSat) {
            setCellToPublicUrl_(sh, row, cSat, fileSat);
            countSat++;
          }
        }
      }

      // 3) Imagem Mapa
      if (cMap) {
        const vMap = String(sh.getRange(row, cMap).getValue() || '').trim();
        if (vMap && !vMap.startsWith('http')) {
          const fileMap = resolveFileFromCell_(vMap, imagensRoot);
          if (fileMap) {
            setCellToPublicUrl_(sh, row, cMap, fileMap);
            countMap++;
          }
        }
      }

      // 4) Street View
      if (cStr) {
        const vStr = String(sh.getRange(row, cStr).getValue() || '').trim();
        if (vStr && !vStr.startsWith('http')) {
          const fileStr = resolveFileFromCell_(vStr, imagensRoot);
          if (fileStr) {
            setCellToPublicUrl_(sh, row, cStr, fileStr);
            countStr++;
          }
        }
      }

      // 5) Croqui (PDF)
      if (cCroq) {
        const vCroq = String(sh.getRange(row, cCroq).getValue() || '').trim();
        if (vCroq && !vCroq.startsWith('http')) {
          const name = vCroq.split('/').pop();
          const it = croquisRoot.getFilesByName(name);
          if (it.hasNext()) {
            const fileCroq = it.next();
            setCellToPublicUrl_(sh, row, cCroq, fileCroq);
            countCroqi++;
          }
        }
      }

      Utilities.sleep(80); // suaviza o uso do Drive

    } catch (e) {
      Logger.log('Erro na normalização da linha ' + row + ': ' + e);
    }
  }

  // Se chegou até aqui, terminou tudo
  props.deleteProperty(NORMALIZAR_STATE_KEY);
  SpreadsheetApp.getUi().alert(
    'Normalização concluída.\n' +
    'Linhas: ' + startRow + '–' + endRow + '\n\n' +
    'Fotos: '   + countFoto  + '\n' +
    'Satélite: '+ countSat   + '\n' +
    'Mapa: '    + countMap   + '\n' +
    'Street: '  + countStr   + '\n' +
    'Croquis: ' + countCroqi
  );
}

function renomearFotosSomenteLinhasFiltradas() {
  const SHEET_NAME = 'Estacoes';

  const COL_NOME_ESTACAO = 4; // Coluna D
  const COL_FOTO = 11;        // Coluna K
  const PASTA_DESTINO = 'Matinhos';

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const lastRow = sheet.getLastRow();

  // Pasta Matinhos
  let pasta;
  const pastas = DriveApp.getFoldersByName(PASTA_DESTINO);
  pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(PASTA_DESTINO);

  for (let row = 2; row <= lastRow; row++) {

    // PULA LINHAS OCULTAS PELO FILTRO
    if (sheet.isRowHiddenByFilter(row)) continue;

    const nomeEstacao = sheet.getRange(row, COL_NOME_ESTACAO).getValue();
    const caminhoFoto = sheet.getRange(row, COL_FOTO).getValue();

    if (!nomeEstacao || !caminhoFoto) continue;

    const nomeArquivoAtual = caminhoFoto.toString().split('/').pop();

    const arquivos = DriveApp.searchFiles(
      `title = "${nomeArquivoAtual}" and trashed = false`
    );

    if (!arquivos.hasNext()) continue;

    const arquivo = arquivos.next();

    const nomeLimpo = nomeEstacao
      .toString()
      .replace(/[\\/:*?"<>|]/g, '')
      .trim();

    const extensao = nomeArquivoAtual.split('.').pop();
    const novoNome = `${nomeLimpo}.${extensao}`;

    arquivo.setName(novoNome);
    pasta.addFile(arquivo);

    sheet.getRange(row, COL_FOTO).setValue(`${PASTA_DESTINO}/${novoNome}`);
  }

  Logger.log('Processo concluído apenas nas linhas filtradas');
}

/****************************WEBAPP*****************************/
function listarEstacoesParaVisualizacao() {
  const ss = getSS_();
  const aba = getSheetAny_(ss, ['Estacoes','Estações','ESTACOES','ESTAÇÕES']);
  if (!aba) throw new Error('Aba de estações não encontrada (esperado: Estacoes/Estações)');


  const dados = aba.getDataRange().getValues();
  const header = dados.shift();

  const idx = nome => header.indexOf(nome);

  const iCodigo = idx('CodigoEstacao');
  const iCidade = idx('Cidade');
  const iTipo = idx('TipoEstacao');
  const iStatus = idx('StatusEstacao');
  const iLat = idx('Latitude');
  const iLng = idx('Longitude');
  const iEndereco = idx('Endereço completo da estação');
  const iBairro = idx('Bairro');
  const iSubpref = idx('Subprefeitura');
  const iCroqui = idx('Croqui');
  const iFoto = idx('Foto da Estação');
  const iMod  = idx('Modalidade');
  const iStreet = idx('Street View');
  const iTipoPublica = idx('TipoPublica');
  const iTPU = idx('TPU');
  const iNomeLocalPrivado   = idx('NomeLocalPrivado');
  const iNomeAutorizante   = idx('NomeAutorizante');
  const iCargoAutorizante  = idx('CargoAutorizante');
  const iTelefoneAutorizante = idx('TelefoneAutorizante');
  const iEmailAutorizante  = idx('EmailAutorizante');
  const iDocumentoAutorizacao = idx('DocumentoAutorizacao');


  // 🔁 DUPLICIDADES
  const iSeqGlobal = idx('SeqGlobal');
  const iAddrNorm  = idx('AddrNorm');
  const iDupGrupo  = idx('DupGrupo');
  const iDupMotivo = idx('DupMotivo');

  const estacoes = [];

  dados.forEach(l => {
    const lat = parseFloat(String(l[iLat]).replace(',', '.'));
    const lng = parseFloat(String(l[iLng]).replace(',', '.'));



    // ignora registros inválidos para mapa
    if (isNaN(lat) || isNaN(lng) ||
      Math.abs(lat) > 90 ||
      Math.abs(lng) > 180) return;


    estacoes.push({
      // ===== IDENTIDADE =====
      codigo: l[iCodigo] || '',
      cidade: l[iCidade] || '',
      tipo: normalizar(l[iTipo] || ''),

      // ===== STATUS / CLASSIFICAÇÃO =====
      statusEstacao: normalizar(l[iStatus] || 'ATIVO'),
      subprefeitura: l[iSubpref] || 'VALIDAR',
      tipoPublica: l[iTipoPublica] || '',

      // ===== LOCALIZAÇÃO =====
      lat: lat,
      lng: lng,
      endereco: l[iEndereco] || '',
      bairro: l[iBairro] || '',
      localizacao: lat + ',' + lng,

      // ===== PRIVADO =====
      nomeLocalPrivado: l[iNomeLocalPrivado] || '',
      nomeAutorizante: l[iNomeAutorizante] || '',
      cargoAutorizante: l[iCargoAutorizante] || '',
      telefoneAutorizante: l[iTelefoneAutorizante] || '',
      emailAutorizante: l[iEmailAutorizante] || '',
      documentoAutorizacao: l[iDocumentoAutorizacao] || '',


      // ===== ARQUIVOS =====
      croqui: l[iCroqui] || '',
      modalidade: iMod >= 0 ? (l[iMod] || 'PATINETE') : 'PATINETE',
      foto: l[iFoto] || '',
      street: iStreet >= 0 ? (l[iStreet] || '') : '',

      // ===== OUTROS =====
      tpu: l[iTPU] || '',

      // ===== DUPLICIDADES =====
      seqGlobal: l[iSeqGlobal] || '',
      addrNorm:  l[iAddrNorm]  || '',
      dupGrupo:  l[iDupGrupo]  || '',
      dupMotivo: l[iDupMotivo] || ''
    });
  });

  return estacoes;
}

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
 
  // Chamada de API dos PWAs via GET
  if (params.action) {
    return handleApiGet_(e);
  }
 
  var page     = String(params.page || 'index').replace(/[^a-zA-Z0-9]/g, '');
  var pageMap  = { 'index': 'index', 'campo': 'pwaCampo', 'gestor': 'pwaGestor' };
  var fileName = pageMap[page] || 'index';
 
  var props    = PropertiesService.getScriptProperties();
  var gmapsKey = props.getProperty('GMAPS_API_KEY') || '';
  var oauthId  = props.getProperty('OAUTH_CLIENT_ID') || '';
  var backend  = ScriptApp.getService().getUrl();
 
  if (!gmapsKey) throw new Error('GMAPS_API_KEY nao definida nas Script Properties');
 
  // index.html -- template original (arquivo pequeno, funciona com createTemplateFromFile)
  if (fileName === 'index') {
    var tpl = HtmlService.createTemplateFromFile('index');
    tpl.GMAPS_API_KEY = gmapsKey;
    return tpl.evaluate()
      .setTitle('App Estacoes -- Mapa')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
 
  // pwaCampo e pwaGestor -- string replace (arquivos grandes, sem template parser)
  var content = HtmlService.createHtmlOutputFromFile(fileName).getContent();
 
  // pwaCampo: chave do Maps injetada via placeholder MAPS_KEY_VALUE
  content = content.replace('MAPS_KEY_VALUE',      gmapsKey);
  // pwaGestor: backend e oauth via placeholders
  content = content.replace('__BACKEND_URL__',     backend);
  content = content.replace('__OAUTH_CLIENT_ID__', oauthId);
 
  var title = fileName === 'pwaGestor' ? 'Estacoes Gestor' : 'Estacoes Campo';
 
  return HtmlService.createHtmlOutput(content)
    .setTitle(title)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
 
/**
 * API via GET -- payload em ?d= (evita CORS de origens externas)
 */
function handleApiGet_(e) {
  var result;
  try {
    var params  = e.parameter || {};
    var action  = String(params.action || '').trim();
    var payload = {};
    try { payload = JSON.parse(decodeURIComponent(params.d || '{}')); } catch(pe) {}
 
    if (action === 'geocodeEndereco') {
      result = geocodeEnderecoPWA_(payload);
    } else if (action === 'loginGestor') {
      result = loginGestor_(payload);
    } else {
      result = dispatchAction_(action, payload);
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
 
/**
 * Forward geocode para busca manual de endereco no campo.
 */
function geocodeEnderecoPWA_(params) {
  try {
    var endereco = String((params && params.endereco) || '').trim();
    if (!endereco) return { ok: false, error: 'Endereco obrigatorio.' };
 
    var key  = getMapsApiKey_();
    var url  = 'https://maps.googleapis.com/maps/api/geocode/json'
             + '?address=' + encodeURIComponent(endereco)
             + '&key='     + key;
 
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    if (!data.results || !data.results.length) return { ok: false, error: 'Nao encontrado.' };
 
    var r     = data.results[0];
    var loc   = r.geometry.location;
    var comps = r.address_components || [];
    var get   = function(type) {
      var f = comps.filter(function(c){ return c.types.indexOf(type) !== -1; })[0];
      return f ? f.long_name : '';
    };
 
    return {
      ok:       true,
      lat:      loc.lat,
      lng:      loc.lng,
      endereco: r.formatted_address || '',
      bairro:   get('sublocality_level_1') || get('neighborhood') || '',
      cidade:   get('locality') || get('administrative_area_level_2') || '',
      estado:   get('administrative_area_level_1') || '',
      pais:     get('country') === 'MX' ? 'MX' : 'BR'
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

  

function getEstacoesWebApp() {
  try {
    return listarEstacoesParaVisualizacao();
  } catch (e) {
    Logger.log('Erro getEstacoesWebApp: ' + e);
    return { ok:false, error: String(e) };
  }
}

  /******************** WEBAPP MOBILE LAZY (CIDADES) ********************/
  // Cole este bloco DEPOIS de getEstacoesWebApp() (ou no final do Código.gs).
  // Ele habilita: getCitiesIndex() e getEstacoesByCity(cityKey) para o mobile carregar por cidade.

  function getCitiesIndex() {
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get('citiesIndex_v1');
      if (cached) return JSON.parse(cached);

      const ss = getSS_();
      const sh = getSheetAny_(ss, ['Estacoes','Estações','ESTACOES','ESTAÇÕES']);
      if (!sh) throw new Error('Aba de estações não encontrada (esperado: Estacoes/Estações)');

      const values = sh.getDataRange().getValues();
      const header = values.shift();

      const idx = (name) => header.indexOf(name);
      const iCidade = idx('Cidade');
      const iLat = idx('Latitude');
      const iLng = idx('Longitude');

      if (iCidade < 0 || iLat < 0 || iLng < 0) {
        throw new Error('Colunas obrigatórias não encontradas (Cidade/Latitude/Longitude).');
      }

      const byKey = {}; // cityKey -> { cidade, cityKey, bounds, count }

      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        const cidade = String(row[iCidade] || '').trim();
        if (!cidade) continue;

        const lat = parseFloat(String(l[iLat]).replace(',', '.'));
        const lng = parseFloat(String(l[iLng]).replace(',', '.'));
        if (isNaN(lat) || isNaN(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;

        const cityKey = normalizeCityKey_(cidade);
        let rec = byKey[cityKey];
        if (!rec) {
          rec = byKey[cityKey] = {
            cidade,
            cityKey,
            count: 0,
            bounds: { n: lat, s: lat, e: lng, w: lng }
          };
        }

        rec.count++;
        if (lat > rec.bounds.n) rec.bounds.n = lat;
        if (lat < rec.bounds.s) rec.bounds.s = lat;
        if (lng > rec.bounds.e) rec.bounds.e = lng;
        if (lng < rec.bounds.w) rec.bounds.w = lng;
      }

      const list = Object.keys(byKey)
        .map(k => byKey[k])
        .sort((a,b) => String(a.cidade).localeCompare(String(b.cidade)));

      // Cache curto para acelerar e economizar leitura
      cache.put('citiesIndex_v1', JSON.stringify(list), 60 * 30); // 30 min
      return list;

    } catch (e) {
      Logger.log('Erro getCitiesIndex: ' + e);
      return [];
    }
  }

  function getEstacoesByCity(cityKey) {
    try {
      cityKey = String(cityKey || '').trim();
      if (!cityKey) return [];

      const cache = CacheService.getScriptCache();
      const ck = 'cityStations_v1_' + cityKey;
      const cached = cache.get(ck);
      if (cached) return JSON.parse(cached);

      const ss = getSS_();
      const sh = getSheetAny_(ss, ['Estacoes','Estações','ESTACOES','ESTAÇÕES']);
      if (!sh) throw new Error('Aba de estações não encontrada (esperado: Estacoes/Estações)');

      const values = sh.getDataRange().getValues();
      const header = values.shift();

      const idx = (name) => header.indexOf(name);
      const iCodigo = idx('CodigoEstacao');
      const iCidade = idx('Cidade');
      const iTipo = idx('TipoEstacao');
      const iStatus = idx('StatusEstacao');
      const iLat = idx('Latitude');
      const iLng = idx('Longitude');
      const iEndereco = idx('Endereço completo da estação');
      const iBairro = idx('Bairro');
      const iSubpref = idx('Subprefeitura');
      const iCroqui = idx('Croqui');
      const iFoto = idx('Foto da Estação');

      if ([iCidade,iLat,iLng].some(i => i < 0)) {
        throw new Error('Colunas obrigatórias não encontradas (Cidade/Latitude/Longitude).');
      }

      const out = [];

      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        const cidade = String(row[iCidade] || '').trim();
        if (!cidade) continue;
        if (normalizeCityKey_(cidade) !== cityKey) continue;

        const lat = parseFloat(String(l[iLat]).replace(',', '.'));
        const lng = parseFloat(String(l[iLng]).replace(',', '.'));
        if (isNaN(lat) || isNaN(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;

        out.push({
          codigo: iCodigo >= 0 ? (row[iCodigo] || '') : '',
          cidade,
          cityKey,
          tipo: iTipo >= 0 ? normalizar(row[iTipo] || '') : '',
          statusEstacao: iStatus >= 0 ? normalizar(row[iStatus] || 'ATIVO') : 'ATIVO',
          subprefeitura: iSubpref >= 0 ? (row[iSubpref] || '') : '',
          lat,
          lng,
          endereco: iEndereco >= 0 ? (row[iEndereco] || '') : '',
          bairro: iBairro >= 0 ? (row[iBairro] || '') : '',
          croqui: iCroqui >= 0 ? (row[iCroqui] || '') : '',
          foto: iFoto >= 0 ? (row[iFoto] || '') : ''
        });
      }

      // Cache curto: cidades grandes podem ter bastante registro
      cache.put(ck, JSON.stringify(out), 60 * 15); // 15 min
      return out;

    } catch (e) {
      Logger.log('Erro getEstacoesByCity: ' + e);
      return [];
    }
  }

  function normalizeCityKey_(cidade) {
    // Normaliza para comparar sem acento/case e com _
    cidade = String(cidade || '').trim().toLowerCase();
    if (!cidade) return '';
    try {
      cidade = cidade.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    } catch (e) {
      // normalize() pode falhar em runtimes antigos; ignora
    }
    cidade = cidade.replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    return cidade;
  }



function exportarFotosCamposDoJordao() {

  const SHEET_NAME = 'Estacoes';
  const CIDADE_ALVO = 'Campos do Jordão';
  const NOME_PASTA_DESTINO = 'Fotos_Estacoes_Campos_do_Jordao';

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Aba Estacoes não encontrada.');

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  const cCidade   = headers.indexOf('Cidade') + 1;
  const cEndereco = headers.indexOf('Endereço completo da estação') + 1;
  const cFoto     = headers.indexOf('Foto da Estação') + 1;

  if (!cCidade || !cEndereco || !cFoto) {
    throw new Error('Colunas obrigatórias não encontradas.');
  }

  // Pasta destino
  let pastaDestino;
  const it = DriveApp.getFoldersByName(NOME_PASTA_DESTINO);
  pastaDestino = it.hasNext() ? it.next() : DriveApp.createFolder(NOME_PASTA_DESTINO);

  const fotosRoot = DriveApp.getFolderById(FOTOS_FOLDER_ID);
  const lastRow = sh.getLastRow();

  let count = 0;

  for (let row = 2; row <= lastRow; row++) {

    const cidade = String(sh.getRange(row, cCidade).getValue() || '').trim();
    if (cidade !== CIDADE_ALVO) continue;

    const endereco = sh.getRange(row, cEndereco).getValue();
    const fotoVal  = sh.getRange(row, cFoto).getValue();

    if (!endereco || !fotoVal) continue;

    const file = resolveFileFromCell_(fotoVal, fotosRoot);
    if (!file) continue;

    // Nome limpo do arquivo
    const nomeBase = cleanName_(endereco);
    const ext = file.getName().match(/(\.[^.]+)$/)?.[1] || '.jpg';
    const novoNome = nomeBase + ext;

    // Copia para a pasta destino
    const copia = file.makeCopy(novoNome, pastaDestino);
    copia.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    count++;
  }

  SpreadsheetApp.getUi().alert(
    'Exportação concluída!\n' +
    'Fotos copiadas: ' + count + '\n' +
    'Pasta: ' + NOME_PASTA_DESTINO
  );
}

function importarEstacoesGenerico(nomeAbaImport = 'IMPORT_MYMAPS') {

  const ss = SpreadsheetApp.getActive();
  const shImport = ss.getSheetByName(nomeAbaImport);
  const shDest = ss.getSheetByName('Estacoes');

  if (!shImport || !shDest) {
    SpreadsheetApp.getUi().alert('Aba não encontrada.');
    return;
  }

  const data = shImport.getDataRange().getValues();
  if (data.length < 2) return;

  const headersImport = data[0].map(h => String(h).toLowerCase().trim());
  const headersDest = shDest.getRange(1,1,1,shDest.getLastColumn()).getValues()[0];

  const idxDest = name => headersDest.indexOf(name);

  const findCol = (possibleNames) => {
    for (let name of possibleNames) {
      const i = headersImport.indexOf(name.toLowerCase());
      if (i !== -1) return i;
    }
    return -1;
  };

  const colNome = findCol(['name','nome','title']);
  const colEndereco = findCol(['address','endereço','endereco','description']);
  const colLat = findCol(['latitude','lat']);
  const colLng = findCol(['longitude','lng','long']);
  const colCoord = findCol(['coordinates','location']);

  let count = 0;

  for (let i = 1; i < data.length; i++) {

    let lat = null;
    let lng = null;

    if (colLat !== -1 && colLng !== -1) {
      lat = Number(data[i][colLat]);
      lng = Number(data[i][colLng]);
    } else if (colCoord !== -1) {
      const coord = String(data[i][colCoord]);
      const parts = coord.split(',');
      if (parts.length === 2) {
        lat = Number(parts[0].trim());
        lng = Number(parts[1].trim());
      }
    }

    // 🔥 AJUSTE AUTOMÁTICO PARA COORDENADA ESCALADA
    if (lat && Math.abs(lat) > 1000) lat = lat / 1000000;
    if (lng && Math.abs(lng) > 1000) lng = lng / 1000000;

    if (!lat || !lng) continue;



    const newRow = Array(headersDest.length).fill('');

    if (colNome !== -1)
      newRow[idxDest('Nome da Estação')] = data[i][colNome];

    if (colEndereco !== -1)
      newRow[idxDest('Endereço completo da estação')] = data[i][colEndereco];

    newRow[idxDest('Latitude')] = lat;
    newRow[idxDest('Longitude')] = lng;
    newRow[idxDest('Localização')] = lat + ', ' + lng;

    newRow[idxDest('TipoEstacao')] = 'PUBLICA';
    newRow[idxDest('StatusEstacao')] = 'ATIVO';
    newRow[idxDest('CriadoPor')] = 'IMPORT';
    newRow[idxDest('OrigemDado')] = 'IMPORT_MYMAPS';
    newRow[idxDest('DataCriacao')] = new Date();

    shDest.appendRow(newRow);
    count++;
  }

  SpreadsheetApp.getUi().alert('Importação concluída: ' + count + ' registros.');
}



function preencherFotosPorEnderecoSelecao() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) {
    ui.alert('Aba "' + SHEET_NAME + '" não encontrada.');
    return;
  }

  const sel = sh.getActiveRange();
  if (!sel) {
    ui.alert('Selecione um intervalo de linhas.');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cEnd  = headers.indexOf(COL.Endereco) + 1;
  const cFoto = headers.indexOf(COL.Foto) + 1;

  if (cEnd === 0 || cFoto === 0) {
    ui.alert('Colunas Endereço ou Foto não encontradas.');
    return;
  }

  const fotosRoot = DriveApp.getFolderById(FOTOS_FOLDER_ID);

  const startRow = Math.max(2, sel.getRow());
  const endRow   = Math.min(sh.getLastRow(), startRow + sel.getNumRows() - 1);

  const HARD_LIMIT_MS = 5 * 60 * 1000;
  const startTime = Date.now();

  let associados = 0;
  let ignorados = 0;

  for (let row = startRow; row <= endRow; row++) {

    if (Date.now() - startTime > HARD_LIMIT_MS) {
      ui.alert(
        'Execução interrompida por tempo.\n\n' +
        'Fotos associadas: ' + associados + '\n' +
        'Ignoradas: ' + ignorados + '\n\n' +
        'Selecione novamente e rode para continuar.'
      );
      return;
    }

    const endereco = sh.getRange(row, cEnd).getValue();
    const fotoVal  = sh.getRange(row, cFoto).getValue();

    if (!endereco || fotoVal) {
      ignorados++;
      continue;
    }

    const nomeBusca = cleanName_(endereco);
    const query = `title contains "${nomeBusca}" and trashed = false`;

    const files = DriveApp.searchFiles(query);
    if (files.hasNext()) {
      const file = files.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sh.getRange(row, cFoto).setValue(file.getUrl());
      associados++;
    }
  }

  ui.alert(
    'Pré-normalização concluída.\n\n' +
    'Fotos associadas: ' + associados + '\n' +
    'Ignoradas: ' + ignorados
  );
}

function reprocessarFotosPorSubprefeitura() {
  const ss = SpreadsheetApp.getActive();
  const aba = ss.getSheetByName('Estacoes');

  if (!aba) throw new Error('Aba Estacoes não encontrada');

  const dados = aba.getDataRange().getValues();
  const header = dados[0];

  const iCidade = header.indexOf('Cidade');
  const iSubpref = header.indexOf('Subprefeitura');
  const iFoto = header.indexOf('Foto da Estação');
  const iNome = header.indexOf('Nome da Estação');

  const pastaRaiz = DriveApp.getFolderById(DRIVE_PASTA_ESTACOES_ID);

  for (let i = 1; i < dados.length; i++) {
    const cidade = dados[i][iCidade];
    const subpref = dados[i][iSubpref];
    const fotoUrl = dados[i][iFoto];
    const nome = dados[i][iNome];

    if (!fotoUrl) continue;
    if (!cidade) continue;
    if (cidade !== 'São Paulo') continue;
    if (!subpref || subpref === 'VALIDAR') continue;

    const fileId = extrairFileId(fotoUrl);
    if (!fileId) continue;

    const file = DriveApp.getFileById(fileId);

    // --- GARANTE PASTA CIDADE ---
    const pastaCidade = getOuCriarPasta(
      pastaRaiz,
      normalizar(cidade)
    );

    // --- GARANTE PASTA SUBPREF ---
    const pastaSubpref = getOuCriarPasta(
      pastaCidade,
      normalizar(subpref)
    );

    // --- MOVE DE VERDADE ---
    pastaSubpref.addFile(file);

    const pais = file.getParents();
    while (pais.hasNext()) {
      const p = pais.next();
      if (p.getId() !== pastaSubpref.getId()) {
        p.removeFile(file);
      }
    }

    // --- RENOMEIA (GARANTE PADRÃO) ---
    const nomeSeguro = nomeArquivoSeguro(nome) + '.jpg';
    file.setName(nomeSeguro);

    // --- ATUALIZA LINK (opcional, mas seguro) ---
    aba.getRange(i + 1, iFoto + 1).setValue(file.getUrl());
  }
}


function preencherSubprefeiturasSP() {
  const ss = SpreadsheetApp.getActive();
  const abaEst = ss.getSheetByName('Estacoes');
  const abaMapa = ss.getSheetByName('MAPA_SUBPREF_SP');

  const dadosEst = abaEst.getDataRange().getValues();
  const dadosMapa = abaMapa.getDataRange().getValues();

  const header = dadosEst[0];
  const iCidade = header.indexOf('Cidade');
  const iBairro = header.indexOf('Bairro');
  const iSubpref = header.indexOf('Subprefeitura');

  // cria dicionário BAIRRO → SUBPREF
  const mapa = {};
  for (let i = 1; i < dadosMapa.length; i++) {
    mapa[
      normalizar(dadosMapa[i][0])
    ] = dadosMapa[i][1];
  }

  for (let r = 1; r < dadosEst.length; r++) {
    const cidade = dadosEst[r][iCidade];
    const bairro = dadosEst[r][iBairro];
    const atual = dadosEst[r][iSubpref];

    if (cidade !== 'São Paulo') continue;
    if (!bairro) continue;
    if (atual && atual !== 'VALIDAR') continue;

    const subpref = mapa[normalizar(bairro)];
    if (subpref) {
      abaEst.getRange(r + 1, iSubpref + 1).setValue(subpref);
    }
  }
}


////////////////////LOGEVENTO//////////////////////////////
function logEvento_(tipo, linha, codigo, mensagem) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Logs');

    if (!sh) {
      sh = ss.insertSheet('Logs');
      sh.appendRow(['Data','Tipo','Linha','Codigo','Mensagem']);
    }

    sh.appendRow([
      new Date(),
      tipo,
      linha || '',
      codigo || '',
      mensagem
    ]);
  } catch (e) {
    // log nunca pode quebrar o sistema
  }
}


///////////////////////RESOLVERDUPLICIDADE/////////////////////
function resolverDuplicidadePlanilha(grupo, status, codigoPrincipal) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Estacoes');
  if (!sh) throw new Error('Aba Estacoes não encontrada');

  const data = sh.getDataRange().getValues();
  const header = data[0];

  const idxGrupo  = header.indexOf('DupGrupo');
  const idxStatus = header.indexOf('DupStatus');
  const idxCodigo = header.indexOf('CodigoEstacao');

  if (idxGrupo === -1 || idxStatus === -1) {
    throw new Error('Colunas DupGrupo ou DupStatus não encontradas');
  }

  let alterou = false;

  for (let i = 1; i < data.length; i++) {
    const linhaGrupo = data[i][idxGrupo];
    if (linhaGrupo !== grupo) continue;

    // ===== IGNORAR GRUPO =====
    if (status === 'IGNORADO') {
      sh.getRange(i + 1, idxStatus + 1).setValue('IGNORADO');
      alterou = true;
      continue;
    }

    // ===== VALIDAR =====
    if (status === 'VALIDADO') {
      if (data[i][idxCodigo] === codigoPrincipal) {
        sh.getRange(i + 1, idxStatus + 1).setValue('VALIDADO');
      } else {
        sh.getRange(i + 1, idxStatus + 1).setValue('IGNORADO');
      }
      alterou = true;
    }
  }

  if (!alterou) {
    throw new Error('Nenhuma linha foi alterada para o grupo ' + grupo);
  }

  return true;
}

/*************************************************
 * MIGRAÇÃO DE FOTOS — NOVA ÁRVORE BR / MX
 *************************************************/

const MIGRACAO_FOTOS_STATE_KEY = 'MIGRAR_FOTOS_ROW';

function migrarFotosParaNovaArvore() {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Estacoes');
  if (!sh) throw new Error('Aba Estacoes não encontrada.');

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  function col(name) {
    const idx = headers.indexOf(name);
    return idx === -1 ? null : idx + 1;
  }

  const colFoto   = col('Foto da Estação');
  const colPais   = col('Pais');
  const colCidade = col('Cidade');
  const colBairro = col('Bairro');
  const colSub    = col('Subprefeitura');
  const colNome   = col('Nome da Estação');

  if (!colFoto || !colCidade) {
    throw new Error(
      'Colunas obrigatórias não encontradas. ' +
      'Verifique: Foto da Estação / Cidade'
    );
  }

  if (!colFoto || !colCidade) {
    throw new Error('Colunas obrigatórias não encontradas.');
  }

  const pastaRaiz = DriveApp.getFolderById(DRIVE_PASTA_ESTACOES_ID);
  const props = PropertiesService.getScriptProperties();

  let startRow = Number(props.getProperty(MIGRACAO_FOTOS_STATE_KEY) || 2);
  const lastRow = sh.getLastRow();

  const HARD_LIMIT_MS = 5 * 60 * 1000;
  const startTime = Date.now();

  let processadas = 0;

  for (let row = startRow; row <= lastRow; row++) {

    if (Date.now() - startTime > HARD_LIMIT_MS) {
      props.setProperty(MIGRACAO_FOTOS_STATE_KEY, String(row));
      SpreadsheetApp.getUi().alert(
        '⏸️ Migração pausada.\n' +
        'Próxima linha: ' + row + '\n' +
        'Fotos migradas: ' + processadas + '\n\n' +
        'Execute novamente para continuar.'
      );
      return;
    }

    const fotoUrl = sh.getRange(row, colFoto).getValue();
    if (!fotoUrl) continue;

    const fileId = extrairFileId(fotoUrl);
    if (!fileId) continue;

    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (e) {
      Logger.log('Arquivo não acessível linha ' + row);
      continue;
    }

    const pais   = colPais   ? normalizar(sh.getRange(row, colPais).getValue() || 'BR') : 'BR';
    const bairro = colBairro ? sh.getRange(row, colBairro).getValue() : '';
    const sub    = colSub    ? sh.getRange(row, colSub).getValue() : '';
    const nome   = colNome   ? sh.getRange(row, colNome).getValue() : '';
    const cidade = colCidade ? sh.getRange(row, colCidade).getValue() : '';

    // ===== PASTA PAÍS =====
    const pastaPais = getOuCriarPasta(pastaRaiz, pais);

    // ===== PASTA CIDADE =====
    const pastaCidade = getOuCriarPasta(
      pastaPais,
      normalizar(cidade || 'SEM_CIDADE')
    );

    let pastaFinal = pastaCidade;

    // ===== BR =====
    if (
      pais === 'BR' &&
      normalizar(cidade) === 'SAO PAULO' &&
      sub &&
      sub !== 'VALIDAR'
    ) {
      pastaFinal = getOuCriarPasta(
        pastaCidade,
        normalizar(sub)
      );
    }

    // ===== MX =====
    if (pais === 'MX' && bairro) {
      pastaFinal = getOuCriarPasta(
        pastaCidade,
        normalizar(bairro)
      );
    }

    // ===== MOVE REAL =====
    pastaFinal.addFile(file);

    const paisAntigos = file.getParents();
    while (paisAntigos.hasNext()) {
      const p = paisAntigos.next();
      if (p.getId() !== pastaFinal.getId()) {
        p.removeFile(file);
      }
    }

    // ===== RENOMEIA (garante padrão) =====
    if (nome) {
      file.setName(nomeArquivoSeguro(nome) + '.jpg');
    }

    // ===== REGRAVA LINK (segurança) =====
    sh.getRange(row, colFoto).setValue(file.getUrl());

    processadas++;
    Utilities.sleep(80);
  }

  props.deleteProperty(MIGRACAO_FOTOS_STATE_KEY);

  SpreadsheetApp.getUi().alert(
    '✅ Migração concluída!\n\n' +
    'Total de fotos migradas: ' + processadas
  );
}




// (Polígonos) Implementação removida deste arquivo.
// A fonte única agora é poligonos.gs (cache versionado + endpoint getPoligonosCidade).



/**************************** ADD MODE (WEB MAPA) ****************************/
/**
 * Salve a senha em Script Properties:
 *   ADD_PASS = "..."
 */
function validateAddPass(pass) {
  pass = String(pass || '').trim();
  if (!pass) return false;

  const expected = PropertiesService.getScriptProperties().getProperty('ADD_PASS') || '';
  if (!expected) {
    // se não estiver configurada, trava por segurança
    return false;
  }
  return pass === expected;
}

/**
 * addEstacaoFromMapa(payload)
 * ------------------------------------------------------------------
 * Recebe payload do WebApp (index.html) e grava na aba Estacoes.
 * Foto (opcional) é salva no Drive e link público é gravado na planilha.
 *
 * Retorno: { ok:true, estacao:{...} } ou { ok:false, error:"..." }
 */
function addEstacaoFromMapa(payload) {
  try {
    payload = payload || {};
    const lat = Number(payload.lat);
    const lng = Number(payload.lng);

    if (!isFinite(lat) || !isFinite(lng)) return { ok:false, error:'Lat/Lng obrigatórios.' };
    if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return { ok:false, error:'Lat/Lng fora do range.' };

    const tipo = String(payload.tipo || '').toUpperCase().trim();
    const modalidade = String(payload.modalidade || '').toUpperCase().trim();
    const status = String(payload.status || '').toUpperCase().trim();
    const nomeConc = String(payload.nomeConcorrente || '').trim();
    const obs = String(payload.observacoes || '').trim();

    if (!tipo) return { ok:false, error:'Tipo obrigatório.' };
    if (!modalidade) return { ok:false, error:'Modalidade obrigatória.' };
    if (tipo === 'PUBLICA' && !status) return { ok:false, error:'Status obrigatório para pública.' };
    if (tipo === 'CONCORRENTE' && !nomeConc) return { ok:false, error:'Nome do concorrente obrigatório.' };
    if (obs.length > 500) return { ok:false, error:'Observações: máximo 500 caracteres.' };

    const geo = payload.geo || {};
    const cidade = String(geo.cidade || '').trim();
    const bairro = String(geo.bairro || '').trim();
    const endereco = String(geo.endereco || '').trim();
    const pais = String(geo.pais || '').trim();
    const estado = String(geo.estado || '').trim();
    const alcaldia = String(geo.alcaldia || '').trim();

    if (!cidade) return { ok:false, error:'Cidade obrigatória (geolocalização).'};
    if (!endereco) return { ok:false, error:'Endereço obrigatório (geolocalização).'};
    if (!bairro) {
      // bairro pode falhar em alguns geocoders; não bloqueia
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return { ok:false, error:'Aba Estacoes não encontrada.' };

    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

    function colIdx(name) {
      const i = headers.indexOf(name);
      return i === -1 ? -1 : i;
    }
    function setCell(rowArr, name, value) {
      const i = colIdx(name);
      if (i !== -1) rowArr[i] = value;
    }

    const tz = Session.getScriptTimeZone();
    const now = new Date();

    const rowKey = Utilities.getUuid();
    const codigo = 'MAPA-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random()*1000);

    const newRow = new Array(headers.length).fill('');

    // Identidade
    setCell(newRow, COL.RowKey, rowKey);
    setCell(newRow, COL.Codigo, codigo);
    setCell(newRow, COL.NomeEstacao, endereco);

    // Localização
    setCell(newRow, 'Latitude', lat);
    setCell(newRow, 'Longitude', lng);
    setCell(newRow, COL.Localizacao, lat + ',' + lng);
    setCell(newRow, COL.Cidade, cidade);
    setCell(newRow, COL.Bairro, bairro);
    setCell(newRow, COL.Endereco, endereco);

    // Classificação
    setCell(newRow, COL.TipoEstacao, tipo);
    setCell(newRow, 'StatusEstacao', (tipo === 'PUBLICA' ? status : ''));
    // Obs / Concorrente
     let obsFinal = obs;
   if (tipo === 'CONCORRENTE') {
     obsFinal = (obsFinal ? obsFinal + '\n' : '') + 'CONCORRENTE: ' + nomeConc;
   }
   setCell(newRow, COL.ObservacaoPrivado, obsFinal);
    setCell(newRow, 'Modalidade', modalidade || 'PATINETE');
    if (payload.dimensoes)
      setCell(newRow, 'Dimensoes da Estacao', payload.dimensoes);
    if (payload.larguraFaixa)
      setCell(newRow, 'Largura da Faixa Livre (m)',
        Number(payload.larguraFaixa) || payload.larguraFaixa);

   if (payload.dimensoes) setCell(newRow, 'Dimensões da Estação', payload.dimensoes);
   if (payload.larguraFaixa) setCell(newRow, 'Largura da Faixa Livre (m)', Number(payload.larguraFaixa) || payload.larguraFaixa);

    // Meta
    setCell(newRow, 'OrigemDado', 'MAPA_WEB');
    setCell(newRow, 'CriadoPor', 'MAPA_WEB');
    setCell(newRow, 'DataCriacao', now);

    // País / alcaldia (se existirem)
    setCell(newRow, 'Pais', pais || '');
    setCell(newRow, 'Estado', estado || '');
    setCell(newRow, 'Alcaldia', alcaldia || '');

    // Foto (opcional)
    let fotoUrl = '';
    try {
      const foto = payload.foto;
      if (foto && foto.base64 && foto.mime) {
        const bytes = Utilities.base64Decode(String(foto.base64));
        // limite defensivo (~3.5MB) para evitar estourar Apps Script
        if (bytes.length > 3.5 * 1024 * 1024) {
          return { ok:false, error:'Foto muito grande. Envie sem foto ou reduza.' };
        }
        const blob = Utilities.newBlob(bytes, String(foto.mime), String(foto.name || 'foto.jpg'));

        const root = DriveApp.getFolderById(FOTOS_FOLDER_ID);
        // organiza em Pais/Cidade/Bairro
        const pastaPais = ensureSubFolder_(root, cleanName_(pais || 'BR'));
        const pastaCidade = ensureSubFolder_(pastaPais, cleanName_(cidade || 'SEM_CIDADE'));
        const pastaFinal = ensureSubFolder_(pastaCidade, cleanName_(bairro || 'Outros'));

        const safeBase = cleanName_(endereco || codigo || 'ESTACAO');
        const ext = String(foto.name || '').match(/(\.[^.]+)$/);
        const fileName = safeBase + '_' + Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss') + (ext ? ext[1] : '.jpg');

        const file = pastaFinal.createFile(blob).setName(fileName);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        fotoUrl = file.getUrl();
        setCell(newRow, COL.Foto, fotoUrl);
      }
    } catch (e) {
      Logger.log('Erro salvando foto (addEstacaoFromMapa): ' + e);
      // foto é opcional
    }

    // grava
    sh.appendRow(newRow);

    // retorno para injeção no mapa (shape compatível com listarEstacoesParaVisualizacao)
    const estacao = {
      codigo: codigo,
      cidade: cidade,
      tipo: tipo,
      statusEstacao: (tipo === 'PUBLICA' ? status : ''),
      subprefeitura: '',

      lat: lat,
      lng: lng,
      endereco: endereco,
      bairro: bairro,
      localizacao: lat + ',' + lng,

      croqui: '',
      foto: fotoUrl || ''
    };

    logEvento_('ADD_MAPA', '', codigo, 'Estação adicionada via mapa: ' + endereco);

    return { ok:true, estacao: estacao };

  } catch (e) {
    Logger.log('addEstacaoFromMapa erro: ' + e);
    return { ok:false, error: String(e && e.message ? e.message : e) };
  }
}

function quemSouEu() {
  Logger.log(Session.getActiveUser().getEmail());
}

function diagnosticarPasta() {
  try {
    var pasta = DriveApp.getFolderById('1hc5-whvQdYNqHbmkzW96MMnOjm4nk964');
    Logger.log('Nome: ' + pasta.getName());
    Logger.log('Acesso: OK');
    
    // Tenta criar arquivo de teste
    var teste = pasta.createFile('TESTE_DELETE.txt', 'teste');
    Logger.log('Criar arquivo: OK — ' + teste.getId());
    teste.setTrashed(true);
    Logger.log('Deletar teste: OK');
  } catch(e) {
    Logger.log('ERRO: ' + e);
  }
}

function diagnosticarCoreResult() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var row = 2; // troque pela linha que estava tentando gerar
  
  try {
    var result = gerarCroqui_CORE(sh, row);
    Logger.log('CORE ok: ' + JSON.stringify(Object.keys(result)));
    Logger.log('imagens: ' + JSON.stringify(Object.keys(result.imagens || {})));
    Logger.log('sat bytes: '    + (result.imagens.sat    ? result.imagens.sat.getBytes().length    : 'NULL'));
    Logger.log('map bytes: '    + (result.imagens.map    ? result.imagens.map.getBytes().length    : 'NULL'));
    Logger.log('street bytes: ' + (result.imagens.street ? result.imagens.street.getBytes().length : 'NULL'));
  } catch(e) {
    Logger.log('ERRO no CORE: ' + e);
  }
}

function salvar(blob, sufixo, col) {
  if (!blob || !col) return;

  var novoBlob = Utilities.newBlob(
    blob.getBytes(),
    'image/png',
    'ESTACAO_' + row + '_' + sufixo + '.png'
  );

  var file = pastaImagens.createFile(novoBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet.getRange(row, col).setValue(file.getUrl());
}

function testeCompletoCroqui() {
  var sh  = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var row = 2; // troque pela linha que quer testar

  try {
    // Passo 1: CORE
    Logger.log('1. Rodando CORE...');
    var result = gerarCroqui_CORE(sh, row);
    Logger.log('1. CORE ok');

    // Passo 2: salvar imagens
    Logger.log('2. Salvando imagens no Drive...');
    salvarImagensNoDrive2_(sh, row, result.imagens);
    Logger.log('2. Imagens salvas ok');

    // Passo 3: template
    Logger.log('3. Resolvendo template...');
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var get = function(n){ var i=headers.indexOf(n); return i<0?'':sh.getRange(row,i+1).getValue(); };
    var resolved = resolverTemplatePorPaisETipo_('BR', 'PUBLICO', {
      cidade: get('Cidade'),
      bairro: get('Bairro')
    });
    Logger.log('3. Template ok: ' + resolved.templateId);
    Logger.log('3. Pasta destino: ' + resolved.pastaDestino.getName());

    // Passo 4: cópia do template
    Logger.log('4. Copiando template...');
    var tmp = DriveApp.getFileById(resolved.templateId)
      .makeCopy('TMP_TESTE_DELETE', resolved.pastaDestino);
    Logger.log('4. Cópia ok: ' + tmp.getId());
    tmp.setTrashed(true);
    Logger.log('4. Cleanup ok');

  } catch(e) {
    Logger.log('ERRO: ' + e);
    Logger.log('Stack: ' + e.stack);
  }
}

function testeSalvarBlob() {
  var sh  = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var result = gerarCroqui_CORE(sh, 2);
  var blob = result.imagens.sat;

  Logger.log('tipo original: ' + blob.getContentType());
  Logger.log('bytes: ' + blob.getBytes().length);

  // Criar blob novo a partir dos bytes
  var novoBlob = Utilities.newBlob(blob.getBytes(), 'image/png', 'TESTE_SAT.png');
  Logger.log('tipo novo: ' + novoBlob.getContentType());

  var pasta = DriveApp.getFolderById(IMAGENS_FOLDER_ID);
  var file = pasta.createFile(novoBlob);
  Logger.log('Arquivo criado: ' + file.getId());
  file.setTrashed(true);
  Logger.log('OK');
}

function verCodigoSalvo() {
  Logger.log('setContentType presente: ' + 
    (salvarImagensNoDrive2_.toString().indexOf('setContentType') !== -1 ? 'SIM' : 'NAO'));
}

function verLinhas531a540() {
  // Não faz nada — só para você copiar o código das linhas 528-538 do Código.gs
  // e colar aqui para eu ver o que está lá de verdade
}

function testeCompletoV2() {
  var sh  = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var row = 2;

  try {
    Logger.log('1. CORE...');
    var result = gerarCroqui_CORE(sh, row);
    Logger.log('1. OK');

    Logger.log('2. Convertendo blobs imediatamente...');
    var sat    = Utilities.newBlob(result.imagens.sat.getBytes(),    'image/png', 'SAT.png');
    var map    = Utilities.newBlob(result.imagens.map.getBytes(),    'image/png', 'MAP.png');
    var street = Utilities.newBlob(result.imagens.street.getBytes(), 'image/png', 'STREET.png');
    Logger.log('2. Blobs convertidos');

    Logger.log('3. Salvando no Drive...');
    var pasta = DriveApp.getFolderById(IMAGENS_FOLDER_ID);
    var f1 = pasta.createFile(sat);    Logger.log('SAT ok: ' + f1.getId());    f1.setTrashed(true);
    var f2 = pasta.createFile(map);    Logger.log('MAP ok: ' + f2.getId());    f2.setTrashed(true);
    var f3 = pasta.createFile(street); Logger.log('STREET ok: ' + f3.getId()); f3.setTrashed(true);
    Logger.log('3. OK');

  } catch(e) {
    Logger.log('ERRO: ' + e);
    Logger.log('Stack: ' + e.stack);
  }
}

function verPastaImagens() {
  Logger.log('IMAGENS_FOLDER_ID: ' + IMAGENS_FOLDER_ID);
  var pasta = DriveApp.getFolderById(IMAGENS_FOLDER_ID);
  Logger.log('Nome: ' + pasta.getName());
  Logger.log('ID: ' + pasta.getId());

  // Testa criar arquivo direto
  try {
    var blob = Utilities.newBlob('teste', 'image/png', 'teste.png');
    var f = pasta.createFile(blob);
    Logger.log('createFile OK: ' + f.getId());
    f.setTrashed(true);
  } catch(e) {
    Logger.log('createFile ERRO: ' + e);
  }
}

function testeCompletoCroqui() {
  var sh  = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var row = 2;

  try {
    Logger.log('1. Rodando CORE...');
    var result = gerarCroqui_CORE(sh, row);
    Logger.log('1. CORE ok');

    Logger.log('2. Convertendo blobs...');
    var imagens = {
      sat:    Utilities.newBlob(result.imagens.sat.getBytes(),    'image/png', 'sat.png'),
      map:    Utilities.newBlob(result.imagens.map.getBytes(),    'image/png', 'map.png'),
      street: Utilities.newBlob(result.imagens.street.getBytes(), 'image/png', 'street.png')
    };
    Logger.log('2. Blobs convertidos');

    Logger.log('3. Salvando imagens no Drive...');
    salvarImagensNoDrive2_(sh, row, imagens);
    Logger.log('3. Imagens salvas ok');

    Logger.log('4. Resolvendo template...');
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var get = function(n){ var i=headers.indexOf(n); return i<0?'':sh.getRange(row,i+1).getValue(); };
    var resolved = resolverTemplatePorPaisETipo_('BR', 'PUBLICO', {
      cidade: get('Cidade'),
      bairro: get('Bairro')
    });
    Logger.log('4. Template ok: ' + resolved.templateId);

    Logger.log('5. Copiando template...');
    var tmp = DriveApp.getFileById(resolved.templateId)
      .makeCopy('TMP_TESTE_DELETE', resolved.pastaDestino);
    Logger.log('5. Copia ok: ' + tmp.getId());
    tmp.setTrashed(true);
    Logger.log('TESTE COMPLETO OK');

  } catch(e) {
    Logger.log('ERRO: ' + e);
    Logger.log('Stack: ' + e.stack);
  }
}

function testeBypassSalvar() {
  var sh  = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  var row = 2;
  var pasta = DriveApp.getFolderById(IMAGENS_FOLDER_ID);

  try {
    Logger.log('1. CORE...');
    var result = gerarCroqui_CORE(sh, row);
    Logger.log('1. OK');

    Logger.log('2. Salvando SAT direto...');
    var sat = Utilities.newBlob(result.imagens.sat.getBytes(), 'image/png', 'SAT_TESTE.png');
    var f1 = pasta.createFile(sat);
    Logger.log('SAT ok: ' + f1.getId());
    f1.setTrashed(true);

    Logger.log('2. Salvando MAP direto...');
    var map = Utilities.newBlob(result.imagens.map.getBytes(), 'image/png', 'MAP_TESTE.png');
    var f2 = pasta.createFile(map);
    Logger.log('MAP ok: ' + f2.getId());
    f2.setTrashed(true);

    Logger.log('2. Salvando STREET direto...');
    var street = Utilities.newBlob(result.imagens.street.getBytes(), 'image/png', 'STREET_TESTE.png');
    var f3 = pasta.createFile(street);
    Logger.log('STREET ok: ' + f3.getId());
    f3.setTrashed(true);

    Logger.log('TUDO OK');

  } catch(e) {
    Logger.log('ERRO: ' + e);
    Logger.log('Stack: ' + e.stack);
  }
}

/******************* MAPAS — GERAÇÃO POR SELEÇÃO *******************/

function gerarImagemSateliteSelecao() {
  gerarImagensSelecao_('SAT');
}

function gerarImagemMapaSelecao() {
  gerarImagensSelecao_('MAP');
}

function gerarImagemStreetViewSelecao() {
  gerarImagensSelecao_('STREET');
}

/*
 * Compatibilidade com chamadas antigas que ainda possam existir
 */
function gerarImagemSatelite() {
  gerarImagemSateliteSelecao();
}

function gerarImagemMapa() {
  gerarImagemMapaSelecao();
}

function gerarImagemStreetView() {
  gerarImagemStreetViewSelecao();
}

function gerarImagensSelecao_(tipo) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) {
    throw new Error('Aba "' + SHEET_NAME + '" não encontrada.');
  }

  const sel = sh.getActiveRange();
  if (!sel) {
    SpreadsheetApp.getUi().alert('Selecione uma ou mais linhas.');
    return;
  }

  const start = Math.max(2, sel.getRow());
  const end = Math.min(sh.getLastRow(), start + sel.getNumRows() - 1);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const cCodigo = headers.indexOf(COL.Codigo) + 1;
  const cLoc = headers.indexOf(COL.Localizacao) + 1;
  const cSat = headers.indexOf(COL.ImgSat) + 1;
  const cMap = headers.indexOf(COL.ImgMapa) + 1;
  const cStreet = headers.indexOf(COL.ImgStreet) + 1;

  if (!cCodigo || !cLoc) {
    throw new Error('Colunas obrigatórias não encontradas para gerar imagens.');
  }

  const pastaImagens = DriveApp.getFolderById(IMAGENS_FOLDER_ID);

  let ok = 0;
  let erro = 0;
  const logs = [];

  for (let row = start; row <= end; row++) {
    try {
      const codigo = String(sh.getRange(row, cCodigo).getValue() || '').trim();
      const loc = String(sh.getRange(row, cLoc).getValue() || '').trim();

      if (!loc) {
        throw new Error('Localização vazia');
      }

      const coords = parseLatLng_CORE_(loc);
      let blob = null;
      let suffix = '';
      let colDestino = 0;

      if (tipo === 'SAT') {
        blob = fetchStaticMap_CORE(coords.lat, coords.lng, { maptype: 'satellite' });
        suffix = 'SAT';
        colDestino = cSat;
      } else if (tipo === 'MAP') {
        blob = fetchStaticMap_CORE(coords.lat, coords.lng, { maptype: 'roadmap' });
        suffix = 'MAP';
        colDestino = cMap;
      } else if (tipo === 'STREET') {
        blob = fetchStreetView_CORE(coords.lat, coords.lng, {});
        suffix = 'STREET';
        colDestino = cStreet;
      } else {
        throw new Error('Tipo de imagem inválido: ' + tipo);
      }

      if (!blob) {
        throw new Error('Blob de imagem vazio');
      }

      const nomeBase = cleanName_(codigo || ('ESTACAO_' + row));
      const nomeArquivo = nomeBase + '_' + suffix + '.png';

      const file = pastaImagens.createFile(blob.setName(nomeArquivo));
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}

      if (!colDestino) {
        throw new Error('Coluna de destino não encontrada para ' + tipo);
      }

      sh.getRange(row, colDestino).setValue(file.getUrl());

      ok++;
      Utilities.sleep(120);

    } catch (e) {
      erro++;
      logs.push('Linha ' + row + ': ' + String(e && e.message ? e.message : e));
      Logger.log('Erro ao gerar imagem [' + tipo + '] linha ' + row + ': ' + e);
    }
  }

  let msg = 'Processo concluído.\n\n';
  msg += 'Tipo: ' + tipo + '\n';
  msg += 'Sucesso: ' + ok + '\n';
  msg += 'Erros: ' + erro;

  if (logs.length) {
    msg += '\n\nPrimeiros erros:\n' + logs.slice(0, 10).join('\n');
  }

  SpreadsheetApp.getUi().alert(msg);
}

function getStreetViewBase64(lat, lng, heading) {
  var key = PropertiesService.getScriptProperties().getProperty('GMAPS_API_KEY');
  var url = 'https://maps.googleapis.com/maps/api/streetview?size=400x300'
    + '&location=' + lat + ',' + lng
    + '&fov=90&pitch=0&heading=' + (heading || 0)
    + '&key=' + key;
  
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var bytes = resp.getContent();
    return Utilities.base64Encode(bytes);
  } catch(e) {
    return null;
  }
}

function analisarCalcadaIA(params) {
  var lat = Number(params && params.lat || 0);
  var lng = Number(params && params.lng || 0);
  if (!lat || !lng) return { ok: false, error: 'Coords invalidas' };

  var geminiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!geminiKey) return { ok: false, error: 'GEMINI_API_KEY nao configurada' };

  var mapsKey = PropertiesService.getScriptProperties().getProperty('GMAPS_API_KEY');

  // 1. Buscar 2 angulos de Street View
  var headings = [0, 90];
  var parts = [{
  text:
    'Voce e um especialista em infraestrutura urbana para micromobilidade.' +
    ' Analise com rigor tecnico as imagens de Street View e determine se existe uma CALCADA REALMENTE ADEQUADA para instalar uma estacao de patinetes eletricos.' +

    '\n\nOBJETIVO:' +
    '\nResponder se o ponto mostrado possui area de calcada suficiente, segura e urbana para instalacao de estacao.' +

    '\n\nREGRA GERAL MAIS IMPORTANTE:' +
    '\nSo aprove se a calcada estiver claramente visivel e se houver evidencia visual suficiente de largura livre adequada.' +
    '\nSe houver qualquer duvida relevante, baixa visibilidade, enquadramento ruim, imagem distante, obstrucao ou impossibilidade de confirmar a calcada, responda aprovado:false.' +
    '\nNao aprove por inferencia.' +

    '\n\nCRITERIOS OBRIGATORIOS PARA APROVACAO (todos devem ser atendidos):' +
    '\n1. Deve existir CALCADA claramente visivel e separada da pista de veiculos.' +
    '\n2. A area analisada deve ser calcada urbana, e nao rua, acostamento, faixa de onibus, ciclovia, estacionamento, entrada de garagem, canteiro, gramado ou sarjeta.' +
    '\n3. A largura livre estimada da calcada deve ser de pelo menos 2.8 metros.' +
    '\n4. A faixa livre deve permitir a estacao + circulacao de pedestres sem bloqueio relevante.' +
    '\n5. Nao pode haver obstrucao dominante por poste, arvore, banca, lixeira grande, gradil, ponto de onibus apertado, mobiliario urbano ou estruturas fixas.' +
    '\n6. O local deve parecer area urbana real com infraestrutura de calcada utilizavel.' +

    '\n\nREGRAS DE REPROVACAO IMEDIATA (aprovado:false, score entre 0 e 5):' +
    '\n- A imagem mostra principalmente pista, asfalto ou faixa de veiculos sem calcada lateral claramente utilizavel.' +
    '\n- O ponto esta sobre meio-fio, sarjeta, canteiro central, area gramada, terra, lote vazio ou area sem urbanizacao adequada.' +
    '\n- A imagem mostra estacionamento, recuo de garagem ou area de embarque/desembarque sem calcada livre suficiente.' +
    '\n- A calcada parece ter menos de 2.0 metros de largura livre.' +
    '\n- A imagem esta noturna, borrada, distante, cortada, obstruida ou sem evidencia suficiente para confirmar a calcada.' +
    '\n- Existe ciclovia ou faixa de onibus, mas nao ha calcada livre adequada e independente.' +

    '\n\nREGRAS IMPORTANTES DE INTERPRETACAO:' +
    '\n- Considere apenas o que esta visualmente evidenciado nas imagens.' +
    '\n- Nao assuma continuidade da calcada fora do enquadramento.' +
    '\n- Nao confunda recuo viario, baia, acostamento ou piso asfaltado com calcada.' +
    '\n- Ponto de onibus so pode ser aprovado se houver calcada lateral ampla e claramente separada da pista.' +
    '\n- Se houver varias imagens, use o conjunto delas, mas seja conservador: se nenhuma comprovar adequadamente a calcada, reprovar.' +
    '\n- A aprovacao exige evidencia positiva, nao ausencia de problema.' +

    '\n\nESCALA DE SCORE:' +
    '\n0 a 5  = reprovado categoricamente, sem calcada adequada ou sem evidencia minima.' +
    '\n6 a 15 = reprovado, ha alguma calcada aparente mas insuficiente, inadequada ou muito duvidosa.' +
    '\n16 a 25 = reprovado, calcada existe mas nao ha seguranca para aprovar por largura, obstrucao ou contexto.' +
    '\n26 a 32 = aprovado com ressalvas, calcada adequada e visivel.' +
    '\n33 a 40 = aprovado com alta confianca, calcada ampla, urbana e claramente adequada.' +

    '\n\nINSTRUCOES DE SAIDA:' +
    '\nRetorne APENAS um JSON valido, sem markdown, sem texto extra, sem comentarios.' +
    '\nUse exatamente este formato:' +
    '\n{"aprovado":true/false,"larguraEstimada":"X metros ou indefinido","observacoes":"motivo objetivo em portugues","confianca":"alta/media/baixa","score":0}' +

    '\n\nREGRAS FINAIS DO JSON:' +
    '\n- "aprovado" deve ser true somente se TODOS os criterios obrigatorios forem atendidos visualmente.' +
    '\n- "larguraEstimada" deve ser conservadora. Se nao for possivel estimar com seguranca, use "indefinido".' +
    '\n- "observacoes" deve explicar de forma curta e objetiva o principal motivo da decisao.' +
    '\n- "confianca" deve refletir a clareza visual real da imagem.' +
    '\n- "score" deve ser inteiro entre 0 e 40 e coerente com aprovado.' +
    '\n- Se houver incerteza relevante, use aprovado:false.'
}];

  headings.forEach(function(h) {
    var url = 'https://maps.googleapis.com/maps/api/streetview?size=400x300'
      + '&location=' + lat + ',' + lng
      + '&fov=90&pitch=0&heading=' + h + '&key=' + mapsKey;
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        var bytes = resp.getContent();
        if (bytes.length > 5000) {
          parts.push({
            inline_data: {
              mime_type: 'image/jpeg',
              data: Utilities.base64Encode(bytes)
            }
          });
        }
      }
    } catch(e) {}
  });

  if (parts.length < 2) return { ok: false, error: 'Street View nao disponivel' };

  // 2. Chamar Gemini com as imagens
  try {
    var payload = {
      contents: [{ parts: parts }],
      generationConfig: { maxOutputTokens: 300, temperature: 0.1 }
    };
    var geminiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + geminiKey;
    var res = UrlFetchApp.fetch(geminiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      return { ok: false, error: 'Gemini HTTP ' + res.getResponseCode() };
    }

    var data = JSON.parse(res.getContentText());
    var text = data.candidates[0].content.parts[0].text;
    var resultado = JSON.parse(text.replace(/```json|```/g, '').trim());
    return { ok: true, resultado: resultado };

  } catch(e) {
    return { ok: false, error: String(e) };
  }
}
