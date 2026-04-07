/*******************************************************
 * POLIGONOS.GS — ARQUITETURA FINAL (CACHE + MENU)
 * Mapa + Multi-fase + Estratégico
 *
 * Frontend V2.1+ chama:
 *   getPoligonosCidade(cidade)
 *
 * Compat:
 *   listarPoligonosPorCidade(cidade) -> alias
 *******************************************************/

const POLIGONOS_SHEET = 'Limites_Mapeamento';

// Cache “global versionado” (permite limpar tudo sem listar chaves)
const POLY_CACHE_VER_PROP = 'POLY_CACHE_VER';
const POLY_CACHE_TTL_SEC = 6 * 60 * 60; // 6h

// lazy state
let polyLazyTimer = null;
let polyLazyLastCity = null;
let polyLazyLoadedCity = null;


/* =====================================================
   MENU (onOpen chama addMenuPoligonos(ui))
===================================================== */
function addMenuPoligonos(ui) {
  ui = ui || SpreadsheetApp.getUi();

  ui.createMenu('Polígonos')
    .addItem('Limpar cache polígonos (cidade...)', 'uiLimparCachePoligonosCidade')
    .addItem('Limpar cache polígonos (tudo)', 'uiLimparCachePoligonosTudo')
    .addSeparator()
    .addItem('Debug: polígonos por cidade', 'uiDebugPoligonosPorCidade')
    .addToUi();
}

/* =====================================================
   ENDPOINT PRINCIPAL — FRONTEND
===================================================== */
function getPoligonosCidade(cidade) {
  if (!cidade) return [];

  const cache = CacheService.getScriptCache();
  const cacheKey = polyCacheKey_(cidade);

  // 1) cache
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {}
  }

  // 2) build
  const polys = listarPoligonosPorCidade_RAW_(cidade);

  // 3) cache put
  try {
    cache.put(cacheKey, JSON.stringify(polys), POLY_CACHE_TTL_SEC);
  } catch (e) {}

  return polys;
}

/* =====================================================
   COMPAT (alias do endpoint antigo)
===================================================== */
function listarPoligonosPorCidade(cidade) {
  return getPoligonosCidade(cidade);
}

/* =====================================================
   CACHE HELPERS
===================================================== */
function getPolyCacheVersion_() {
  return PropertiesService.getScriptProperties().getProperty(POLY_CACHE_VER_PROP) || 'v1';
}

function polyCacheKey_(cidade) {
  return 'POLY:' + getPolyCacheVersion_() + ':' + cidade;
}

function limparCachePoligonosCidade(cidade) {
  if (!cidade) return;
  CacheService.getScriptCache().remove(polyCacheKey_(cidade));
}

function limparCachePoligonosTudo_() {
  const v = 'v' + Date.now();
  PropertiesService.getScriptProperties().setProperty(POLY_CACHE_VER_PROP, v);
  return v;
}

/* =====================================================
   UI ACTIONS (menu)
===================================================== */
function uiLimparCachePoligonosCidade() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Limpar cache de polígonos — Cidade',
    'Digite a cidade EXATAMENTE como aparece no filtro do mapa (ex: São Paulo / Ciudad de México):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const cidade = (resp.getResponseText() || '').trim();
  if (!cidade) {
    ui.alert('Cidade vazia. Cancelado.');
    return;
  }

  limparCachePoligonosCidade(cidade);
  ui.alert('Cache de polígonos limpo para: ' + cidade);
}

function uiLimparCachePoligonosTudo() {
  const ui = SpreadsheetApp.getUi();
  const v = limparCachePoligonosTudo_();
  ui.alert('Cache global de polígonos resetado.\nVersão: ' + v);
}

function uiDebugPoligonosPorCidade() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(POLIGONOS_SHEET);
  if (!sh) return ui.alert('Aba "' + POLIGONOS_SHEET + '" não encontrada.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return ui.alert('Aba vazia.');

  const header = data[0].map(h => (h || '').toString().trim());
  const iCidade = header.indexOf('Cidade');
  const iPol = header.indexOf('Poligono');
  if (iCidade === -1 || iPol === -1) return ui.alert('Headers "Cidade" e/ou "Poligono" não encontrados.');

  const mapCount = {};
  for (let i = 1; i < data.length; i++) {
    const c = (data[i][iCidade] || '').toString().trim();
    const p = (data[i][iPol] || '').toString().trim();
    if (!c || !p) continue;
    mapCount[c] = (mapCount[c] || 0) + 1;
  }

  const linhas = Object.keys(mapCount).sort().map(c => `${c}: ${mapCount[c]}`).join('\n');
  ui.alert('Polígonos por cidade:\n\n' + (linhas || '(nenhum)'));
}

/* =====================================================
   LEITURA RAW DA PLANILHA (normaliza payload p/ frontend)
===================================================== */
function listarPoligonosPorCidade_RAW_(cidade) {
  if (!cidade) return [];
 
  const sh = SpreadsheetApp.getActive().getSheetByName(POLIGONOS_SHEET);
  if (!sh) return [];
 
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
 
  const header = data.shift().map(h => (h || '').toString().trim());
  const idx = nome => header.indexOf(nome);
 
  const iCidade     = idx('Cidade');
  const iGrupo      = idx('Grupo');
  const iFase       = idx('Fase');
  const iNome       = idx('Nome Área');
  const iTipo       = idx('Tipo');
  const iPrioridade = idx('Prioridade');
  const iAtivo      = idx('Ativo');
  const iCor        = idx('Cor');
  const iPol        = idx('Poligono');
 
  if (iCidade === -1 || iPol === -1) return [];
 
  const out = [];
 
  // i = índice no array data[] (0-based, após shift do header)
  // rowId = i + 2 (linha 1 é header, dados começam na linha 2)
  data.forEach((r, i) => {
    const c = String(r[iCidade] || '').trim();
    if (c !== cidade) return;
 
    const polStr = r[iPol];
    if (!polStr) return;
 
    const ativo  = (iAtivo === -1) ? true : normalizeBool_(r[iAtivo], true);
    const pontos = parsePoligono_(polStr);
    if (!pontos.length) return;
 
    out.push({
      rowId:      i + 2,                                           // ← NOVO
      id:         Utilities.getUuid(),
      cidade:     c,
      grupo:      iGrupo      !== -1 ? (r[iGrupo]      || 'Geral')   : 'Geral',
      fase:       iFase       !== -1 ? (r[iFase]       || 'Fase 1')  : 'Fase 1',
      nome:       iNome       !== -1 ? (r[iNome]       || '')         : '',
      tipo:       iTipo       !== -1 ? (r[iTipo]       || '')         : '',
      prioridade: iPrioridade !== -1 ? (Number(r[iPrioridade]) || 1)  : 1,
      ativo:      ativo,
      cor:        (iCor !== -1 && r[iCor]) ? r[iCor] : '#2563eb',
      poligono:   pontos
    });
  });
 
  out.sort((a, b) => (a.prioridade || 99) - (b.prioridade || 99));
  return out;
}


/* =====================================================
   PARSE STRING → ARRAY DE COORDENADAS
   Suporta "|" e ";"
===================================================== */
function parsePoligono_(str) {
  if (!str) return [];
 
  var s = String(str).trim();
 
  // Normalizar separadores alternativos:
  // 1. underscore entre coordenadas: -23.xxx_-23.yyy → -23.xxx|-23.yyy
  s = s.replace(/(\d)_(-?\d)/g, '$1|$2');
  // 2. newline como separador (cópia do Maps)
  s = s.replace(/\r?\n/g, '|');
  // 3. ponto-e-vírgula → pipe
  s = s.replace(/;/g, '|');
 
  // Remover pipes duplicados
  s = s.replace(/\|+/g, '|');
 
  return s
    .split('|')
    .map(function(p) { return p.trim(); })
    .filter(Boolean)
    .map(function(p) {
      var parts = p.split(',').map(function(x) { return x.trim(); });
      return { lat: Number(parts[0]), lng: Number(parts[1]) };
    })
    .filter(function(p) {
      return !isNaN(p.lat) &&
             !isNaN(p.lng) &&
             Math.abs(p.lat) <= 90 &&
             Math.abs(p.lng) <= 180;
    });
}

function normalizeBool_(v, defaultValue) {
  if (typeof v === 'boolean') return v;
  const s = (v || '').toString().trim().toLowerCase();
  if (s === 'true' || s === '1' || s === 'sim' || s === 'yes') return true;
  if (s === 'false' || s === '0' || s === 'não' || s === 'nao' || s === 'no') return false;
  return defaultValue;
}

/* =====================================================
   AGRUPAR POLÍGONOS POR FASE (mantido)
===================================================== */
function listarPoligonosPorFase(cidade) {
  const lista = getPoligonosCidade(cidade);
  const fases = {};

  lista.forEach(p => {
    if (!fases[p.fase]) fases[p.fase] = [];
    fases[p.fase].push(p);
  });

  return fases;
}

/* =====================================================
   CONTAR ESTAÇÕES DENTRO DE UM POLÍGONO (mantido)
   OBS: isso é útil para relatórios server-side, mas o front já faz via geometry.
===================================================== */
function contarEstacoesNoPoligono(poligono, estacoes) {
  if (!poligono || !estacoes) return 0;

  return estacoes.filter(e =>
    pontoDentroPoligono_(e.lat, e.lng, poligono)
  ).length;
}

/* =====================================================
   POINT IN POLYGON (Ray Casting) (mantido)
===================================================== */
function pontoDentroPoligono_(lat, lng, vertices) {
  let inside = false;

  for (let i = 0, j = vertices.length - 1; i < vertices.length; j = i++) {
    const xi = vertices[i].lat, yi = vertices[i].lng;
    const xj = vertices[j].lat, yj = vertices[j].lng;

    const intersect =
      ((yi > lng) !== (yj > lng)) &&
      (lat < (xj - xi) * (lng - yi) / (yj - yi) + xi);

    if (intersect) inside = !inside;
  }

  return inside;
}



/**
 * onEdit(e) — limpa cache de polígonos ao editar a aba Limites_Mapeamento
 *
 * Estratégia:
 * - Se editou a aba Limites_Mapeamento:
 *   - tenta descobrir a(s) cidade(s) afetadas
 *   - limpa cache só dessas cidades
 * - Se não der pra determinar (colagem grande / edição massiva):
 *   - bump de versão (limpa tudo)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh) return;

    if (sh.getName() !== POLIGONOS_SHEET) return;

    // cabeçalho
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(h => (h || '').toString().trim());

    const iCidade = header.indexOf('Cidade') + 1; // 1-based
    if (iCidade <= 0) {
      // sem coluna Cidade => melhor reset global
      limparCachePoligonosTudo_();
      return;
    }

    const r = e.range;
    const rowStart = r.getRow();
    const rowEnd = r.getLastRow();

    // se mexeu no header, reset global
    if (rowStart === 1) {
      limparCachePoligonosTudo_();
      return;
    }

    // Edição massiva? (cola grande / várias linhas)
    const rowsTouched = rowEnd - rowStart + 1;
    const colsTouched = r.getLastColumn() - r.getColumn() + 1;

    // threshold conservador (ajuste se quiser)
    if (rowsTouched > 50 || colsTouched > 10) {
      limparCachePoligonosTudo_();
      return;
    }

    // Coletar cidades afetadas (das linhas alteradas)
    const cidades = new Set();
    const cidadeVals = sh.getRange(rowStart, iCidade, rowsTouched, 1).getValues();

    for (let i = 0; i < cidadeVals.length; i++) {
      const c = (cidadeVals[i][0] || '').toString().trim();
      if (c) cidades.add(c);
    }

    // Se a edição foi na própria coluna Cidade, também considerar valor antigo e novo
    // (Apps Script não fornece valor antigo facilmente; então limpamos as cidades atuais + fallback)
    if (cidades.size === 0) {
      // não conseguiu identificar => reset global seguro
      limparCachePoligonosTudo_();
      return;
    }

    // Limpa cache apenas das cidades afetadas
    cidades.forEach(cidade => limparCachePoligonosCidade(cidade));

  } catch (err) {
    // fallback seguro
    try { limparCachePoligonosTudo_(); } catch (e2) {}
    console.error('onEdit polígonos erro:', err);
  }
}


/*******************************************************
 * COMPAT EXTRA — manter nomes antigos do Código.gs
 *******************************************************/

// alguns lugares podem chamar sem underscore
function limparCachePoligonosTudo() {
  return limparCachePoligonosTudo_();
}

// compat opcional caso algum script antigo chame com underscore
function limparCachePoligonosCidade_(cidade) {
  return limparCachePoligonosCidade(cidade);
}

