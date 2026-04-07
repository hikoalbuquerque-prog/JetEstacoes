/**
 * utils_math.gs
 * ─────────────────────────────────────────────────────────────────
 * ARQUIVO NOVO — criar no Editor do Apps Script.
 *
 * FONTE ÚNICA para funções matemáticas e de texto compartilhadas.
 * Elimina 3 cópias de distanciaMetros() e 4+ cópias de normalizar().
 *
 * Após criar este arquivo:
 *   - Apagar distanciaMetros() de jet_cross.gs        (linha 7)
 *   - Apagar distanciaMetros() de cruzarJetComEstacoes.gs (linha 42)
 *   - Apagar normalizar()       de cruzarJetComEstacoes.gs (linha 29)
 *   - Apagar parseNumeroSeguro() de jet_cross.gs       (linha 1)
 *   - forms.gs já será substituído pelo arquivo 3_forms_gs_corrigido.gs
 */

/**
 * Distância em metros entre dois pontos (Haversine).
 *
 * @param {number} lat1
 * @param {number} lng1
 * @param {number} lat2
 * @param {number} lng2
 * @returns {number} metros
 */
function distanciaMetros(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var toRad = function(x) { return x * Math.PI / 180; };
  var dLat = toRad(lat2 - lat1);
  var dLng = toRad(lng2 - lng1);
  var a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2)   * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

/**
 * Normaliza texto: remove acentos, upper-case, trim.
 * Canônico para todo o projeto.
 *
 * @param {*} txt
 * @returns {string}
 */
function normalizarTexto(txt) {
  return (txt || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

/**
 * Alias de compatibilidade — não remover.
 * jet_cross.gs, cruzarJetComEstacoes.gs e Código.gs chamam normalizar().
 */
function normalizar(txt) {
  return normalizarTexto(txt);
}

/**
 * Parse seguro de número (aceita vírgula decimal).
 * Substitui parseNumeroSeguro() em jet_cross.gs e parseNum() em cruzarJetComEstacoes.gs.
 *
 * @param {*} v
 * @returns {number}
 */
function parseNumSafe(v) {
  if (v === null || v === undefined) return NaN;
  return Number(String(v).trim().replace(',', '.'));
}