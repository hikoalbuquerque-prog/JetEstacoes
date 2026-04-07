/**
 * utils_auth.gs — ARQUIVO NOVO
 * ─────────────────────────────────────────────────────────────────
 * Autenticacao e controle de acesso para o PWA de campo.
 *
 * SETUP (fazer uma vez):
 *   1. No menu Configuracoes > "Criar aba USUARIOS (PWA auth)"
 *   2. No menu Configuracoes > Salvar Google Maps API Key
 *   3. Salvar o OAUTH_CLIENT_ID nas Script Properties:
 *      PropertiesService.getScriptProperties().setProperty('OAUTH_CLIENT_ID','...')
 *
 * ROLES: CAMPO | SUPERVISOR | OPERACOES | ADMIN | VIEWER
 * Hierarquia: ADMIN > OPERACOES > SUPERVISOR > CAMPO > VIEWER
 *
 * ATENÇÃO: validateUser() faz requisicao externa ao Google (tokeninfo).
 * Em producao, considere mover para verificacao local com biblioteca
 * google-auth-library se a latencia for critica.
 */

var AUTH_SHEET_NAME = 'USUARIOS';
var ROLES_VALIDOS   = ['CAMPO', 'SUPERVISOR', 'OPERACOES', 'ADMIN', 'VIEWER'];

/**
 * Valida token Google OAuth e verifica permissao na aba USUARIOS.
 * Endpoint principal chamado pelo PWA apos login.
 *
 * @param {string} tokenId  ID token JWT do Google Identity Services
 * @returns {{ ok:boolean, role?:string, nome?:string, cidade?:string, pais?:string, error?:string }}
 */
function validateUser(tokenId) {
  try {
    if (!tokenId) return { ok: false, error: 'Token ausente.' };

    var email = verificarTokenGoogle_(tokenId);
    if (!email) return { ok: false, error: 'Token invalido ou expirado.' };

    var usuario = buscarUsuario_(email);
    if (!usuario)    return { ok: false, error: 'Usuario nao autorizado.' };
    if (!usuario.ativo) return { ok: false, error: 'Usuario inativo. Contate o administrador.' };

    try { registrarUltimoAcesso_(email); } catch(e) { /* nao bloqueia o login */ }

    return {
      ok:     true,
      email:  email,
      nome:   usuario.nome,
      role:   usuario.role,
      cidade: usuario.cidade,
      pais:   usuario.pais
    };

  } catch (e) {
    Logger.log('validateUser erro: ' + e);
    return { ok: false, error: 'Erro interno de autenticacao.' };
  }
}

/**
 * Verifica se o usuario tem permissao para uma acao.
 *
 * @param {string} roleUsuario
 * @param {string} roleMinimo   role minimo necessario
 * @returns {boolean}
 */
function temPermissao(roleUsuario, roleMinimo) {
  var hierarquia = ['VIEWER', 'CAMPO', 'SUPERVISOR', 'OPERACOES', 'ADMIN'];
  var nU = hierarquia.indexOf(String(roleUsuario || '').toUpperCase());
  var nM = hierarquia.indexOf(String(roleMinimo  || '').toUpperCase());
  if (nM < 0 || nU < 0) return false;
  return nU >= nM;
}

/**
 * Lista usuarios (somente ADMIN).
 * @param {string} roleRequisitante
 * @returns {Array|{ok:false, error:string}}
 */
function listarUsuarios(roleRequisitante) {
  if (!temPermissao(roleRequisitante, 'ADMIN')) return { ok: false, error: 'Acesso negado.' };

  var sh = getAbaUsuarios_();
  if (!sh) return { ok: false, error: 'Aba USUARIOS nao encontrada.' };

  var dados  = sh.getDataRange().getValues();
  var header = dados.shift();
  var idx    = function(n) { return header.indexOf(n); };

  return dados.map(function(r) {
    return {
      email:        r[idx('Email')]        || '',
      nome:         r[idx('Nome')]         || '',
      role:         r[idx('Role')]         || '',
      cidade:       r[idx('Cidade')]       || '',
      pais:         r[idx('Pais')]         || '',
      ativo:        r[idx('Ativo')] === true || String(r[idx('Ativo')]).toUpperCase() === 'TRUE',
      dataCriacao:  r[idx('DataCriacao')]  || '',
      ultimoAcesso: r[idx('UltimoAcesso')]|| ''
    };
  }).filter(function(u) { return u.email; });
}

/**
 * Adiciona ou atualiza usuario (somente ADMIN).
 * @param {string} roleRequisitante
 * @param {{ email, nome, role, cidade, pais, ativo }} dados
 */
function salvarUsuario(roleRequisitante, dados) {
  if (!temPermissao(roleRequisitante, 'ADMIN')) return { ok: false, error: 'Acesso negado.' };

  dados      = dados || {};
  var email  = String(dados.email || '').trim().toLowerCase();
  if (!email) return { ok: false, error: 'Email obrigatorio.' };

  var role = String(dados.role || '').toUpperCase();
  if (ROLES_VALIDOS.indexOf(role) === -1) return { ok: false, error: 'Role invalido: ' + role };

  var sh = getAbaUsuarios_();
  if (!sh) return { ok: false, error: 'Aba USUARIOS nao encontrada.' };

  var header   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var iEmail   = header.indexOf('Email') + 1;
  var lastRow  = sh.getLastRow();

  for (var r = 2; r <= lastRow; r++) {
    var cel = String(sh.getRange(r, iEmail).getValue() || '').trim().toLowerCase();
    if (cel === email) {
      atualizarLinha_(sh, r, header, dados);
      return { ok: true, acao: 'atualizado' };
    }
  }

  var novaLinha = header.map(function(col) {
    switch (col) {
      case 'Email':        return email;
      case 'Nome':         return dados.nome   || '';
      case 'Role':         return role;
      case 'Cidade':       return dados.cidade || '';
      case 'Pais':         return dados.pais   || 'BR';
      case 'Ativo':        return dados.ativo  !== false;
      case 'DataCriacao':  return new Date();
      case 'UltimoAcesso': return '';
      default:             return '';
    }
  });

  sh.appendRow(novaLinha);
  return { ok: true, acao: 'criado' };
}

// ─────────── FUNCOES PRIVADAS ───────────────────────────

function verificarTokenGoogle_(tokenId) {
  try {
    var url  = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(tokenId);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;

    var payload  = JSON.parse(resp.getContentText());
    var clientId = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID') || '';

    if (clientId && payload.aud !== clientId) {
      Logger.log('Token audience invalido: ' + payload.aud);
      return null;
    }
    return payload.email || null;
  } catch (e) {
    Logger.log('verificarTokenGoogle_ erro: ' + e);
    return null;
  }
}

function buscarUsuario_(email) {
  var sh = getAbaUsuarios_();
  if (!sh) return null;

  var dados    = sh.getDataRange().getValues();
  var header   = dados.shift();
  var idx      = function(n) { return header.indexOf(n); };
  var emailNorm = String(email || '').trim().toLowerCase();

  for (var i = 0; i < dados.length; i++) {
    var r         = dados[i];
    var rowEmail  = String(r[idx('Email')] || '').trim().toLowerCase();
    if (rowEmail !== emailNorm) continue;

    return {
      email:  rowEmail,
      nome:   r[idx('Nome')]   || '',
      role:   String(r[idx('Role')] || '').toUpperCase(),
      cidade: r[idx('Cidade')] || '',
      pais:   r[idx('Pais')]   || 'BR',
      ativo:  r[idx('Ativo')]  === true || String(r[idx('Ativo')]).toUpperCase() === 'TRUE'
    };
  }
  return null;
}

function registrarUltimoAcesso_(email) {
  var sh = getAbaUsuarios_();
  if (!sh) return;

  var header  = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var iEmail  = header.indexOf('Email')        + 1;
  var iAcesso = header.indexOf('UltimoAcesso') + 1;
  if (!iEmail || !iAcesso) return;

  var emailNorm = String(email || '').trim().toLowerCase();
  var lastRow   = sh.getLastRow();

  for (var r = 2; r <= lastRow; r++) {
    var cel = String(sh.getRange(r, iEmail).getValue() || '').trim().toLowerCase();
    if (cel === emailNorm) {
      sh.getRange(r, iAcesso).setValue(new Date());
      return;
    }
  }
}

function atualizarLinha_(sh, row, header, dados) {
  var idx  = function(n) { return header.indexOf(n); };
  var sets = {
    'Nome':   dados.nome,
    'Role':   String(dados.role || '').toUpperCase(),
    'Cidade': dados.cidade,
    'Pais':   dados.pais,
    'Ativo':  dados.ativo !== false
  };

  Object.keys(sets).forEach(function(col) {
    var i = idx(col);
    if (i >= 0 && sets[col] !== undefined && sets[col] !== null) {
      sh.getRange(row, i + 1).setValue(sets[col]);
    }
  });
}

function getAbaUsuarios_() {
  return SpreadsheetApp.getActive().getSheetByName(AUTH_SHEET_NAME);
}

/**
 * Cria aba USUARIOS com headers e linha de admin.
 * Chamar pelo menu Configuracoes > "Criar aba USUARIOS".
 */
function criarAbaUsuarios() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(AUTH_SHEET_NAME);
  if (sh) {
    SpreadsheetApp.getUi().alert('Aba USUARIOS ja existe.');
    return;
  }

  sh = ss.insertSheet(AUTH_SHEET_NAME);
  var headers = ['Email','Nome','Role','Cidade','Pais','Ativo','DataCriacao','UltimoAcesso'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#1E3A5F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  var adminEmail = Session.getEffectiveUser().getEmail();
  sh.appendRow([adminEmail, 'Administrador', 'ADMIN', '', 'BR', true, new Date(), '']);

  SpreadsheetApp.getUi().alert(
    'Aba USUARIOS criada!\n\n' +
    'Admin: ' + adminEmail + '\n\n' +
    'Adicione os agentes de campo com role CAMPO.'
  );
}

/**
 * Testa busca de usuario pelo email — rodar no Editor para debug.
 */
function debugBuscarUsuarioPorEmail() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt('Debug Auth', 'Email do usuario:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var u = buscarUsuario_(resp.getResponseText().trim());
  ui.alert(u ? JSON.stringify(u, null, 2) : 'Usuario nao encontrado.');
}