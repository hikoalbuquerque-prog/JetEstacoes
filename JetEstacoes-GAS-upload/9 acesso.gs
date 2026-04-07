/**
 * acesso.gs -- ARQUIVO NOVO
 *
 * Fluxo de solicitacao de acesso para novos usuarios do PWA Campo.
 *
 * 1. Usuario tenta logar -> validateUser retorna ok:false, error:'NAO_AUTORIZADO'
 * 2. PWA exibe tela "Solicitar acesso"
 * 3. Usuario preenche nome + cidade e clica "Solicitar"
 * 4. solicitarAcesso() grava na aba SOLICITACOES e envia email ao admin
 * 5. Admin aprova/rejeita pelo menu no Sheets OU pela aba SOLICITACOES
 * 6. Ao aprovar, usuario e movido para USUARIOS com role CAMPO
 *
 * SETUP:
 *   - A aba SOLICITACOES e criada automaticamente na primeira solicitacao
 *   - Salvar email do admin nas Script Properties:
 *     Chave: ADMIN_EMAIL  Valor: seu@email.com
 *   - Adicionar no menu onOpen():
 *     .addItem('Aprovar solicitacoes de acesso', 'uiAprovarSolicitacoes')
 */

var SOLICITACOES_SHEET = 'SOLICITACOES';

/**
 * Solicitacao de acesso vinda do PWA.
 * Nao requer auth -- o usuario ainda nao tem acesso.
 *
 * @param {{ email:string, nome:string, cidade:string, pais:string }} params
 */
function solicitarAcesso(params) {
  try {
    params = params || {};
    var email  = String(params.email  || '').trim().toLowerCase();
    var nome   = String(params.nome   || '').trim();
    var cidade = String(params.cidade || '').trim();
    var pais   = String(params.pais   || 'BR').toUpperCase();

    if (!email) return { ok: false, error: 'Email obrigatorio.' };
    if (!nome)  return { ok: false, error: 'Nome obrigatorio.' };

    // Verificar se ja e usuario ativo
    var ja = buscarUsuario_(email);
    if (ja && ja.ativo) {
      return { ok: false, error: 'JA_AUTORIZADO' };
    }

    // Verificar se ja tem solicitacao pendente
    var sh = garantirAbaSolicitacoes_();
    var dados  = sh.getDataRange().getValues();
    var header = dados[0];
    var iEmail  = header.indexOf('Email');
    var iStatus = header.indexOf('Status');

    for (var i = 1; i < dados.length; i++) {
      var rowEmail  = String(dados[i][iEmail]  || '').trim().toLowerCase();
      var rowStatus = String(dados[i][iStatus] || '').toUpperCase();
      if (rowEmail === email && rowStatus === 'PENDENTE') {
        return { ok: true, status: 'JA_PENDENTE',
          msg: 'Voce ja tem uma solicitacao pendente. Aguarde a aprovacao do administrador.' };
      }
    }

    // Gravar solicitacao
    sh.appendRow([
      email, nome, cidade, pais, 'CAMPO', 'PENDENTE',
      new Date(), '', ''
    ]);

    // Notificar admin por email
    try { notificarAdmin_(email, nome, cidade); } catch(e) { /* nao bloqueia */ }

    return {
      ok: true,
      status: 'SOLICITADO',
      msg: 'Solicitacao enviada! O administrador sera notificado e liberara seu acesso em breve.'
    };

  } catch (e) {
    Logger.log('solicitarAcesso erro: ' + e);
    return { ok: false, error: String(e) };
  }
}

/**
 * Aprova uma solicitacao e move o usuario para USUARIOS.
 * Requer role ADMIN.
 *
 * @param {{ tokenId:string, email:string, role?:string }} params
 */
function aprovarSolicitacao(params) {
  try {
    params = params || {};
    var auth = validateUser(params.tokenId);
    if (!auth.ok) return auth;
    if (!temPermissao(auth.role, 'ADMIN')) return { ok: false, error: 'Acesso negado.' };

    var email = String(params.email || '').trim().toLowerCase();
    var role  = String(params.role  || 'CAMPO').toUpperCase();

    if (!email) return { ok: false, error: 'Email obrigatorio.' };

    var sh = garantirAbaSolicitacoes_();
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var iEmail  = header.indexOf('Email')  + 1;
    var iNome   = header.indexOf('Nome')   + 1;
    var iCidade = header.indexOf('Cidade') + 1;
    var iPais   = header.indexOf('Pais')   + 1;
    var iStatus = header.indexOf('Status') + 1;
    var iAprov  = header.indexOf('AprovadoPor') + 1;
    var iDataAp = header.indexOf('DataAprovacao') + 1;

    var lastRow = sh.getLastRow();
    for (var r = 2; r <= lastRow; r++) {
      var rowEmail = String(sh.getRange(r, iEmail).getValue() || '').trim().toLowerCase();
      if (rowEmail !== email) continue;

      var rowStatus = String(sh.getRange(r, iStatus).getValue() || '').toUpperCase();
      if (rowStatus !== 'PENDENTE') {
        return { ok: false, error: 'Solicitacao ja processada: ' + rowStatus };
      }

      // Atualizar status na aba SOLICITACOES
      sh.getRange(r, iStatus).setValue('APROVADO');
      if (iAprov  > 0) sh.getRange(r, iAprov).setValue(auth.email);
      if (iDataAp > 0) sh.getRange(r, iDataAp).setValue(new Date());

      // Ler dados para criar o usuario
      var nome   = sh.getRange(r, iNome).getValue()   || '';
      var cidade = sh.getRange(r, iCidade).getValue() || '';
      var pais   = sh.getRange(r, iPais).getValue()   || 'BR';

      // Criar na aba USUARIOS
      var resultado = salvarUsuario(auth.role, {
        email: email, nome: nome, role: role,
        cidade: cidade, pais: pais, ativo: true
      });

      if (!resultado.ok) return resultado;

      // Notificar usuario aprovado
      try { notificarUsuarioAprovado_(email, nome); } catch(e) { /* nao bloqueia */ }

      return { ok: true, email: email, role: role };
    }

    return { ok: false, error: 'Solicitacao nao encontrada para: ' + email };

  } catch (e) {
    Logger.log('aprovarSolicitacao erro: ' + e);
    return { ok: false, error: String(e) };
  }
}

/**
 * Rejeita uma solicitacao.
 * Requer role ADMIN.
 */
function rejeitarSolicitacao(params) {
  try {
    params = params || {};
    var auth = validateUser(params.tokenId);
    if (!auth.ok) return auth;
    if (!temPermissao(auth.role, 'ADMIN')) return { ok: false, error: 'Acesso negado.' };

    var email  = String(params.email  || '').trim().toLowerCase();
    var motivo = String(params.motivo || '').trim();

    var sh     = garantirAbaSolicitacoes_();
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var iEmail  = header.indexOf('Email')  + 1;
    var iStatus = header.indexOf('Status') + 1;
    var iAprov  = header.indexOf('AprovadoPor') + 1;
    var iDataAp = header.indexOf('DataAprovacao') + 1;

    var lastRow = sh.getLastRow();
    for (var r = 2; r <= lastRow; r++) {
      var rowEmail = String(sh.getRange(r, iEmail).getValue() || '').trim().toLowerCase();
      if (rowEmail !== email) continue;

      sh.getRange(r, iStatus).setValue('REJEITADO');
      if (iAprov  > 0) sh.getRange(r, iAprov).setValue(auth.email);
      if (iDataAp > 0) sh.getRange(r, iDataAp).setValue(new Date());

      return { ok: true, email: email };
    }
    return { ok: false, error: 'Solicitacao nao encontrada.' };

  } catch (e) {
    Logger.log('rejeitarSolicitacao erro: ' + e);
    return { ok: false, error: String(e) };
  }
}

/**
 * Lista solicitacoes pendentes (ADMIN/SUPERVISOR).
 */
function listarSolicitacoesPendentes(params) {
  try {
    params = params || {};
    var auth = validateUser(params.tokenId);
    if (!auth.ok) return auth;
    if (!temPermissao(auth.role, 'SUPERVISOR')) return { ok: false, error: 'Acesso negado.' };

    var sh = garantirAbaSolicitacoes_();
    var dados  = sh.getDataRange().getValues();
    var header = dados.shift();
    var idx    = function(n) { return header.indexOf(n); };

    var pendentes = dados
      .filter(function(r) { return String(r[idx('Status')]||'').toUpperCase() === 'PENDENTE'; })
      .map(function(r) {
        return {
          email:      r[idx('Email')]      || '',
          nome:       r[idx('Nome')]       || '',
          cidade:     r[idx('Cidade')]     || '',
          pais:       r[idx('Pais')]       || '',
          roleDesej:  r[idx('RoleDesejado')] || 'CAMPO',
          status:     r[idx('Status')]     || '',
          dataSolic:  r[idx('DataSolicita')] || ''
        };
      });

    return { ok: true, pendentes: pendentes, total: pendentes.length };

  } catch (e) {
    Logger.log('listarSolicitacoesPendentes erro: ' + e);
    return { ok: false, error: String(e) };
  }
}

// --- UI (menu Sheets) -------------------------------------------

/**
 * Abre sidebar para aprovar/rejeitar solicitacoes pendentes.
 * Chamar pelo menu: Configuracoes > Aprovar solicitacoes de acesso
 */
function uiAprovarSolicitacoes() {
  var sh = garantirAbaSolicitacoes_();
  var dados  = sh.getDataRange().getValues();
  var header = dados.shift();
  var idx    = function(n) { return header.indexOf(n); };

  var pendentes = dados.filter(function(r) {
    return String(r[idx('Status')]||'').toUpperCase() === 'PENDENTE';
  });

  if (!pendentes.length) {
    SpreadsheetApp.getUi().alert('Nenhuma solicitacao pendente.');
    return;
  }

  var ui  = SpreadsheetApp.getUi();
  var msg = 'Solicitacoes pendentes (' + pendentes.length + '):\n\n';

  pendentes.forEach(function(r, i) {
    msg += (i+1) + '. ' + r[idx('Nome')] + ' <' + r[idx('Email')] + '>'
      + ' - ' + (r[idx('Cidade')]||'sem cidade') + '\n';
  });

  msg += '\nAbra a aba SOLICITACOES para aprovar ou rejeitar manualmente,\n'
       + 'ou use o Painel do Gestor (aba Usuarios > Solicitacoes).';

  ui.alert('Solicitacoes de Acesso', msg, ui.ButtonSet.OK);
}

// --- Privadas ---------------------------------------------------

function garantirAbaSolicitacoes_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SOLICITACOES_SHEET);
  if (sh) return sh;

  sh = ss.insertSheet(SOLICITACOES_SHEET);
  var headers = [
    'Email','Nome','Cidade','Pais','RoleDesejado',
    'Status','DataSolicita','AprovadoPor','DataAprovacao'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#1E3A5F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  return sh;
}

function notificarAdmin_(email, nome, cidade) {
  var adminEmail = PropertiesService.getScriptProperties()
    .getProperty('ADMIN_EMAIL');
  if (!adminEmail) return;

  var assunto = '[App Estacoes] Nova solicitacao de acesso: ' + nome;
  var corpo   = 'Nova solicitacao de acesso ao App Estacoes Campo:\n\n'
    + 'Nome:   ' + nome   + '\n'
    + 'Email:  ' + email  + '\n'
    + 'Cidade: ' + cidade + '\n\n'
    + 'Para aprovar, acesse a planilha > aba SOLICITACOES\n'
    + 'ou use o menu Configuracoes > Aprovar solicitacoes de acesso.';

  GmailApp.sendEmail(adminEmail, assunto, corpo);
}

function notificarUsuarioAprovado_(email, nome) {
  var corpo = 'Ola ' + nome + ',\n\n'
    + 'Seu acesso ao App Estacoes Campo foi aprovado!\n\n'
    + 'Acesse pelo link enviado pela sua equipe e faca login com esta conta Google.\n\n'
    + 'Bom mapeamento!';
  GmailApp.sendEmail(email, '[App Estacoes] Acesso aprovado!', corpo);
}