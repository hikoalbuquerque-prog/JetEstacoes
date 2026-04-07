/**
 * login_senha_gs.gs
 * Adicionar em qualquer .gs do projeto (ex: utils_auth.gs)
 *
 * Login por e-mail + senha para substituir o OAuth Google.
 * A senha fica na coluna SENHA da aba USUARIOS do Sheets.
 *
 * SETUP: Adicionar coluna "Senha" na aba USUARIOS
 * (pode ser a ultima coluna, ex: col I)
 */

// Ajuste se a sua aba tiver colunas em ordem diferente
var LOGIN_COLS = {
  EMAIL:     1,   // col A
  NOME:      2,   // col B
  ROLE:      3,   // col C
  CIDADE:    4,   // col D
  PAIS:      5,   // col E
  ATIVO:     6,   // col F
  SENHA:     9    // col I -- ADICIONAR esta coluna na aba USUARIOS
};

/**
 * Login para o PWA Campo (chamado via google.script.run).
 * Retorna { ok, email, nome, role, cidade, pais } ou { ok:false, error }
 */
function loginCampo(email, senha) {
  return _loginUsuario_(email, senha, ['CAMPO','SUPERVISOR','OPERACOES','ADMIN','VIEWER']);
}

/**
 * Login para o Gestor (chamado via GET ?action=loginGestor).
 * Retorna { ok, email, nome, role, cidade, pais } ou { ok:false, error }
 */
function loginGestor_(payload) {
  var email = String((payload && payload.email) || '').trim().toLowerCase();
  var senha = String((payload && payload.senha) || '').trim();
  return _loginUsuario_(email, senha, ['SUPERVISOR','OPERACOES','ADMIN']);
}

/**
 * Logica central de autenticacao por senha.
 */
function _loginUsuario_(email, senha, rolesPermitidos) {
  try {
    if (!email || !senha) return { ok: false, error: 'E-mail e senha obrigatorios.' };

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('USUARIOS');
    if (!aba) return { ok: false, error: 'Aba USUARIOS nao encontrada.' };

    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      var row   = dados[i];
      var rowEmail = String(row[LOGIN_COLS.EMAIL - 1] || '').trim().toLowerCase();
      if (rowEmail !== email) continue;

      // Verificar se esta ativo
      var ativo = String(row[LOGIN_COLS.ATIVO - 1] || '').trim().toUpperCase();
      if (ativo === 'FALSE' || ativo === 'NAO' || ativo === '0') {
        return { ok: false, error: 'Usuario inativo. Contate o administrador.' };
      }

      // Verificar senha
      var senhaArmazenada = String(row[LOGIN_COLS.SENHA - 1] || '').trim();
      if (!senhaArmazenada) return { ok: false, error: 'Senha nao configurada. Contate o administrador.' };
      if (senhaArmazenada !== senha) return { ok: false, error: 'E-mail ou senha incorretos.' };

      // Verificar role
      var role = String(row[LOGIN_COLS.ROLE - 1] || '').trim().toUpperCase();
      if (rolesPermitidos.length && rolesPermitidos.indexOf(role) === -1) {
        return { ok: false, error: 'Perfil ' + role + ' sem permissao para este acesso.' };
      }

      // Atualizar ultimo acesso
      try {
        aba.getRange(i + 1, 8).setValue(new Date());
      } catch(e) {}

      return {
        ok:     true,
        email:  rowEmail,
        nome:   String(row[LOGIN_COLS.NOME   - 1] || '').trim(),
        role:   role,
        cidade: String(row[LOGIN_COLS.CIDADE - 1] || '').trim(),
        pais:   String(row[LOGIN_COLS.PAIS   - 1] || '').trim()
      };
    }

    // Email nao encontrado na aba USUARIOS -- sinalizar para o frontend oferecer solicitacao
    return { ok: false, naoAutorizado: true, email: email, error: 'Email nao cadastrado. Solicite acesso.' };

  } catch(e) {
    Logger.log('_loginUsuario_ erro: ' + e);
    return { ok: false, error: 'Erro interno: ' + String(e) };
  }
}

/**
 * Definir ou resetar senha de um usuario (rodar pelo Editor como admin).
 * Exemplo: definirSenhaUsuario('campo@email.com', 'MinhasSenha123')
 */
function definirSenhaUsuario(email, novaSenha) {
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba  = ss.getSheetByName('USUARIOS');
  if (!aba) { Logger.log('Aba USUARIOS nao encontrada'); return; }

  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var rowEmail = String(dados[i][LOGIN_COLS.EMAIL - 1] || '').trim().toLowerCase();
    if (rowEmail === email.trim().toLowerCase()) {
      aba.getRange(i + 1, LOGIN_COLS.SENHA).setValue(novaSenha);
      Logger.log('Senha definida para: ' + email);
      return;
    }
  }
  Logger.log('Usuario nao encontrado: ' + email);
}

/**
 * solicitarAcesso
 * Grava solicitacao e envia email ao admin.
 */
function solicitarAcesso(payload) {
  try {
    var email    = String((payload && payload.email)    || '').trim().toLowerCase();
    var nome     = String((payload && payload.nome)     || '').trim();
    var cidade   = String((payload && payload.cidade)   || '').trim();
    var telefone = String((payload && payload.telefone) || '').trim();
    var pais     = String((payload && payload.pais)     || 'BR').trim().toUpperCase();

    if (!email) return { ok: false, error: 'Email obrigatorio.' };
    if (!nome)  return { ok: false, error: 'Nome obrigatorio.' };

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Criar aba SOLICITACOES se nao existir
    var aba = ss.getSheetByName('SOLICITACOES');
    if (!aba) {
      aba = ss.insertSheet('SOLICITACOES');
      aba.appendRow(['Timestamp','Email','Nome','Cidade','Pais','Telefone','Status','AprovadoPor','DataAprovacao']);
      aba.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a2236').setFontColor('white');
    }

    // Verificar duplicata pendente
    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][1]||'').trim().toLowerCase() === email &&
          String(dados[i][6]||'').toUpperCase() === 'PENDENTE') {
        return { ok: false, jaPendente: true, error: 'Ja existe uma solicitacao pendente para este email.' };
      }
    }

    // Gravar solicitacao
    aba.appendRow([new Date(), email, nome, cidade, pais, telefone, 'PENDENTE', '', '']);

    // Enviar email ao admin
    var emailSent = false;
    var emailErr  = '';
    try {
      // Tentar ADMIN_EMAIL das Properties, senao usar o dono do script
      var adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
                    || Session.getEffectiveUser().getEmail();

      var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:480px;background:#0b1220;color:white;border-radius:12px;padding:24px">'
        + '<h2 style="color:#60a5fa;margin-top:0">Nova solicitacao de acesso</h2>'
        + '<table style="width:100%;border-collapse:collapse">'
        + '<tr><td style="padding:6px 0;color:rgba(255,255,255,.5);width:100px">Nome</td><td style="color:white;font-weight:bold">' + nome + '</td></tr>'
        + '<tr><td style="padding:6px 0;color:rgba(255,255,255,.5)">Email</td><td style="color:#60a5fa">' + email + '</td></tr>'
        + '<tr><td style="padding:6px 0;color:rgba(255,255,255,.5)">Telefone</td><td>' + (telefone||'--') + '</td></tr>'
        + '<tr><td style="padding:6px 0;color:rgba(255,255,255,.5)">Cidade</td><td>' + (cidade||'--') + '</td></tr>'
        + '<tr><td style="padding:6px 0;color:rgba(255,255,255,.5)">Pais</td><td>' + pais + '</td></tr>'
        + '<tr><td style="padding:6px 0;color:rgba(255,255,255,.5)">Data</td><td>' + new Date().toLocaleString('pt-BR') + '</td></tr>'
        + '</table>'
        + '<div style="margin-top:20px;padding:14px;background:rgba(255,255,255,.05);border-radius:8px;font-size:13px;color:rgba(255,255,255,.6)">'
        + 'Para aprovar: execute <code style="color:#4ade80">aprovarAcesso("' + email + '")</code> no Editor do Apps Script.<br>'
        + 'Ou adicione manualmente na aba <b>USUARIOS</b> e rode <code style="color:#fbbf24">definirSenhaUsuario("' + email + '", "senha")</code>'
        + '</div></div>';

      MailApp.sendEmail({
        to: adminEmail,
        subject: '[App Estacoes] Acesso solicitado: ' + nome,
        htmlBody: htmlBody,
        body: 'Nova solicitacao de acesso. Nome: ' + nome + ' Email: ' + email + ' Telefone: ' + telefone + ' Cidade: ' + cidade
      });
      emailSent = true;
    } catch(mailErr) {
      emailErr = String(mailErr);
      Logger.log('Email admin falhou: ' + mailErr);
    }

    return { ok: true, emailSent: emailSent, emailErr: emailErr };

  } catch(e) {
    Logger.log('solicitarAcesso erro: ' + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/**
 * aprovarAcesso
 * Chamado pelo admin no Editor para aprovar uma solicitacao.
 * Cria o usuario em USUARIOS e envia email com senha temporaria.
 * Uso: aprovarAcesso("email@dominio.com")
 */
function aprovarAcesso(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) return Logger.log('Email obrigatorio.');

    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var abaSol  = ss.getSheetByName('SOLICITACOES');
    var abaUsu  = ss.getSheetByName('USUARIOS');
    if (!abaUsu) { Logger.log('Aba USUARIOS nao encontrada.'); return; }

    // Verificar se ja existe na USUARIOS
    var usuDados = abaUsu.getDataRange().getValues();
    for (var i = 1; i < usuDados.length; i++) {
      if (String(usuDados[i][0]||'').trim().toLowerCase() === email) {
        Logger.log('AVISO: ' + email + ' ja existe em USUARIOS.');
        return;
      }
    }

    // Buscar dados da solicitacao
    var nome = '', cidade = '', pais = 'BR';
    if (abaSol) {
      var solDados = abaSol.getDataRange().getValues();
      for (var j = 1; j < solDados.length; j++) {
        if (String(solDados[j][1]||'').trim().toLowerCase() === email &&
            String(solDados[j][6]||'').toUpperCase() === 'PENDENTE') {
          nome   = solDados[j][2] || '';
          cidade = solDados[j][3] || '';
          pais   = solDados[j][4] || 'BR';
          // Marcar como aprovado
          abaSol.getRange(j+1, 7).setValue('APROVADO');
          abaSol.getRange(j+1, 8).setValue(Session.getEffectiveUser().getEmail());
          abaSol.getRange(j+1, 9).setValue(new Date());
          break;
        }
      }
    }

    // Gerar senha temporaria
    var chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
    var senhaTemp = '';
    for (var k = 0; k < 8; k++) {
      senhaTemp += chars.charAt(Math.floor(Math.random() * chars.length));
    }

    // Adicionar na aba USUARIOS
    var headers = abaUsu.getRange(1, 1, 1, abaUsu.getLastColumn()).getValues()[0];
    var newRow  = new Array(headers.length).fill('');

    function setH(h, val) {
      var idx = headers.indexOf(h);
      if (idx >= 0) newRow[idx] = val;
    }
    setH('Email', email);
    setH('Nome',  nome || email);
    setH('Role',  'CAMPO');
    setH('Ativo', true);
    setH('Cidade', cidade);
    setH('Pais',  pais);
    setH('Senha', senhaTemp);
    setH('DataCriacao', new Date());

    abaUsu.appendRow(newRow);

    // Enviar email ao novo usuario com senha temporaria
    var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:420px;background:#0b1220;color:white;border-radius:12px;padding:24px">'
      + '<h2 style="color:#4ade80;margin-top:0">Acesso aprovado!</h2>'
      + '<p style="color:rgba(255,255,255,.7)">Ola ' + (nome || email) + ', seu acesso ao <b>App Estacoes Campo</b> foi aprovado.</p>'
      + '<div style="background:rgba(255,255,255,.05);border-radius:10px;padding:16px;margin:16px 0">'
      + '<div style="font-size:12px;color:rgba(255,255,255,.4);margin-bottom:4px">Seu email</div>'
      + '<div style="font-size:15px;color:#60a5fa">' + email + '</div>'
      + '<div style="font-size:12px;color:rgba(255,255,255,.4);margin:12px 0 4px">Senha temporaria</div>'
      + '<div style="font-size:22px;font-weight:bold;color:#fbbf24;letter-spacing:3px">' + senhaTemp + '</div>'
      + '</div>'
      + '<p style="font-size:13px;color:rgba(255,255,255,.5)">Esta e uma senha temporaria. Altere sua senha apos o primeiro acesso.</p>'
      + '</div>';

    MailApp.sendEmail({
      to: email,
      subject: '[App Estacoes] Seu acesso foi aprovado',
      htmlBody: htmlBody,
      body: 'Acesso aprovado! Email: ' + email + ' Senha: ' + senhaTemp
    });

    Logger.log('Acesso aprovado para ' + email + '. Senha temporaria enviada por email.');

  } catch(e) {
    Logger.log('aprovarAcesso erro: ' + e);
  }
}

/**
 * recuperarSenha
 * Gera nova senha temporaria e envia ao usuario por email.
 * Uso: recuperarSenha("email@dominio.com")
 * Tambem chamado pelo frontend via google.script.run
 */
function recuperarSenha(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) return { ok: false, error: 'Email obrigatorio.' };

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('USUARIOS');
    if (!aba) return { ok: false, error: 'Aba USUARIOS nao encontrada.' };

    var dados   = aba.getDataRange().getValues();
    var headers = dados[0];
    var iEmail  = headers.indexOf('Email');
    var iSenha  = headers.indexOf('Senha');
    var iAtivo  = headers.indexOf('Ativo');

    for (var i = 1; i < dados.length; i++) {
      var rowEmail = String(dados[i][iEmail] || '').trim().toLowerCase();
      if (rowEmail !== email) continue;

      // Verificar se ativo
      var ativo = String(dados[i][iAtivo] || '').trim().toUpperCase();
      if (ativo === 'FALSE' || ativo === 'NAO' || ativo === '0') {
        return { ok: false, error: 'Usuario inativo. Contate o administrador.' };
      }

      // Gerar nova senha temporaria
      var chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
      var novaSenha = '';
      for (var k = 0; k < 8; k++) {
        novaSenha += chars.charAt(Math.floor(Math.random() * chars.length));
      }

      // Gravar nova senha
      aba.getRange(i + 1, iSenha + 1).setValue(novaSenha);

      // Enviar email
      MailApp.sendEmail({
        to: email,
        subject: '[App Estacoes] Nova senha de acesso',
        htmlBody: '<div style="font-family:Arial,sans-serif;max-width:420px;background:#0b1220;color:white;border-radius:12px;padding:24px">'
          + '<h2 style="color:#fbbf24;margin-top:0">Nova senha gerada</h2>'
          + '<p style="color:rgba(255,255,255,.7)">Voce solicitou a recuperacao de senha.</p>'
          + '<div style="background:rgba(255,255,255,.05);border-radius:10px;padding:16px">'
          + '<div style="font-size:12px;color:rgba(255,255,255,.4);margin-bottom:4px">Nova senha temporaria</div>'
          + '<div style="font-size:24px;font-weight:bold;color:#fbbf24;letter-spacing:3px">' + novaSenha + '</div>'
          + '</div>'
          + '<p style="font-size:13px;color:rgba(255,255,255,.5);margin-top:16px">Use esta senha para fazer login. Recomendamos altera-la em seguida.</p>'
          + '</div>',
        body: 'Nova senha temporaria: ' + novaSenha
      });

      return { ok: true };
    }

    // Email nao encontrado -- nao revelar se existe ou nao (seguranca)
    return { ok: true }; // Sempre retorna ok para nao expor quais emails existem

  } catch(e) {
    Logger.log('recuperarSenha erro: ' + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}
