/**
 * monitor_gs.gs
 * Gerencia a aba MONITOR do projeto.
 *
 * ESTRUTURA DA ABA MONITOR:
 *   CodigoEstacao | Ativo | Motivo | DataInicio | DataFim | Obs | CriadoPor | DataCriacao
 *
 * COMO USAR:
 *   - Marcar/desmarcar via campo.html (painel de edicao da estacao)
 *   - listarMonitor()  retorna lista de codigos ativos para o mapa
 *   - toggleMonitor()  ativa/desativa monitoramento de uma estacao
 *
 * O mapa enriquece cada estacao com { monitor: true/false }
 * antes de renderizar, comparando com a lista retornada por listarMonitor()
 */


/* ── utilitario: obter/criar aba MONITOR ───────────────────────────────── */
function getAbaMonitor_() {
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName('MONITOR');
  if (!aba) {
    aba = ss.insertSheet('MONITOR');
    var headers = ['CodigoEstacao','Ativo','Motivo','DataInicio','DataFim','Obs','CriadoPor','DataCriacao'];
    aba.appendRow(headers);
    aba.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold')
       .setBackground('#1a2236')
       .setFontColor('#ffffff');
    aba.setFrozenRows(1);
    aba.setColumnWidth(1, 180);
    aba.setColumnWidth(3, 220);
    Logger.log('Aba MONITOR criada.');
  }
  return aba;
}


/* ── listarMonitor ─────────────────────────────────────────────────────── */
/**
 * Retorna array de codigos que estao ATIVOS no MONITOR.
 * Chamado pelo frontend via google.script.run para enriquecer as estacoes.
 */
function listarMonitor() {
  try {
    var aba   = getAbaMonitor_();
    var dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return [];

    var headers = dados[0].map(function(h){ return String(h).trim(); });
    var iCod    = headers.indexOf('CodigoEstacao');
    var iAtivo  = headers.indexOf('Ativo');

    if (iCod < 0) return [];

    var ativos = [];
    for (var i = 1; i < dados.length; i++) {
      var ativo = String(dados[i][iAtivo] || '').toUpperCase().trim();
      if (ativo === 'TRUE' || ativo === 'SIM' || ativo === '1') {
        var cod = String(dados[i][iCod] || '').trim();
        if (cod) ativos.push(cod);
      }
    }
    return ativos;

  } catch(e) {
    Logger.log('listarMonitor erro: ' + e);
    return [];
  }
}


/* ── toggleMonitor ─────────────────────────────────────────────────────── */
/**
 * Ativa ou desativa o monitoramento de uma estacao.
 * Payload: { codigo, ativo, motivo, obs, email }
 * ativo = true  → monitora
 * ativo = false → remove monitoramento
 */
function toggleMonitor(payload) {
  try {
    var codigo = String((payload && payload.codigo) || '').trim();
    if (!codigo) return { ok: false, error: 'Codigo obrigatorio.' };

    var ativo  = payload.ativo === true || payload.ativo === 'true';
    var motivo = String((payload && payload.motivo) || '').trim();
    var obs    = String((payload && payload.obs)    || '').trim();
    var email  = String((payload && payload.email)  || '').trim();

    var aba    = getAbaMonitor_();
    var dados  = aba.getDataRange().getValues();
    var headers= dados[0].map(function(h){ return String(h).trim(); });

    var iCod   = headers.indexOf('CodigoEstacao');
    var iAtivo = headers.indexOf('Ativo');
    var iMot   = headers.indexOf('Motivo');
    var iIni   = headers.indexOf('DataInicio');
    var iFim   = headers.indexOf('DataFim');
    var iObs   = headers.indexOf('Obs');
    var iUser  = headers.indexOf('CriadoPor');

    // Verificar se ja existe linha para este codigo
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][iCod] || '').trim() === codigo) {
        var row = i + 1;
        aba.getRange(row, iAtivo + 1).setValue(ativo);
        if (iMot  >= 0) aba.getRange(row, iMot  + 1).setValue(motivo || dados[i][iMot]);
        if (iObs  >= 0) aba.getRange(row, iObs  + 1).setValue(obs    || dados[i][iObs]);
        if (iUser >= 0) aba.getRange(row, iUser + 1).setValue(email);
        if (!ativo && iFim >= 0) aba.getRange(row, iFim + 1).setValue(new Date());
        if ( ativo && iIni >= 0) aba.getRange(row, iIni + 1).setValue(new Date());
        return { ok: true, codigo: codigo, ativo: ativo };
      }
    }

    // Nao existe -- criar nova linha (so se ativando)
    if (!ativo) return { ok: true, codigo: codigo, ativo: false }; // nada a fazer

    var novaLinha = new Array(headers.length).fill('');
    novaLinha[iCod]  = codigo;
    novaLinha[iAtivo]= true;
    if (iMot  >= 0) novaLinha[iMot]  = motivo;
    if (iIni  >= 0) novaLinha[iIni]  = new Date();
    if (iObs  >= 0) novaLinha[iObs]  = obs;
    if (iUser >= 0) novaLinha[iUser] = email;

    var iDataCriacao = headers.indexOf('DataCriacao');
    if (iDataCriacao >= 0) novaLinha[iDataCriacao] = new Date();

    aba.appendRow(novaLinha);
    return { ok: true, codigo: codigo, ativo: true };

  } catch(e) {
    Logger.log('toggleMonitor erro: ' + e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}
