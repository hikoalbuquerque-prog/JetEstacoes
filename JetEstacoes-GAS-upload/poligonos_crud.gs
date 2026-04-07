function salvarPoligono(params) {
  try {
    if (!params || !params.cidade || !params.nome || !params.pipe) {
      return { ok: false, error: 'Campos obrigatorios ausentes (cidade, nome, pipe).' };
    }
 
    var sh = SpreadsheetApp.getActive().getSheetByName(POLIGONOS_SHEET);
    if (!sh) return { ok: false, error: 'Aba ' + POLIGONOS_SHEET + ' nao encontrada.' };
 
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim(); });
 
    var iCidade     = header.indexOf('Cidade');
    var iGrupo      = header.indexOf('Grupo');
    var iFase       = header.indexOf('Fase');
    var iNome       = header.indexOf('Nome Área');
    var iTipo       = header.indexOf('Tipo');
    var iPrioridade = header.indexOf('Prioridade');
    var iAtivo      = header.indexOf('Ativo');
    var iCor        = header.indexOf('Cor');
    var iPol        = header.indexOf('Poligono');
 
    if (iCidade === -1 || iNome === -1 || iPol === -1) {
      return { ok: false, error: 'Colunas obrigatorias ausentes no cabecalho.' };
    }
 
    // Montar nova linha com base no header
    var lastCol = sh.getLastColumn();
    var novaLinha = new Array(lastCol).fill('');
 
    novaLinha[iCidade]     = String(params.cidade     || '').trim();
    novaLinha[iGrupo]      = String(params.grupo      || '').trim();
    novaLinha[iFase]       = String(params.fase       || 'Fase 1').trim();
    novaLinha[iNome]       = String(params.nome       || '').trim();
    novaLinha[iTipo]       = String(params.tipo       || '').trim();
    novaLinha[iPrioridade] = Number(params.prioridade)  || 1;
    novaLinha[iAtivo]      = (params.ativo === false || params.ativo === 'FALSE') ? 'FALSE' : 'TRUE';
    novaLinha[iCor]        = String(params.cor        || '#2563eb').trim();
    novaLinha[iPol]        = String(params.pipe       || '').trim();
 
    sh.appendRow(novaLinha);
    var rowId = sh.getLastRow();
 
    // Invalidar cache da cidade
    limparCachePoligonosCidade(params.cidade);
 
    Logger.log('salvarPoligono: linha ' + rowId + ' — ' + params.nome);
    return { ok: true, rowId: rowId };
 
  } catch (e) {
    Logger.log('salvarPoligono erro: ' + e);
    return { ok: false, error: String(e.message || e) };
  }
}
 
 
/**
 * atualizarPoligono(params)
 * Atualiza linha existente identificada por rowId.
 * Retorna { ok: true } ou { ok: false, error: '...' }
 *
 * params: { rowId, cidade, grupo, fase, nome, tipo, prioridade, ativo, cor, pipe }
 */
function atualizarPoligono(params) {
  try {
    if (!params || !params.rowId) {
      return { ok: false, error: 'rowId obrigatorio para atualizarPoligono.' };
    }
 
    var sh = SpreadsheetApp.getActive().getSheetByName(POLIGONOS_SHEET);
    if (!sh) return { ok: false, error: 'Aba ' + POLIGONOS_SHEET + ' nao encontrada.' };
 
    var rowId = Number(params.rowId);
    if (rowId < 2 || rowId > sh.getLastRow()) {
      return { ok: false, error: 'rowId invalido: ' + rowId };
    }
 
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim(); });
 
    var campos = {
      'Cidade'     : params.cidade,
      'Grupo'      : params.grupo,
      'Fase'       : params.fase,
      'Nome Área'  : params.nome,
      'Tipo'       : params.tipo,
      'Prioridade' : params.prioridade,
      'Ativo'      : (params.ativo === false || params.ativo === 'FALSE') ? 'FALSE' : 'TRUE',
      'Cor'        : params.cor,
      'Poligono'   : params.pipe
    };
 
    // Atualizar apenas colunas presentes no header
    Object.keys(campos).forEach(function(colName) {
      var col = header.indexOf(colName);
      if (col === -1) return;
      var val = campos[colName];
      if (val === undefined || val === null) return;
      sh.getRange(rowId, col + 1).setValue(
        colName === 'Prioridade' ? (Number(val) || 1) : String(val)
      );
    });
 
    // Invalidar cache da cidade
    if (params.cidade) limparCachePoligonosCidade(params.cidade);
 
    Logger.log('atualizarPoligono: linha ' + rowId + ' — ' + params.nome);
    return { ok: true };
 
  } catch (e) {
    Logger.log('atualizarPoligono erro: ' + e);
    return { ok: false, error: String(e.message || e) };
  }
}
 
 
/**
 * excluirPoligono(params)
 * Remove linha identificada por rowId.
 * Retorna { ok: true } ou { ok: false, error: '...' }
 *
 * params: { rowId, cidade }
 */
function excluirPoligono(params) {
  try {
    if (!params || !params.rowId) {
      return { ok: false, error: 'rowId obrigatorio para excluirPoligono.' };
    }
 
    var sh = SpreadsheetApp.getActive().getSheetByName(POLIGONOS_SHEET);
    if (!sh) return { ok: false, error: 'Aba ' + POLIGONOS_SHEET + ' nao encontrada.' };
 
    var rowId = Number(params.rowId);
    if (rowId < 2 || rowId > sh.getLastRow()) {
      return { ok: false, error: 'rowId invalido: ' + rowId };
    }
 
    // Guardar cidade para invalidar cache antes de deletar
    var cidade = params.cidade;
    if (!cidade) {
      var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim(); });
      var iCidade = header.indexOf('Cidade');
      if (iCidade >= 0) {
        cidade = sh.getRange(rowId, iCidade + 1).getValue();
      }
    }
 
    sh.deleteRow(rowId);
 
    if (cidade) limparCachePoligonosCidade(String(cidade).trim());
 
    Logger.log('excluirPoligono: linha ' + rowId);
    return { ok: true };
 
  } catch (e) {
    Logger.log('excluirPoligono erro: ' + e);
    return { ok: false, error: String(e.message || e) };
  }
}
