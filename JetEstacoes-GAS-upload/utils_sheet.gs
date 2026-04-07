function makeRowAccessor_(sheet, row) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const vals = sheet.getRange(row,1,1,sheet.getLastColumn()).getValues()[0];

  const idx = name => {
    const i = headers.indexOf(name);
    if (i === -1) throw new Error('Coluna não encontrada: ' + name);
    return i + 1;
  };

  const get = name => vals[idx(name) - 1];
  return { headers, idx, get };
}

function validarColunasObrigatorias_(sheet, colNames) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const faltando = colNames.filter(c => headers.indexOf(c) === -1);
  if (faltando.length) {
    throw new Error('Colunas ausentes:\n' + faltando.join('\n'));
  }
}

function validarSchemaEstacoes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Estacoes');
  validarColunasObrigatorias_(sh, [
    'TipoEstacao',
    'Cidade',
    'Endereço completo da estação',
    'Localização',
    'CodigoEstacao'
  ]);
}

