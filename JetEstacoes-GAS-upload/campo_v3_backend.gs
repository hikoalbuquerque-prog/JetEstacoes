Skip to content
hikoalbuquerque-prog
JetEstacoes
Repository navigation
Code
Issues
Pull requests
Actions
Projects
Wiki
Security and quality
2
 (2)
JetEstacoes/JetEstacoes-GAS-upload
/
campo_v3_backend.gs
in
main

Edit

Preview
Indent mode

Spaces
Indent size

2
Line wrap mode

No wrap
Editing campo_v3_backend.gs file contents
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
/**
 * campo_v3_backend_gs.gs
 * Adicionar como arquivo novo no projeto GAS.
 * Contem: getRankingCampoHoje, editarEstacaoFromMapa, geocodeEndereco
 */

/* ----------------------------------------------------------------
 * getRankingCampoHoje
 * Retorna array { email, nome, total } ordenado desc por cadastros do dia
 * ---------------------------------------------------------------- */
function getRankingCampoHoje() {
  try {
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('Estacoes') || ss.getSheetByName('ESTACOES');
    if (!aba) return [];

    var data    = aba.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });

    // Encontrar colunas
    var iOp   = headers.indexOf('CriadoPor');
    if (iOp < 0) iOp = headers.indexOf('Operador');
    var iData = headers.indexOf('DataCriacao');

    if (iOp < 0) return [];

    var hoje = new Date(); hoje.setHours(0,0,0,0);
    var ranking = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var op  = String(row[iOp] || '').trim();
      if (!op) continue;
      if (iData >= 0) {
        var d = new Date(row[iData]);
        if (isNaN(d.getTime()) || d < hoje) continue;
Use Control + Shift + m to toggle the tab key moving focus. Alternatively, use esc then tab to move to the next interactive element on the page.
