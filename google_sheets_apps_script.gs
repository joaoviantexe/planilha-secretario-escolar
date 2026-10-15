/**
 * CHECKLIST POR TURMA - Google Sheets (Apps Script)
 *
 * Como instalar:
 *   1. Envie Planilha_Checklist_Turmas.xlsx para o Google Drive e abra com
 *      o Google Planilhas (ou Arquivo > Importar > Fazer upload).
 *   2. Menu Extensoes > Apps Script.
 *   3. Apague o conteudo de Codigo.gs, cole TODO este texto e clique em Salvar.
 *   4. Volte para a planilha e recarregue a pagina (F5). Pronto: nao precisa
 *      de autorizacao nem de gatilho, o onEdit simples roda sozinho.
 *
 * Funcionamento:
 *   - Ao trocar a turma em A1 da aba Checklist, as respostas dessa turma sao
 *     carregadas da aba oculta BaseDados.
 *   - Ao alterar Situacao, Ultima Atualizacao, Prazo, Responsavel ou
 *     Observacao, a resposta e guardada para a turma selecionada e A1 mostra
 *     a nova contagem de itens pendentes.
 */
var ABA_CHECKLIST = 'Checklist';
var ABA_BASE = 'BaseDados';
var ABA_TURMAS = 'Turmas';

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== ABA_CHECKLIST) return;

  var r = e.range;
  var ss = sh.getParent();
  var touchesA1 = r.getRow() <= 1 && r.getLastRow() >= 1 &&
                  r.getColumn() <= 1 && r.getLastColumn() >= 1;

  if (touchesA1) {
    carregarTurma_(ss, sh);
    return;
  }
  if (r.getLastRow() >= 3 && r.getRow() <= 40 && r.getLastColumn() >= 2 && r.getColumn() <= 6) {
    salvarCelulas_(ss, sh, r);
    atualizarTitulo_(ss, sh);
  }
}

// Codigo da turma (ex.: C21101) lido do texto de A1
function codigoTurma_(ss, sh) {
  var txt = String(sh.getRange('A1').getValue());
  var cods = ss.getSheetByName(ABA_TURMAS).getRange('M2:M200').getValues();
  for (var i = 0; i < cods.length; i++) {
    var c = String(cods[i][0]).trim();
    if (c && txt.indexOf('| ' + c) !== -1) return c;
  }
  return '';
}

// Linhas da BaseDados: [{linha: n, idx: posicao}] da turma
function linhasDaTurma_(base, cod) {
  var n = base.getLastRow();
  if (n < 2) return [];
  var col = base.getRange(2, 1, n - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]) === cod) out.push({ linhaChecklist: Number(col[i][1]), linhaBase: i + 2 });
  }
  return out;
}

function carregarTurma_(ss, sh) {
  var cod = codigoTurma_(ss, sh);
  if (!cod) return;
  var base = ss.getSheetByName(ABA_BASE);
  var linhas = linhasDaTurma_(base, cod);
  linhas.forEach(function (l) {
    var v = base.getRange(l.linhaBase, 4, 1, 5).getValues();  // D..H
    sh.getRange(l.linhaChecklist, 2, 1, 5).setValues(v);       // B..F
  });
  atualizarTitulo_(ss, sh);
}

function salvarCelulas_(ss, sh, range) {
  var cod = codigoTurma_(ss, sh);
  if (!cod) return;
  var base = ss.getSheetByName(ABA_BASE);
  var linhas = linhasDaTurma_(base, cod);
  var porLinha = {};
  linhas.forEach(function (l) { porLinha[l.linhaChecklist] = l.linhaBase; });
  var valores = range.getValues();
  for (var i = 0; i < valores.length; i++) {
    var linha = range.getRow() + i;
    if (!porLinha[linha]) continue;           // cabecalhos de secao
    for (var j = 0; j < valores[i].length; j++) {
      var col = range.getColumn() + j;
      if (col < 2 || col > 6) continue;
      base.getRange(porLinha[linha], col + 2).setValue(valores[i][j]);
    }
  }
  SpreadsheetApp.flush();
}

function atualizarTitulo_(ss, sh) {
  var cod = codigoTurma_(ss, sh);
  if (!cod) return;
  SpreadsheetApp.flush();
  var t = ss.getSheetByName(ABA_TURMAS);
  var cods = t.getRange('M2:M200').getValues();
  for (var i = 0; i < cods.length; i++) {
    if (String(cods[i][0]).trim() === cod) {
      sh.getRange('A1').setValue(t.getRange(i + 2, 21).getValue());
      return;
    }
  }
}
