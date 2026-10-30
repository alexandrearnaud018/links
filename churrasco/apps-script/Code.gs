// Ponte entre o site e a planilha do churrasco.
// Script criado DENTRO da planilha (Extensões → Apps Script), então não precisa do ID dela.
// A senha fica em Configurações do projeto → Propriedades do script → SENHA (nunca no código).
// Toda chamada é um POST com corpo JSON: { senha, acao, ... }

const ABA = 'Itens';
// tipo: pessoa | compra | config
// ok: pagou (pessoa) ou comprado (compra); ref: preço do churrasco do hexa; pago: valor pago (ou valor da config)
const COLUNAS = ['id', 'tipo', 'nome', 'ok', 'ref', 'pago', 'ordem'];

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (_) { return resposta({ erro: 'requisicao_invalida' }); }

  const senha = PropertiesService.getScriptProperties().getProperty('SENHA');
  if (!senha || req.senha !== senha) return resposta({ erro: 'senha' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (req.acao === 'listar') return resposta({ itens: lerItens() });
    if (req.acao === 'salvar') { salvar(req.item); return resposta({ ok: true }); }
    if (req.acao === 'excluir') { excluir(req.id); return resposta({ ok: true }); }
    return resposta({ erro: 'acao_desconhecida' });
  } catch (err) {
    return resposta({ erro: 'falha', detalhe: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function aba() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
}

function lerItens() {
  const linhas = aba().getDataRange().getValues();
  linhas.shift(); // cabeçalho
  return linhas
    .filter((l) => l[0] !== '')
    .map((l) => {
      const it = {};
      COLUNAS.forEach((c, i) => { it[c] = l[i] === '' ? null : l[i]; });
      it.id = String(it.id);
      it.ok = it.ok === true || it.ok === 'TRUE';
      return it;
    });
}

// Linha (1-based) de cada id
function indice() {
  const sh = aba();
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const map = {};
  ids.forEach((r, i) => { if (i > 0 && r[0] !== '') map[String(r[0])] = i + 1; });
  return map;
}

// Cria ou atualiza; só os campos enviados mudam
function salvar(item) {
  if (!item || !item.id) throw new Error('item sem id');
  const sh = aba();
  const linha = indice()[item.id];
  if (linha) {
    const atual = sh.getRange(linha, 1, 1, COLUNAS.length).getValues()[0];
    const nova = COLUNAS.map((c, i) => (c in item ? (item[c] == null ? '' : item[c]) : atual[i]));
    sh.getRange(linha, 1, 1, COLUNAS.length).setValues([nova]);
  } else {
    sh.appendRow(COLUNAS.map((c) => (item[c] == null ? '' : item[c])));
  }
}

function excluir(id) {
  const linha = indice()[id];
  if (linha) aba().deleteRow(linha);
}

// Rode UMA vez pelo editor (botão Executar): cria a aba "Itens" com a lista inicial.
// Não apaga nada: se a aba já existir, para sem mexer.
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ABA)) throw new Error('A aba "' + ABA + '" já existe. Nada foi alterado.');
  const sh = ss.insertSheet(ABA);
  const linhas = [COLUNAS].concat(INICIAL);
  sh.getRange(1, 1, linhas.length, COLUNAS.length).setValues(linhas);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, COLUNAS.length).setFontWeight('bold');
}

// Lista inicial: 87 pessoas (ninguém pagou) e 35 compras (nada comprado), com o preço do churrasco do hexa como referência
const INICIAL = [
  ["valorPessoa", "config", "Valor por pessoa", false, "", 52, 0],
  ["salao", "config", "Salão", false, "", 200, 0],
  ["p001", "pessoa", "Hani", false, "", "", 1],
  ["p002", "pessoa", "Rafael", false, "", "", 2],
  ["p003", "pessoa", "Ana Júlia", false, "", "", 3],
  ["p004", "pessoa", "Andreyna", false, "", "", 4],
  ["p005", "pessoa", "Babs", false, "", "", 5],
  ["p006", "pessoa", "Bernardo", false, "", "", 6],
  ["p007", "pessoa", "Bernardo Rosa", false, "", "", 7],
  ["p008", "pessoa", "Carla", false, "", "", 8],
  ["p009", "pessoa", "Carol", false, "", "", 9],
  ["p010", "pessoa", "Clarice", false, "", "", 10],
  ["p011", "pessoa", "Danilo", false, "", "", 11],
  ["p012", "pessoa", "Day", false, "", "", 12],
  ["p013", "pessoa", "Elaine", false, "", "", 13],
  ["p014", "pessoa", "Eunice", false, "", "", 14],
  ["p015", "pessoa", "Fê", false, "", "", 15],
  ["p016", "pessoa", "Fernanda", false, "", "", 16],
  ["p017", "pessoa", "Filipe Bellio", false, "", "", 17],
  ["p018", "pessoa", "Filipe Oliveira", false, "", "", 18],
  ["p019", "pessoa", "Renato", false, "", "", 19],
  ["p020", "pessoa", "Giovanna", false, "", "", 20],
  ["p021", "pessoa", "Handel", false, "", "", 21],
  ["p022", "pessoa", "Hellen", false, "", "", 22],
  ["p023", "pessoa", "Jady", false, "", "", 23],
  ["p024", "pessoa", "Jamila", false, "", "", 24],
  ["p025", "pessoa", "João", false, "", "", 25],
  ["p026", "pessoa", "João Carnevale", false, "", "", 26],
  ["p027", "pessoa", "Juliana", false, "", "", 27],
  ["p028", "pessoa", "Kogut", false, "", "", 28],
  ["p029", "pessoa", "Letícia", false, "", "", 29],
  ["p030", "pessoa", "Lindi", false, "", "", 30],
  ["p031", "pessoa", "Lorrein", false, "", "", 31],
  ["p032", "pessoa", "Lucas Arnaud", false, "", "", 32],
  ["p033", "pessoa", "Lucas Pavanelli", false, "", "", 33],
  ["p034", "pessoa", "Luiz Paulo", false, "", "", 34],
  ["p035", "pessoa", "Luiza", false, "", "", 35],
  ["p036", "pessoa", "Luíze", false, "", "", 36],
  ["p037", "pessoa", "Luly", false, "", "", 37],
  ["p038", "pessoa", "Luquinhas", false, "", "", 38],
  ["p039", "pessoa", "Luz", false, "", "", 39],
  ["p040", "pessoa", "Marmota", false, "", "", 40],
  ["p041", "pessoa", "Mateus", false, "", "", 41],
  ["p042", "pessoa", "Mi", false, "", "", 42],
  ["p043", "pessoa", "Miguel", false, "", "", 43],
  ["p044", "pessoa", "Mirtoc", false, "", "", 44],
  ["p045", "pessoa", "Misha", false, "", "", 45],
  ["p046", "pessoa", "Nando", false, "", "", 46],
  ["p047", "pessoa", "Natan", false, "", "", 47],
  ["p048", "pessoa", "Nic", false, "", "", 48],
  ["p049", "pessoa", "Paula", false, "", "", 49],
  ["p050", "pessoa", "Carlos", false, "", "", 50],
  ["p051", "pessoa", "Pedrinho", false, "", "", 51],
  ["p052", "pessoa", "PPS", false, "", "", 52],
  ["p053", "pessoa", "Rodrigo", false, "", "", 53],
  ["p054", "pessoa", "Rebeca", false, "", "", 54],
  ["p055", "pessoa", "Rodrigo T.", false, "", "", 55],
  ["p056", "pessoa", "Rufião", false, "", "", 56],
  ["p057", "pessoa", "Sari", false, "", "", 57],
  ["p058", "pessoa", "Sergio", false, "", "", 58],
  ["p059", "pessoa", "Tamir", false, "", "", 59],
  ["p060", "pessoa", "Thiengo", false, "", "", 60],
  ["p061", "pessoa", "Vanessa", false, "", "", 61],
  ["p062", "pessoa", "Vitor", false, "", "", 62],
  ["p063", "pessoa", "Vivi", false, "", "", 63],
  ["p064", "pessoa", "Danielzinho", false, "", "", 64],
  ["p065", "pessoa", "Elisa", false, "", "", 65],
  ["p066", "pessoa", "Liz", false, "", "", 66],
  ["p067", "pessoa", "Isabela Belio", false, "", "", 67],
  ["p068", "pessoa", "Nathalia", false, "", "", 68],
  ["p069", "pessoa", "Vic", false, "", "", 69],
  ["p070", "pessoa", "Isa", false, "", "", 70],
  ["p071", "pessoa", "Carol Mattos", false, "", "", 71],
  ["p072", "pessoa", "Allan H.", false, "", "", 72],
  ["p073", "pessoa", "Fernanda W", false, "", "", 73],
  ["p074", "pessoa", "Matheus B", false, "", "", 74],
  ["p075", "pessoa", "Aline", false, "", "", 75],
  ["p076", "pessoa", "Matheus M", false, "", "", 76],
  ["p077", "pessoa", "Martinelli", false, "", "", 77],
  ["p078", "pessoa", "Pinho", false, "", "", 78],
  ["p079", "pessoa", "Marcelo", false, "", "", 79],
  ["p080", "pessoa", "Emily", false, "", "", 80],
  ["p081", "pessoa", "Rogerinho", false, "", "", 81],
  ["p082", "pessoa", "Dani Vitor", false, "", "", 82],
  ["p083", "pessoa", "Gabriel Barros", false, "", "", 83],
  ["p084", "pessoa", "Osorio", false, "", "", 84],
  ["p085", "pessoa", "Vinicius", false, "", "", 85],
  ["p086", "pessoa", "Virginia", false, "", "", 86],
  ["p087", "pessoa", "Andre dock", false, "", "", 87],
  ["c001", "compra", "2 kg arroz", false, 8.28, "", 1],
  ["c002", "compra", "2 kg farofa", false, 16.35, "", 2],
  ["c003", "compra", "Alho zs", false, 16.9, "", 3],
  ["c004", "compra", "Cebola 3", false, 6.16, "", 4],
  ["c005", "compra", "Tomate 1kg", false, 14.16, "", 5],
  ["c006", "compra", "Pimentão 4", false, 9.64, "", 6],
  ["c007", "compra", "Vinagre", false, 7.25, "", 7],
  ["c008", "compra", "Alcatra", false, 365.65, "", 8],
  ["c009", "compra", "Picanha e Filet Mignon suíno", false, 56.61, "", 9],
  ["c010", "compra", "Costela janelinha 4,5kg", false, 144.08, "", 10],
  ["c011", "compra", "Picanha", false, 89.98, "", 11],
  ["c012", "compra", "Linguiça", false, 142.87, "", 12],
  ["c013", "compra", "Coração", false, 53.97, "", 13],
  ["c014", "compra", "Fósforo acendedor", false, 12.99, "", 14],
  ["c015", "compra", "Azeite", false, 24.99, "", 15],
  ["c016", "compra", "Papel celofane/alumínio", false, 4.59, "", 16],
  ["c017", "compra", "Gelo", false, 143.99, "", 17],
  ["c018", "compra", "Sorvete", false, 84.0, "", 18],
  ["c019", "compra", "Banana", false, 35.91, "", 19],
  ["c020", "compra", "Drumet", false, 25.96, "", 20],
  ["c021", "compra", "Queijo coalho", false, 53.98, "", 21],
  ["c022", "compra", "Pão de alho", false, 73.65, "", 22],
  ["c023", "compra", "Carvão", false, 64.5, "", 23],
  ["c024", "compra", "Prato descartável", false, 85.27, "", 24],
  ["c025", "compra", "Copo descartável", false, 23.97, "", 25],
  ["c026", "compra", "Talher descartável", false, 43.94, "", 26],
  ["c027", "compra", "Acendedor", false, 9.98, "", 27],
  ["c028", "compra", "Refrigerante", false, 36.46, "", 28],
  ["c029", "compra", "Matte", false, 15.8, "", 29],
  ["c030", "compra", "Peixe kogut", false, 20.56, "", 30],
  ["c031", "compra", "Óleo", false, 21.84, "", 31],
  ["c032", "compra", "Papel toalha/Guardanapo", false, 21.44, "", 32],
  ["c033", "compra", "Fraldinha", false, 82.17, "", 33],
  ["c034", "compra", "Sal de churrasco/parrilha", false, 37.49, "", 34],
  ["c035", "compra", "Chimichurri", false, 4.39, "", 35],
];
