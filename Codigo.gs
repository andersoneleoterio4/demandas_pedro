// ===== Configuração =====
// Ajuste para o nome EXATO da aba da sua planilha.
const NOME_ABA = "Clientes";

// Coluna usada para ordenação alfabética (2 = "Cliente", pois a coluna 1 é o id).
const COLUNA_ORDENACAO = 2;

function getSheet() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(NOME_ABA);

  if (!sheet) {
    throw new Error('Aba "' + NOME_ABA + '" não encontrada na planilha.');
  }

  return sheet;
}

function respostaJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ordenarPlanilha(sheet) {
  const ultimaColuna = sheet.getLastColumn();
  const totalLinhas = sheet.getLastRow();

  if (totalLinhas <= 2) return;

  sheet
    .getRange(2, 1, totalLinhas - 1, ultimaColuna)
    .sort({ column: COLUNA_ORDENACAO, ascending: true });
}

// ===== Consulta (listar dados) =====
function doGet() {
  try {
    return respostaJson(listarClientes());
  } catch (erro) {
    return respostaJson({ erro: erro.message });
  }
}

function listarClientes() {
  const sheet = getSheet();
  ordenarPlanilha(sheet);

  const dados = sheet.getDataRange().getValues();
  const clientes = [];

  for (let i = 1; i < dados.length; i++) {
    if (dados[i].join("") === "") continue;

    const linha = dados[i];
    const formatoNovo = linha.length >= 10;

    if (formatoNovo) {
      clientes.push({
        id: linha[0],
        cliente: linha[1],
        cnpjCpf: linha[2],
        processoAnm: linha[3],
        processoAnmLink: linha[4] || "",
        nup: linha[5],
        nupLink: linha[6] || "",
        demanda: linha[7],
        demandaLink: linha[8] || "",
        status: linha[9],
      });
    } else {
      // Compatível com planilhas no formato antigo (sem colunas de link).
      clientes.push({
        id: linha[0],
        cliente: linha[1],
        cnpjCpf: linha[2],
        processoAnm: linha[3],
        processoAnmLink: "",
        nup: linha[4],
        nupLink: "",
        demanda: linha[5],
        demandaLink: "",
        status: linha[6],
      });
    }
  }

  return clientes;
}

// ===== Roteador de gravação (inserir / atualizar / excluir) =====
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const acao = body.acao;
    let resultado;

    switch (acao) {
      case "inserir":
        resultado = inserirCliente(body.cliente);
        break;
      case "atualizar":
        resultado = atualizarCliente(body.cliente);
        break;
      case "excluir":
        resultado = excluirCliente(body.id);
        break;
      default:
        throw new Error("Ação inválida: " + acao);
    }

    return respostaJson({ sucesso: resultado });
  } catch (erro) {
    return respostaJson({ erro: erro.message });
  }
}

function linhaCliente(cliente) {
  return [
    new Date().getTime(),
    cliente.cliente || "",
    cliente.cnpjCpf || "",
    cliente.processoAnm || "",
    cliente.processoAnmLink || "",
    cliente.nup || "",
    cliente.nupLink || "",
    cliente.demanda || "",
    cliente.demandaLink || "",
    cliente.status || "",
  ];
}

function valoresCliente(cliente) {
  return [
    cliente.cliente || "",
    cliente.cnpjCpf || "",
    cliente.processoAnm || "",
    cliente.processoAnmLink || "",
    cliente.nup || "",
    cliente.nupLink || "",
    cliente.demanda || "",
    cliente.demandaLink || "",
    cliente.status || "",
  ];
}

// ===== Inserir Registro =====
function inserirCliente(cliente) {
  const sheet = getSheet();
  sheet.appendRow(linhaCliente(cliente));
  ordenarPlanilha(sheet);
  return true;
}

// ===== Atualizar Registro =====
function atualizarCliente(cliente) {
  const sheet = getSheet();
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] == cliente.id) {
      sheet.getRange(i + 1, 2, 1, 9).setValues([valoresCliente(cliente)]);
      ordenarPlanilha(sheet);
      return true;
    }
  }

  return false;
}

// ===== Excluir Registro =====
function excluirCliente(id) {
  const sheet = getSheet();
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }

  return false;
}
