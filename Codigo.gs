// ===== Configuração =====
const NOME_ABA = "Clientes";

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
  const dados = sheet.getDataRange().getValues();
  const clientes = [];

  for (let i = 1; i < dados.length; i++) {
    // Ignora linhas totalmente vazias
    if (dados[i].join("") === "") continue;

    clientes.push({
      id: dados[i][0],
      cliente: dados[i][1],
      cnpjCpf: dados[i][2],
      processoAnm: dados[i][3],
      nup: dados[i][4],
      demanda: dados[i][5],
      status: dados[i][6],
    });
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

// ===== Inserir Registro =====
function inserirCliente(cliente) {
  const sheet = getSheet();

  sheet.appendRow([
    new Date().getTime(),
    cliente.cliente,
    cliente.cnpjCpf,
    cliente.processoAnm,
    cliente.nup,
    cliente.demanda,
    cliente.status,
  ]);

  return true;
}

// ===== Atualizar Registro =====
function atualizarCliente(cliente) {
  const sheet = getSheet();
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] == cliente.id) {
      sheet.getRange(i + 1, 2).setValue(cliente.cliente);
      sheet.getRange(i + 1, 3).setValue(cliente.cnpjCpf);
      sheet.getRange(i + 1, 4).setValue(cliente.processoAnm);
      sheet.getRange(i + 1, 5).setValue(cliente.nup);
      sheet.getRange(i + 1, 6).setValue(cliente.demanda);
      sheet.getRange(i + 1, 7).setValue(cliente.status);

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
