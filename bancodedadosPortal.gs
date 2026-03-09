// =============================================================
// SISTEMA NIO PESQUISA - BACKEND INSTITUCIONAL
// =============================================================

var ss = SpreadsheetApp.openById('1zFEf9Sq9FQDLvsEbxzItWa0Qn0NIkQiXQrdcVuCbSMA');

// Definição das abas pelos nomes exatos dos seus prints
var sheetPedidos = ss.getSheetByName("Página1"); 
var sheetRefs    = ss.getSheetByName("Referencias");
var sheetConfig  = ss.getSheetByName("config");
var sheetUsers   = ss.getSheetByName("usuarios");

// --- FUNÇÃO GET (Leitura de dados) ---
function doGet(e) {
  var acao = e.parameter.acao;

  // 1. LISTAR PEDIDOS (Painel Principal)
  if (acao == "listar_pedidos") {
    if (!sheetPedidos) return outputJSON({ "result": "error", "message": "Aba Página1 não encontrada." });
    var dados = sheetPedidos.getDataRange().getDisplayValues();
    dados.shift(); // Remove cabeçalho
    return outputJSON({ "result": "success", "pedidos": dados });
  }

  // 2. LISTAR REFERÊNCIAS (Guia de Referências)
  else if (acao == "listar_referencias_completa" || acao == "listar_referencias") {
    if (!sheetRefs) return outputJSON({ "result": "error", "message": "Aba Referencias não encontrada." });
    var dadosRef = sheetRefs.getDataRange().getDisplayValues();
    dadosRef.shift(); // Remove cabeçalho
    return outputJSON({ "result": "success", "referencias": dadosRef });
  }

  // 3. BUSCAR LINK DO DIÁRIO (Para o botão da Index)
  else if (acao == "link_diario") {
    if (!sheetConfig) return outputJSON({ "result": "empty", "link": "" });
    var linkDoDia = sheetConfig.getRange("A2").getValue();
    if (linkDoDia && linkDoDia.toString().includes("http")) {
      return outputJSON({ "result": "success", "link": linkDoDia });
    } else {
      return outputJSON({ "result": "empty", "link": "" });
    }
  }

  return outputJSON({ "result": "error", "message": "Ação inválida." });
}

// --- FUNÇÃO POST (Gravação e Alteração) ---
function doPost(e) {
  var dados = JSON.parse(e.postData.contents);

  // 1. LOGIN (Validação na aba usuarios)
  if (dados.acao == "login") {
    var usuario = (dados.usuario || "").toString().trim().toLowerCase();
    var senha = (dados.senha || "").toString().trim();
    if (!sheetUsers) return outputJSON({ result: "error" });

    var dadosUsuarios = sheetUsers.getDataRange().getValues();
    for (var i = 1; i < dadosUsuarios.length; i++) {
      if (dadosUsuarios[i][0].toString().toLowerCase() === usuario && 
          dadosUsuarios[i][1].toString() === senha) {
        return outputJSON({ result: "success", nome: dadosUsuarios[i][2] });
      }
    }
    return outputJSON({ result: "error" });
  }

  // 2. NOVO PEDIDO (Vindo do site ou do botão Novo Pedido)
  else if (dados.acao == "novo_pedido") {
    var id = "ID_" + Math.random().toString(16).substr(2, 8);
    var dataHora = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
    
    sheetPedidos.appendRow([
      id, 
      dataHora, 
      (dados.solicitante || "").toUpperCase(), 
      (dados.nome_pesquisado || "").toUpperCase(), 
      (dados.periodo || "").toUpperCase(), 
      dados.whatsapp || "", 
      "EM ANÁLISE", 
      (dados.responsavel || "SISTEMA").toUpperCase(), 
      (dados.observacoes || "").toUpperCase()
    ]);
    
    formatarUltimaLinha(sheetPedidos);
    return outputJSON({ "result": "success" });
  }

  // 3. ATUALIZAR PEDIDO (Botão Atender)
  else if (dados.acao == "atualizar_pedido") {
    var allData = sheetPedidos.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][0].toString() == dados.id.toString()) {
        var linhaReal = i + 1;
        // Coluna G (7): Status
        sheetPedidos.getRange(linhaReal, 7).setValue(String(dados.status || "").toUpperCase());
        // Coluna H (8): Responsável
        sheetPedidos.getRange(linhaReal, 8).setValue(String(dados.responsavel || "").toUpperCase());
        
        // GRAVAÇÃO NA COLUNA K (11) - RETORNO AO CLIENTE
        if (dados.retorno !== undefined) {
           sheetPedidos.getRange(linhaReal, 11).setValue(dados.retorno); 
        }
        return outputJSON({ "result": "success" });
      }
    }
  }

  // 4. EXCLUIR PEDIDO
  else if (dados.acao == "excluir_pedido") {
    return excluirLinhaPorID(sheetPedidos, dados.id);
  }

  // 5. NOVA REFERÊNCIA (Guia de Referências)
  else if (dados.acao == "nova_referencia") {
    var idRef = "REF_" + Math.random().toString(36).substr(2, 9).toUpperCase();
    sheetRefs.appendRow([
      (dados.organizacao || "").toUpperCase(), 
      dados.data || "", 
      dados.referencia || "", 
      idRef
    ]);
    return outputJSON({ "result": "success" });
  }

  // 6. EXCLUIR REFERÊNCIA
  else if (dados.acao == "excluir_referencia") {
    return excluirLinhaPorID(sheetRefs, dados.id);
  }

  return outputJSON({ "result": "error" });
}

// --- FUNÇÕES DE APOIO ---

function outputJSON(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto)).setMimeType(ContentService.MimeType.JSON);
}

function formatarUltimaLinha(aba) {
  var lastRow = aba.getLastRow();
  var range = aba.getRange(lastRow, 1, 1, aba.getLastColumn());
  range.setBorder(true, true, true, true, true, true)
       .setVerticalAlignment("middle")
       .setFontSize(10)
       .setWrap(true);
}

function excluirLinhaPorID(aba, id) {
  var allData = aba.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0].toString() == id.toString()) {
      aba.deleteRow(i + 1);
      return outputJSON({ "result": "success" });
    }
  }
  return outputJSON({ "result": "error" });
}

// --- ROBÔ DE RASPAGEM AUTOMÁTICA ---

function atualizarLinkDiarioAutomaticamente() {
  var agora = new Date();
  var horaAtual = parseInt(Utilities.formatDate(agora, "GMT-3", "HH"));
  var horasPermitidas = [9, 10, 11, 14, 15, 20, 21, 22, 23];
  if (horasPermitidas.indexOf(horaAtual) !== -1) { nucleoDoRobo(); }
}

function FORCAR_ATUALIZACAO_AGORA() { nucleoDoRobo(); }

function nucleoDoRobo() {
  if (!sheetConfig) return;
  var linkAtual = sheetConfig.getRange("A2").getValue();
  var idAtual = 0;
  var match = linkAtual.match(/\/(\d+)$/);
  idAtual = match ? parseInt(match[1], 10) : 9824;

  var novoID = idAtual;
  var encontrouNovo = false;

  for (var i = 1; i <= 30; i++) {
    var idTeste = idAtual + i;
    var urlTeste = "https://diofe.portal.ap.gov.br/portal/edicoes/download/" + idTeste;
    if (ehUmPDFReal(urlTeste)) { novoID = idTeste; encontrouNovo = true; }
  }

  if (encontrouNovo) {
    var linkFinal = "https://diofe.portal.ap.gov.br/portal/edicoes/download/" + novoID;
    sheetConfig.getRange("A2").setValue(linkFinal);
  }
}

function ehUmPDFReal(url) {
  try {
    var options = { 'muteHttpExceptions': true, 'followRedirects': true };
    var response = UrlFetchApp.fetch(url, options);
    var contentType = response.getHeaders()['Content-Type'] || response.getHeaders()['content-type'] || "";
    return (contentType.indexOf("pdf") !== -1);
  } catch (e) { return false; }
}