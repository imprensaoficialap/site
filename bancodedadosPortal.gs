// =============================================================
// SISTEMA NIO PESQUISA - BACKEND INSTITUCIONAL
// =============================================================

var ss = SpreadsheetApp.openById('1zFEf9Sq9FQDLvsEbxzItWa0Qn0NIkQiXQrdcVuCbSMA');

// Definição das abas pelos nomes exatos dos seus prints
var sheetPedidos = ss.getSheetByName("Página1");
var sheetRefs    = ss.getSheetByName("Referencias");
var sheetConfig  = ss.getSheetByName("config");
var sheetUsers   = ss.getSheetByName("usuarios");
var sheetAcervo  = ss.getSheetByName("acervo_doe");

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

  // 3. ACERVO DOE — edições de um mês específico
  else if (acao == "acervo_doe") {
    var ano = parseInt(e.parameter.ano);
    var mes = parseInt(e.parameter.mes);
    return getAcervoPorMes(ano, mes);
  }

  // 4. ÚLTIMA EDIÇÃO DO DOE — para o card de destaque
  else if (acao == "ultima_doe") {
    return getUltimaEdicao();
  }

  // 5. BUSCAR LINK DO DIÁRIO (Para o botão da Index)
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


// =============================================================
// ACERVO DOE — leitura via doGet
// =============================================================

function getAcervoPorMes(ano, mes) {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return outputJSON({ result: "error", message: "Aba acervo_doe não encontrada." });

  var dados = sheet.getDataRange().getDisplayValues();
  var edicoes = [];

  for (var i = 1; i < dados.length; i++) {
    if (parseInt(dados[i][3]) === ano && parseInt(dados[i][4]) === mes) {
      edicoes.push({
        n:    parseInt(dados[i][0]),
        pub:  dados[i][1],
        circ: dados[i][2],
        url:  dados[i][5]
      });
    }
  }

  // Mais recentes primeiro
  edicoes.sort(function(a, b) { return b.n - a.n; });
  return outputJSON({ result: "success", edicoes: edicoes });
}

function getUltimaEdicao() {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return outputJSON({ result: "empty" });

  var dados = sheet.getDataRange().getDisplayValues();
  if (dados.length < 2) return outputJSON({ result: "empty" });

  // Última linha inserida = edição mais recente
  var u = dados[dados.length - 1];
  return outputJSON({
    result: "success",
    n:    parseInt(u[0]),
    pub:  u[1],
    circ: u[2],
    ano:  parseInt(u[3]),
    mes:  parseInt(u[4]),
    url:  u[5]
  });
}


// =============================================================
// RASPAGEM HISTÓRICA — rode manualmente, 1 vez por bloco de anos
// Progresso salvo na aba "config", célula B2
// =============================================================

var MESES_URL  = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
var SEAD_BASE  = 'https://sead.portal.ap.gov.br/diario_oficial/';
var ANOS_POR_EXEC = 5; // quantos anos processa por execução (ajuste se precisar)

function rasparAcervoHistorico() {
  var sheet = ss.getSheetByName("acervo_doe");

  // Cria a aba com cabeçalho se não existir
  if (!sheet) {
    sheet = ss.insertSheet("acervo_doe");
    sheet.appendRow(["numero_doe","data_pub","data_circ","ano","mes","url_pdf"]);
    sheet.setFrozenRows(1);
  }

  var anoInicio = getProgressoRaspagem();
  var anoFim    = Math.min(anoInicio + ANOS_POR_EXEC - 1, new Date().getFullYear());

  Logger.log("▶ Raspando " + anoInicio + " → " + anoFim);

  for (var ano = anoInicio; ano <= anoFim; ano++) {
    var mesMax = (ano === new Date().getFullYear()) ? new Date().getMonth() + 1 : 12;

    for (var mes = 1; mes <= mesMax; mes++) {
      if (mesJaRaspado(sheet, ano, mes)) {
        Logger.log("  [OK] " + ano + "/" + MESES_URL[mes-1] + " já existe, pulando");
        continue;
      }
      rasparMes(sheet, ano, mes);
      Utilities.sleep(1000); // pausa para não sobrecarregar o SEAD
    }
  }

  salvarProgressoRaspagem(anoFim + 1);
  Logger.log("✔ Concluído! Próxima execução começa em: " + (anoFim + 1));
}

function rasparMes(sheet, ano, mes) {
  var mesStr = MESES_URL[mes - 1];
  var url    = SEAD_BASE + ano + '/' + mesStr;

  try {
    var opcoes = {
      muteHttpExceptions: true,
      validateHttpsCertificates: false,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9',
        'Referer': 'https://sead.portal.ap.gov.br/diario_oficial'
      }
    };

    var resp = UrlFetchApp.fetch(url, opcoes);
    if (resp.getResponseCode() !== 200) {
      Logger.log("  [" + resp.getResponseCode() + "] " + ano + "/" + mesStr);
      return;
    }

    var html    = resp.getContentText("UTF-8");
    var edicoes = parsearTabelaSEAD(html);

    for (var i = 0; i < edicoes.length; i++) {
      var e = edicoes[i];
      if (e.numero && e.url) {
        sheet.appendRow([e.numero, e.pub, e.circ, ano, mes, e.url]);
      }
    }

    Logger.log("  [" + edicoes.length + "] " + ano + "/" + mesStr);
  } catch (err) {
    Logger.log("  [ERRO] " + ano + "/" + mesStr + " — " + err.message);
  }
}

function parsearTabelaSEAD(html) {
  var edicoes   = [];
  var rowRegex  = /<tr[\s\S]*?>([\s\S]*?)<\/tr>/gi;
  var tdRegex   = /<td[^>]*>([\s\S]*?)<\/td>/gi;
  var linkRegex = /href=["'](https?:\/\/[^"']+\.pdf[^"']*)/i;
  var numRegex  = /DOE[^n]*n[°º\.\s]*(\d+)/i;
  var dateRegex = /(\d{2}\/\d{2}\/\d{4})/;

  var rowMatch;
  while ((rowMatch = rowRegex.exec(html)) !== null) {
    var rowHtml = rowMatch[1];
    var tds = [];
    var tdMatch;

    while ((tdMatch = tdRegex.exec(rowHtml)) !== null) {
      // Remove tags internas para pegar o texto limpo
      tds.push(tdMatch[1].replace(/<[^>]+>/g, ' ').trim());
    }

    if (tds.length < 3) continue;

    var pub      = (tds[0].match(dateRegex) || [])[1] || '';
    var circ     = (tds[1].match(dateRegex) || [])[1] || '';
    var numMatch = (tds[2] || '').match(numRegex);
    if (!numMatch) continue;

    var numero = parseInt(numMatch[1]);

    // Link PDF — busca em toda a linha (pode estar em qualquer <td>)
    var urlMatch = rowHtml.match(linkRegex);
    if (!urlMatch) continue;

    edicoes.push({ numero: numero, pub: pub, circ: circ, url: urlMatch[1] });
  }

  return edicoes;
}

function mesJaRaspado(sheet, ano, mes) {
  var dados = sheet.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (parseInt(dados[i][3]) === ano && parseInt(dados[i][4]) === mes) return true;
  }
  return false;
}

function getProgressoRaspagem() {
  var config = ss.getSheetByName("config");
  if (!config) return 1964;
  var val = config.getRange("B2").getValue();
  return (val && !isNaN(parseInt(val))) ? parseInt(val) : 1964;
}

function salvarProgressoRaspagem(proximoAno) {
  var config = ss.getSheetByName("config");
  if (config) config.getRange("B2").setValue(proximoAno);
}


// =============================================================
// ROBÔ DIÁRIO — atualiza automaticamente via trigger
// Configure: Triggers → roboDiarioOficial → Por hora → 23h–00h
// =============================================================

function roboDiarioOficial() {
  var agora = new Date();
  var ano   = agora.getFullYear();
  var mes   = agora.getMonth() + 1;

  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return;

  // Registra quantas linhas havia antes
  var linhasAntes = sheet.getLastRow();

  rasparMes(sheet, ano, mes);

  var linhasDepois = sheet.getLastRow();
  var novas = linhasDepois - linhasAntes;

  Logger.log(
    "Robô DOE — " +
    Utilities.formatDate(agora, "GMT-3", "dd/MM/yyyy HH:mm") +
    " — " + novas + " nova(s) edição(ões) adicionada(s)."
  );
}