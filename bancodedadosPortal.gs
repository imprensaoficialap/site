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
// DIAGNÓSTICO PLANILHA — execute para investigar o acervo_doe
// =============================================================
function DIAGNOSTICO_PLANILHA() {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) { Logger.log("ERRO: aba acervo_doe não encontrada"); return; }

  var dados = sheet.getDataRange().getValues();
  Logger.log("Total de linhas (com cabeçalho): " + dados.length);
  Logger.log("Primeiras 3 linhas:");
  for (var i = 0; i < Math.min(3, dados.length); i++) {
    Logger.log("  [" + i + "] col0=" + JSON.stringify(dados[i][0]) + " tipo=" + typeof dados[i][0]);
  }

  var maxN = 0; var maxRow = -1;
  for (var i = 1; i < dados.length; i++) {
    var n = parseInt(dados[i][0]);
    if (!isNaN(n) && n > maxN) { maxN = n; maxRow = i; }
  }
  Logger.log("Maior número encontrado: " + maxN + " (linha índice " + maxRow + ")");
  if (maxRow > 0) {
    Logger.log("Conteúdo da linha com maior número: " + JSON.stringify(dados[maxRow]));
  }
}

// =============================================================
// DIAGNÓSTICO DIOFE — execute para ver o que o servidor retorna
// =============================================================
function DIAGNOSTICO_DIOFE() {
  // Pega o ID atual salvo na config
  var linkAtual = sheetConfig ? sheetConfig.getRange("A2").getValue() : "";
  var match = linkAtual.match(/\/(\d+)$/);
  var idAtual = match ? parseInt(match[1], 10) : 9824;

  Logger.log("Link atual na config: " + linkAtual);
  Logger.log("ID atual: " + idAtual);

  var url = "https://diofe.portal.ap.gov.br/portal/edicoes/download/" + idAtual;
  try {
    var options = { muteHttpExceptions: true, followRedirects: true };
    var resp = UrlFetchApp.fetch(url, options);
    var headers = resp.getHeaders();
    Logger.log("HTTP: " + resp.getResponseCode());
    Logger.log("Content-Type: " + (headers['Content-Type'] || headers['content-type'] || 'n/a'));
    Logger.log("Content-Disposition: " + (headers['Content-Disposition'] || headers['content-disposition'] || 'n/a'));
    Logger.log("Todos os headers: " + JSON.stringify(headers));
  } catch(e) {
    Logger.log("ERRO: " + e.message);
  }
}


// =============================================================
// ACERVO DOE — leitura via doGet
// =============================================================

function getAcervoPorMes(ano, mes) {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return outputJSON({ result: "error", message: "Aba acervo_doe não encontrada." });

  var dados = sheet.getDataRange().getDisplayValues();
  var edicoes = [];
  var vistos = {};

  for (var i = 1; i < dados.length; i++) {
    if (parseInt(dados[i][3]) === ano && parseInt(dados[i][4]) === mes) {
      var numero = parseInt(dados[i][0]);
      var url = (dados[i][5] || '').toString().trim();
      // Contenção de base poluída: mantém 1 registro por data de publicação/circulação.
      // Isso evita "multiplicação" de março quando o robô gravou vários números para os mesmos dias.
      var pub = (dados[i][1] || '').toString().trim();
      var circ = (dados[i][2] || '').toString().trim();
      var chave = pub + '|' + circ;
      if (vistos[chave]) continue;
      vistos[chave] = true;

      edicoes.push({
        n: numero,
        pub: pub,
        circ: circ,
        url: url
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

  // Usa data_pub (DD/MM/YYYY) como critério principal — mais confiável que as colunas ano/mes
  var melhor = null;
  var melhorVal = -1;
  var melhorN = -1;

  for (var i = 1; i < dados.length; i++) {
    var n   = parseInt(dados[i][0]);
    var pub = (dados[i][1] || '').trim();
    if (isNaN(n) || !pub) continue;

    var partes = pub.split('/');
    if (partes.length !== 3) continue;

    var dia = parseInt(partes[0]);
    var mes = parseInt(partes[1]);
    var ano = parseInt(partes[2]);
    if (isNaN(dia) || isNaN(mes) || isNaN(ano)) continue;

    var val = ano * 10000 + mes * 100 + dia;

    if (val > melhorVal || (val === melhorVal && n > melhorN)) {
      melhorVal = val;
      melhorN   = n;
      melhor    = { n: n, pub: pub, circ: dados[i][2], ano: ano, mes: mes, url: dados[i][5] };
    }
  }

  if (!melhor) return outputJSON({ result: "empty" });
  return outputJSON({ result: "success", n: melhor.n, pub: melhor.pub, circ: melhor.circ, ano: melhor.ano, mes: melhor.mes, url: melhor.url });
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

    // Coleta números já existentes para este mês/ano (evita duplicatas)
    var existentes = sheet.getDataRange().getValues();
    var numerosExistentes = {};
    for (var k = 1; k < existentes.length; k++) {
      if (parseInt(existentes[k][3]) === ano && parseInt(existentes[k][4]) === mes) {
        numerosExistentes[parseInt(existentes[k][0])] = true;
      }
    }

    for (var i = 0; i < edicoes.length; i++) {
      var e = edicoes[i];
      if (e.numero && e.url && !numerosExistentes[e.numero]) {
        sheet.appendRow([e.numero, e.pub, e.circ, ano, mes, e.url]);
        numerosExistentes[e.numero] = true;
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
// ROBÔ DIÁRIO — atualiza acervo_doe via diofe.portal.ap.gov.br
// Estratégia: sonda IDs do diofe, extrai data do filename, atribui DOE sequencial
// Configure: Triggers → roboDiarioOficialAgendado → a cada 30 min
// =============================================================

var DIOFE_BASE = 'https://diofe.portal.ap.gov.br/portal/edicoes/download/';

// Trigger: a cada 30 minutos — executa entre 07h e 13h (Brasília)
function roboDiarioOficialAgendado() {
  var hora = parseInt(Utilities.formatDate(new Date(), "GMT-3", "HH"));
  if (hora >= 7 && hora < 13) { roboDiarioOficial(); }
}

function FORCAR_ROBO_AGORA() { roboDiarioOficial(); }

function roboDiarioOficial() {
  var agora = new Date();

  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return;

  var ultimoN = getUltimoNumeroDOE(sheet);
  if (!ultimoN) { Logger.log("Robô DOE — sem registros recentes na planilha."); return; }

  var ultimaDataVal = getUltimaDataDOE(sheet);

  // Garante que config!A2 aponte para o ID mais recente do diofe
  nucleoDoRobo();

  var linkAtual = sheetConfig ? sheetConfig.getRange("A2").getValue() : "";
  var matchId = linkAtual.match(/\/(\d+)$/);
  if (!matchId) { Logger.log("Robô DOE — config!A2 sem ID diofe válido."); return; }
  var idAtual = parseInt(matchId[1], 10);

  Logger.log("Último DOE no acervo: " + ultimoN + " | data val: " + ultimaDataVal + " | diofe ID: " + idAtual);

  // Sonda os 25 IDs anteriores ao atual para encontrar edições mais recentes
  var novasEdicoes = [];
  for (var id = idAtual - 25; id <= idAtual; id++) {
    if (id <= 0) continue;
    var info = getPDFInfo(DIOFE_BASE + id);
    if (!info) continue;
    var dm = info.filename.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (!dm) continue;
    var dv = parseInt(dm[1]) * 10000 + parseInt(dm[2]) * 100 + parseInt(dm[3]);
    if (dv > ultimaDataVal) {
      novasEdicoes.push({
        id: id, url: DIOFE_BASE + id,
        ano: parseInt(dm[1]), mes: parseInt(dm[2]),
        dataStr: dm[3] + '/' + dm[2] + '/' + dm[1],
        dataVal: dv
      });
    }
  }

  // Ordena por data (mais antigas primeiro) e remove duplicatas de data
  novasEdicoes.sort(function(a, b) { return a.dataVal - b.dataVal; });
  var visto = {};
  novasEdicoes = novasEdicoes.filter(function(e) {
    if (visto[e.dataVal]) return false;
    visto[e.dataVal] = true;
    return true;
  });

  var n = ultimoN;
  for (var i = 0; i < novasEdicoes.length; i++) {
    n++;
    var e = novasEdicoes[i];
    sheet.appendRow([n, e.dataStr, e.dataStr, e.ano, e.mes, e.url]);
    Logger.log("  [+] DOE nº " + n + " (" + e.dataStr + ") — diofe ID " + e.id);
  }

  Logger.log("Robô DOE — " + Utilities.formatDate(agora, "GMT-3", "dd/MM/yyyy HH:mm") +
             " — " + novasEdicoes.length + " nova(s) edição(ões) adicionada(s).");
}

// Retorna informações do PDF (filename do Content-Disposition)
function getPDFInfo(url) {
  try {
    var options = { muteHttpExceptions: true, followRedirects: true };
    var resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() !== 200) return null;
    var ct = resp.getHeaders()['Content-Type'] || resp.getHeaders()['content-type'] || '';
    if (ct.indexOf('pdf') === -1) return null;
    var cd = resp.getHeaders()['Content-Disposition'] || resp.getHeaders()['content-disposition'] || '';
    var fn = cd.match(/filename="([^"]+)"/);
    return { filename: fn ? fn[1] : '' };
  } catch(e) { return null; }
}

// Retorna o valor numérico da data mais recente no acervo (YYYYMMDD)
function getUltimaDataDOE(sheet) {
  var anoAtual = new Date().getFullYear();
  var dados = sheet.getDataRange().getValues();
  var maxData = 0;
  for (var i = 1; i < dados.length; i++) {
    var ano = parseInt(dados[i][3]);
    if (ano < anoAtual - 1) continue;
    var pub = (dados[i][1] || '').toString().trim();
    var p = pub.split('/');
    if (p.length !== 3) continue;
    var dia = parseInt(p[0]), mes = parseInt(p[1]), anoD = parseInt(p[2]);
    if (isNaN(dia) || isNaN(mes) || isNaN(anoD)) continue;
    var val = anoD * 10000 + mes * 100 + dia;
    if (val > maxData) maxData = val;
  }
  return maxData;
}

// Retorna o maior número de DOE no acervo (apenas anos 2025-2026+)
function getUltimoNumeroDOE(sheet) {
  var anoAtual = new Date().getFullYear();
  var dados = sheet.getDataRange().getValues();
  var maxN = 0;
  for (var i = 1; i < dados.length; i++) {
    var ano = parseInt(dados[i][3]);
    if (ano < anoAtual - 1) continue;
    var n = parseInt(dados[i][0]);
    if (!isNaN(n) && n > maxN) maxN = n;
  }
  return maxN > 0 ? maxN : null;
}
