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
// SANEAMENTO acervo_doe — executa manualmente, gera log e aba backup
// Chave canônica: pub|circ  (uma edição por par de datas de pub/circ)
// Critério de "vencedor": registro com maior numero_doe (oriundo do SEAD histórico)
// Rollback: aba acervo_doe_backup_YYYYMMDD criada antes de qualquer remoção
// =============================================================
function SANEAR_ACERVO_DOE() {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) { Logger.log("ERRO: aba acervo_doe não encontrada"); return; }

  // 1. Backup antes de qualquer alteração
  var dataBackup = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd_HHmm");
  var nomeBackup = "acervo_doe_backup_" + dataBackup;
  var backup = sheet.copyTo(ss);
  backup.setName(nomeBackup);
  Logger.log("Backup criado: " + nomeBackup);

  // 2. Lê todos os dados (display para evitar Date objects)
  var dados = sheet.getDataRange().getDisplayValues();
  var cabecalho = dados[0];
  Logger.log("Total linhas (com cabeçalho): " + dados.length);

  // 3. Agrupa por chave pub|circ — mantém o registro com maior numero_doe
  var grupos = {};  // chave → { indice, numero_doe, linha }
  var duplicatas = [];

  for (var i = 1; i < dados.length; i++) {
    var pub  = (dados[i][1] || '').toString().trim();
    var circ = (dados[i][2] || '').toString().trim();
    var n    = parseInt(dados[i][0]);
    var chave = pub + '|' + circ;

    if (!pub && !circ) {
      Logger.log("  [SKIP] linha " + (i+1) + ": pub e circ vazios, ignorado");
      continue;
    }

    if (!grupos[chave]) {
      grupos[chave] = { indice: i, n: isNaN(n) ? 0 : n };
    } else {
      // Regra de vencedor:
      // 1) número > 0 (real) vence número 0 (pendente)
      // 2) entre dois números > 0, mantém o menor (evita preservar inflado sintético)
      // 3) entre dois 0, mantém o primeiro
      var atual = grupos[chave];
      var nAtual = atual.n;
      var nNovo = isNaN(n) ? 0 : n;
      var deveAtualizar =
        (nAtual === 0 && nNovo > 0) ||
        (nAtual > 0 && nNovo > 0 && nNovo < nAtual);

      if (deveAtualizar) {
        duplicatas.push({ linha: atual.indice + 1, chave: chave, n: nAtual, motivo: "substituído por n=" + nNovo });
        grupos[chave] = { indice: i, n: nNovo };
      } else {
        duplicatas.push({ linha: i + 1, chave: chave, n: nNovo, motivo: "duplicata de " + chave + " (vencedor n=" + grupos[chave].n + ")" });
      }
    }
  }

  Logger.log("Registros únicos: " + Object.keys(grupos).length);
  Logger.log("Duplicatas identificadas: " + duplicatas.length);

  // 4. Monta o conjunto de índices a MANTER (vencedores)
  var manter = {};
  for (var chave in grupos) { manter[grupos[chave].indice] = true; }

  // 5. Reconstrói em lote (muito mais rápido que deleteRow em loop)
  var saida = [cabecalho];
  for (var i = 1; i < dados.length; i++) {
    if (manter[i]) saida.push(dados[i]);
  }

  var removidos = (dados.length - 1) - (saida.length - 1);

  sheet.clearContents();
  sheet.getRange(1, 1, saida.length, cabecalho.length).setValues(saida);
  if (sheet.getMaxRows() > saida.length) {
    sheet.deleteRows(saida.length + 1, sheet.getMaxRows() - saida.length);
  }

  Logger.log("Linhas removidas: " + removidos);
  Logger.log("Linhas restantes: " + (saida.length - 1));

  // 6. Log detalhado das duplicatas
  if (duplicatas.length > 0) {
    Logger.log("--- DUPLICATAS REMOVIDAS ---");
    for (var d = 0; d < Math.min(duplicatas.length, 100); d++) {
      Logger.log("  Linha " + duplicatas[d].linha + " | n=" + duplicatas[d].n + " | " + duplicatas[d].chave + " | " + duplicatas[d].motivo);
    }
    if (duplicatas.length > 100) Logger.log("  ... e mais " + (duplicatas.length - 100) + " omitidas.");
  }

  Logger.log("SANEAMENTO CONCLUÍDO. Rollback disponível na aba: " + nomeBackup);
}

// =============================================================
// OBSERVABILIDADE — detecta crescimento anômalo por mês
// Execute manualmente ou agende mensalmente
// =============================================================
function ALERTAR_CRESCIMENTO_ANOMALO() {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return;
  var dados = sheet.getDataRange().getDisplayValues();

  var contadores = {};  // "YYYY-MM" → count
  for (var i = 1; i < dados.length; i++) {
    var ano = (dados[i][3] || '').toString().trim();
    var mes = (dados[i][4] || '').toString().trim();
    if (!ano || !mes) continue;
    var chave = ano + '-' + (mes.length === 1 ? '0' + mes : mes);
    contadores[chave] = (contadores[chave] || 0) + 1;
  }

  var LIMITE_NORMAL = 35; // DOEs por mês — acima disso é suspeito
  var alertas = [];
  for (var k in contadores) {
    if (contadores[k] > LIMITE_NORMAL) {
      alertas.push("ANOMALIA: " + k + " tem " + contadores[k] + " registros (limite=" + LIMITE_NORMAL + ")");
    }
  }

  if (alertas.length > 0) {
    Logger.log("=== ALERTAS DE CRESCIMENTO ANÔMALO ===");
    alertas.forEach(function(a) { Logger.log(a); });
  } else {
    Logger.log("OK — nenhum mês acima de " + LIMITE_NORMAL + " registros.");
  }

  // Log geral dos últimos 6 meses
  var chaves = Object.keys(contadores).sort().slice(-6);
  Logger.log("Últimos 6 meses: " + chaves.map(function(k){ return k + "=" + contadores[k]; }).join(", "));
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
  var cache = CacheService.getScriptCache();
  var cacheKey = "acervo_" + ano + "_" + mes;
  var cached = cache.get(cacheKey);
  if (cached) {
    return outputJSON(JSON.parse(cached));
  }

  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) return outputJSON({ result: "error", message: "Aba acervo_doe não encontrada." });

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return outputJSON({ result: "success", edicoes: [] });
  // Lê somente A:F (colunas usadas pelo acervo) para reduzir latência.
  var dados = sheet.getRange(1, 1, lastRow, 6).getDisplayValues();
  var porData = {};

  for (var i = 1; i < dados.length; i++) {
    if (parseInt(dados[i][3]) === ano && parseInt(dados[i][4]) === mes) {
      var numero = parseInt(dados[i][0]);
      var url = (dados[i][5] || '').toString().trim();
      var pub = (dados[i][1] || '').toString().trim();
      var circ = (dados[i][2] || '').toString().trim();
      var chave = pub + '|' + circ;
      var atual = porData[chave];
      var cand = {
        n: numero,
        pub: pub,
        circ: circ,
        url: url
      };

      if (!atual) {
        porData[chave] = cand;
      } else {
        var atualSead = urlEhSead(atual.url);
        var candSead = urlEhSead(cand.url);
        if (candSead && !atualSead) {
          porData[chave] = cand;
        } else if (candSead === atualSead) {
          // Empate de fonte: mantém menor número > 0 (evita preservar inflado sintético)
          var nAtual = parseInt(atual.n);
          var nCand = parseInt(cand.n);
          if (nAtual <= 0 && nCand > 0) {
            porData[chave] = cand;
          } else if (nAtual > 0 && nCand > 0 && nCand < nAtual) {
            porData[chave] = cand;
          }
        }
      }
    }
  }

  var edicoes = [];
  for (var k in porData) edicoes.push(porData[k]);

  // Mais recentes primeiro
  edicoes.sort(function(a, b) { return b.n - a.n; });
  var payload = { result: "success", edicoes: edicoes };
  cache.put(cacheKey, JSON.stringify(payload), 300); // 5 minutos
  return outputJSON(payload);
}

function urlEhSead(url) {
  var u = (url || '').toString().toLowerCase();
  return u.indexOf('sead.portal.ap.gov.br') !== -1 || u.indexOf('seadantigo.portal.ap.gov.br') !== -1;
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

  for (var i = 0; i < novasEdicoes.length; i++) {
    var e = novasEdicoes[i];
    sheet.appendRow([0, e.dataStr, e.dataStr, e.ano, e.mes, e.url]);
    Logger.log("  [+] DOE pendente (" + e.dataStr + ") — diofe ID " + e.id);
  }

  Logger.log("Robô DOE — " + Utilities.formatDate(agora, "GMT-3", "dd/MM/yyyy HH:mm") +
             " — " + novasEdicoes.length + " nova(s) edição(ões) adicionada(s).");

  // Observabilidade: conta registros do mês atual para detectar explosão
  var mesAtual = agora.getMonth() + 1;
  var anoAtual2 = agora.getFullYear();
  var dadosObs = sheet.getDataRange().getDisplayValues();
  var contMes = 0;
  for (var oi = 1; oi < dadosObs.length; oi++) {
    if (parseInt(dadosObs[oi][3]) === anoAtual2 && parseInt(dadosObs[oi][4]) === mesAtual) contMes++;
  }
  Logger.log("  [OBS] Registros em " + mesAtual + "/" + anoAtual2 + ": " + contMes + (contMes > 35 ? " ← ANOMALIA" : ""));
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
  var dados = sheet.getDataRange().getDisplayValues(); // getDisplayValues evita Date objects do Sheets
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
  var dados = sheet.getDataRange().getDisplayValues();
  var maxN = 0;
  for (var i = 1; i < dados.length; i++) {
    var ano = parseInt(dados[i][3]);
    if (ano < anoAtual - 1) continue;
    var n = parseInt(dados[i][0]);
    if (!isNaN(n) && n > maxN) maxN = n;
  }
  return maxN > 0 ? maxN : null;
}

// =============================================================
// REPROCESSAMENTO DE MÊS (SEAD) — corrige mês contaminado
// Exemplo: REPROCESSAR_MES_SEAD(2026, 3)
// =============================================================
function REPROCESSAR_MES_SEAD(ano, mes) {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) { Logger.log("ERRO: aba acervo_doe não encontrada"); return; }
  if (!ano || !mes) { Logger.log("ERRO: informe ano e mes"); return; }

  // Backup antes de alterar
  var dataBackup = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd_HHmm");
  var nomeBackup = "acervo_doe_backup_reproc_" + ano + "_" + mes + "_" + dataBackup;
  var backup = sheet.copyTo(ss);
  backup.setName(nomeBackup);
  Logger.log("Backup criado: " + nomeBackup);

  // Remove somente o mês/ano alvo
  var dados = sheet.getDataRange().getValues();
  var manter = [dados[0]];
  var removidas = 0;
  for (var i = 1; i < dados.length; i++) {
    if (parseInt(dados[i][3]) === ano && parseInt(dados[i][4]) === mes) {
      removidas++;
      continue;
    }
    manter.push(dados[i]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, manter.length, manter[0].length).setValues(manter);
  if (sheet.getMaxRows() > manter.length) {
    sheet.deleteRows(manter.length + 1, sheet.getMaxRows() - manter.length);
  }
  Logger.log("Linhas removidas do mês alvo: " + removidas);

  // Recoleta oficial do mês pelo parser SEAD
  rasparMes(sheet, ano, mes);
  Logger.log("REPROCESSAMENTO CONCLUÍDO: " + ano + "/" + mes);
}

function REPROCESSAR_MAR_2026() {
  REPROCESSAR_MES_SEAD(2026, 3);
}

function CORRIGIR_MAR_2026_MANTER_SEAD() {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) { Logger.log("ERRO: aba acervo_doe não encontrada"); return; }

  var dataBackup = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd_HHmm");
  var nomeBackup = "acervo_doe_backup_fix_mar_2026_" + dataBackup;
  sheet.copyTo(ss).setName(nomeBackup);
  Logger.log("Backup criado: " + nomeBackup);

  var dados = sheet.getDataRange().getDisplayValues();
  var cab = dados[0];
  var saida = [cab];
  var removidas = 0;

  for (var i = 1; i < dados.length; i++) {
    var ano = parseInt(dados[i][3]);
    var mes = parseInt(dados[i][4]);
    var url = (dados[i][5] || '').toString();
    if (ano === 2026 && mes === 3 && !urlEhSead(url)) {
      removidas++;
      continue;
    }
    saida.push(dados[i]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, saida.length, cab.length).setValues(saida);
  if (sheet.getMaxRows() > saida.length) {
    sheet.deleteRows(saida.length + 1, sheet.getMaxRows() - saida.length);
  }

  Logger.log("MAR/2026 não-SEAD removidas: " + removidas);
  Logger.log("Concluído. Rollback: " + nomeBackup);
}

function parseDDMMYYYY_(s) {
  var p = (s || '').toString().trim().split('/');
  if (p.length !== 3) return 0;
  var d = parseInt(p[0], 10), m = parseInt(p[1], 10), a = parseInt(p[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(a)) return 0;
  return a * 10000 + m * 100 + d;
}

function normalizarLinha_(linha, totalCols) {
  var out = (linha || []).slice(0, totalCols);
  while (out.length < totalCols) out.push('');
  return out;
}

// Restaura somente MAR/2026 a partir de uma aba de backup.
// Ex.: RESTAURAR_MARCO_DO_BACKUP('acervo_doe_backup_20260401_1121')
function RESTAURAR_MARCO_DO_BACKUP(nomeAbaBackup) {
  var sheet = ss.getSheetByName("acervo_doe");
  var backup = ss.getSheetByName(nomeAbaBackup);
  if (!sheet) { Logger.log("ERRO: aba acervo_doe não encontrada"); return; }
  if (!backup) { Logger.log("ERRO: aba backup não encontrada: " + nomeAbaBackup); return; }

  var stamp = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd_HHmm");
  var nomeSeg = "acervo_doe_backup_pre_restore_mar_2026_" + stamp;
  sheet.copyTo(ss).setName(nomeSeg);
  Logger.log("Backup de segurança criado: " + nomeSeg);

  var dadosAtual = sheet.getDataRange().getDisplayValues();
  var totalCols = sheet.getLastColumn();
  var cab = normalizarLinha_(dadosAtual[0], totalCols);
  var saida = [cab];

  for (var i = 1; i < dadosAtual.length; i++) {
    var ano = parseInt(dadosAtual[i][3], 10);
    var mes = parseInt(dadosAtual[i][4], 10);
    if (ano === 2026 && mes === 3) continue;
    saida.push(normalizarLinha_(dadosAtual[i], totalCols));
  }

  var dadosBkp = backup.getDataRange().getDisplayValues();
  var inseridas = 0;
  for (var j = 1; j < dadosBkp.length; j++) {
    var anoB = parseInt(dadosBkp[j][3], 10);
    var mesB = parseInt(dadosBkp[j][4], 10);
    if (anoB === 2026 && mesB === 3) {
      saida.push(normalizarLinha_(dadosBkp[j], totalCols));
      inseridas++;
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, saida.length, totalCols).setValues(saida);
  if (sheet.getMaxRows() > saida.length) {
    sheet.deleteRows(saida.length + 1, sheet.getMaxRows() - saida.length);
  }

  Logger.log("MAR/2026 restaurado do backup. Linhas inseridas: " + inseridas);
}

// Renumera MAR/2026 em sequência por data de publicação (ascendente).
// Ex.: RENUMERAR_MARCO_2026(8605)
function RENUMERAR_MARCO_2026(primeiroNumero) {
  var sheet = ss.getSheetByName("acervo_doe");
  if (!sheet) { Logger.log("ERRO: aba acervo_doe não encontrada"); return; }
  if (!primeiroNumero || isNaN(parseInt(primeiroNumero, 10))) {
    Logger.log("ERRO: informe primeiroNumero válido. Ex.: 8605");
    return;
  }

  var stamp = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd_HHmm");
  var nomeSeg = "acervo_doe_backup_pre_renum_mar_2026_" + stamp;
  sheet.copyTo(ss).setName(nomeSeg);
  Logger.log("Backup de segurança criado: " + nomeSeg);

  var dados = sheet.getDataRange().getDisplayValues();
  var totalCols = sheet.getLastColumn();
  var cab = normalizarLinha_(dados[0], totalCols);
  var marco = [];
  var outros = [cab];

  for (var i = 1; i < dados.length; i++) {
    var ano = parseInt(dados[i][3], 10);
    var mes = parseInt(dados[i][4], 10);
    if (ano === 2026 && mes === 3) {
      marco.push({
        pub: (dados[i][1] || '').toString().trim(),
        circ: (dados[i][2] || '').toString().trim(),
        ano: dados[i][3],
        mes: dados[i][4],
        url: (dados[i][5] || '').toString().trim()
      });
    } else {
      outros.push(normalizarLinha_(dados[i], totalCols));
    }
  }

  // Dedup por dia (pub|circ), preservando a primeira ocorrência.
  var seen = {};
  var unicas = [];
  for (var k = 0; k < marco.length; k++) {
    var key = marco[k].pub + '|' + marco[k].circ;
    if (seen[key]) continue;
    seen[key] = true;
    unicas.push(marco[k]);
  }

  unicas.sort(function(a, b) { return parseDDMMYYYY_(a.pub) - parseDDMMYYYY_(b.pub); });

  var n = parseInt(primeiroNumero, 10) - 1;
  for (var u = 0; u < unicas.length; u++) {
    n++;
    outros.push(normalizarLinha_([String(n), unicas[u].pub, unicas[u].circ, '2026', '3', unicas[u].url], totalCols));
  }

  sheet.clearContents();
  sheet.getRange(1, 1, outros.length, totalCols).setValues(outros);
  if (sheet.getMaxRows() > outros.length) {
    sheet.deleteRows(outros.length + 1, sheet.getMaxRows() - outros.length);
  }

  Logger.log("MAR/2026 renumerado. Total edições: " + unicas.length + ". Faixa DOE: " + primeiroNumero + " até " + n);
}

function RESTAURAR_MAR_2026_AGORA() {
  RESTAURAR_MARCO_DO_BACKUP('acervo_doe_backup_20260401_1121');
}

function RENUMERAR_MAR_2026_AGORA() {
  RENUMERAR_MARCO_2026(8605);
}
