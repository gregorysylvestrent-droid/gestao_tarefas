// ============================================================
//  NORTE TECH — GESTÃO DE TAREFAS DELEGADAS
//  Code.gs — Apps Script Backend
//  Cole este código no Apps Script vinculado à sua planilha.
// ============================================================

// ============================================================
//  CONFIGURAÇÕES — ajuste conforme necessário
// ============================================================
var CONFIG = {
  ABA_TAREFAS: "TAREFAS",
  ABA_RESPONSAVEIS: "RESPONSÁVEIS",
  EMAIL_REMETENTE: Session.getActiveUser().getEmail(),
  TITULO_PLANILHA: "Gestão de Tarefas — Norte Tech Logística",
  // Colunas da aba TAREFAS (índice base 1)
  COL: {
    ID:            1,
    TITULO:        2,
    DESCRICAO:     3,
    RESPONSAVEL:   4,
    CATEGORIA:     5,
    PRIORIDADE:    6,
    STATUS:        7,
    DATA_CRIACAO:  8,
    DATA_PREVISTA: 9,
    DATA_CONCLUSAO:10,
    OBSERVACOES:   11,
    ALERTAS:       12  // contador de alertas enviados
  }
};

var CATEGORIAS   = ["Administrativo","Armazém","Compras","Dashboards", "Frota", "Instalação", "Logística", "Oficina"];
var PRIORIDADES  = ["Alta", "Média", "Baixa"];
var STATUS_LIST  = ["Pendente", "Em Andamento", "Concluída", "Cancelada"];

// ============================================================
//  doGet — serve o HTML do dashboard
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("dashboard_tarefas")
    .setTitle("Gestão de Tarefas — Norte Tech")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
//  SETUP — cria abas e cabeçalhos se não existirem
// ============================================================
function setupPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- ABA TAREFAS ---
  var abaTarefas = ss.getSheetByName(CONFIG.ABA_TAREFAS);
  if (!abaTarefas) {
    abaTarefas = ss.insertSheet(CONFIG.ABA_TAREFAS);
  }

  var cabecalho = [
    "ID", "Título", "Descrição", "Responsável", "Categoria",
    "Prioridade", "Status", "Data Criação", "Data Prevista",
    "Data Conclusão", "Observações", "Nº Alertas"
  ];

  if (abaTarefas.getLastRow() === 0) {
    abaTarefas.appendRow(cabecalho);
    formatarCabecalho_(abaTarefas);
  }

  // --- ABA RESPONSÁVEIS ---
  var abaResp = ss.getSheetByName(CONFIG.ABA_RESPONSAVEIS);
  if (!abaResp) {
    abaResp = ss.insertSheet(CONFIG.ABA_RESPONSAVEIS);
    abaResp.appendRow(["Nome", "Email", "Setor"]);
    formatarCabecalho_(abaResp);
    // Exemplos iniciais — edite direto na planilha
    abaResp.appendRow(["Mauricio", "mauricio@nortetech.net", "Logística"]);
    abaResp.appendRow(["Gregory",  "gregory@nortetech.net",  "Logística"]);
  }

  SpreadsheetApp.getUi().alert("✅ Planilha configurada com sucesso!");
}

// ============================================================
//  MENU PERSONALIZADO
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔧 Norte Tech — Tarefas")
    .addItem("⚙️ Configurar planilha",           "setupPlanilha")
    .addItem("🔔 Enviar alertas de vencimento",   "enviarAlertasVencimento")
    .addItem("📊 Abrir Dashboard",                "abrirDashboard")
    .addItem("🗑️  Limpar tarefas concluídas (>30d)","limparConcluidas")
    .addToUi();
}

function abrirDashboard() {
  var html = HtmlService
    .createHtmlOutputFromFile("dashboard_tarefas")
    .setWidth(1200)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, "📊 Dashboard de Tarefas");
}

// ============================================================
//  LEITURA DE DADOS — chamado pelo frontend
// ============================================================
function getTarefasData() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var sheet    = ss.getSheetByName(CONFIG.ABA_TAREFAS);
    var sheetR   = ss.getSheetByName(CONFIG.ABA_RESPONSAVEIS);

    if (!sheet) return JSON.stringify({ erro: "Aba TAREFAS não encontrada. Execute '⚙️ Configurar planilha' primeiro." });

    // ---- Tarefas ----
    var dados = sheet.getDataRange().getValues();
    var tarefas = [];
    var hoje = new Date();
    hoje.setHours(0,0,0,0);

    for (var i = 1; i < dados.length; i++) {
      var row = dados[i];
      if (!row[0]) continue; // linha vazia

      var dataPrevista  = row[CONFIG.COL.DATA_PREVISTA  - 1];
      var dataConclusao = row[CONFIG.COL.DATA_CONCLUSAO - 1];
      var status        = row[CONFIG.COL.STATUS - 1] || "Pendente";

      var diasRestantes = null;
      var vencida       = false;

      if (dataPrevista instanceof Date && dataPrevista.getTime() > 0) {
        var dp = new Date(dataPrevista);
        dp.setHours(0,0,0,0);
        diasRestantes = Math.round((dp - hoje) / 86400000);
        vencida = (diasRestantes < 0 && status !== "Concluída" && status !== "Cancelada");
      }

      tarefas.push({
        id:            row[CONFIG.COL.ID            - 1],
        titulo:        row[CONFIG.COL.TITULO        - 1],
        descricao:     row[CONFIG.COL.DESCRICAO     - 1],
        responsavel:   row[CONFIG.COL.RESPONSAVEL   - 1],
        categoria:     row[CONFIG.COL.CATEGORIA     - 1],
        prioridade:    row[CONFIG.COL.PRIORIDADE    - 1],
        status:        status,
        dataCriacao:   dataPrevista  instanceof Date ? Utilities.formatDate(row[CONFIG.COL.DATA_CRIACAO  - 1], Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
        dataPrevista:  dataPrevista  instanceof Date ? Utilities.formatDate(dataPrevista,  Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
        dataConclusao: dataConclusao instanceof Date ? Utilities.formatDate(dataConclusao, Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
        observacoes:   row[CONFIG.COL.OBSERVACOES   - 1],
        numAlertas:    row[CONFIG.COL.ALERTAS       - 1] || 0,
        diasRestantes: diasRestantes,
        vencida:       vencida,
        linha:         i + 1 // para edição direta
      });
    }

    // ---- Responsáveis ----
    var responsaveis = [];
    if (sheetR) {
      var dadosR = sheetR.getDataRange().getValues();
      for (var j = 1; j < dadosR.length; j++) {
        if (dadosR[j][0]) {
          responsaveis.push({
            nome:  dadosR[j][0],
            email: dadosR[j][1],
            setor: dadosR[j][2]
          });
        }
      }
    }

    // ---- Resumo por categoria ----
    var porCategoria = {};
    CATEGORIAS.forEach(function(cat) { porCategoria[cat] = { total:0, pendentes:0, andamento:0, concluidas:0, vencidas:0 }; });

    tarefas.forEach(function(t) {
      var cat = t.categoria;
      if (!porCategoria[cat]) porCategoria[cat] = { total:0, pendentes:0, andamento:0, concluidas:0, vencidas:0 };
      porCategoria[cat].total++;
      if (t.status === "Pendente")      porCategoria[cat].pendentes++;
      if (t.status === "Em Andamento")  porCategoria[cat].andamento++;
      if (t.status === "Concluída")     porCategoria[cat].concluidas++;
      if (t.vencida)                    porCategoria[cat].vencidas++;
    });

    // ---- Resumo por responsável ----
    var porResponsavel = {};
    tarefas.forEach(function(t) {
      var resp = t.responsavel || "Sem responsável";
      if (!porResponsavel[resp]) porResponsavel[resp] = { total:0, pendentes:0, andamento:0, concluidas:0, vencidas:0, alta:0 };
      porResponsavel[resp].total++;
      if (t.status === "Pendente")     porResponsavel[resp].pendentes++;
      if (t.status === "Em Andamento") porResponsavel[resp].andamento++;
      if (t.status === "Concluída")    porResponsavel[resp].concluidas++;
      if (t.vencida)                   porResponsavel[resp].vencidas++;
      if (t.prioridade === "Alta")     porResponsavel[resp].alta++;
    });

    return JSON.stringify({
      tarefas:         tarefas,
      responsaveis:    responsaveis,
      porCategoria:    porCategoria,
      porResponsavel:  porResponsavel,
      categorias:      CATEGORIAS,
      prioridades:     PRIORIDADES,
      statusList:      STATUS_LIST,
      totalTarefas:    tarefas.length,
      totalPendentes:  tarefas.filter(function(t){ return t.status === "Pendente"; }).length,
      totalAndamento:  tarefas.filter(function(t){ return t.status === "Em Andamento"; }).length,
      totalConcluidas: tarefas.filter(function(t){ return t.status === "Concluída"; }).length,
      totalVencidas:   tarefas.filter(function(t){ return t.vencida; }).length,
      totalAlta:       tarefas.filter(function(t){ return t.prioridade === "Alta" && t.status !== "Concluída"; }).length,
      dataAtualizacao: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    });

  } catch(e) {
    return JSON.stringify({ erro: e.message });
  }
}

// ============================================================
//  SALVAR TAREFA (nova ou edição)
// ============================================================
function salvarTarefa(dados) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.ABA_TAREFAS);
    if (!sheet) return { sucesso: false, mensagem: "Aba TAREFAS não encontrada." };

    var agora = new Date();

    // Converte string de data para Date
    function parseData(str) {
      if (!str) return "";
      // Aceita dd/MM/yyyy ou yyyy-MM-dd
      var partes = str.split(/[\/\-]/);
      if (partes.length !== 3) return "";
      if (str.indexOf("/") !== -1) {
        return new Date(Number(partes[2]), Number(partes[1])-1, Number(partes[0]));
      }
      return new Date(Number(partes[0]), Number(partes[1])-1, Number(partes[2]));
    }

    if (dados.linha && dados.linha > 1) {
      // EDIÇÃO — atualiza linha existente
      var row = dados.linha;
      sheet.getRange(row, CONFIG.COL.TITULO       ).setValue(dados.titulo       || "");
      sheet.getRange(row, CONFIG.COL.DESCRICAO    ).setValue(dados.descricao    || "");
      sheet.getRange(row, CONFIG.COL.RESPONSAVEL  ).setValue(dados.responsavel  || "");
      sheet.getRange(row, CONFIG.COL.CATEGORIA    ).setValue(dados.categoria    || "");
      sheet.getRange(row, CONFIG.COL.PRIORIDADE   ).setValue(dados.prioridade   || "");
      sheet.getRange(row, CONFIG.COL.STATUS       ).setValue(dados.status       || "Pendente");
      sheet.getRange(row, CONFIG.COL.DATA_PREVISTA).setValue(parseData(dados.dataPrevista));
      sheet.getRange(row, CONFIG.COL.OBSERVACOES  ).setValue(dados.observacoes  || "");
      if (dados.status === "Concluída") {
        sheet.getRange(row, CONFIG.COL.DATA_CONCLUSAO).setValue(agora);
      }
      aplicarFormatacaoLinha_(sheet, row, dados.prioridade, dados.status);
      return { sucesso: true, mensagem: "Tarefa atualizada com sucesso!" };

    } else {
      // NOVA TAREFA — gera ID e insere
      var novoId = gerarId_(sheet);
      var novaLinha = [
        novoId,
        dados.titulo       || "",
        dados.descricao    || "",
        dados.responsavel  || "",
        dados.categoria    || "",
        dados.prioridade   || "Média",
        dados.status       || "Pendente",
        agora,
        parseData(dados.dataPrevista),
        "",
        dados.observacoes  || "",
        0
      ];
      sheet.appendRow(novaLinha);
      var ultimaLinha = sheet.getLastRow();
      aplicarFormatacaoLinha_(sheet, ultimaLinha, dados.prioridade, dados.status);
      return { sucesso: true, mensagem: "Tarefa criada com sucesso! ID: " + novoId, id: novoId };
    }

  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

// ============================================================
//  EXCLUIR TAREFA
// ============================================================
function excluirTarefa(linha) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.ABA_TAREFAS);
    if (!sheet || linha < 2) return { sucesso: false, mensagem: "Linha inválida." };
    sheet.deleteRow(linha);
    return { sucesso: true, mensagem: "Tarefa excluída." };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

// ============================================================
//  ATUALIZAR STATUS RÁPIDO
// ============================================================
function atualizarStatus(linha, novoStatus) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.ABA_TAREFAS);
    if (!sheet || linha < 2) return { sucesso: false };
    sheet.getRange(linha, CONFIG.COL.STATUS).setValue(novoStatus);
    if (novoStatus === "Concluída") {
      sheet.getRange(linha, CONFIG.COL.DATA_CONCLUSAO).setValue(new Date());
    }
    var prioridade = sheet.getRange(linha, CONFIG.COL.PRIORIDADE).getValue();
    aplicarFormatacaoLinha_(sheet, linha, prioridade, novoStatus);
    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  }
}

// ============================================================
//  ALERTAS DE VENCIMENTO — pode ser acionado por trigger diário
// ============================================================
function enviarAlertasVencimento() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var sheet    = ss.getSheetByName(CONFIG.ABA_TAREFAS);
    var sheetR   = ss.getSheetByName(CONFIG.ABA_RESPONSAVEIS);
    if (!sheet) return;

    var dados    = sheet.getDataRange().getValues();
    var dadosR   = sheetR ? sheetR.getDataRange().getValues() : [];
    var hoje     = new Date();
    hoje.setHours(0,0,0,0);

    // Mapa email dos responsáveis
    var emailMap = {};
    for (var j = 1; j < dadosR.length; j++) {
      if (dadosR[j][0] && dadosR[j][1]) {
        emailMap[dadosR[j][0].toLowerCase()] = dadosR[j][1];
      }
    }

    var enviados = 0;
    var alertasPorResponsavel = {};

    for (var i = 1; i < dados.length; i++) {
      var row = dados[i];
      if (!row[0]) continue;

      var status       = row[CONFIG.COL.STATUS - 1];
      var dataPrevista = row[CONFIG.COL.DATA_PREVISTA - 1];
      var responsavel  = row[CONFIG.COL.RESPONSAVEL - 1];
      var titulo       = row[CONFIG.COL.TITULO - 1];
      var prioridade   = row[CONFIG.COL.PRIORIDADE - 1];

      if (status === "Concluída" || status === "Cancelada") continue;
      if (!(dataPrevista instanceof Date)) continue;

      var dp = new Date(dataPrevista);
      dp.setHours(0,0,0,0);
      var diasRestantes = Math.round((dp - hoje) / 86400000);

      // Alerta: vencida, hoje, ou em 1/2/3 dias (prioridade alta) ou 1/3 dias (média)
      var deveAlertar = false;
      if (diasRestantes < 0)  deveAlertar = true; // vencida
      if (diasRestantes === 0) deveAlertar = true; // vence hoje
      if (prioridade === "Alta"  && diasRestantes <= 3 && diasRestantes > 0) deveAlertar = true;
      if (prioridade === "Média" && diasRestantes <= 1 && diasRestantes > 0) deveAlertar = true;

      if (!deveAlertar) continue;

      if (!alertasPorResponsavel[responsavel]) alertasPorResponsavel[responsavel] = [];
      alertasPorResponsavel[responsavel].push({
        titulo: titulo,
        prioridade: prioridade,
        diasRestantes: diasRestantes,
        dataPrevista: Utilities.formatDate(dataPrevista, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        linha: i + 1
      });
    }

    // Envia um e-mail consolidado por responsável
    for (var resp in alertasPorResponsavel) {
      var email = emailMap[resp.toLowerCase()];
      if (!email) continue;

      var tarefasResp = alertasPorResponsavel[resp];
      var corpo = montarEmailAlerta_(resp, tarefasResp);

      MailApp.sendEmail({
        to:      email,
        cc:      CONFIG.EMAIL_REMETENTE,
        subject: "🔔 [Norte Tech] Alerta de Tarefas — " + tarefasResp.length + " tarefa(s) requer(em) atenção",
        htmlBody: corpo
      });

      // Incrementa contador de alertas
      tarefasResp.forEach(function(t) {
        var cel = sheet.getRange(t.linha, CONFIG.COL.ALERTAS);
        cel.setValue((cel.getValue() || 0) + 1);
      });

      enviados += tarefasResp.length;
    }

    var msg = enviados > 0
      ? "✅ " + enviados + " alerta(s) enviado(s) com sucesso!"
      : "ℹ️ Nenhuma tarefa requer alerta no momento.";

    try { SpreadsheetApp.getUi().alert(msg); } catch(e) { Logger.log(msg); }

  } catch(e) {
    Logger.log("Erro enviarAlertasVencimento: " + e.message);
  }
}

// ============================================================
//  TRIGGER DIÁRIO — configure via: Extensões > Apps Script > Acionadores
//  Função: enviarAlertasVencimento | Tipo: Temporizador por tempo > Dia
// ============================================================
function criarTriggerDiario() {
  // Remove triggers existentes da função
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "enviarAlertasVencimento") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("enviarAlertasVencimento")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  SpreadsheetApp.getUi().alert("✅ Trigger diário configurado para 08h!");
}

// ============================================================
//  LIMPAR TAREFAS CONCLUÍDAS (+30 DIAS)
// ============================================================
function limparConcluidas() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ABA_TAREFAS);
  if (!sheet) return;

  var dados = sheet.getDataRange().getValues();
  var hoje  = new Date();
  var removidas = 0;

  for (var i = dados.length - 1; i >= 1; i--) {
    var status    = dados[i][CONFIG.COL.STATUS - 1];
    var dtConc    = dados[i][CONFIG.COL.DATA_CONCLUSAO - 1];
    if (status === "Concluída" && dtConc instanceof Date) {
      var dias = (hoje - dtConc) / 86400000;
      if (dias > 30) {
        sheet.deleteRow(i + 1);
        removidas++;
      }
    }
  }

  SpreadsheetApp.getUi().alert("🗑️ " + removidas + " tarefa(s) concluídas há mais de 30 dias foram removidas.");
}

// ============================================================
//  HELPERS INTERNOS
// ============================================================
function gerarId_(sheet) {
  var dados = sheet.getDataRange().getValues();
  var max   = 0;
  for (var i = 1; i < dados.length; i++) {
    var id = parseInt(dados[i][0]);
    if (!isNaN(id) && id > max) max = id;
  }
  return max + 1;
}

function formatarCabecalho_(sheet) {
  var range = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 12);
  range.setBackground("#0a1628")
       .setFontColor("#f0f6ff")
       .setFontWeight("bold")
       .setFontSize(10);
  sheet.setFrozenRows(1);
}

function aplicarFormatacaoLinha_(sheet, linha, prioridade, status) {
  var range = sheet.getRange(linha, 1, 1, 12);
  var cor   = "#ffffff";
  if (status === "Concluída")    cor = "#f0fff4";
  else if (status === "Cancelada") cor = "#f5f5f5";
  else if (prioridade === "Alta")  cor = "#fff5f5";
  else if (prioridade === "Média") cor = "#fffbf0";
  range.setBackground(cor);
}

function montarEmailAlerta_(responsavel, tarefas) {
  var html = '<div style="font-family:Sora,Arial,sans-serif;max-width:620px;margin:0 auto;background:#0a1628;color:#f0f6ff;border-radius:12px;overflow:hidden">';
  html += '<div style="background:linear-gradient(135deg,#1a56db,#e4c62a);padding:24px 28px;display:flex;align-items:center;gap:12px">';
  html += '<img src="https://compliance.nortesistech.com/static/compliance/img/logo%20norte%20full.png" height="36" style="object-fit:contain"/>';
  html += '<div><div style="font-size:1.1rem;font-weight:700;color:#fff">Gestão de Tarefas</div>';
  html += '<div style="font-size:0.75rem;color:rgba(255,255,255,0.8)">Alerta de Vencimento — Norte Tech</div></div></div>';
  html += '<div style="padding:24px 28px">';
  html += '<p style="font-size:0.9rem;color:#8da5c6;margin-bottom:20px">Olá <strong style="color:#f0f6ff">' + responsavel + '</strong>, você possui <strong style="color:#E4C62A">' + tarefas.length + '</strong> tarefa(s) que requerem atenção:</p>';
  html += '<table style="width:100%;border-collapse:collapse">';
  html += '<thead><tr style="background:rgba(255,255,255,0.05)">';
  html += '<th style="padding:10px 12px;text-align:left;font-size:0.7rem;color:#8da5c6;text-transform:uppercase;letter-spacing:.1em">Tarefa</th>';
  html += '<th style="padding:10px 12px;text-align:left;font-size:0.7rem;color:#8da5c6;text-transform:uppercase;letter-spacing:.1em">Prioridade</th>';
  html += '<th style="padding:10px 12px;text-align:left;font-size:0.7rem;color:#8da5c6;text-transform:uppercase;letter-spacing:.1em">Prazo</th>';
  html += '<th style="padding:10px 12px;text-align:left;font-size:0.7rem;color:#8da5c6;text-transform:uppercase;letter-spacing:.1em">Situação</th>';
  html += '</tr></thead><tbody>';

  tarefas.forEach(function(t, idx) {
    var bg       = idx % 2 === 0 ? "rgba(255,255,255,0.02)" : "transparent";
    var corPrio  = t.prioridade === "Alta" ? "#ef4444" : t.prioridade === "Média" ? "#eab308" : "#22c55e";
    var sitLabel = t.diasRestantes < 0  ? "VENCIDA (" + Math.abs(t.diasRestantes) + "d)" :
                   t.diasRestantes === 0 ? "VENCE HOJE" :
                   "Vence em " + t.diasRestantes + "d";
    var sitCor   = t.diasRestantes < 0  ? "#ef4444" :
                   t.diasRestantes === 0 ? "#eab308" : "#3b82f6";

    html += '<tr style="background:' + bg + ';border-bottom:1px solid rgba(255,255,255,0.05)">';
    html += '<td style="padding:10px 12px;font-size:0.82rem">' + t.titulo + '</td>';
    html += '<td style="padding:10px 12px"><span style="background:' + corPrio + '22;color:' + corPrio + ';font-size:0.7rem;font-weight:700;padding:3px 8px;border-radius:99px">' + t.prioridade.toUpperCase() + '</span></td>';
    html += '<td style="padding:10px 12px;font-size:0.78rem;font-family:monospace;color:#8da5c6">' + t.dataPrevista + '</td>';
    html += '<td style="padding:10px 12px;font-size:0.75rem;font-weight:700;color:' + sitCor + '">' + sitLabel + '</td>';
    html += '</tr>';
  });

  html += '</tbody></table>';
  html += '<p style="margin-top:20px;font-size:0.8rem;color:#8da5c6">Acesse a planilha de gestão para atualizar o status das suas tarefas.</p>';
  html += '</div>';
  html += '<div style="padding:16px 28px;background:rgba(255,255,255,0.03);border-top:1px solid rgba(59,130,246,.18);font-size:0.7rem;color:#8da5c6;display:flex;justify-content:space-between">';
  html += '<span>Norte Tech Logística</span><span>Enviado automaticamente pelo sistema de gestão de tarefas</span>';
  html += '</div></div>';
  return html;
}
