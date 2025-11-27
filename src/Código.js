/**
 * ============================================================================
 * 🚀 SMART MANAGER PRO - CORE CONTROLLER
 * Versão: 8.0.1 (White Label Release)
 * ============================================================================
 */

// --- CONFIGURAÇÃO ---
const DEFAULT_APP_NAME = "Smart Manager";

// Mapeamento das Abas
const DB_SHEETS = {
  CLIENTES: "CLIENTES", SERVICOS: "SERVICOS", PRODUTOS: "PRODUTOS",
  AGENDAMENTOS: "AGENDAMENTOS", PEDIDOS: "PEDIDOS", FINANCEIRO: "FINANCEIRO",
  CONFIG: "CONFIG", INSUMOS: "INSUMOS", FORMAS_PAGAMENTO: "FORMAS_PAGAMENTO",
  TAXAS_ENTREGA: "TAXAS_ENTREGA", ENTREGADORES: "ENTREGADORES", USUARIOS: "USUARIOS"
};

const STATUS_PEDIDO = { RECEBIDO: "RECEBIDO", EM_PREPARO: "EM PREPARO", SAIU_ENTREGA: "SAIU PARA ENTREGA", CONCLUIDO: "CONCLUIDO", CANCELADO: "CANCELADO" };
const STATUS_AGENDA = { PENDENTE: "PENDENTE", CONFIRMADO: "CONFIRMADO", FINALIZADO: "FINALIZADO", CANCELADO: "CANCELADO" };

// 1. SYSTEM CORE
function doGet(e) {
  // Busca o nome da empresa dinamicamente na aba CONFIG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appName = getConf(ss, 'NOME_EMPRESA', DEFAULT_APP_NAME);

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(appName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Admin Sistema')
    .addItem('📱 Abrir Painel Gestão', 'showSidebar')
    .addItem('🔄 Atualizar Banco de Dados', 'configurarPlanilha')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('Painel Admin').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 2. UTILS & SECURITY
function getConf(ss, key, def) {
  const sheet = ss.getSheetByName(DB_SHEETS.CONFIG);
  if (!sheet) return def;
  const data = sheet.getDataRange().getDisplayValues();
  const row = data.find(r => r[0] === key);
  return row ? String(row[1]) : def;
}

function verificarLoginAdmin(email, senha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_SHEETS.USUARIOS);
  const data = sheet.getDataRange().getDisplayValues();
  // Validação segura ignorando case no email
  const user = data.slice(1).find(r => String(r[1]).trim().toLowerCase() === String(email).trim().toLowerCase() && String(r[2]).trim() === String(senha).trim());
  
  if (user) return { success: true, nivel: user[3] };
  return { success: false, message: "Acesso Negado. Verifique suas credenciais." };
}

// 2. HELPERS & AUTH
function parseMoney(value) {
  if (!value) return 0;
  if (typeof value === 'number') return value;
  let str = String(value).replace("R$", "").replace(/\s/g, "").replace(",", ".");
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function sheetToJSON(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    row.forEach((cell, i) => obj[headers[i]] = cell);
    return obj;
  });
}

function getConf(ss, key, def) {
  const sheet = ss.getSheetByName(DB_SHEETS.CONFIG);
  if (!sheet) return def;
  const data = sheet.getDataRange().getDisplayValues();
  const row = data.find(r => r[0] === key);
  return row ? String(row[1]) : def;
}



/* Atualização na função getCatalogo para suportar White Label no Front */
function getCatalogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toJson = (n) => sheetToJSON(ss.getSheetByName(n));
  const active = (l) => l.filter(i => String(i.Ativo).toLowerCase() === 'true');

  return {
    appName: getConf(ss, 'NOME_EMPRESA', DEFAULT_APP_NAME), // Novo campo
    produtos: active(toJson(DB_SHEETS.PRODUTOS)).filter(p => parseMoney(p.Estoque_Atual) > 0).map(p => ({ ...p, Preco: parseMoney(p.Preco) })),
    servicos: active(toJson(DB_SHEETS.SERVICOS)).map(s => ({ ...s, Preco: parseMoney(s.Preco), Duracao_Minutos: parseMoney(s.Duracao_Minutos) })),
    pagamentos: active(toJson(DB_SHEETS.FORMAS_PAGAMENTO)),
    taxasEntrega: active(toJson(DB_SHEETS.TAXAS_ENTREGA)).map(t => ({ ...t, Valor_Taxa: parseMoney(t.Valor_Taxa) })),
    config: { whatsappLoja: getConf(ss, 'WHATSAPP_LOJA', '') }
  };
}

// 4. TRANSAÇÕES
function criarPedidoProduto(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(DB_SHEETS.PEDIDOS);
      const id = "PED-" + Utilities.getUuid().slice(0, 6).toUpperCase();
      const subtotal = parseMoney(payload.subtotal);
      const taxa = parseMoney(payload.taxaEntrega);

      // Histórico Inicial
      const hist = [{ status: STATUS_PEDIDO.RECEBIDO, data: new Date(), obs: "Pedido Criado via App" }];

      sheet.appendRow([
        id, new Date(), JSON.stringify(payload.cliente), JSON.stringify(payload.itens),
        subtotal, taxa, subtotal + taxa, STATUS_PEDIDO.RECEBIDO, payload.pagamento, payload.obs || "", JSON.stringify(hist)
      ]);
      return { success: true, id: id, whatsappLoja: payload.whatsappLoja };
    } catch (e) { return { success: false, message: e.message }; }
    finally { lock.releaseLock(); }
  }
}

function criarAgendamentoServico(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(DB_SHEETS.AGENDAMENTOS);
      const id = "AGD-" + Utilities.getUuid().slice(0, 6).toUpperCase();

      const hist = [{ status: STATUS_AGENDA.PENDENTE, data: new Date(), obs: "Solicitação via App" }];

      sheet.appendRow([
        id, new Date(), payload.data, payload.horaInicio, payload.horaFim,
        JSON.stringify(payload.cliente), JSON.stringify(payload.itens),
        parseMoney(payload.total), STATUS_AGENDA.PENDENTE, "",
        payload.tipoAtendimento, payload.endereco, payload.pagamento, JSON.stringify(hist)
      ]);

      return { success: true, id: id };
    } catch (e) { return { success: false, message: e.message }; }
    finally { lock.releaseLock(); }
  }
}

// 5. KDS PRO (DADOS COMPLETOS)
function getKDSData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Retorna TUDO (o frontend filtra)
  const pedidos = sheetToJSON(ss.getSheetByName(DB_SHEETS.PEDIDOS));
  const agendamentos = sheetToJSON(ss.getSheetByName(DB_SHEETS.AGENDAMENTOS));
  const entregadores = sheetToJSON(ss.getSheetByName(DB_SHEETS.ENTREGADORES));

  return { pedidos, agendamentos, entregadores };
}

// 6. ATUALIZAÇÃO DE STATUS COM HISTÓRICO
function updateStatusKDS(type, id, newStatus, userLog) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetName = type === 'PEDIDO' ? DB_SHEETS.PEDIDOS : DB_SHEETS.AGENDAMENTOS;
      const sheet = ss.getSheetByName(sheetName);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      let rowIndex = -1;
      let rowData = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
          rowIndex = i + 1;
          rowData = data[i];
          break;
        }
      }

      if (rowIndex === -1) return { success: false, message: "ID não encontrado." };

      // Índices dinâmicos
      const statusIdx = headers.indexOf("Status");
      const histIdx = headers.indexOf("Historico_Json");
      const eventIdx = headers.indexOf("ID_Evento_Calendar");

      // 1. Atualizar Histórico
      let historico = [];
      if (histIdx > -1) {
        try { historico = JSON.parse(rowData[histIdx] || "[]"); } catch (e) { }
        historico.push({
          status: newStatus,
          data: new Date(),
          obs: `Alterado por ${userLog || 'Admin'}`
        });
        sheet.getRange(rowIndex, histIdx + 1).setValue(JSON.stringify(historico));
      }

      // 2. Atualizar Status
      sheet.getRange(rowIndex, statusIdx + 1).setValue(newStatus);

      // 3. Regras de Negócio

      // Regra: Confirmar Agendamento -> Criar Calendar
      if (type === 'AGENDAMENTO' && newStatus === STATUS_AGENDA.CONFIRMADO) {
        // Se não tiver evento ainda, cria
        const currentEventId = eventIdx > -1 ? rowData[eventIdx] : "";
        if (!currentEventId) {
          const map = {}; headers.forEach((h, i) => map[h] = rowData[i]);
          const evtId = criarEventoCalendar(ss, map);
          if (evtId && eventIdx > -1) sheet.getRange(rowIndex, eventIdx + 1).setValue(evtId);
        }
      }

      // Regra: Baixa de Estoque/Financeiro
      if (newStatus === STATUS_PEDIDO.CONCLUIDO || newStatus === STATUS_AGENDA.FINALIZADO) {
        executarBaixaEstoqueEFinanceiro(ss, type, rowData, headers);
      }

      return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
    finally { lock.releaseLock(); }
  }
}

function criarEventoCalendar(ss, payload) {
  try {
    const calendarId = getConf(ss, 'CALENDAR_ID', 'primary');
    const [ano, mes, dia] = safeDateStr(payload.Data_Agendada || payload.Data);
    const [hI, mI] = payload.Hora_Inicio.split(':').map(Number);
    const [hF, mF] = payload.Hora_Fim.split(':').map(Number);

    const start = new Date(ano, mes - 1, dia, hI, mI);
    const end = new Date(ano, mes - 1, dia, hF, mF);

    const cli = JSON.parse(payload.Cliente_Json);
    const itens = JSON.parse(payload.Itens_Json);

    const calendar = CalendarApp.getCalendarById(calendarId);
    if (calendar) {
      const event = calendar.createEvent(`💇‍♀️ ${cli.nome}`, start, end, {
        description: `Tel: ${cli.telefone}\nServiços: ${itens.map(i => i.Nome).join(', ')}`,
        location: payload.Tipo_Atendimento === 'DOMICILIO' ? payload.Endereco_Domicilio : "No Studio"
      });
      return event.getId();
    }
  } catch (e) { Logger.log("Cal Error: " + e); }
  return "";
}

function safeDateStr(dateVal) {
  // Trata se vier objeto Date ou String "YYYY-MM-DD"
  if (dateVal instanceof Date) {
    return [dateVal.getFullYear(), dateVal.getMonth() + 1, dateVal.getDate()];
  }
  return String(dateVal).split('-').map(Number);
}

function executarBaixaEstoqueEFinanceiro(ss, type, rowData, headers) {
  const map = {}; headers.forEach((h, i) => map[h] = rowData[i]);

  // Financeiro
  const sheetFin = ss.getSheetByName(DB_SHEETS.FINANCEIRO);
  const val = type === 'PEDIDO' ? map.Total : map.Total_Valor;
  sheetFin.appendRow([Utilities.getUuid().slice(0, 8), new Date(), "RECEITA", `${type} #${map.ID}`, val, map.Forma_Pagamento, map.ID]);

  // Estoque
  const itens = JSON.parse(map.Itens_Json || "[]");
  if (type === 'PEDIDO') {
    const sheetProd = ss.getSheetByName(DB_SHEETS.PRODUTOS);
    const prodData = sheetProd.getDataRange().getValues();
    itens.forEach(item => {
      for (let r = 1; r < prodData.length; r++) {
        if (String(prodData[r][0]) === String(item.ID)) {
          const novo = Number(prodData[r][4]) - 1;
          sheetProd.getRange(r + 1, 5).setValue(novo);
          break;
        }
      }
    });
  } else {
    const sheetServ = ss.getSheetByName(DB_SHEETS.SERVICOS);
    const servAll = sheetToJSON(sheetServ);
    const sheetIns = ss.getSheetByName(DB_SHEETS.INSUMOS);
    const insData = sheetIns.getDataRange().getValues();
    itens.forEach(sv => {
      const full = servAll.find(s => String(s.ID) === String(sv.ID));
      if (full && full.Ficha_Tecnica_Json) {
        try {
          JSON.parse(full.Ficha_Tecnica_Json).forEach(ing => {
            for (let r = 1; r < insData.length; r++) {
              if (String(insData[r][0]) === String(ing.id_insumo)) {
                sheetIns.getRange(r + 1, 5).setValue(Number(insData[r][4]) - Number(ing.qtd));
                break;
              }
            }
          });
        } catch (e) { }
      }
    });
  }
}

function getHorariosDisponiveis(dataStr, dur) {
  // Mesma lógica de sempre
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calId = getConf(ss, "CALENDAR_ID", "primary");
  const dias = getConf(ss, "DIAS_FUNCIONAMENTO", "1,2,3,4,5,6").split(',').map(Number);
  const [ano, m, d] = dataStr.split('-').map(Number);
  const data = new Date(ano, m - 1, d);

  if (!dias.includes(data.getDay())) return [];

  const [hA, mA] = getConf(ss, "HORARIO_ABERTURA", "09:00").split(':').map(Number);
  const [hF, mF] = getConf(ss, "HORARIO_FECHAMENTO", "19:00").split(':').map(Number);

  const ini = new Date(data); ini.setHours(hA, mA, 0);
  const fim = new Date(data); fim.setHours(hF, mF, 0);

  let busy = [];
  try {
    const c = CalendarApp.getCalendarById(calId);
    if (c) busy = c.getEvents(ini, fim).map(e => ({ s: e.getStartTime().getTime(), e: e.getEndTime().getTime() }));
  } catch (e) { }

  let slots = [];
  const durMs = dur * 60000;
  let cur = new Date(ini);

  while (cur.getTime() + durMs <= fim.getTime()) {
    const s = cur.getTime(), e = s + durMs;
    if (!busy.some(b => s < b.e && e > b.s)) slots.push(Utilities.formatDate(cur, Session.getScriptTimeZone(), "HH:mm"));
    cur.setMinutes(cur.getMinutes() + 30);
  }
  return slots;
}

function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Adicionei Historico_Json nos Schemas
  const schemas = [
    { name: DB_SHEETS.TAXAS_ENTREGA, headers: ["ID", "Nome_Regiao", "Valor_Taxa", "Tempo_Estimado_Min", "Ativo"] },
    { name: DB_SHEETS.ENTREGADORES, headers: ["ID", "Nome", "Telefone", "Placa_Veiculo", "Ativo"] },
    { name: DB_SHEETS.FORMAS_PAGAMENTO, headers: ["ID", "Nome", "Instrucao", "Ativo"] },
    { name: DB_SHEETS.CONFIG, headers: ["Chave", "Valor", "Descricao"] },
    { name: DB_SHEETS.PEDIDOS, headers: ["ID", "Data", "Cliente_Json", "Itens_Json", "Subtotal", "Taxa_Entrega", "Total", "Status", "Forma_Pagamento", "Obs", "Historico_Json"] },
    { name: DB_SHEETS.AGENDAMENTOS, headers: ["ID", "Data_Criacao", "Data_Agendada", "Hora_Inicio", "Hora_Fim", "Cliente_Json", "Itens_Json", "Total_Valor", "Status", "ID_Evento_Calendar", "Tipo_Atendimento", "Endereco_Domicilio", "Forma_Pagamento", "Historico_Json"] },
    { name: DB_SHEETS.PRODUTOS, headers: ["ID", "Nome", "Categoria", "Preco", "Estoque_Atual", "Foto_Url", "Ativo"] },
    { name: DB_SHEETS.SERVICOS, headers: ["ID", "Nome", "Categoria", "Preco", "Duracao_Minutos", "Foto_Url", "Ativo", "Ficha_Tecnica_Json"] },
    { name: DB_SHEETS.CLIENTES, headers: ["ID", "Nome", "Telefone", "Email", "Data_Cadastro", "Obs", "Endereco_Padrao"] },
    { name: DB_SHEETS.INSUMOS, headers: ["ID", "Nome", "Unidade", "Custo", "Estoque_Atual"] },
    { name: DB_SHEETS.FINANCEIRO, headers: ["ID", "Data", "Tipo", "Descricao", "Valor", "Forma_Pagamento", "Ref_ID"] },
    { name: DB_SHEETS.USUARIOS, headers: ["ID", "Email", "Senha", "Nivel"] }
  ];

  schemas.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) { sheet = ss.insertSheet(s.name); sheet.appendRow(s.headers); }
    else {
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      s.headers.forEach(h => { if (!currentHeaders.includes(h)) sheet.getRange(1, currentHeaders.length + 1).setValue(h); });
    }
  });
}

// CRUD SideBar
function crudGetTableData(n) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); if (!s) return { headers: [], items: [] }; const d = s.getDataRange().getDisplayValues(); if (d.length < 2) return { headers: [], items: [] }; let e = null; if (n === DB_SHEETS.SERVICOS) { const i = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DB_SHEETS.INSUMOS); if (i) { const id = i.getDataRange().getDisplayValues(); if (id.length > 1) e = id.slice(1).map(r => ({ ID: r[0], Nome: r[1], Unidade: r[2] })); } } const h = d[0]; const i = d.slice(1).map(r => { let o = {}; r.forEach((c, x) => o[h[x]] = c); return o; }); return { headers: h, items: i, extraData: e }; }
function crudSaveItem(n, o) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); const h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; if (!o.ID) o.ID = Utilities.getUuid().slice(0, 8); const d = s.getDataRange().getValues(); let idx = -1; for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(o.ID)) { idx = i + 1; break; } } const r = h.map(k => o[k] === undefined ? "" : o[k]); if (idx > 0) s.getRange(idx, 1, 1, r.length).setValues([r]); else s.appendRow(r); return { success: true }; }
function crudDeleteItem(n, id) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); const d = s.getDataRange().getValues(); for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(id)) { s.deleteRow(i + 1); return { success: true }; } } return { success: false }; }