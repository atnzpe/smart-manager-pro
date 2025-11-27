/**
 * ============================================================================
 * 🚀 SMART MANAGER PRO - V1.0 (COMMERCIAL RELEASE)
 * ============================================================================
 */

// --- CONFIGURAÇÃO CENTRAL ---
const APP_VERSION = "1.0.0";
const DEFAULT_NAME = "Smart Manager";

const DB_SHEETS = {
  CLIENTES: "CLIENTES", SERVICOS: "SERVICOS", PRODUTOS: "PRODUTOS",
  AGENDAMENTOS: "AGENDAMENTOS", PEDIDOS: "PEDIDOS", FINANCEIRO: "FINANCEIRO",
  CONFIG: "CONFIG", INSUMOS: "INSUMOS", FORMAS_PAGAMENTO: "FORMAS_PAGAMENTO",
  TAXAS_ENTREGA: "TAXAS_ENTREGA", ENTREGADORES: "ENTREGADORES", USUARIOS: "USUARIOS"
};

const STATUS_PEDIDO = { RECEBIDO: "RECEBIDO", EM_PREPARO: "EM PREPARO", SAIU_ENTREGA: "SAIU PARA ENTREGA", CONCLUIDO: "CONCLUIDO", CANCELADO: "CANCELADO" };
const STATUS_AGENDA = { PENDENTE: "PENDENTE", CONFIRMADO: "CONFIRMADO", FINALIZADO: "FINALIZADO", CANCELADO: "CANCELADO" };

// 1. CORE
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appName = getConf(ss, 'NOME_EMPRESA', DEFAULT_NAME);
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(appName).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Admin Sistema')
    .addItem('Abrir Painel Gestão', 'showSidebar')
    .addItem('🔧 Instalar/Resetar Banco de Dados', 'configurarPlanilha')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('Painel Admin').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 2. HELPERS
function getConf(ss, key, def) {
  const sheet = ss.getSheetByName(DB_SHEETS.CONFIG);
  if (!sheet) return def;
  const data = sheet.getDataRange().getDisplayValues();
  const row = data.find(r => r[0] === key);
  return row ? String(row[1]) : def;
}

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
    let obj = {}; row.forEach((cell, i) => obj[headers[i]] = cell); return obj;
  });
}

// 3. AUTH (Login Real)
function verificarLoginAdmin(email, senha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_SHEETS.USUARIOS);
  const data = sheet.getDataRange().getDisplayValues();
  const user = data.slice(1).find(r => String(r[1]).trim().toLowerCase() === String(email).trim().toLowerCase() && String(r[2]).trim() === String(senha).trim());
  if (user) return { success: true, nivel: user[3] };
  return { success: false, message: "Acesso Negado." };
}

// 4. API FRONTEND
function getCatalogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toJson = (n) => sheetToJSON(ss.getSheetByName(n));
  const active = (l) => l.filter(i => String(i.Ativo).toLowerCase() === 'true');

  return {
    appName: getConf(ss, 'NOME_EMPRESA', DEFAULT_NAME),
    produtos: active(toJson(DB_SHEETS.PRODUTOS)).filter(p => parseMoney(p.Estoque_Atual) > 0).map(p => ({ ...p, Preco: parseMoney(p.Preco) })),
    servicos: active(toJson(DB_SHEETS.SERVICOS)).map(s => ({ ...s, Preco: parseMoney(s.Preco), Duracao_Minutos: parseMoney(s.Duracao_Minutos) })),
    pagamentos: active(toJson(DB_SHEETS.FORMAS_PAGAMENTO)),
    taxasEntrega: active(toJson(DB_SHEETS.TAXAS_ENTREGA)).map(t => ({ ...t, Valor_Taxa: parseMoney(t.Valor_Taxa) })),
    config: { whatsappLoja: getConf(ss, 'WHATSAPP_LOJA', '') }
  };
}

// 5. TRANSAÇÕES
function criarPedidoProduto(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(DB_SHEETS.PEDIDOS);
      const id = "PED-" + Utilities.getUuid().slice(0, 6).toUpperCase();
      const sub = parseMoney(payload.subtotal);
      const tax = parseMoney(payload.taxaEntrega);
      const hist = [{ status: STATUS_PEDIDO.RECEBIDO, data: new Date(), obs: "Via App" }];

      sheet.appendRow([id, new Date(), JSON.stringify(payload.cliente), JSON.stringify(payload.itens), sub, tax, sub + tax, STATUS_PEDIDO.RECEBIDO, payload.pagamento, payload.obs || "", JSON.stringify(hist)]);
      return { success: true, id: id, whatsappLoja: payload.whatsappLoja };
    } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
  }
}

function criarAgendamentoServico(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(DB_SHEETS.AGENDAMENTOS);
      const id = "AGD-" + Utilities.getUuid().slice(0, 6).toUpperCase();
      const hist = [{ status: STATUS_AGENDA.PENDENTE, data: new Date(), obs: "Via App" }];

      // Calendar Integration
      let evtId = "";
      try {
        const calId = getConf(ss, 'CALENDAR_ID', 'primary');
        const [a, m, d] = payload.data.split('-').map(Number);
        const [hI, mI] = payload.horaInicio.split(':').map(Number);
        const [hF, mF] = payload.horaFim.split(':').map(Number);
        const cal = CalendarApp.getCalendarById(calId);
        if (cal) evtId = cal.createEvent(`💇‍♀️ ${payload.cliente.nome}`, new Date(a, m - 1, d, hI, mI), new Date(a, m - 1, d, hF, mF)).getId();
      } catch (e) { }

      sheet.appendRow([id, new Date(), payload.data, payload.horaInicio, payload.horaFim, JSON.stringify(payload.cliente), JSON.stringify(payload.itens), parseMoney(payload.total), STATUS_AGENDA.CONFIRMADO, evtId, payload.tipoAtendimento, payload.endereco, payload.pagamento, JSON.stringify(hist)]);
      return { success: true, id: id };
    } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
  } else { return { success: false, message: "Ocupado" }; }
}

// 6. KDS & STATUS
function getKDSData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidos = sheetToJSON(ss.getSheetByName(DB_SHEETS.PEDIDOS)).filter(p => p.Status !== STATUS_PEDIDO.CONCLUIDO && p.Status !== STATUS_PEDIDO.CANCELADO);
  const agendamentos = sheetToJSON(ss.getSheetByName(DB_SHEETS.AGENDAMENTOS)).filter(a => a.Status !== STATUS_AGENDA.FINALIZADO && a.Status !== STATUS_AGENDA.CANCELADO);
  return { pedidos, agendamentos };
}

function updateStatusKDS(type, id, newStatus, user) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(type === 'PEDIDO' ? DB_SHEETS.PEDIDOS : DB_SHEETS.AGENDAMENTOS);
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { rowIndex = i + 1; break; } }
      if (rowIndex === -1) return { success: false, message: "404" };

      const headers = data[0];
      const statusIdx = headers.indexOf("Status");
      const histIdx = headers.indexOf("Historico_Json");

      // Update Log
      if (histIdx > -1) {
        let h = []; try { h = JSON.parse(data[rowIndex - 1][histIdx] || "[]"); } catch (e) { }
        h.push({ status: newStatus, data: new Date(), obs: `Por ${user}` });
        sheet.getRange(rowIndex, histIdx + 1).setValue(JSON.stringify(h));
      }
      sheet.getRange(rowIndex, statusIdx + 1).setValue(newStatus);

      // Finalização
      if (newStatus === STATUS_PEDIDO.CONCLUIDO || newStatus === STATUS_AGENDA.FINALIZADO) {
        const map = {}; headers.forEach((k, x) => map[k] = data[rowIndex - 1][x]);

        // Financeiro
        ss.getSheetByName(DB_SHEETS.FINANCEIRO).appendRow([Utilities.getUuid().slice(0, 8), new Date(), "RECEITA", `${type} #${id}`, type === 'PEDIDO' ? map.Total : map.Total_Valor, map.Forma_Pagamento, id]);

        // Estoque
        const itens = JSON.parse(map.Itens_Json || "[]");
        if (type === 'PEDIDO') {
          const sP = ss.getSheetByName(DB_SHEETS.PRODUTOS); const dP = sP.getDataRange().getValues();
          itens.forEach(it => { for (let r = 1; r < dP.length; r++) if (String(dP[r][0]) === String(it.ID)) sP.getRange(r + 1, 5).setValue(Number(dP[r][4]) - 1); });
        } else {
          const sS = ss.getSheetByName(DB_SHEETS.SERVICOS); const sI = ss.getSheetByName(DB_SHEETS.INSUMOS);
          const allS = sheetToJSON(sS); const dI = sI.getDataRange().getValues();
          itens.forEach(sv => {
            const s = allS.find(x => String(x.ID) === String(sv.ID));
            if (s && s.Ficha_Tecnica_Json) JSON.parse(s.Ficha_Tecnica_Json).forEach(ing => {
              for (let r = 1; r < dI.length; r++) if (String(dI[r][0]) === String(ing.id_insumo)) sI.getRange(r + 1, 5).setValue(Number(dI[r][4]) - Number(ing.qtd));
            });
          });
        }
      }
      return { success: true };
    } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
  }
}

function getHorariosDisponiveis(dataStr, dur) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const [a, m, d] = dataStr.split('-').map(Number);
  const data = new Date(a, m - 1, d);
  const dias = getConf(ss, "DIAS_FUNCIONAMENTO", "1,2,3,4,5,6").split(',').map(Number);
  if (!dias.includes(data.getDay())) return [];

  const [hA, mA] = getConf(ss, "HORARIO_ABERTURA", "09:00").split(':').map(Number);
  const [hF, mF] = getConf(ss, "HORARIO_FECHAMENTO", "19:00").split(':').map(Number);
  const ini = new Date(data); ini.setHours(hA, mA, 0);
  const fim = new Date(data); fim.setHours(hF, mF, 0);

  let busy = [];
  try {
    const cal = CalendarApp.getCalendarById(getConf(ss, "CALENDAR_ID", "primary"));
    if (cal) busy = cal.getEvents(ini, fim).map(e => ({ s: e.getStartTime().getTime(), e: e.getEndTime().getTime() }));
  } catch (e) { }

  let slots = []; let cur = new Date(ini);
  while (cur.getTime() + (dur * 60000) <= fim.getTime()) {
    const s = cur.getTime(), e = s + (dur * 60000);
    if (!busy.some(b => s < b.e && e > b.s)) slots.push(Utilities.formatDate(cur, Session.getScriptTimeZone(), "HH:mm"));
    cur.setMinutes(cur.getMinutes() + 30);
  }
  return slots;
}

// 7. SETUP
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = [
    { name: DB_SHEETS.CONFIG, headers: ["Chave", "Valor", "Descricao"] },
    { name: DB_SHEETS.USUARIOS, headers: ["ID", "Email", "Senha", "Nivel"] },
    { name: DB_SHEETS.CLIENTES, headers: ["ID", "Nome", "Telefone", "Email", "Data_Cadastro", "Obs", "Endereco_Padrao"] },
    { name: DB_SHEETS.PRODUTOS, headers: ["ID", "Nome", "Categoria", "Preco", "Estoque_Atual", "Foto_Url", "Ativo"] },
    { name: DB_SHEETS.SERVICOS, headers: ["ID", "Nome", "Categoria", "Preco", "Duracao_Minutos", "Foto_Url", "Ativo", "Ficha_Tecnica_Json"] },
    { name: DB_SHEETS.INSUMOS, headers: ["ID", "Nome", "Unidade", "Custo", "Estoque_Atual"] },
    { name: DB_SHEETS.PEDIDOS, headers: ["ID", "Data", "Cliente_Json", "Itens_Json", "Subtotal", "Taxa_Entrega", "Total", "Status", "Forma_Pagamento", "Obs", "Historico_Json"] },
    { name: DB_SHEETS.AGENDAMENTOS, headers: ["ID", "Data_Criacao", "Data_Agendada", "Hora_Inicio", "Hora_Fim", "Cliente_Json", "Itens_Json", "Total_Valor", "Status", "ID_Evento_Calendar", "Tipo_Atendimento", "Endereco_Domicilio", "Forma_Pagamento", "Historico_Json"] },
    { name: DB_SHEETS.FINANCEIRO, headers: ["ID", "Data", "Tipo", "Descricao", "Valor", "Forma_Pagamento", "Ref_ID"] },
    { name: DB_SHEETS.TAXAS_ENTREGA, headers: ["ID", "Nome_Regiao", "Valor_Taxa", "Tempo_Estimado_Min", "Ativo"] },
    { name: DB_SHEETS.FORMAS_PAGAMENTO, headers: ["ID", "Nome", "Instrucao", "Ativo"] },
    { name: DB_SHEETS.ENTREGADORES, headers: ["ID", "Nome", "Telefone", "Placa_Veiculo", "Ativo"] }
  ];

  schemas.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) { sheet = ss.insertSheet(s.name); sheet.appendRow(s.headers); }
    else {
      const cur = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      s.headers.forEach(h => { if (!cur.includes(h)) sheet.getRange(1, cur.length + 1).setValue(h); });
    }
  });

  const cfg = ss.getSheetByName(DB_SHEETS.CONFIG);
  const hasKey = (k) => cfg.getDataRange().getDisplayValues().some(r => r[0] === k);
  if (!hasKey('NOME_EMPRESA')) cfg.appendRow(['NOME_EMPRESA', 'Smart Manager', 'Nome do App']);
  if (!hasKey('WHATSAPP_LOJA')) cfg.appendRow(['WHATSAPP_LOJA', '5511999999999', 'DDD+Número']);

  const usr = ss.getSheetByName(DB_SHEETS.USUARIOS);
  if (usr.getLastRow() === 1) usr.appendRow([Utilities.getUuid().slice(0, 8), 'admin', 'admin', 'ADMIN']);
}

// CRUDs (Sidebar)
function crudGetTableData(n) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); if (!s) return { headers: [], items: [] }; const d = s.getDataRange().getDisplayValues(); if (d.length < 2) return { headers: [], items: [] }; let e = null; if (n === DB_SHEETS.SERVICOS) { const i = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DB_SHEETS.INSUMOS); if (i) { const id = i.getDataRange().getDisplayValues(); if (id.length > 1) e = id.slice(1).map(r => ({ ID: r[0], Nome: r[1], Unidade: r[2] })); } } const h = d[0]; const i = d.slice(1).map(r => { let o = {}; r.forEach((c, x) => o[h[x]] = c); return o; }); return { headers: h, items: i, extraData: e }; }
function crudSaveItem(n, o) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); const h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; if (!o.ID) o.ID = Utilities.getUuid().slice(0, 8); const d = s.getDataRange().getValues(); let idx = -1; for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(o.ID)) { idx = i + 1; break; } } const r = h.map(k => o[k] === undefined ? "" : o[k]); if (idx > 0) s.getRange(idx, 1, 1, r.length).setValues([r]); else s.appendRow(r); return { success: true }; }
function crudDeleteItem(n, id) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); const d = s.getDataRange().getValues(); for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(id)) { s.deleteRow(i + 1); return { success: true }; } } return { success: false }; }