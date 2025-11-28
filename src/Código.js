/**
 * ============================================================================
 * 🏢 STUDIO AFROSTYLE MANAGER - V8 (DASHBOARD & CONFIG)
 * ============================================================================
 */

const APP_NAME = "Studio AfroStyle Manager";

const DB_SHEETS = {
  CLIENTES: "CLIENTES", SERVICOS: "SERVICOS", PRODUTOS: "PRODUTOS",
  AGENDAMENTOS: "AGENDAMENTOS", PEDIDOS: "PEDIDOS", FINANCEIRO: "FINANCEIRO",
  CONFIG: "CONFIG", INSUMOS: "INSUMOS", FORMAS_PAGAMENTO: "FORMAS_PAGAMENTO",
  TAXAS_ENTREGA: "TAXAS_ENTREGA", ENTREGADORES: "ENTREGADORES", USUARIOS: "USUARIOS"
};

// ... (MANTENHA AS CONSTANTES DE STATUS AQUI) ...
const STATUS_PEDIDO = { RECEBIDO: "RECEBIDO", EM_PREPARO: "EM PREPARO", SAIU_ENTREGA: "SAIU PARA ENTREGA", CONCLUIDO: "CONCLUIDO", CANCELADO: "CANCELADO" };
const STATUS_AGENDA = { PENDENTE: "PENDENTE", CONFIRMADO: "CONFIRMADO", FINALIZADO: "FINALIZADO", CANCELADO: "CANCELADO" };

// 1. SETUP & CORE
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appName = getConf(ss, 'NOME_EMPRESA', 'Smart Manager');
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(appName).addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🌟 Admin')
    .addItem('Abrir Painel Lateral', 'showSidebar')
    .addItem('Atualizar Estrutura DB', 'configurarPlanilha')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('Gestão Admin').setWidth(400);
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

// --- NOVO: SALVAR NOME DA EMPRESA (CONFIG) ---
// --- CORREÇÃO: SALVAR NOME DA EMPRESA (BLINDADO) ---
function setAppName(newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_SHEETS.CONFIG);

  // Garante que lê tudo como string para comparar
  const data = sheet.getDataRange().getDisplayValues();
  let found = false;

  for (let i = 0; i < data.length; i++) {
    // Compara removendo espaços e ignorando maiúsculas/minúsculas por segurança
    if (String(data[i][0]).trim().toUpperCase() === 'NOME_EMPRESA') {
      sheet.getRange(i + 1, 2).setValue(newName);
      found = true;
      break;
    }
  }

  // Se não achou, cria uma nova linha
  if (!found) {
    sheet.appendRow(['NOME_EMPRESA', newName, 'Nome exibido no App']);
  }

  // Força atualização rápida do cache do Google
  SpreadsheetApp.flush();

  return { success: true };
}

// --- DASHBOARD ANALYTICS ---
function getDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_SHEETS.FINANCEIRO);
  const appName = getConf(ss, 'NOME_EMPRESA', 'Minha Loja'); // Pega nome atual

  if (!sheet) return { totalMonth: 0, totalToday: 0, byType: { 'Vendas': 0, 'Serviços': 0 }, appName: appName };

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  const todayStr = now.toDateString();

  let totalMonth = 0; let totalToday = 0; let byType = { 'Vendas': 0, 'Serviços': 0 };

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][1]);
    const desc = String(data[i][3]);
    const val = Number(data[i][4]);

    if (rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
      totalMonth += val;
      if (rowDate.toDateString() === todayStr) totalToday += val;
      if (desc.includes('Pedido')) byType['Vendas'] += val; else byType['Serviços'] += val;
    }
  }
  return { totalMonth, totalToday, byType, appName: appName };
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

// --- NOVO: LOGIN REAL ---
function verificarLoginAdmin(email, senha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_SHEETS.USUARIOS);
  // Schema USUARIOS: ID, Email, Senha, Nivel
  const data = sheet.getDataRange().getDisplayValues();

  // Procura usuário (case insensitive para email)
  const user = data.slice(1).find(r =>
    String(r[1]).trim().toLowerCase() === String(email).trim().toLowerCase() &&
    String(r[2]).trim() === String(senha).trim()
  );

  if (user) return { success: true, nivel: user[3] };
  return { success: false, message: "Credenciais inválidas." };
}

// 3. API PÚBLICA
function getCatalogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toJson = (name) => sheetToJSON(ss.getSheetByName(name));
  const filterActive = (list) => list.filter(i => String(i.Ativo).toLowerCase() === 'true');

  return {
    produtos: filterActive(toJson(DB_SHEETS.PRODUTOS))
      .filter(p => parseMoney(p.Estoque_Atual) > 0)
      .map(p => ({ ...p, Preco: parseMoney(p.Preco) })),
    servicos: filterActive(toJson(DB_SHEETS.SERVICOS))
      .map(s => ({ ...s, Preco: parseMoney(s.Preco), Duracao_Minutos: parseMoney(s.Duracao_Minutos) })),
    pagamentos: filterActive(toJson(DB_SHEETS.FORMAS_PAGAMENTO)),
    taxasEntrega: filterActive(toJson(DB_SHEETS.TAXAS_ENTREGA)).map(t => ({ ...t, Valor_Taxa: parseMoney(t.Valor_Taxa) })),
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
      const hist = [{ status: STATUS_AGENDA.PENDENTE, data: new Date(), obs: "Solicitado via App" }];

      // MUDANÇA: NÃO CRIA EVENTO NO CALENDAR AQUI MAIS. SÓ SALVA PENDENTE.
      const evtId = "";

      sheet.appendRow([id, new Date(), payload.data, payload.horaInicio, payload.horaFim, JSON.stringify(payload.cliente), JSON.stringify(payload.itens), parseMoney(payload.total), STATUS_AGENDA.PENDENTE, evtId, payload.tipoAtendimento, payload.endereco, payload.pagamento, JSON.stringify(hist)]);

      return { success: true, id: id, whatsappLoja: getConf(ss, 'WHATSAPP_LOJA', '') };
    } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
  } else { return { success: false, message: "Servidor Ocupado" }; }
}

// 5. KDS PRO (DADOS COMPLETOS)
function getKDSData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidos = sheetToJSON(ss.getSheetByName(DB_SHEETS.PEDIDOS)).filter(p => p.Status !== STATUS_PEDIDO.CONCLUIDO && p.Status !== STATUS_PEDIDO.CANCELADO);
  const agendamentos = sheetToJSON(ss.getSheetByName(DB_SHEETS.AGENDAMENTOS)).filter(a => a.Status !== STATUS_AGENDA.FINALIZADO && a.Status !== STATUS_AGENDA.CANCELADO);
  const entregadores = sheetToJSON(ss.getSheetByName(DB_SHEETS.ENTREGADORES));
  return { pedidos, agendamentos, entregadores };
}

// 6. ATUALIZAÇÃO DE STATUS COM HISTÓRICO
function updateStatusKDS(type, id, newStatus, userLog) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(type === 'PEDIDO' ? DB_SHEETS.PEDIDOS : DB_SHEETS.AGENDAMENTOS);
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      let rowData = null;
      for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { rowIndex = i + 1; rowData = data[i]; break; } }
      if (rowIndex === -1) return { success: false, message: "ID não encontrado" };

      const headers = data[0];
      const statusIdx = headers.indexOf("Status");
      const histIdx = headers.indexOf("Historico_Json");
      const eventIdx = headers.indexOf("ID_Evento_Calendar"); // Para salvar o ID do Google Agenda

      // Mapeia dados da linha para objeto útil
      const map = {}; headers.forEach((h, i) => map[h] = rowData[i]);

      // Log Histórico
      if (histIdx > -1) {
        let h = []; try { h = JSON.parse(rowData[histIdx] || "[]"); } catch (e) { }
        h.push({ status: newStatus, data: new Date(), obs: `Por ${userLog || 'Admin'}` });
        sheet.getRange(rowIndex, histIdx + 1).setValue(JSON.stringify(h));
      }
      sheet.getRange(rowIndex, statusIdx + 1).setValue(newStatus);

      // === REGRA DE NEGÓCIO: CONFIRMAÇÃO DE AGENDAMENTO ===
      let whatsappLink = null;

      if (type === 'AGENDAMENTO' && newStatus === STATUS_AGENDA.CONFIRMADO) {
        // 1. Cria evento no Google Calendar AGORA
        const currentEventId = eventIdx > -1 ? rowData[eventIdx] : "";
        if (!currentEventId) {
          const evtId = criarEventoCalendar(ss, map); // Passa o objeto mapeado
          if (evtId && eventIdx > -1) sheet.getRange(rowIndex, eventIdx + 1).setValue(evtId);
        }

        // 2. Gera Link WhatsApp para o Cliente (Confirmando)
        try {
          const cli = JSON.parse(map.Cliente_Json);
          const itens = JSON.parse(map.Itens_Json);
          const nomesServicos = itens.map(i => i.Nome).join(', ');
          const dataFormatada = new Date(map.Data_Agendada || map.Data).toLocaleDateString('pt-BR'); // Ajuste conforme formato da data salva

          const msg = `Olá ${cli.nome}! Temos ótimas notícias! ✨\n\n` +
            `Seu agendamento foi *CONFIRMADO* com sucesso! ✅\n\n` +
            `🗓 *Data:* ${map.Data_Agendada} às ${map.Hora_Inicio}\n` +
            `✂️ *Serviço:* ${nomesServicos}\n` +
            `📍 *Local:* ${map.Tipo_Atendimento === 'DOMICILIO' ? 'Em Domicílio' : 'No Studio'}\n\n` +
            `Já estamos preparando tudo para te receber com todo carinho. Até lá! 💖`;

          whatsappLink = `https://wa.me/${cli.telefone.replace(/\D/g, '')}?text=${encodeURIComponent(msg)}`;
        } catch (e) { Logger.log("Erro montando zap: " + e); }
      }

      // Regra: Finalização = Baixa de Estoque + Financeiro
      if (newStatus === STATUS_PEDIDO.CONCLUIDO || newStatus === STATUS_AGENDA.FINALIZADO) {
        executarBaixaEstoqueEFinanceiro(ss, type, rowData, headers);
      }

      return { success: true, whatsappLink: whatsappLink };
    } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
  }
}

function criarEventoCalendar(ss, map) {
  try {
    const calendarId = getConf(ss, 'CALENDAR_ID', 'primary');
    // Ajuste de data: Se vier YYYY-MM-DD ou Date object
    let ano, mes, dia;
    if (map.Data_Agendada instanceof Date) {
      ano = map.Data_Agendada.getFullYear(); mes = map.Data_Agendada.getMonth() + 1; dia = map.Data_Agendada.getDate();
    } else {
      [ano, mes, dia] = map.Data_Agendada.split('-').map(Number);
    }

    const [hI, mI] = map.Hora_Inicio.split(':').map(Number);
    const [hF, mF] = map.Hora_Fim.split(':').map(Number);

    const start = new Date(ano, mes - 1, dia, hI, mI);
    const end = new Date(ano, mes - 1, dia, hF, mF);

    const cli = JSON.parse(map.Cliente_Json);
    const itens = JSON.parse(map.Itens_Json);

    const calendar = CalendarApp.getCalendarById(calendarId);
    if (calendar) {
      const event = calendar.createEvent(`💇‍♀️ ${cli.nome} - AfroStyle`, start, end, {
        description: `Tel: ${cli.telefone}\nServiços: ${itens.map(i => i.Nome).join(', ')}`,
        location: map.Tipo_Atendimento === 'DOMICILIO' ? map.Endereco_Domicilio : "No Studio"
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Força o fuso horário de SP para evitar confusão de servidor
  const timeZone = ss.getSpreadsheetTimeZone();

  const calId = getConf(ss, "CALENDAR_ID", "primary");
  const dias = getConf(ss, "DIAS_FUNCIONAMENTO", "1,2,3,4,5,6").split(',').map(Number);

  // Parse da data vinda do input (YYYY-MM-DD)
  const [ano, mes, dia] = dataStr.split('-').map(Number);
  const dataAlvo = new Date(ano, mes - 1, dia, 0, 0, 0);

  if (!dias.includes(dataAlvo.getDay())) return [];

  const [hA, mA] = getConf(ss, "HORARIO_ABERTURA", "09:00").split(':').map(Number);
  const [hF, mF] = getConf(ss, "HORARIO_FECHAMENTO", "19:00").split(':').map(Number);

  const ini = new Date(dataAlvo); ini.setHours(hA, mA, 0, 0);
  const fim = new Date(dataAlvo); fim.setHours(hF, mF, 0, 0);

  // Agora real no fuso horário da planilha
  const agoraStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
  const agora = new Date(agoraStr);

  // Margem de 30min
  const margem = new Date(agora.getTime() + (30 * 60000));

  // Se a data alvo for passado (ontem), retorna vazio
  const hojeZeroStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  const hojeZero = new Date(hojeZeroStr + " 00:00:00"); // Zera hora

  // Comparação de datas simples
  if (dataAlvo.getTime() < hojeZero.getTime()) return [];

  let busy = [];
  try {
    const c = CalendarApp.getCalendarById(calId);
    if (c) busy = c.getEvents(ini, fim).map(e => ({ s: e.getStartTime().getTime(), e: e.getEndTime().getTime() }));
  } catch (e) { }

  let slots = [];
  let cur = new Date(ini);

  while (cur.getTime() + (dur * 60000) <= fim.getTime()) {
    const s = cur.getTime();
    const e = s + (dur * 60000);

    // Verifica se o slot é no futuro
    const ehFuturo = s > margem.getTime();
    const livre = !busy.some(b => (s < b.e && e > b.s));

    if (livre && ehFuturo) {
      slots.push(Utilities.formatDate(cur, timeZone, "HH:mm"));
    }
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
// 8. CRUD ADMIN (CORREÇÃO DE TELA BRANCA APLICADA AQUI)
function crudGetTableData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { headers: [], items: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return { headers: [], items: [] };

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const items = lastRow < 2 ? [] : data.slice(1).map(row => {
    let obj = {}; row.forEach((c, i) => { if (headers[i]) obj[headers[i]] = c; }); return obj;
  });

  let extraData = null;
  if (sheetName === DB_SHEETS.SERVICOS) {
    const si = ss.getSheetByName(DB_SHEETS.INSUMOS);
    if (si && si.getLastRow() > 1) {
      const id = si.getDataRange().getDisplayValues();
      extraData = id.slice(1).map(r => ({ ID: r[0], Nome: r[1], Unidade: r[2] }));
    }
  }
  return { headers: headers, items: items, extraData: extraData };
}
function crudSaveItem(n, o) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); const h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; if (!o.ID) o.ID = Utilities.getUuid().slice(0, 8); const d = s.getDataRange().getValues(); let idx = -1; for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(o.ID)) { idx = i + 1; break; } } const r = h.map(k => o[k] === undefined ? "" : o[k]); if (idx > 0) s.getRange(idx, 1, 1, r.length).setValues([r]); else s.appendRow(r); return { success: true }; }
function crudDeleteItem(n, id) { const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); const d = s.getDataRange().getValues(); for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(id)) { s.deleteRow(i + 1); return { success: true }; } } return { success: false }; }