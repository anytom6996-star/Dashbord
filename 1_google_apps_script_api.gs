/**
 * ═══════════════════════════════════════════════════════════════════
 * SALES DASHBOARD API — Google Apps Script
 * ═══════════════════════════════════════════════════════════════════
 *
 * СТРУКТУРА GOOGLE SHEETS (потрібні такі листи):
 *   📄 Budget    — A=Місяць | B=Рік | C=План New Rev | D=План Corp Rev | G=№ місяця
 *   📄 Payments  — A=Дата | B=Менеджер | C=Сума | D=Статус | E=Тип угоди
 *   📄 Managers  — A=Ім'я | B=Роль | C=Місяць старту | D=Рік старту | E=План/міс
 *   📄 KPI       — A=Менеджер | B=Місяць | C=Рік | D=План дзв. | E=Факт дзв.
 *                  F=План зуст. | G=Факт зуст. | H=План угоди | I=Факт угоди
 *
 * ЯК ЗАДЕПЛОЇТИ:
 *   1. Extensions → Apps Script → вставити цей код → Зберегти
 *   2. Запустити createSampleSheets() — створить всі листи з прикладом
 *   3. Deploy → New deployment → Web app
 *      Execute as: Me  |  Who has access: Anyone
 *   4. Скопіювати URL → вставити в dashboard.html (поле Налаштування)
 * ═══════════════════════════════════════════════════════════════════
 */

const SHEET_BUDGET   = 'Budget';
const SHEET_PAYMENTS = 'Payments';
const SHEET_MANAGERS = 'Managers';
const SHEET_KPI      = 'KPI';

// ─── ENTRY POINT ─────────────────────────────────────────────────────
// ВАЖЛИВО: Google Apps Script Web App автоматично додає CORS-заголовки
// коли Access = "Anyone". Не потрібно вручну додавати — це робить Google.
function doGet(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = buildDashboardData(ss);
    // ContentService з MimeType.JSON = правильний CORS від Google
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── MAIN DATA BUILDER ───────────────────────────────────────────────
function buildDashboardData(ss) {
  const budget   = readBudget(ss);
  const payments = readPayments(ss);
  const managers = readManagers(ss);
  const kpi      = readKPI(ss);

  const monthly  = aggregateMonthly(payments, budget);
  const mgrStats = aggregateManagers(payments, managers);
  const forecast = buildForecast(monthly, budget);
  const funnel   = buildFunnel(payments);

  return {
    meta: {
      updatedAt:     new Date().toISOString(),
      spreadsheet:   ss.getName(),
      totalManagers: managers.length,
    },
    budget,
    monthly,
    managers: mgrStats,
    forecast,
    funnel,
    kpi,
  };
}

// ─── BUDGET READER ───────────────────────────────────────────────────
function readBudget(ss) {
  const sh = ss.getSheetByName(SHEET_BUDGET);
  if (!sh) return defaultBudget();
  const rows = sh.getDataRange().getValues().slice(1);
  return rows.map(r => ({
    month:       r[0],
    year:        toNum(r[1]),
    planNewRev:  toNum(r[2]),
    planCorpRev: toNum(r[3]),
    planProjRev: toNum(r[4]) || 0,
    planAccRev:  toNum(r[5]) || 0,
    monthNum:    r[6] ? toNum(r[6]) : getMonthNum(r[0]),
  })).filter(r => r.month && r.planNewRev > 0);
}

// ─── PAYMENTS READER ─────────────────────────────────────────────────
function readPayments(ss) {
  const sh = ss.getSheetByName(SHEET_PAYMENTS);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[0] && r[2])
    .map(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return {
        date:     d,
        manager:  String(r[1] || '').trim(),
        amount:   toNum(r[2]),
        status:   String(r[3] || 'paid').trim().toLowerCase(),
        dealType: String(r[4] || 'new').trim(),
        month:    d.getMonth() + 1,
        year:     d.getFullYear(),
      };
    });
}

// ─── MANAGERS READER ─────────────────────────────────────────────────
function readManagers(ss) {
  const sh = ss.getSheetByName(SHEET_MANAGERS);
  if (!sh) return defaultManagers();
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[0])
    .map(r => ({
      name:       String(r[0]).trim(),
      role:       String(r[1] || 'Менеджер').trim(),
      startMonth: toNum(r[2]) || 1,
      startYear:  toNum(r[3]) || 2026,
      planMonth:  toNum(r[4]) || 452717,
    }));
}

// ─── KPI READER ──────────────────────────────────────────────────────
function readKPI(ss) {
  const sh = ss.getSheetByName(SHEET_KPI);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[0])
    .map(r => ({
      manager:      String(r[0]).trim(),
      month:        toNum(r[1]),
      year:         toNum(r[2]),
      planCalls:    toNum(r[3]),
      factCalls:    toNum(r[4]),
      planMeetings: toNum(r[5]),
      factMeetings: toNum(r[6]),
      planDeals:    toNum(r[7]),
      factDeals:    toNum(r[8]),
    }));
}

// ─── AGGREGATE MONTHLY ───────────────────────────────────────────────
function aggregateMonthly(payments, budget) {
  const map = {};
  payments.forEach(p => {
    if (p.status === 'cancelled') return;
    const key = p.year + '-' + p.month;
    if (!map[key]) map[key] = { year: p.year, month: p.month, fact: 0, deals: 0 };
    map[key].fact  += p.amount;
    map[key].deals += 1;
  });

  return budget.map(b => {
    const key  = b.year + '-' + b.monthNum;
    const fact = map[key];
    const f    = fact ? fact.fact  : null;
    const d    = fact ? fact.deals : null;
    return {
      month:       b.month,
      year:        b.year,
      monthNum:    b.monthNum,
      planNewRev:  b.planNewRev,
      planCorpRev: b.planCorpRev,
      factRev:     f,
      deals:       d,
      pct:         f != null ? Math.round(f / b.planNewRev * 1000) / 10 : null,
    };
  });
}

// ─── AGGREGATE MANAGERS ──────────────────────────────────────────────
function aggregateManagers(payments, managers) {
  const map = {};
  payments.forEach(p => {
    if (p.status === 'cancelled' || !p.manager) return;
    if (!map[p.manager]) map[p.manager] = { byMonth: {}, total: 0, deals: 0 };
    const key = p.year + '-' + p.month;
    map[p.manager].byMonth[key] = (map[p.manager].byMonth[key] || 0) + p.amount;
    map[p.manager].total  += p.amount;
    map[p.manager].deals  += 1;
  });

  return managers.map(m => {
    const stats  = map[m.name] || { byMonth: {}, total: 0, deals: 0 };
    const monthly = Object.entries(stats.byMonth)
      .map(([k, v]) => { const [y, mo] = k.split('-').map(Number); return { year:y, month:mo, amount:v }; })
      .sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);

    const amounts = monthly.map(x => x.amount);
    const last    = amounts[amounts.length - 1] || 0;
    const prev    = amounts[amounts.length - 2] || 0;
    const trend   = amounts.length < 2 ? 'neutral'
                  : last > prev * 1.1  ? 'up'
                  : last < prev * 0.9  ? 'down' : 'neutral';

    return {
      ...m,
      total:    stats.total,
      deals:    stats.deals,
      monthly,
      trend,
      avgMonth: monthly.filter(x => x.amount > 0).length > 0
        ? Math.round(stats.total / monthly.filter(x => x.amount > 0).length)
        : 0,
    };
  });
}

// ─── FORECAST ────────────────────────────────────────────────────────
function buildForecast(monthly, budget) {
  const completed = monthly.filter(m => m.factRev !== null && m.factRev > 0);
  const sumFact   = completed.reduce((a, m) => a + m.factRev, 0);
  const sumPlan   = completed.reduce((a, m) => a + m.planNewRev, 0);
  const avgRate   = sumPlan > 0 ? sumFact / sumPlan : 0;

  const now          = new Date();
  const curMonth     = now.getMonth() + 1;
  const curYear      = now.getFullYear();
  const daysElapsed  = now.getDate();
  const daysInMonth  = new Date(curYear, curMonth, 0).getDate();
  const partialRate  = daysElapsed / daysInMonth;

  const rows = monthly.map(m => {
    const isCurrent = m.monthNum === curMonth && m.year === curYear;
    let projected, type;
    if (m.factRev !== null && !isCurrent) {
      projected = m.factRev; type = 'fact';
    } else if (isCurrent && m.factRev !== null) {
      projected = Math.round(m.factRev / partialRate); type = 'partial';
    } else {
      projected = Math.round(m.planNewRev * avgRate); type = 'forecast';
    }
    return { month: m.month, monthNum: m.monthNum, plan: m.planNewRev,
             fact: m.factRev, projected, type, pct: Math.round(projected / m.planNewRev * 10) / 10 };
  });

  const totalPlan      = budget.reduce((a, b) => a + b.planNewRev, 0);
  const totalProjected = rows.reduce((a, r) => a + r.projected, 0);

  return {
    rows,
    avgRate:        Math.round(avgRate * 1000) / 10,
    totalPlan,
    totalProjected,
    totalPct:       Math.round(totalProjected / totalPlan * 1000) / 10,
    scenarios: {
      current:    { projected: totalProjected,               pct: Math.round(totalProjected / totalPlan * 10) / 10 },
      optimistic: { projected: Math.round(totalPlan * 0.80), pct: 80  },
      target:     { projected: totalPlan,                    pct: 100 },
    }
  };
}

// ─── FUNNEL ──────────────────────────────────────────────────────────
function buildFunnel(payments) {
  const map = {};
  payments.forEach(p => {
    if (p.status === 'cancelled') return;
    const key = p.year + '-' + p.month;
    if (!map[key]) map[key] = { year: p.year, month: p.month, amount: 0, deals: 0 };
    map[key].amount += p.amount;
    map[key].deals  += 1;
  });
  return Object.values(map).sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);
}

// ─── HELPERS ─────────────────────────────────────────────────────────
function toNum(v) {
  const n = parseFloat(String(v).replace(/\s/g,'').replace(',','.').replace(/[^0-9.\-]/g,''));
  return isNaN(n) ? 0 : n;
}

function getMonthNum(name) {
  const map = { 'Січень':1,'Лютий':2,'Березень':3,'Квітень':4,'Травень':5,'Червень':6,
    'Липень':7,'Серпень':8,'Вересень':9,'Жовтень':10,'Листопад':11,'Грудень':12 };
  return map[String(name).trim()] || 0;
}

// ─── DEFAULT DATA ────────────────────────────────────────────────────
function defaultBudget() {
  const names = ['Січень','Лютий','Березень','Квітень','Травень','Червень',
    'Липень','Серпень','Вересень','Жовтень','Листопад','Грудень'];
  const plans = [1388289,1358151,1339317,1483050,2390744,2108000,2318800,2523400,2796200,3069000,3410000,3726732];
  const corp  = [21654441,21184333,20890563,20010925,19629328,20671409,23258422,31546033,31001417,28856826,36607854,36803369];
  return names.map((m,i) => ({ month:m, year:2026, planNewRev:plans[i], planCorpRev:corp[i], monthNum:i+1 }));
}

function defaultManagers() {
  return [
    { name:'Мельничук',      role:'Менеджер',          startMonth:2, startYear:2026, planMonth:452717 },
    { name:'Онофрійчук',     role:'Менеджер',          startMonth:2, startYear:2026, planMonth:452717 },
    { name:'Настаченко',     role:'Керівник напрямку', startMonth:2, startYear:2026, planMonth:452717 },
    { name:'Купельський О.', role:'Менеджер',          startMonth:3, startYear:2026, planMonth:452717 },
    { name:'Купельський Є.', role:'Менеджер',          startMonth:3, startYear:2026, planMonth:452717 },
  ];
}

// ─── CREATE SAMPLE SHEETS (запусти один раз!) ────────────────────────
function createSampleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Budget
  let sh = ss.getSheetByName(SHEET_BUDGET) || ss.insertSheet(SHEET_BUDGET);
  sh.clearContents();
  sh.getRange(1,1,1,7).setValues([['Місяць','Рік','План New Rev (UAH)','План Corp Rev (UAH)','План Проект','План Акаунт','Номер місяця']]);
  const bd = defaultBudget();
  sh.getRange(2,1,bd.length,7).setValues(bd.map(b => [b.month,b.year,b.planNewRev,b.planCorpRev,0,0,b.monthNum]));

  // ── Payments  (з реальними тестовими даними)
  sh = ss.getSheetByName(SHEET_PAYMENTS) || ss.insertSheet(SHEET_PAYMENTS);
  sh.clearContents();
  sh.getRange(1,1,1,6).setValues([['Дата','Менеджер','Сума (UAH)','Статус','Тип угоди','Нотатки']]);
  const payments = [
    [new Date('2026-02-10'),'Мельничук',     417074,'paid','new',''],
    [new Date('2026-02-14'),'Онофрійчук',    191989,'paid','new',''],
    [new Date('2026-02-20'),'Настаченко',    485183,'paid','new',''],
    [new Date('2026-03-05'),'Настаченко',    400118,'paid','new',''],
    [new Date('2026-03-12'),'Купельський О.',  74358,'paid','new','Старт'],
    [new Date('2026-03-18'),'Купельський Є.',  16990,'paid','new','Старт'],
    [new Date('2026-03-22'),'Мельничук',      36153,'paid','new',''],
    [new Date('2026-04-08'),'Купельський Є.', 487502,'paid','new',''],
    [new Date('2026-04-15'),'Купельський О.', 874884,'paid','new',''],
    [new Date('2026-05-07'),'Купельський Є.',  67764,'paid','new',''],
    [new Date('2026-05-12'),'Купельський О.', 254105,'paid','new',''],
    [new Date('2026-05-16'),'Мельничук',     110964,'paid','new',''],
    [new Date('2026-05-19'),'Онофрійчук',    205468,'paid','new',''],
  ];
  sh.getRange(2,1,payments.length,6).setValues(payments);

  // ── Managers
  sh = ss.getSheetByName(SHEET_MANAGERS) || ss.insertSheet(SHEET_MANAGERS);
  sh.clearContents();
  sh.getRange(1,1,1,5).setValues([["Ім'я","Роль","Місяць старту","Рік старту","План/міс (UAH)"]]);
  const mgrs = defaultManagers();
  sh.getRange(2,1,mgrs.length,5).setValues(mgrs.map(m => [m.name,m.role,m.startMonth,m.startYear,m.planMonth]));

  // ── KPI
  sh = ss.getSheetByName(SHEET_KPI) || ss.insertSheet(SHEET_KPI);
  sh.clearContents();
  sh.getRange(1,1,1,9).setValues([['Менеджер','Місяць','Рік','План дзвінки','Факт дзвінки','План зустрічі','Факт зустрічі','План угоди','Факт угоди']]);
  // Приклад даних квітня
  sh.getRange(2,1,2,9).setValues([
    ['Купельський О.',4,2026,700,168,12,7,7,2],
    ['Купельський Є.',4,2026,700,262,12,2,7,0],
  ]);

  SpreadsheetApp.getUi().alert('✅ Листи створено з тестовими даними!\n\nТепер зроби Deploy → New deployment → Web app\n(Execute as: Me · Who has access: Anyone)\nта скопіюй URL у дашборд.');
}
