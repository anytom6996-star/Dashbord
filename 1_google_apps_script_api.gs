/**
 * ═══════════════════════════════════════════════════════════════════
 * SALES DASHBOARD API — Google Apps Script v2
 * ═══════════════════════════════════════════════════════════════════
 *
 * СТРУКТУРА GOOGLE SHEETS:
 *
 *   📄 Budget
 *      A=Місяць | B=Рік | C=План New Rev (UAH) | D=План Corp Rev (UAH)
 *      E=План Проект | F=План Акаунт | G=Номер місяця
 *
 *   📄 Payments  ← сюди додаєш нові оплати щомісяця
 *      A=Дата | B=Менеджер | C=Сума (UAH) | D=Статус | E=Тип угоди | F=Нотатки
 *      Статус: paid / cancelled
 *      Тип угоди: new / renewal
 *
 *   📄 Managers
 *      A=Ім'я | B=Роль | C=Місяць старту | D=Рік старту | E=План/міс (UAH)
 *      Роль: "Менеджер" або "Керівник напрямку"
 *
 *   📄 KPI  ← дзвінки, зустрічі, угоди по менеджерах
 *      A=Менеджер | B=Місяць | C=Рік
 *      D=План дзвінки | E=Факт дзвінки
 *      F=Plan зустрічі | G=Факт зустрічі
 *      H=Plan угоди (шт) | I=Факт угоди (шт)
 *      J=Plan виручка (UAH) | K=Факт виручка (UAH)  ← необов'язково, рахується з Payments
 *
 * ЯК ЗАДЕПЛОЇТИ:
 *   1. Extensions → Apps Script → вставити цей код → Ctrl+S
 *   2. Запустити createSampleSheets() → створить листи з даними
 *   3. Deploy → New deployment → Web app
 *      Execute as: Me  |  Who has access: Anyone
 *   4. Скопіювати URL → вставити в index.html
 *
 * ЩО ТАКЕ PAYMENTS (Оплати):
 *   Це лист де ти фіксуєш кожну нову угоду/оплату від клієнта.
 *   Кожен рядок = одна оплата. Менеджер отримав гроші → додаєш рядок.
 *   Дата | Хто продав | Скільки грн | paid | new | (коментар)
 *
 * ═══════════════════════════════════════════════════════════════════
 */

const SHEET_BUDGET   = 'Budget';
const SHEET_PAYMENTS = 'Payments';
const SHEET_MANAGERS = 'Managers';
const SHEET_KPI      = 'KPI';

// ─── ENTRY POINT ─────────────────────────────────────────────────────
function doGet(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = buildDashboardData(ss);
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── MAIN ────────────────────────────────────────────────────────────
function buildDashboardData(ss) {
  const budget   = readBudget(ss);
  const payments = readPayments(ss);
  const managers = readManagers(ss);
  const kpiRaw   = readKPI(ss);

  const monthly  = aggregateMonthly(payments, budget);
  const mgrStats = aggregateManagers(payments, managers, budget);
  const forecast = buildForecast(monthly, budget, managers);
  const kpi      = buildKPI(kpiRaw, payments, managers);

  return {
    meta: {
      updatedAt:     new Date().toISOString(),
      spreadsheet:   ss.getName(),
      totalManagers: managers.filter(m => m.role === 'Менеджер').length,
      totalTeam:     managers.length,
    },
    budget,
    monthly,
    managers: mgrStats,
    forecast,
    kpi,
  };
}

// ─── BUDGET ──────────────────────────────────────────────────────────
function readBudget(ss) {
  const sh = ss.getSheetByName(SHEET_BUDGET);
  if (!sh) return defaultBudget();
  return sh.getDataRange().getValues().slice(1)
    .filter(r => r[0] && toNum(r[2]) > 0)
    .map(r => ({
      month:      String(r[0]).trim(),
      year:       toNum(r[1]),
      planNewRev: toNum(r[2]),   // UAH — єдиний бюджет що використовується
      monthNum:   r[6] ? toNum(r[6]) : getMonthNum(r[0]),
    }));
}

// ─── PAYMENTS ────────────────────────────────────────────────────────
// Payments = журнал оплат. Кожен рядок — одна закрита угода/оплата.
// Додаєш рядок коли клієнт заплатив. Статус "paid" = рахується, "cancelled" = ні.
function readPayments(ss) {
  const sh = ss.getSheetByName(SHEET_PAYMENTS);
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[2])
    .map(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return {
        date:    d,
        manager: String(r[1] || '').trim(),
        amount:  toNum(r[2]),
        status:  String(r[3] || 'paid').trim().toLowerCase(),
        type:    String(r[4] || 'new').trim(),
        note:    String(r[5] || ''),
        month:   d.getMonth() + 1,
        year:    d.getFullYear(),
      };
    });
}

// ─── MANAGERS ────────────────────────────────────────────────────────
// Роль "Менеджер" — продає, рахується у KPI по продажах
// Роль "Керівник напрямку" — управляє, окремий блок у звіті
// startMonth/startYear — з якого місяця рахувати план (важливо для нових!)
function readManagers(ss) {
  const sh = ss.getSheetByName(SHEET_MANAGERS);
  if (!sh) return defaultManagers();
  return sh.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({
      name:       String(r[0]).trim(),
      role:       String(r[1] || 'Менеджер').trim(),
      startMonth: toNum(r[2]) || 1,
      startYear:  toNum(r[3]) || 2026,
      planMonth:  toNum(r[4]) || 452717,
    }));
}

// ─── KPI RAW ─────────────────────────────────────────────────────────
function readKPI(ss) {
  const sh = ss.getSheetByName(SHEET_KPI);
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1)
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
      month:      b.month,
      year:       b.year,
      monthNum:   b.monthNum,
      planNewRev: b.planNewRev,
      factRev:    f,
      deals:      d,
      pct:        f != null ? Math.round(f / b.planNewRev * 1000) / 10 : null,
    };
  });
}

// ─── AGGREGATE MANAGERS ──────────────────────────────────────────────
// ВАЖЛИВО:
//   - Купельські О. і Є. стартували з березня → план рахується тільки з березня
//   - Настаченко: до березня — Менеджер (рахується у продажах),
//                 з березня — Керівник напрямку (не рахується у плані менеджерів)
//   - Цю логіку береться зі startMonth у листі Managers та поля role
function aggregateManagers(payments, managers, budget) {
  // Зібрати оплати по менеджерах
  const payMap = {};
  payments.forEach(p => {
    if (p.status === 'cancelled') return;
    if (!payMap[p.manager]) payMap[p.manager] = { byMonth: {}, total: 0, deals: 0 };
    const key = p.year + '-' + p.month;
    payMap[p.manager].byMonth[key] = (payMap[p.manager].byMonth[key] || 0) + p.amount;
    payMap[p.manager].total  += p.amount;
    payMap[p.manager].deals  += 1;
  });

  return managers.map(m => {
    const stats   = payMap[m.name] || { byMonth: {}, total: 0, deals: 0 };
    const monthly = Object.entries(stats.byMonth)
      .map(([k, v]) => { const [y, mo] = k.split('-').map(Number); return { year:y, month:mo, amount:v }; })
      .sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);

    // Активні місяці (коли реально стартував)
    const activeMonths = budget.filter(b =>
      b.year > m.startYear || (b.year === m.startYear && b.monthNum >= m.startMonth)
    );

    // Для керівника напрямку — план за місяці коли ще був менеджером
    // Настаченко: менеджер у лютому (startMonth=2), керівник з березня
    // Тому план рахуємо тільки за місяці до зміни ролі
    const isHead = m.role === 'Керівник напрямку';

    // Знаходимо місяць переходу (перший місяць у ролі керівника)
    // Якщо роль вже "Керівник напрямку" і є продажі до певного місяця —
    // вважаємо що перехід стався у місяць startMonth наступний після продажів
    // Спрощення: для керівника план = 0 починаючи з місяця переходу
    // Місяць переходу = перший місяць де немає оплат, або визначений вручну
    // Для Настаченко: продажі є у лютому та березні, далі — 0
    // Вважаємо що перейшов на керівника з квітня (березень — останній місяць продажів)
    let planMonths = 0;
    if (!isHead) {
      planMonths = activeMonths.length;
    } else {
      // Рахуємо місяці де є фактичні продажі (до переходу)
      planMonths = monthly.filter(x => x.amount > 0).length;
    }

    const totalPlan = m.planMonth * planMonths;

    // Тренд: порівняти останні два місяці з продажами
    const amounts = monthly.map(x => x.amount).filter(v => v > 0);
    const last    = amounts[amounts.length - 1] || 0;
    const prev    = amounts[amounts.length - 2] || 0;
    const trend   = amounts.length < 2 ? 'neutral'
                  : last > prev * 1.1  ? 'up'
                  : last < prev * 0.9  ? 'down' : 'neutral';

    return {
      name:       m.name,
      role:       m.role,
      isHead,
      startMonth: m.startMonth,
      startYear:  m.startYear,
      planMonth:  m.planMonth,
      totalPlan,
      total:      stats.total,
      deals:      stats.deals,
      monthly,
      trend,
      pct:        totalPlan > 0 ? Math.round(stats.total / totalPlan * 1000) / 10 : null,
      avgMonth:   monthly.filter(x => x.amount > 0).length > 0
        ? Math.round(stats.total / monthly.filter(x => x.amount > 0).length) : 0,
    };
  });
}

// ─── FORECAST ────────────────────────────────────────────────────────
// Прогноз рахується тільки по менеджерах (не керівниках)
// Для нових менеджерів (старт з березня) — враховуємо менше місяців плану
function buildForecast(monthly, budget, managers) {
  const completed = monthly.filter(m => m.factRev != null && m.factRev > 0);
  const sumFact   = completed.reduce((a, m) => a + m.factRev, 0);
  const sumPlan   = completed.reduce((a, m) => a + m.planNewRev, 0);
  const avgRate   = sumPlan > 0 ? sumFact / sumPlan : 0;

  const now         = new Date();
  const curMonth    = now.getMonth() + 1;
  const curYear     = now.getFullYear();
  const daysElapsed = now.getDate();
  const daysInMonth = new Date(curYear, curMonth, 0).getDate();
  const partialRate = daysElapsed / daysInMonth;

  const rows = monthly.map(m => {
    const isCurrent = m.monthNum === curMonth && m.year === curYear;
    let projected, type;
    if (m.factRev != null && !isCurrent) {
      projected = m.factRev; type = 'fact';
    } else if (isCurrent && m.factRev != null) {
      projected = Math.round(m.factRev / partialRate); type = 'partial';
    } else {
      projected = Math.round(m.planNewRev * avgRate); type = 'forecast';
    }
    return {
      month: m.month, monthNum: m.monthNum,
      plan: m.planNewRev, fact: m.factRev,
      projected, type,
      pct: Math.round(projected / m.planNewRev * 10) / 10,
    };
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

// ─── KPI BUILD ───────────────────────────────────────────────────────
// Будує KPI для ВСІХ менеджерів і керівника напрямку
// Факт угод (шт) і виручка (UAH) — рахується з Payments автоматично
// Дзвінки та зустрічі — вводяться вручну у лист KPI
function buildKPI(kpiRaw, payments, managers) {
  // Знаходимо місяці де є дані
  const monthsWithData = [...new Set(
    payments.filter(p => p.status !== 'cancelled').map(p => p.year + '-' + p.month)
  )].sort();

  const result = [];

  monthsWithData.forEach(key => {
    const [yr, mo] = key.split('-').map(Number);

    managers.forEach(m => {
      // Пропускаємо менеджерів що ще не стартували
      const started = yr > m.startYear || (yr === m.startYear && mo >= m.startMonth);
      if (!started) return;

      // Факт з Payments
      const mgrPays = payments.filter(p =>
        p.manager === m.name && p.year === yr && p.month === mo && p.status !== 'cancelled'
      );
      const factRev   = mgrPays.reduce((a, p) => a + p.amount, 0);
      const factDeals = mgrPays.length;

      // KPI з листа KPI (дзвінки, зустрічі)
      const kpiRow = kpiRaw.find(k => k.manager === m.name && k.month === mo && k.year === yr);

      // Для керівника напрямку — план 0 (не продає)
      const isHead   = m.role === 'Керівник напрямку';
      const planRev  = isHead ? 0 : m.planMonth;
      const planDeal = isHead ? 0 : (kpiRow?.planDeals || 0);

      result.push({
        manager:      m.name,
        role:         m.role,
        isHead,
        month:        mo,
        year:         yr,
        planRev,
        factRev,
        planDeals:    planDeal,
        factDeals,
        pctRev:       planRev > 0 ? Math.round(factRev / planRev * 1000) / 10 : null,
        pctDeals:     planDeal > 0 ? Math.round(factDeals / planDeal * 1000) / 10 : null,
        // З листа KPI:
        planCalls:    kpiRow?.planCalls    || 0,
        factCalls:    kpiRow?.factCalls    || 0,
        planMeetings: kpiRow?.planMeetings || 0,
        factMeetings: kpiRow?.factMeetings || 0,
      });
    });
  });

  return result;
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

function defaultBudget() {
  const names = ['Січень','Лютий','Березень','Квітень','Травень','Червень',
    'Липень','Серпень','Вересень','Жовтень','Листопад','Грудень'];
  const plans = [1388289,1358151,1339317,1483050,2390744,2108000,2318800,2523400,2796200,3069000,3410000,3726732];
  return names.map((m,i) => ({ month:m, year:2026, planNewRev:plans[i], monthNum:i+1 }));
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

// ─── CREATE SAMPLE SHEETS ────────────────────────────────────────────
function createSampleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Budget
  let sh = ss.getSheetByName(SHEET_BUDGET) || ss.insertSheet(SHEET_BUDGET);
  sh.clearContents();
  sh.getRange(1,1,1,7).setValues([['Місяць','Рік','План New Rev (UAH)','План Corp Rev (UAH)','План Проект','План Акаунт','Номер місяця']]);
  const bd = defaultBudget();
  sh.getRange(2,1,bd.length,7).setValues(bd.map(b=>[b.month,b.year,b.planNewRev,0,0,0,b.monthNum]));

  // Payments — з поясненням
  sh = ss.getSheetByName(SHEET_PAYMENTS) || ss.insertSheet(SHEET_PAYMENTS);
  sh.clearContents();
  sh.getRange(1,1,1,6).setValues([['Дата','Менеджер','Сума (UAH)','Статус (paid/cancelled)','Тип угоди (new/renewal)','Нотатки']]);
  const pays = [
    [new Date('2026-02-10'),'Мельничук',     417074,'paid','new',''],
    [new Date('2026-02-14'),'Онофрійчук',    191989,'paid','new',''],
    [new Date('2026-02-20'),'Настаченко',    485183,'paid','new',''],
    [new Date('2026-03-05'),'Настаченко',    400118,'paid','new','Останній місяць як менеджер'],
    [new Date('2026-03-12'),'Купельський О.',  74358,'paid','new','Перший місяць'],
    [new Date('2026-03-18'),'Купельський Є.',  16990,'paid','new','Перший місяць'],
    [new Date('2026-03-22'),'Мельничук',      36153,'paid','new',''],
    [new Date('2026-04-08'),'Купельський Є.', 487502,'paid','new',''],
    [new Date('2026-04-15'),'Купельський О.', 874884,'paid','new',''],
    [new Date('2026-05-07'),'Купельський Є.',  67764,'paid','new',''],
    [new Date('2026-05-12'),'Купельський О.', 254105,'paid','new',''],
    [new Date('2026-05-16'),'Мельничук',     110964,'paid','new',''],
    [new Date('2026-05-19'),'Онофрійчук',    205468,'paid','new',''],
  ];
  sh.getRange(2,1,pays.length,6).setValues(pays);

  // Managers
  sh = ss.getSheetByName(SHEET_MANAGERS) || ss.insertSheet(SHEET_MANAGERS);
  sh.clearContents();
  sh.getRange(1,1,1,5).setValues([["Ім'я","Роль","Місяць старту","Рік старту","План/міс (UAH)"]]);
  sh.getRange(2,1,5,5).setValues([
    ['Мельничук',     'Менеджер',          2,2026,452717],
    ['Онофрійчук',    'Менеджер',          2,2026,452717],
    ['Настаченко',    'Керівник напрямку', 2,2026,452717],
    ['Купельський О.','Менеджер',          3,2026,452717],
    ['Купельський Є.','Менеджер',          3,2026,452717],
  ]);

  // KPI — всі менеджери + керівник
  sh = ss.getSheetByName(SHEET_KPI) || ss.insertSheet(SHEET_KPI);
  sh.clearContents();
  sh.getRange(1,1,1,9).setValues([['Менеджер','Місяць','Рік','План дзвінки','Факт дзвінки','План зустрічі','Факт зустрічі','План угоди (шт)','Факт угоди (шт)']]);
  sh.getRange(2,1,5,9).setValues([
    ['Мельничук',      4,2026,700,0, 12,3,7,2],
    ['Онофрійчук',     4,2026,700,0, 12,0,7,0],
    ['Настаченко',     4,2026,0,0,   0,0,0,0],
    ['Купельський О.', 4,2026,700,168,12,7,7,2],
    ['Купельський Є.', 4,2026,700,262,12,2,7,0],
  ]);

  SpreadsheetApp.getUi().alert(
    '✅ Листи створено!\n\n' +
    'Наступний крок:\n' +
    'Deploy → New deployment → Web app\n' +
    'Execute as: Me\n' +
    'Who has access: Anyone\n\n' +
    'Потім скопіюй URL у дашборд.'
  );
}
