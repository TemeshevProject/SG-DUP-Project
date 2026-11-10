/* План закупок по графику устава 1.05 от 22.05.2025.
   Проект: услуги ИС ДОБ, Алматинская область. Ввод 15.09.2025.
   Срез 02.06.2025. Процедуры — Правила закупок, приказ №13-О/Д-ОД от 15.12.2025.
   Суммы строк в уставе не выделены: воронка считается строками. */
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.DigitalPm = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const TODAY = "2025-06-02";
  const COMMISSION = "2025-09-15";
  const CHART_START = "2024-09-01";
  const CHART_END = "2025-10-01";
  const PROJECT_COST = "26 116 061 400 тенге";

  const STEPS = {
    zcp: 35,
    zcpComplex: 40,
    oi: 23,
    invoice: 3,
  };
  const METHOD = {
    zcp: "Запрос цен · 35 р.д.",
    zcpComplex: "Запрос цен, сложное ТЗ · 40 р.д.",
    oi: "Один источник · 23 р.д.",
    invoice: "Счёт до 100 МРП · 3 р.д.",
  };

  function parse(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  function iso(d) {
    const z = (n) => String(n).padStart(2, "0");
    return d.getFullYear() + "-" + z(d.getMonth() + 1) + "-" + z(d.getDate());
  }
  function addCal(d, days) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate() + days);
  }
  function weekend(d) {
    const w = d.getDay();
    return w === 0 || w === 6;
  }
  function addWd(d, days) {
    const step = days < 0 ? -1 : 1;
    let left = Math.abs(days);
    let x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    while (left > 0) {
      x = addCal(x, step);
      if (!weekend(x)) left -= 1;
    }
    return x;
  }
  function calBetween(a, b) {
    const ua = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    const ub = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
    return Math.round((ub - ua) / 86400000);
  }
  function onOrBeforeWeekday(d) {
    let x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    while (weekend(x)) x = addCal(x, -1);
    return x;
  }

  function requestDate(line) {
    let d = parse(line.need);
    d = addCal(d, -line.lead);
    if (line.pay === "advance" || line.pay === "prepay") {
      d = onOrBeforeWeekday(d);
      d = addWd(d, -3);
    }
    d = addWd(d, -STEPS[line.method]);
    return d;
  }

  const LINES = [
    {
      id: "apk",
      name: "Поставка оборудования АПК и ОВН",
      who: "ОЗиС",
      method: "zcpComplex",
      pay: "advance",
      lead: 45,
      need: "2025-02-01",
      charterStart: "2024-10-01",
      state: "done",
      source: "План проекта, поставка АПК и ОВН",
    },
    {
      id: "patrol",
      name: "Мобильный АПК Sergek Patrol",
      who: "Филиал / ST",
      method: "zcp",
      pay: "advance",
      lead: 40,
      need: "2025-04-28",
      charterStart: "2024-10-07",
      state: "done",
      source: "План проекта: поставка и настройка до старта поверки",
    },
    {
      id: "server",
      name: "Серверное, сетевое и коммутационное оборудование",
      who: "VMO FT",
      method: "zcp",
      pay: "advance",
      lead: 30,
      need: "2025-06-30",
      charterStart: "2024-10-23",
      state: "transit",
      source: "План проекта, закуп серверного и сетевого оборудования",
    },
    {
      id: "link",
      name: "Каналы связи",
      who: "FT",
      method: "zcp",
      pay: "post",
      lead: 20,
      need: "2025-07-15",
      charterStart: "2024-11-01",
      state: "service",
      source: "План проекта, организация каналов связи",
    },
    {
      id: "soft",
      name: "ПО ЦОУ 2.0",
      who: "АзиаСофт",
      method: "oi",
      pay: "prepay",
      lead: 15,
      need: "2025-07-15",
      charterStart: "2025-05-15",
      state: "procedure",
      source: "План проекта, запуск ПО ЦОУ 2.0",
    },
    {
      id: "video",
      name: "Видеоаналитика на камеры ОВН",
      who: "ОЗиС / ST",
      method: "zcp",
      pay: "advance",
      lead: 10,
      need: "2025-07-15",
      charterStart: "2025-06-20",
      state: "notStarted",
      cause: "Окно плана 20.06–15.07 короче процедуры. Заявку нужно было подать раньше, к 15.07.25 уже не успеть, если стартовать 20.06.",
      source: "План проекта, внедрение видеоаналитики.",
    },
    {
      id: "verify",
      name: "Поверка АПК (Sergek Patrol)",
      who: "Подрядчик",
      method: "zcp",
      pay: "post",
      lead: 45,
      need: "2025-07-15",
      charterStart: "2025-04-28",
      state: "slip",
      slipCal: 25,
      cause: "Подрядчик сдвигает поверку на 25 календарных дней. Запуск штрафования, который в плане идёт следом, под угрозой.",
      source: "План проекта, поверка АПК.",
    },
  ];

  function analyze(line) {
    const today = parse(TODAY);
    const need = parse(line.need);
    const required = requestDate(line);
    let forecast = need;
    let stage = "В графике";
    let contract = false;
    let paid = false;
    let delivered = false;

    if (line.state === "done") {
      stage = "Принято";
      contract = true;
      paid = true;
      delivered = true;
    } else if (line.state === "notStarted") {
      let cursor = today;
      cursor = addWd(cursor, STEPS[line.method]);
      if (line.pay === "advance" || line.pay === "prepay") cursor = addWd(cursor, 3);
      forecast = addCal(cursor, line.lead);
      stage = "Заявку не подали";
    } else if (line.state === "slip") {
      forecast = addCal(need, line.slipCal);
      stage = "Срок сдвигается";
      contract = true;
      paid = line.pay !== "post";
    } else if (line.state === "transit") {
      stage = "В пути";
      contract = true;
      paid = true;
    } else if (line.state === "service") {
      stage = "Оказывается";
      contract = true;
      paid = false;
    } else if (line.state === "procedure") {
      stage = "Идёт процедура";
      contract = false;
      paid = false;
    }

    const slip = calBetween(need, forecast);
    return {
      ...line,
      methodTitle: METHOD[line.method],
      norm: STEPS[line.method],
      requiredIso: iso(required),
      forecastIso: iso(forecast),
      slip,
      late: slip > 0,
      stage,
      contract,
      paid,
      delivered,
      cause: line.cause || "",
    };
  }

  const PLAN = [
    { id: "fin", name: "Финансовое закрытие", who: "Фин. блок", start: "2024-10-01", end: "2024-12-31" },
    { id: "ppp", name: "Регистрация договора ГЧП", who: "ДРБ", start: "2024-09-29", end: "2024-10-14" },
    { id: "tz", name: "Подготовка и согласование ТЗ", who: "ДКС", start: "2024-10-10", end: "2024-12-10" },
    { id: "apk", name: "Поставка оборудования АПК, ОВН", who: "ОЗиС", start: "2024-10-01", end: "2025-02-01", zakup: "apk" },
    { id: "smr-apk", name: "СМР и ПНР АПК", who: "Подрядная организация", start: "2024-10-20", end: "2025-03-17" },
    { id: "smr-ovn", name: "СМР и ПНР ОВН", who: "Подрядная организация", start: "2024-10-20", end: "2025-03-17" },
    { id: "patrol", name: "Поставка и настройка мобильного АПК", who: "Филиал / ST", start: "2024-10-07", end: "2025-04-28", zakup: "patrol" },
    { id: "verify", name: "Поверка АПК (Sergek Patrol)", who: "Подрядная организация", start: "2025-04-28", end: "2025-07-15", risk: true, zakup: "verify" },
    { id: "twin", name: "Транспортный двойник города", who: "ST", start: "2025-01-05", end: "2025-05-31" },
    { id: "link", name: "Организация каналов связи", who: "FT", start: "2024-11-01", end: "2025-07-15", zakup: "link" },
    { id: "video", name: "Внедрение видеоаналитики на камеры ОВН", who: "ОЗиС / ST", start: "2025-06-20", end: "2025-07-15", risk: true, zakup: "video" },
    { id: "server", name: "Закуп серверного и сетевого оборудования", who: "VMO FT", start: "2024-10-23", end: "2025-06-30", zakup: "server" },
    { id: "cou", name: "Оснащение ЦОУ (9 районов и Алатау)", who: "Филиал", start: "2025-04-21", end: "2025-06-10" },
    { id: "cou-link", name: "Подключение ЦОУ к КС", who: "FT", start: "2025-05-15", end: "2025-07-15" },
    { id: "soft", name: "Запуск ПО ЦОУ 2.0", who: "АзиаСофт", start: "2025-05-15", end: "2025-07-15", zakup: "soft" },
    { id: "cert", name: "Сертификаты поверки и запуск штрафования", who: "Филиал", start: "2025-02-01", end: "2025-08-31", risk: true, riskFrom: "2025-04-28" },
    { id: "dc", name: "Организация проектного ЦОД", who: "FT", start: "2025-05-22", end: "2025-06-15" },
    { id: "pnr", name: "Пусконаладочные работы ИТ", who: "Зерек", start: "2025-06-01", end: "2025-06-30" },
    { id: "gts", name: "ГТС КНБ РК, требования ИБ", who: "Филиал / подряд", start: "2025-06-01", end: "2025-08-31" },
    { id: "act", name: "Подписание акта ввода в ПЭ", who: "Филиал", start: "2025-09-01", end: "2025-09-15" },
    { id: "ops", name: "Оказание услуг", who: "Филиал", start: "2025-09-15", end: "2030-05-15" },
  ];

  function planTasks() {
    const today = parse(TODAY);
    return PLAN.map((task) => ({
      ...task,
      done: parse(task.end) < today,
      started: parse(task.start) <= today,
      late: !!task.risk,
    }));
  }

  function rows() {
    return LINES.map(analyze);
  }

  function notices(list) {
    const all = list || rows();
    const video = all.find((r) => r.id === "video");
    const verify = all.find((r) => r.id === "verify");
    return [
      {
        tone: "red",
        tag: "Отклонение от плана",
        text: "Поверка АПК: финиш " + shortDate(verify.forecastIso) + " вместо " + shortDate(verify.need) + ". Под угрозой запуск штрафования.",
        go: "proekt/verify",
      },
      {
        tone: "gold",
        tag: "ОЗиС · план закупок",
        text: "Видеоаналитика ОВН: заявку не подали. Старт " + shortDate(TODAY) + " даёт приёмку " + shortDate(video.forecastIso) + ", веха " + shortDate(video.need) + ".",
        go: "zakupki/video",
      },
      {
        tone: "blue",
        tag: "Финблок · платежи",
        text: "После приёмки к оплате каналы связи и поверка АПК. Сдвиг поверки оплату не двигает: в договоре оплата после поставки.",
        go: "platezhi/verify",
      },
      {
        tone: "green",
        tag: "Выполнено",
        text: "Регистрация договора ГЧП закрыта 14.10.24. Поставка АПК и ОВН принята 01.02.25, мобильный АПК — 28.04.25.",
        go: "proekt/ppp",
      },
    ];
  }

  function summary(list) {
    const all = list || rows();
    return {
      total: all.length,
      owners: new Set(all.map((r) => r.who)).size,
      late: all.filter((r) => r.late).length,
      contract: all.filter((r) => r.contract).length,
      paid: all.filter((r) => r.paid).length,
      delivered: all.filter((r) => r.delivered).length,
    };
  }

  const MONTHS = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];
  function shortDate(isoText) {
    if (!isoText) return "—";
    const d = parse(isoText);
    const z = (n) => String(n).padStart(2, "0");
    return z(d.getDate()) + "." + z(d.getMonth() + 1) + "." + String(d.getFullYear()).slice(2);
  }
  function pos(isoText) {
    const t0 = parse(CHART_START).getTime();
    const t1 = parse(CHART_END).getTime();
    const t = parse(isoText).getTime();
    return Math.max(0, Math.min(100, ((t - t0) / (t1 - t0)) * 100));
  }
  function months() {
    const out = [];
    let d = parse(CHART_START);
    const end = parse(CHART_END);
    while (d < end) {
      out.push({ label: MONTHS[d.getMonth()], year: d.getFullYear(), start: iso(d) });
      d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }
    return out;
  }

  return {
    TODAY,
    COMMISSION,
    CHART_START,
    CHART_END,
    PROJECT_COST,
    rows,
    planTasks,
    notices,
    summary,
    shortDate,
    pos,
    months,
    STEPS,
  };
});
