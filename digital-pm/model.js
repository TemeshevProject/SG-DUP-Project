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
  const PROJECT_COST_VALUE = 26116061400;

  const CAUSE = {
    lateStart: "Поздно начали закупочную процедуру",
    shortWindow: "Окно плана короче процедуры",
    lateContract: "Договор заключили не вовремя",
    noPrepay: "Не внесли предоплату",
    noBalance: "Не оплатили остаток",
    contractor: "Срок сдвигает подрядчик",
    noDate: "Нет даты потребности",
    noBudget: "Лимит строки не утверждён",
  };
  const RESPONSE = {
    lateStart: "Подать заявку сегодня, зафиксировать новый прогноз и, если он позже вехи, вынести изменение даты потребности.",
    shortWindow: "Окно устава короче процедуры. Двигать веху, дробить объём или менять способ, если правила это позволяют.",
    lateContract: "Эскалация закупщику: ДУП держит дату заявки и дату договора, процедуру ведёт закупщик.",
    noPrepay: "Платёж аванса за 3 рабочих дня. Пока денег нет, отгрузка не начинается и прогноз поставки сдвигается.",
    noBalance: "Остаток после приёмки дату поставки не двигает. Поставить платёж в календарь. Если поставщик держит товар из-за долга, записать это отдельной причиной.",
    contractor: "Претензия подрядчику, новая дата в графике и проверка следующего пакета, который стоит на этой работе.",
    noDate: "Менеджер подтверждает дату потребности. Пока её нет, строка не входит в график и заявку не считают.",
    noBudget: "Утвердить лимит строки до подачи заявки. В уставе задана только стоимость всего проекта.",
  };

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

  function forecastFrom(startIso, line) {
    if (!startIso || !line.need || !STEPS[line.method]) return "";
    let cursor = parse(startIso);
    cursor = addWd(cursor, STEPS[line.method]);
    if (line.pay === "advance" || line.pay === "prepay") cursor = addWd(cursor, 3);
    return iso(addCal(cursor, line.lead || 0));
  }

  const LINES = [
    { id: "apk-cam", plan: "apk", name: "Камеры и контроллеры АПК и ОВН", who: "ОЗиС", method: "zcpComplex", pay: "advance", lead: 45, need: "2025-02-01", charterStart: "2024-10-01", state: "done", limit: 2400000000, source: "Лот поставки АПК" },
    { id: "apk-pole", plan: "apk", name: "Опоры, кронштейны и фундаменты", who: "Подряд СМР", method: "zcp", pay: "advance", lead: 30, need: "2025-01-20", charterStart: "2024-10-01", state: "done", limit: 640000000, source: "Лот строительной части АПК" },
    { id: "apk-cable", plan: "apk", name: "Кабель, муфты и кабельная канализация", who: "Подряд СМР", method: "zcp", pay: "advance", lead: 25, need: "2025-01-31", charterStart: "2024-10-05", state: "done", limit: 380000000, source: "Лот трассы" },
    { id: "apk-box", plan: "apk", name: "Шкафы уличного исполнения", who: "ОЗиС", method: "zcp", pay: "advance", lead: 20, need: "2025-02-15", charterStart: "2024-10-10", state: "done", limit: 210000000, source: "Лот шкафов" },
    { id: "patrol", plan: "patrol", name: "Мобильные комплексы Sergek Patrol", who: "Филиал / ST", method: "zcp", pay: "advance", lead: 40, need: "2025-04-28", charterStart: "2024-10-07", state: "done", limit: 1150000000, source: "Лот мобильного АПК" },
    { id: "patrol-link", plan: "patrol", name: "Связь и планшеты патрульных экипажей", who: "Филиал / ST", method: "zcp", pay: "advance", lead: 20, need: "2025-04-15", charterStart: "2024-11-01", state: "done", limit: 180000000, source: "Лот оснащения экипажей" },
    { id: "smr-apk", plan: "smr-apk", name: "СМР и ПНР стационарного АПК", who: "Подряд СМР", method: "zcpComplex", pay: "post", lead: 10, need: "2025-03-17", charterStart: "2024-10-20", state: "done", limit: 720000000, source: "Подряд на монтаж АПК" },
    { id: "smr-ovn", plan: "smr-ovn", name: "СМР и ПНР камер ОВН", who: "Подряд СМР", method: "zcpComplex", pay: "post", lead: 10, need: "2025-03-17", charterStart: "2024-10-20", state: "done", limit: 510000000, source: "Подряд на монтаж ОВН" },
    { id: "racks", plan: "dc", name: "Серверные стойки и организация рядов", who: "FT", method: "zcp", pay: "post", lead: 20, need: "2025-03-01", charterStart: "2024-11-15", state: "done", limit: 45000000, source: "Лот ЦОД" },
    { id: "twin", plan: "twin", name: "Данные для транспортного двойника", who: "ST", method: "oi", pay: "prepay", lead: 10, need: "2025-05-31", charterStart: "2025-01-05", state: "done", limit: 90000000, source: "Поставка данных двойника" },
    { id: "server", plan: "server", name: "Серверы вычислительного контура", who: "VMO FT", method: "zcp", pay: "advance", lead: 30, need: "2025-06-30", charterStart: "2024-10-23", state: "transit", limit: 890000000, source: "Лот серверов" },
    { id: "storage", plan: "server", name: "СХД и дисковые полки", who: "VMO FT", method: "zcp", pay: "advance", lead: 30, need: "2025-06-30", charterStart: "2024-11-01", state: "transit", limit: 760000000, source: "Лот хранения" },
    { id: "ups", plan: "dc", name: "ИБП и распределение питания", who: "FT", method: "zcp", pay: "advance", lead: 25, need: "2025-06-20", charterStart: "2024-12-01", state: "transit", limit: 150000000, source: "Лот электропитания ЦОД" },
    { id: "arm", plan: "cou", name: "АРМ операторов ЦОУ", who: "Филиал", method: "zcp", pay: "advance", lead: 15, need: "2025-06-25", charterStart: "2025-03-01", state: "awaitPay", limit: 96000000, source: "Договор есть, аванс не ушёл" },
    { id: "soft", plan: "soft", name: "ПО ЦОУ 2.0, лицензии и внедрение", who: "АзиаСофт", method: "oi", pay: "prepay", lead: 15, need: "2025-07-15", charterStart: "2025-05-15", state: "procedure", limit: 540000000, source: "Процедура одного источника" },
    { id: "switch", plan: "server", name: "Коммутаторы ядра и доступа", who: "VMO FT", method: "zcp", pay: "advance", lead: 20, need: "2025-07-01", charterStart: "2025-04-01", state: "procedure", limit: 240000000, source: "Запрос цен по сети" },
    { id: "insure", plan: "fin", name: "Страхование поставленного оборудования", who: "Фин. блок", method: "oi", pay: "prepay", lead: 10, need: "2025-08-15", charterStart: "2025-05-01", state: "procedure", limit: 55000000, source: "Страхование имущества проекта" },
    { id: "furn", plan: "cou", name: "Мебель ситуационного центра", who: "Филиал", method: "zcp", pay: "post", lead: 15, need: "2025-06-20", charterStart: "2025-02-01", state: "procedure", limit: 35000000, source: "Процедура затянулась дольше норматива" },
    { id: "link", plan: "link", name: "Магистральные каналы связи", who: "FT", method: "zcp", pay: "post", lead: 20, need: "2025-07-15", charterStart: "2024-11-01", state: "service", limit: 410000000, source: "Услуга связи, оплата после периода" },
    { id: "mile", plan: "cou-link", name: "Каналы последней мили, 9 районов", who: "FT", method: "zcp", pay: "post", lead: 20, need: "2025-07-15", charterStart: "2025-01-15", state: "service", limit: 270000000, source: "Районные каналы" },
    { id: "care", plan: "link", name: "Сопровождение каналов до ввода", who: "FT", method: "zcp", pay: "post", lead: 5, need: "2025-09-15", charterStart: "2025-05-01", state: "service", limit: 140000000, source: "Сервис до ввода в ПЭ" },
    { id: "video", plan: "video", name: "Лицензии видеоаналитики на камеры ОВН", who: "ОЗиС / ST", method: "zcp", pay: "advance", lead: 10, need: "2025-07-15", charterStart: "2025-06-20", state: "notStarted", limit: 680000000, cause: "Окно плана 20.06–15.07 короче процедуры. Заявку нужно было подать раньше, к 15.07.25 уже не успеть, если стартовать 20.06.", source: "Лот видеоаналитики" },
    { id: "wall", plan: "cou", name: "Видеостена ситуационного центра", who: "Филиал", method: "zcp", pay: "advance", lead: 20, need: "2025-07-10", charterStart: "2025-06-16", state: "notStarted", limit: 320000000, source: "Лот отображения ЦОУ" },
    { id: "scs", plan: "cou", name: "СКС, лотки и розетки ЦОУ", who: "Зерек", method: "zcp", pay: "post", lead: 10, need: "2025-06-30", charterStart: "2025-05-20", state: "notStarted", limit: 70000000, source: "Лот слаботочки ЦОУ" },
    { id: "zip", plan: "ops", name: "ЗИП на гарантийный год", who: "ОЗиС", method: "zcp", pay: "post", lead: 20, need: "2025-09-01", charterStart: "2025-07-01", state: "notStarted", limit: 60000000, source: "Запасные части после ввода" },
    { id: "train", plan: "cou", name: "Обучение смен операторов", who: "Филиал", method: "invoice", pay: "post", lead: 5, need: "2025-08-20", charterStart: "2025-07-01", state: "notStarted", limit: 18000000, source: "Подготовка персонала" },
    { id: "pnr", plan: "pnr", name: "Пусконаладка ИТ-контура", who: "Зерек", method: "zcp", pay: "post", lead: 10, need: "2025-06-30", charterStart: "2025-06-01", state: "notStarted", limit: 85000000, source: "ПНР ещё не начаты заявкой" },
    { id: "gts", plan: "gts", name: "Работы по требованиям информационной безопасности", who: "Филиал / подряд", method: "zcp", pay: "post", lead: 15, need: "2025-08-31", charterStart: "2025-06-01", state: "notStarted", limit: 120000000, source: "Контур ИБ" },
    { id: "dc-fit", plan: "dc", name: "Подготовка машинного зала", who: "FT", method: "zcpComplex", pay: "post", lead: 10, need: "2025-06-15", charterStart: "2025-05-22", state: "notStarted", limit: 260000000, source: "Строительная готовность ЦОД" },
    { id: "verify", plan: "verify", name: "Поверка стационарных АПК", who: "Подрядчик", method: "zcp", pay: "post", lead: 45, need: "2025-07-15", charterStart: "2025-04-28", state: "slip", slipCal: 25, limit: 95000000, cause: "Подрядчик сдвигает поверку на 25 календарных дней. Запуск штрафования, который в плане идёт следом, под угрозой.", source: "Поверка стационарных комплексов" },
    { id: "verify-p", plan: "verify", name: "Поверка мобильных комплексов", who: "Подрядчик", method: "zcp", pay: "post", lead: 30, need: "2025-08-01", charterStart: "2025-05-15", state: "slip", slipCal: 12, limit: 40000000, cause: "Подрядчик сдвигает поверку мобильных комплексов на 12 календарных дней.", source: "Поверка Sergek Patrol" },
    { id: "tool", plan: "smr-apk", name: "Инструмент и расходники монтажа", who: "Подряд СМР", method: "invoice", pay: "post", lead: 5, need: "2025-03-01", charterStart: "2024-12-01", state: "unpaid", limit: 25000000, source: "Товар принят, остаток по счёту не оплачен" },
  ];

  function analyze(line) {
    const today = parse(TODAY);
    if (!line.need) {
      return {
        ...line,
        methodTitle: METHOD[line.method] || "",
        norm: STEPS[line.method] || 0,
        requiredIso: "",
        forecastIso: "",
        scenarioToday: "",
        scenarioPlan: "",
        slip: 0,
        late: true,
        stage: "Нет даты потребности",
        contract: false,
        paid: false,
        delivered: false,
        causeCode: "noDate",
        causeTitle: CAUSE.noDate,
        shortWindow: false,
        response: RESPONSE.noDate,
        cause: line.cause || RESPONSE.noDate,
        limit: line.limit == null ? null : line.limit,
      };
    }
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
      forecast = addCal(need, line.slipCal || 0);
      stage = "Срок сдвигается";
      contract = true;
      paid = line.pay !== "post";
    } else if (line.state === "transit") {
      stage = "В пути";
      contract = true;
      paid = true;
    } else if (line.state === "awaitPay") {
      stage = "Ждёт предоплату";
      contract = true;
      paid = false;
      let cursor = addWd(today, 3);
      forecast = addCal(cursor, line.lead || 0);
    } else if (line.state === "unpaid") {
      stage = "Остаток не оплачен";
      contract = true;
      paid = false;
      delivered = true;
      forecast = need;
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
    const shortWindow = !!(line.charterStart && calBetween(parse(line.charterStart), need) < (STEPS[line.method] || 0));
    let causeCode = line.causeCode || "";
    if (!causeCode) {
      if (line.state === "slip" || line.slipCal) causeCode = "contractor";
      else if (line.state === "notStarted" && required < today) causeCode = "lateStart";
      else if (line.state === "notStarted" && shortWindow) causeCode = "shortWindow";
      else if (line.state === "procedure" && today > addWd(required, STEPS[line.method])) causeCode = "lateContract";
      else if ((line.pay === "advance" || line.pay === "prepay") && contract && !paid && !delivered) causeCode = "noPrepay";
      else if (line.pay === "post" && delivered && !paid) causeCode = "noBalance";
    }
    const response = causeCode === "lateStart" && shortWindow
      ? "Подать заявку сегодня и сразу вынести изменение вехи: даже старт в дату плана не попадает в потребность, окно короче процедуры."
      : (RESPONSE[causeCode] || "");
    return {
      ...line,
      methodTitle: METHOD[line.method],
      norm: STEPS[line.method],
      requiredIso: iso(required),
      forecastIso: iso(forecast),
      scenarioToday: forecastFrom(TODAY, line),
      scenarioPlan: forecastFrom(line.charterStart, line),
      slip,
      late: slip > 0,
      stage,
      contract,
      paid,
      delivered,
      causeCode,
      causeTitle: CAUSE[causeCode] || "",
      shortWindow,
      response,
      cause: line.cause || response,
      limit: line.limit == null || line.limit === "" ? null : Number(line.limit),
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

  function applyEdits(edits, drafts) {
    const lines = LINES.map((line) => {
      const edit = edits && edits[line.id];
      if (!edit) return line;
      const next = { ...line };
      ["need", "method", "pay", "state", "slipCal", "limit"].forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(edit, key) && edit[key] !== "") next[key] = edit[key];
      });
      return next;
    });
    (drafts || []).forEach((draft) => lines.push(draft));
    return lines;
  }

  function planTasks(edits) {
    const today = parse(TODAY);
    return PLAN.map((task) => {
      const linked = LINES.filter((line) => line.plan === task.id).map((line) => line.id);
      const edited = linked.map((id) => edits && edits[id] && edits[id].need).find(Boolean);
      const end = edited || task.end;
      return {
        ...task,
        end,
        baselineEnd: task.end,
        moved: end !== task.end,
        linked,
        done: parse(end) < today,
        started: parse(task.start) <= today,
        late: !!task.risk,
      };
    });
  }

  function rows() {
    return LINES.map(analyze);
  }

  function dueSoon(isoText, days) {
    if (!isoText) return false;
    return calBetween(parse(TODAY), parse(isoText)) <= days;
  }

  function queue(list) {
    const items = [];
    list.forEach((r) => {
      if (r.state === "notStarted" || r.causeCode === "noDate") {
        items.push({
          id: r.id,
          tone: r.late || r.causeCode === "noDate" ? "bad" : "",
          title: r.causeCode === "noDate" ? "Подтвердить дату" : "Подать заявку",
          who: "ДУП",
          text: r.name,
          due: r.requiredIso ? shortDate(r.requiredIso) : "даты нет",
          go: "zakupki/" + r.id,
        });
      } else if (r.causeCode === "noBalance") {
        items.push({
          id: r.id,
          tone: "bad",
          title: "Оплатить остаток",
          who: "Финблок",
          text: r.name,
          due: "после приёмки",
          go: "platezhi/" + r.id,
        });
      } else if (r.state === "procedure" || r.causeCode === "lateContract") {
        if (r.causeCode !== "lateContract" && !dueSoon(r.need, 21)) return;
        items.push({
          id: r.id,
          tone: r.causeCode === "lateContract" ? "bad" : "",
          title: "Довести до договора",
          who: "Закупщик",
          text: r.name,
          due: r.need ? shortDate(r.need) : "—",
          go: "dogovory/" + r.id,
        });
      } else if (r.causeCode === "noPrepay") {
        items.push({
          id: r.id,
          tone: "bad",
          title: "Внести предоплату",
          who: "Финблок",
          text: r.name,
          due: "3 р.д.",
          go: "platezhi/" + r.id,
        });
      } else if (r.causeCode === "contractor") {
        items.push({
          id: r.id,
          tone: "bad",
          title: "Претензия и новая дата",
          who: r.who,
          text: r.name,
          due: r.forecastIso ? shortDate(r.forecastIso) : "—",
          go: "risk/" + r.id,
        });
      } else if (r.state === "transit" && (r.late || dueSoon(r.need, 21))) {
        items.push({
          id: r.id,
          tone: "",
          title: "Принять поставку",
          who: r.who,
          text: r.name,
          due: shortDate(r.need),
          go: "zakupki/" + r.id,
        });
      } else if (r.state === "service" && !r.paid && (r.late || dueSoon(r.need, 21))) {
        items.push({
          id: r.id,
          tone: "",
          title: "Оплата после приёмки",
          who: "Финблок",
          text: r.name,
          due: shortDate(r.need),
          go: "platezhi/" + r.id,
        });
      }
    });
    return items;
  }

  function risks(list) {
    const items = [];
    list.forEach((r) => {
      if (!r.causeCode || r.state === "done") return;
      if (r.causeCode === "noBalance" && !r.delivered) return;
      const high = r.causeCode === "shortWindow" || r.causeCode === "contractor" || r.slip > 14 || r.causeCode === "noDate";
      items.push({
        id: r.id,
        level: high ? "высокий" : "средний",
        tone: high ? "red" : "gold",
        title: r.name,
        cause: r.causeTitle,
        text: r.cause || r.response,
        response: r.response,
        go: "zakupki/" + r.id,
      });
    });
    const verify = list.find((r) => r.id === "verify");
    if (verify && verify.late) {
      items.push({
        id: "cert",
        level: "высокий",
        tone: "red",
        title: "Сертификаты и запуск штрафования",
        cause: "Зависимость от поверки",
        text: "Поверка финиширует " + shortDate(verify.forecastIso) + " вместо " + shortDate(verify.need) + ". Штрафование стоит следом и сдвигается вместе с ней.",
        response: "Не запускать штрафование по старой дате. Сначала закрыть поверку, затем сдвинуть сертификаты и сообщить филиалу новую дату.",
        go: "proekt/cert",
      });
    }
    if (list.some((r) => r.limit == null && r.state !== "done")) {
      items.push({
        id: "budget",
        level: "средний",
        tone: "gold",
        title: "Бюджет закупок по строкам",
        cause: CAUSE.noBudget,
        text: "Потолок проекта " + PROJECT_COST + ". Суммы строк в уставе не выделены, лимиты пока не утверждены.",
        response: RESPONSE.noBudget,
        go: "budget",
      });
    }
    return items;
  }

  function sumLimits(rows) {
    return rows.reduce((acc, r) => acc + Number(r.limit), 0);
  }

  function budget(list) {
    const filled = list.filter((r) => r.limit != null && !isNaN(r.limit));
    const sum = sumLimits(filled);
    return {
      ceiling: PROJECT_COST_VALUE,
      ceilingText: PROJECT_COST,
      filled: filled.length,
      open: list.filter((r) => r.limit == null).length,
      sum,
      over: filled.length > 0 && sum > PROJECT_COST_VALUE,
    };
  }

  function money(list) {
    const valued = (pred) => list.filter((r) => r.limit != null && !isNaN(r.limit) && pred(r));
    const committed = valued((r) => r.contract);
    const paid = valued((r) => r.paid);
    const delivered = valued((r) => r.delivered);
    return {
      ceiling: PROJECT_COST_VALUE,
      lots: sumLimits(valued(() => true)),
      committed: sumLimits(committed),
      paid: sumLimits(paid),
      delivered: sumLimits(delivered),
      committedCount: committed.length,
      paidCount: paid.length,
      deliveredCount: delivered.length,
      daysToCommission: calBetween(parse(TODAY), parse(COMMISSION)),
    };
  }

  function notices(list) {
    const all = list || rows();
    const book = risks(all);
    const out = [];
    book.slice(0, 2).forEach((risk) => {
      out.push({ tone: risk.tone === "red" ? "red" : "gold", tag: risk.cause, text: risk.title + ". " + risk.text, go: risk.go });
    });
    const post = all.filter((r) => r.state === "service");
    if (post.length && out.length < 4) {
      out.push({
        tone: "blue",
        tag: "Финблок · платежи",
        text: "После приёмки к оплате: " + post.map((r) => r.name).join(", ") + ". Эта оплата дату потребности не двигает.",
        go: "platezhi/" + post[0].id,
      });
    }
    const done = all.filter((r) => r.state === "done");
    if (done.length && out.length < 4) {
      out.push({
        tone: "green",
        tag: "Выполнено",
        text: "Принято строк: " + done.length + ". " + done.map((r) => r.name).join("; ") + ".",
        go: "proekt/ppp",
      });
    }
    return out;
  }

  function control(edits, drafts) {
    const list = applyEdits(edits, drafts).map(analyze);
    return {
      list,
      plan: planTasks(edits),
      queue: queue(list),
      risks: risks(list),
      budget: budget(list),
      money: money(list),
      notices: notices(list),
      summary: summary(list),
    };
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
  function formatTenge(value) {
    if (value == null || isNaN(value)) return "не задан";
    return String(Math.round(value)).replace(/\B(?=(\d{3})+(?!\d))/g, " ") + " тенге";
  }
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
    control,
    queue,
    risks,
    budget,
    money,
    CAUSE,
    RESPONSE,
    METHOD,
    shortDate,
    pos,
    months,
    formatTenge,
    STEPS,
  };
});
