/* Перечень ТРУ, обратный счёт от даты потребности и сценарии срыва.
   Сроки процедур — Правила закупок ТОО «Көркем Телеком»,
   приказ №13-О/Д-ОД от 15.12.2025. Праздники РК в счёт не входят: только пн–пт.
   Проект и суммы вымышлены. */
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.BudgetShow = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const TODAY = "2026-11-09";
  const PAY_LAG = 3;

  const STEP_BOOK = {
    zcp: [
      ["ztru", "Согласование заявки", "Заявка", 3, "п. 6.7.6"],
      ["cp", "Сбор ценовых предложений", "Сбор предложений", 5, "п. 6.3.3"],
      ["tech", "Техническая оценка", "Техническая оценка", 3, "п. 6.3.8"],
      ["improve", "Улучшение предложения, один круг", "Улучшение цены", 4, "п. 6.3.9–6.3.10"],
      ["trust", "Проверка благонадёжности", "Благонадёжность", 3, "п. 6.3.15"],
      ["protocol", "Протокол комиссии", "Протокол", 5, "п. 6.7.7"],
      ["draft", "Запуск согласования договора", "Договор", 2, "п. 7.1.1"],
      ["legal", "Правовая экспертиза", "Экспертиза", 5, "п. 7.1.1"],
      ["sign", "Подписание поставщиком", "Подписание", 5, "п. 7.1.5"],
    ],
    zcpComplex: null,
    oi: [
      ["ztru", "Согласование заявки", "Заявка", 3, "п. 6.7.6"],
      ["trust", "Проверка благонадёжности", "Благонадёжность", 3, "п. 6.4.3"],
      ["protocol", "Протокол комиссии", "Протокол", 5, "п. 6.7.7"],
      ["draft", "Запуск согласования договора", "Договор", 2, "п. 7.1.1"],
      ["legal", "Правовая экспертиза", "Экспертиза", 5, "п. 7.1.1"],
      ["sign", "Подписание поставщиком", "Подписание", 5, "п. 7.1.5"],
    ],
    invoice: [
      ["ztru", "Согласование заявки, договор не заключается", "Заявка", 3, "п. 6.5"],
    ],
  };
  STEP_BOOK.zcpComplex = STEP_BOOK.zcp.map((s) =>
    s[0] === "cp" ? ["cp", "Сбор предложений по сложному ТЗ", "Сбор предложений", 10, "п. 6.3.3"] : s.slice()
  );

  const METHOD = {
    zcp: { title: "Запрос цен", norm: 35 },
    zcpComplex: { title: "Запрос цен, сложное ТЗ", norm: 40 },
    oi: { title: "Один источник", norm: 23 },
    invoice: { title: "Счёт до 100 МРП", norm: 3 },
  };

  function stepsOf(method) {
    return STEP_BOOK[method].map(([id, label, short, days, clause]) => ({ id, label, short, days, clause }));
  }
  function procedureWd(method) {
    return stepsOf(method).reduce((s, x) => s + x.days, 0);
  }

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
  function wdBetween(a, b) {
    if (iso(a) === iso(b)) return 0;
    const sign = b > a ? 1 : -1;
    let n = 0;
    let x = new Date(a.getFullYear(), a.getMonth(), a.getDate());
    while (iso(x) !== iso(b)) {
      x = addCal(x, sign);
      if (!weekend(x)) n += sign;
    }
    return n;
  }
  function weekdayOnOrBefore(d) {
    let x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    while (weekend(x)) x = addCal(x, -1);
    return x;
  }

  function payDate(line) {
    if (line.pay === "post" || !line.need) return null;
    return weekdayOnOrBefore(addCal(parse(line.need), -line.lead));
  }
  function requestDate(line) {
    let d = parse(line.need);
    d = addCal(d, -line.lead);
    if (line.pay === "advance" || line.pay === "prepay") {
      d = weekdayOnOrBefore(d);
      d = addWd(d, -PAY_LAG);
    }
    d = addWd(d, -procedureWd(line.method));
    return d;
  }

  function analyze(line, hit, todayIso) {
    const today = parse(todayIso);
    const base = {
      id: line.id,
      name: line.name,
      spec: line.spec,
      kind: line.kind,
      method: line.method,
      methodTitle: METHOD[line.method].title,
      norm: procedureWd(line.method),
      amount: line.amount,
      source: line.source,
      quote: line.quote,
      milestone: line.milestone,
      pay: line.pay,
      lead: line.lead,
      steps: stepsOf(line.method),
    };
    if (!line.need) {
      return Object.assign(base, {
        undated: true,
        need: null,
        planStart: null,
        forecast: null,
        slip: null,
        stage: "Нет даты",
        cause: line.quote,
        tone: "mute",
        inScenario: false,
        committed: 0,
        paid: 0,
        delivered: 0,
        planStartIso: null,
        forecastIso: null,
        needIso: null,
        actualStartIso: null,
      });
    }
    const need = parse(line.need);
    const planStart = requestDate(line);
    const m = hit || {};
    const extraMap = m.extra || {};
    const extraWd =
      (m.startShiftWd || 0) +
      (m.blockedWd || 0) +
      Object.values(extraMap).reduce((s, n) => s + n, 0) +
      (m.holdWd || 0);
    let forecast = addWd(need, extraWd);
    const extraCal = (m.supplierSlipCal || 0) + (m.customsCal || 0) + (m.siteLagCal || 0);
    forecast = addCal(forecast, extraCal);
    const due = payDate(line);
    let advanceOpen = false;
    if (m.advanceUnpaid && due && due <= today) {
      const late = calBetween(due, today);
      forecast = addCal(need, late);
      advanceOpen = true;
    }
    if (m.crewUntil) {
      const crew = parse(m.crewUntil);
      if (crew > forecast) forecast = crew;
    }
    let rejected = false;
    if (m.remedyCal && need <= today) {
      forecast = addCal(today, m.remedyCal);
      rejected = true;
    }
    const slip = calBetween(need, forecast);
    const actualStart = extraWd || m.blockedWd || m.startShiftWd ? addWd(planStart, (m.startShiftWd || 0) + (m.blockedWd || 0)) : planStart;

    let stage;
    if (rejected) stage = "Приёмка отклонена";
    else if (forecast <= today) stage = slip > 0 ? "Принято с опозданием" : "Принято";
    else if (advanceOpen) stage = "Ждёт аванс";
    else if (today < actualStart) {
      if (slip > 0) stage = "Прогноз позже вехи";
      else if (planStart < today) stage = "Старт сдвинут";
      else stage = "Запас " + wdBetween(today, actualStart) + " р.д.";
    }
    else {
      let cursor = actualStart;
      let found = null;
      for (const step of stepsOf(line.method)) {
        const end = addWd(cursor, step.days + (extraMap[step.id] || 0));
        if (today < end) {
          found = step.short;
          break;
        }
        cursor = end;
      }
      if (!found && m.holdWd) found = "Пауза";
      stage = found || (line.kind === "Товар" ? "В пути" : "Идёт выполнение");
    }

    const started = actualStart <= today;
    const procEnd = addWd(actualStart, procedureWd(line.method) + Object.values(extraMap).reduce((s, n) => s + n, 0) + (m.holdWd || 0));
    const committed = !line.need ? 0 : line.method === "invoice" ? (addWd(actualStart, 3) <= today ? line.amount : 0) : procEnd <= today || forecast <= today ? line.amount : 0;
    const accepted = !rejected && forecast <= today;
    let paid = 0;
    if (line.pay === "post") paid = accepted ? line.amount : 0;
    else if (advanceOpen) paid = 0;
    else if (line.pay === "prepay") paid = (due && due <= today) || accepted ? line.amount : 0;
    else if (accepted) paid = line.amount;
    else if (due && due <= today && started) paid = Math.round(line.amount * 0.4);
    const delivered = accepted ? line.amount : 0;

    return Object.assign(base, {
      undated: false,
      needIso: line.need,
      planStartIso: iso(planStart),
      actualStartIso: iso(actualStart),
      forecastIso: iso(forecast),
      slip,
      stage,
      cause: m.cause || "",
      tone: m.tone || (m.cause ? "bad" : ""),
      inScenario: Boolean(hit),
      committed,
      paid,
      delivered,
      advanceOpen,
      rejected,
      dueIso: due ? iso(due) : null,
      illustrate: m.illustrate || "",
      extras: describeExtras(m),
    });
  }

  function describeExtras(m) {
    if (!m) return [];
    const out = [];
    if (m.blockedWd) out.push({ label: "Ожидание техкарт до заявки", days: m.blockedWd, unit: "р.д." });
    if (m.startShiftWd) out.push({ label: "Заявку подали позже рассчитанной даты", days: m.startShiftWd, unit: "р.д." });
    if (m.extra) {
      for (const [id, days] of Object.entries(m.extra)) {
        out.push({ label: m.extraLabel || "Сверх норматива на шаге " + id, days, unit: "р.д." });
      }
    }
    if (m.holdWd) out.push({ label: m.holdLabel || "Пауза", days: m.holdWd, unit: "р.д." });
    if (m.supplierSlipCal) out.push({ label: "Завод сдвинул отгрузку", days: m.supplierSlipCal, unit: "к.д." });
    if (m.customsCal) out.push({ label: "Груз задержан на таможне", days: m.customsCal, unit: "к.д." });
    if (m.siteLagCal) out.push({ label: "Площадка не готова, работы ждут", days: m.siteLagCal, unit: "к.д." });
    if (m.advanceUnpaid) out.push({ label: "Аванс не оплачен, отгрузка не началась", days: null, unit: "" });
    if (m.remedyCal) out.push({ label: "Повторная поставка после отклонения", days: m.remedyCal, unit: "к.д." });
    if (m.crewUntil) out.push({ label: "Бригада свободна только " + m.crewUntil, days: null, unit: "" });
    if (m.illustrate === "postpay") out.push({ label: "Оплата после приёмки не двигает поставку", days: null, unit: "" });
    return out;
  }

  const LINES_A = [
    { id: "survey", name: "Обследование опор и геодезия", spec: "12 перекрёстков", kind: "Услуга", method: "zcp", amount: 4200000, pay: "post", lead: 15, need: "2026-09-25", milestone: "Обследование", source: "Договор ГЧП, график", quote: "Обследование существующих опор завершить до 25.09.2026." },
    { id: "design", name: "Рабочая документация и сметы", spec: "комплект на участок", kind: "Услуга", method: "zcp", amount: 7800000, pay: "post", lead: 20, need: "2026-10-30", milestone: "Выпуск рабочей документации", source: "Договор ГЧП, график", quote: "Рабочая документация выдаётся до 30.10.2026." },
    { id: "dc", name: "Шкафы связи на узле ЦОД", spec: "доля проекта", kind: "Товар", method: "oi", amount: 24000000, pay: "post", lead: 20, need: "2026-10-15", milestone: "Узел связи готов", source: "Информационный листок, приложение 2", quote: "Доля шкафов ЦОД поставляется до 15.10.2026, исключительный производитель." },
    { id: "fast", name: "Крепёж, метизы и закладные", spec: "партия на 12 опор", kind: "Товар", method: "zcp", amount: 3150000, pay: "advance", lead: 14, need: "2026-11-06", milestone: "Комплектация до монтажа опор", source: "Состав проекта, ведомость", quote: "Закладные детали поставить до начала монтажа опор, 06.11.2026." },
    { id: "poles", name: "Опоры и фундаментные блоки", spec: "12 комплектов", kind: "Товар", method: "zcp", amount: 41200000, pay: "advance", lead: 40, need: "2026-12-04", milestone: "Опоры установлены", source: "ФЭМ, статья «конструкции»", quote: "Опоры должны стоять к 04.12.2026. Отгрузка после аванса 40%." },
    { id: "switch", name: "Коммутаторы и оптические кроссы", spec: "12 комплектов", kind: "Товар", method: "zcp", amount: 11400000, pay: "post", lead: 25, need: "2026-11-30", milestone: "Комплектация узла", source: "ТЗ, спецификация п. 5", quote: "Коммутаторы поставить до 30.11.2026. Оплата после приёмки." },
    { id: "radar", name: "Радары скорости", spec: "12 шт., ввоз", kind: "Товар", method: "zcp", amount: 19700000, pay: "advance", lead: 32, need: "2026-12-11", milestone: "Оборудование на площадке", source: "ТЗ, спецификация п. 7", quote: "Радары, ввозная поставка, на площадке до 11.12.2026." },
    { id: "cams", name: "Камеры распознавания и обзорные", spec: "48 шт.", kind: "Товар", method: "zcp", amount: 86400000, pay: "advance", lead: 30, need: "2027-02-26", milestone: "Оборудование на площадке", source: "ТЗ, спецификация п. 4.2", quote: "Камеры поставить до начала монтажа, 26.02.2027." },
    { id: "cabs", name: "Шкафы управления, ИБП, грозозащита", spec: "12 комплектов", kind: "Товар", method: "zcpComplex", amount: 27600000, pay: "advance", lead: 35, need: "2027-02-26", milestone: "Оборудование на площадке", source: "ТЗ, раздел 6", quote: "Шкафы со сложным ТЗ поставить до 26.02.2027." },
    { id: "cable", name: "Кабель оптический и медный", spec: "8,4 км", kind: "Товар", method: "zcp", amount: 18900000, pay: "post", lead: 21, need: "2027-03-05", milestone: "До начала прокладки", source: "Состав проекта, ведомость объёмов", quote: "Кабель на трассе до начала прокладки, 05.03.2027." },
    { id: "server", name: "Сервер видеоаналитики и СХД", spec: "1 комплект", kind: "Товар", method: "oi", amount: 34500000, pay: "advance", lead: 45, need: "2027-04-02", milestone: "Пусконаладка", source: "Информационный листок, приложение 2", quote: "Сервер единственного совместимого производителя — до пусконаладки, 02.04.2027." },
    { id: "lic", name: "Лицензии ПО на 48 каналов", spec: "ключи", kind: "Товар", method: "oi", amount: 22800000, pay: "prepay", lead: 10, need: "2027-04-02", milestone: "Пусконаладка", source: "Бизнес-план, лист лицензий", quote: "Ключи выдаются после предоплаты 100%, до 02.04.2027." },
    { id: "duct", name: "Прокладка кабельной канализации", spec: "8,4 км", kind: "Работа", method: "zcp", amount: 46800000, pay: "advance", lead: 30, need: "2027-03-05", milestone: "Канализация сдана", source: "Адресная программа", quote: "Канализация по 12 адресам сдаётся до 05.03.2027." },
    { id: "mount", name: "Монтаж оборудования на перекрёстках", spec: "12 адресов", kind: "Работа", method: "zcp", amount: 31200000, pay: "advance", lead: 20, need: "2027-03-26", milestone: "Монтаж завершён", source: "Адресная программа", quote: "Монтаж по адресной программе завершить до 26.03.2027." },
    { id: "pnr", name: "Пусконаладка и калибровка", spec: "48 камер", kind: "Услуга", method: "zcp", amount: 9600000, pay: "post", lead: 15, need: "2027-04-16", milestone: "Готовность к вводу", source: "ТЗ, программа испытаний", quote: "Калибровка заканчивается до 16.04.2027." },
    { id: "link", name: "Каналы связи", spec: "12 точек", kind: "Услуга", method: "zcp", amount: 14400000, pay: "post", lead: 20, need: "2027-04-02", milestone: "Пусконаладка", source: "Бизнес-план, связь", quote: "Каналы включены до пусконаладки, 02.04.2027." },
    { id: "guard", name: "Охрана площадок на период работ", spec: "февраль–март", kind: "Услуга", method: "zcp", amount: 4320000, pay: "post", lead: 10, need: "2027-02-01", milestone: "Начало монтажа", source: "ФЭМ, накладные", quote: "Охрана заступает к 01.02.2027." },
    { id: "signs", name: "Таблички и паспорта узлов", spec: "12 комплектов", kind: "Товар", method: "invoice", amount: 360000, pay: "post", lead: 10, need: "2027-04-23", milestone: "Ввод", source: "ТЗ, маркировка", quote: "Маркировка узлов до ввода, 23.04.2027. Сумма ниже 100 МРП." },
    { id: "train", name: "Обучение операторов филиала", spec: "2 группы", kind: "Услуга", method: "zcp", amount: 2160000, pay: "post", lead: 5, need: "2027-04-23", milestone: "Ввод", source: "Бизнес-план", quote: "Обучение до ввода участка, 23.04.2027." },
    { id: "zip", name: "ЗИП первого года", spec: "комплект", kind: "Товар", method: "zcp", amount: 5760000, pay: "post", lead: 20, need: "2027-04-30", milestone: "Ввод участка", source: "Информационный листок", quote: "ЗИП на складе филиала к вводу, 30.04.2027." },
    { id: "land", name: "Восстановление благоустройства", spec: "объём не разбит по адресам", kind: "Работа", method: "zcp", amount: 8900000, pay: "post", lead: 20, need: null, milestone: "—", source: "Бизнес-план", quote: "В бизнес-плане строка есть, вехи и даты нет. В расчёт не берём." },
    { id: "author", name: "Авторский надзор проектировщика", spec: "на весь срок", kind: "Услуга", method: "oi", amount: 3100000, pay: "post", lead: 10, need: null, milestone: "—", source: "Договор ГЧП", quote: "Формулировка «в ходе реализации». Даты потребности в пакете нет." },
  ];

  const SCENARIOS = [
    {
      id: "plan",
      group: "План",
      title: "План держится",
      lead: "Заявки уходят в рассчитанный день. Прогноз совпадает с датой потребности. Исполнение — только то, что уже принято.",
      hits: {},
    },
    {
      id: "late",
      group: "Заявка и процедура",
      title: "Заявку подали поздно",
      lead: "Камеры, опоры, монтаж и охрану вспомнили на 12 рабочих дней позже. Прогноз уходит за веху ещё до договора. Срочная закупка по п. 6.4.2 здесь не выход: Правила не разрешают её из-за позднего планирования.",
      hits: {
        cams: { startShiftWd: 12, cause: "Заявку подали на 12 рабочих дней позже рассчитанной даты." },
        poles: { startShiftWd: 12, cause: "Заявку подали на 12 рабочих дней позже. Аванс и отгрузка сдвигаются вместе с ней." },
        mount: { startShiftWd: 12, cause: "Подряд на монтаж запустили на 12 рабочих дней позже." },
        guard: { startShiftWd: 12, cause: "Охрану заказали, когда площадки уже должны были быть закрыты." },
      },
    },
    {
      id: "cards",
      group: "Заявка и процедура",
      title: "Нет техкарты от СТ",
      lead: "Сервисная служба не отдала техкарты на камеры. Заявку не принимают, п. 6.7.4. Процедура ещё не началась, а дата уже поехала.",
      hits: {
        cams: { blockedWd: 8, cause: "8 рабочих дней заявку не принимают: нет техкарт СТ. П. 6.7.4." },
      },
    },
    {
      id: "round",
      group: "Заявка и процедура",
      title: "Второй круг снижения цены",
      lead: "По шкафам ушли на второй круг улучшения цены. Норматив включает один круг. Правила разрешают повторять его без ограничения, каждый круг добавляет дни сверх 40.",
      hits: {
        cabs: { extra: { improve: 4 }, extraLabel: "Второй круг улучшения цены сверх одного, заложенного в норму", cause: "Второй круг улучшения цены: +4 рабочих дня. Норматив этого круга не содержит." },
      },
    },
    {
      id: "onebid",
      group: "Заявка и процедура",
      title: "Пришло одно предложение",
      lead: "На каналы связи поступило одно ценовое предложение. По п. 6.3.17 сбор объявляют заново. Это ещё одно окно, не «чуть дольше обычного».",
      hits: {
        link: { extra: { cp: 10 }, extraLabel: "Повторный сбор предложений: поступило одно ЦП, п. 6.3.17", cause: "Одно ценовое предложение, закупку объявили заново. П. 6.3.17, +10 рабочих дней." },
      },
    },
    {
      id: "quorum",
      group: "Заявка и процедура",
      title: "Комиссия без кворума",
      lead: "Протокол по обучению лежит дольше пяти рабочих дней: комиссию не собрали. П. 6.7.7 ограничивает согласование протокола пятью днями.",
      hits: {
        train: { extra: { protocol: 8 }, extraLabel: "Протокол лежит дольше пяти рабочих дней, нет кворума", cause: "Протокол не подписан: нет кворума. Сверх п. 6.7.7 ещё 8 рабочих дней." },
      },
    },
    {
      id: "legal",
      group: "Заявка и процедура",
      title: "Договор завис у юристов",
      lead: "По монтажу контрагент прислал протокол разногласий. Правовая экспертиза идёт дольше пяти рабочих дней из п. 7.1.1, договор не подписан.",
      hits: {
        mount: { extra: { legal: 9 }, extraLabel: "Правовая экспертиза дольше пяти рабочих дней: протокол разногласий", cause: "Экспертиза и протокол разногласий: +9 рабочих дней сверх п. 7.1.1." },
      },
    },
    {
      id: "price",
      group: "Заявка и процедура",
      title: "Цена выше строки бюджета",
      lead: "Победитель по шкафам дороже строки финансово-экономической модели. Присуждение стоит, пока финансовый блок не пересоберёт бюджет. Календарь процедуры в это время не идёт.",
      hits: {
        cabs: { holdWd: 10, holdLabel: "Пауза: цена выше строки ФЭМ, ждём пересбор бюджета", cause: "Цена выше строки ФЭМ. Пауза 10 рабочих дней, пока бюджет не пересоберут." },
      },
    },
    {
      id: "advance",
      group: "Деньги",
      title: "Аванс не ушёл",
      lead: "По опорам аванс должен был уйти до отгрузки и до сих пор не ушёл: поставка сдвигается на число дней просрочки платежа. Коммутаторы оплачиваются после приёмки — пустая касса их дату не двигает.",
      hits: {
        poles: { advanceUnpaid: true, cause: "Аванс 40% не оплачен. Поставщик не отгружает, дата уезжает на срок просрочки платежа." },
        switch: { illustrate: "postpay", tone: "ok", cause: "Оплата после приёмки. Деньги ещё не ушли, дата потребности на месте." },
      },
    },
    {
      id: "factory",
      group: "Поставка и приёмка",
      title: "Завод сдвинул отгрузку",
      lead: "Договор и деньги по опорам в порядке. Завод перенёс отгрузку на 21 календарный день. Закупка тут уже ни при чём.",
      hits: {
        poles: { supplierSlipCal: 21, cause: "Завод сдвинул отгрузку на 21 календарный день при живом договоре и оплаченном авансе." },
      },
    },
    {
      id: "customs",
      group: "Поставка и приёмка",
      title: "Груз на таможне",
      lead: "Радары уже в пути. Выпуск задержан на 18 календарных дней. Процедура закупки к этой задержке не относится.",
      hits: {
        radar: { customsCal: 18, cause: "Груз на таможне, +18 календарных дней к поставке." },
      },
    },
    {
      id: "reject",
      group: "Поставка и приёмка",
      title: "Партию отклонили",
      lead: "Крепёж приехал в срок, приёмка отклонила партию. В исполнении ноль, аванс уже у поставщика. Повторная поставка — 18 календарных дней.",
      hits: {
        fast: { remedyCal: 18, cause: "Приёмка отклонила партию. Исполнение не засчитано, повтор 18 календарных дней. Аванс уже уплачен." },
      },
    },
    {
      id: "site",
      group: "Площадка и люди",
      title: "Площадка не готова",
      lead: "Канализацию закупили вовремя. Филиал не согласовал отключения и котлованы, работы стоят 15 календарных дней. Это предшественник на площадке, не срыв закупки.",
      hits: {
        duct: { siteLagCal: 15, cause: "Площадка не готова: нет отключений и котлованов. +15 календарных дней при вовремя запущенной закупке." },
      },
    },
    {
      id: "crew",
      group: "Площадка и люди",
      title: "Бригада на другом участке",
      lead: "Калибровку ещё не закупали, но единственная бригада занята до 28 апреля, а веха — 16 апреля. Просрочка видна уже в ноябре.",
      hits: {
        pnr: { crewUntil: "2027-04-28", cause: "Бригада калибровки свободна только 28.04.2027, веха — 16.04.2027." },
      },
    },
    {
      id: "combo",
      group: "Вместе",
      title: "Несколько причин сразу",
      lead: "Так выглядит месяц, когда причины сложились: поздние камеры, неоплаченный аванс по опорам, таможня по радарам, отклонённый крепёж, неготовая площадка, договор монтажа у юристов и одно предложение на каналы.",
      hits: {
        cams: { startShiftWd: 12, cause: "Заявку на камеры подали на 12 рабочих дней позже." },
        poles: { advanceUnpaid: true, cause: "Аванс по опорам не оплачен, отгрузка стоит." },
        radar: { customsCal: 18, cause: "Радары задержаны на таможне на 18 календарных дней." },
        fast: { remedyCal: 18, cause: "Партию крепежа отклонили на приёмке." },
        duct: { siteLagCal: 15, cause: "Канализация ждёт готовности площадки, +15 календарных дней." },
        mount: { extra: { legal: 9 }, extraLabel: "Договор монтажа завис на правовой экспертизе", cause: "Договор монтажа завис у юристов, +9 рабочих дней." },
        link: { extra: { cp: 10 }, extraLabel: "Повторный сбор: одно предложение", cause: "По каналам связи одно предложение, сбор объявили заново." },
      },
    },
  ];

  const LINES_B = [
    { id: "b1", name: "Обследование 6 перекрёстков", spec: "выполнено", kind: "Услуга", method: "zcp", amount: 2100000, pay: "post", lead: 10, need: "2026-08-20", milestone: "Обследование", source: "Договор ГЧП", quote: "Обследование закрыто 20.08.2026." },
    { id: "b2", name: "Опоры, 6 комплектов", spec: "приняты", kind: "Товар", method: "zcp", amount: 18600000, pay: "advance", lead: 30, need: "2026-10-01", milestone: "Опоры", source: "ФЭМ", quote: "Опоры приняты 01.10.2026." },
    { id: "b3", name: "Камеры, 24 шт.", spec: "приняты", kind: "Товар", method: "zcp", amount: 39200000, pay: "advance", lead: 25, need: "2026-10-20", milestone: "Поставка", source: "ТЗ", quote: "Камеры приняты 20.10.2026." },
    { id: "b4", name: "Монтаж 6 адресов", spec: "идёт по графику", kind: "Работа", method: "zcp", amount: 12400000, pay: "advance", lead: 15, need: "2026-11-27", milestone: "Монтаж", source: "Адресная программа", quote: "Монтаж завершается 27.11.2026, запас есть." },
    { id: "b5", name: "Каналы связи, 6 точек", spec: "в графике", kind: "Услуга", method: "zcp", amount: 4800000, pay: "post", lead: 12, need: "2026-12-10", milestone: "Включение", source: "Бизнес-план", quote: "Каналы к 10.12.2026." },
    { id: "b6", name: "Ввод участка", spec: "калибровка", kind: "Услуга", method: "zcp", amount: 2400000, pay: "post", lead: 8, need: "2026-12-18", milestone: "Ввод", source: "ТЗ", quote: "Калибровка до ввода 18.12.2026." },
  ];

  const LINES_C = [
    { id: "c1", name: "Проект развязок", spec: "просрочен выпуск", kind: "Услуга", method: "zcp", amount: 9600000, pay: "post", lead: 20, need: "2026-09-01", milestone: "Проект", source: "Договор ГЧП", quote: "Проект должен был выйти 01.09.2026." },
    { id: "c2", name: "Металлоконструкции пролётов", spec: "завод", kind: "Товар", method: "zcp", amount: 64000000, pay: "advance", lead: 40, need: "2026-10-20", milestone: "Поставка металла", source: "ФЭМ", quote: "Металл на площадке до 20.10.2026." },
    { id: "c3", name: "Опоры освещения", spec: "в пути", kind: "Товар", method: "zcp", amount: 15200000, pay: "advance", lead: 20, need: "2026-11-02", milestone: "Освещение", source: "Состав проекта", quote: "Опоры до 02.11.2026." },
    { id: "c4", name: "Монтаж пролётных строений", spec: "площадка", kind: "Работа", method: "zcp", amount: 28000000, pay: "advance", lead: 25, need: "2026-11-25", milestone: "Монтаж", source: "График производства работ", quote: "Монтаж до 25.11.2026." },
    { id: "c5", name: "Камеры на развязках", spec: "заявка", kind: "Товар", method: "zcp", amount: 22100000, pay: "advance", lead: 20, need: "2026-12-15", milestone: "Оснащение", source: "ТЗ", quote: "Камеры до 15.12.2026." },
  ];

  const LOCK_C = {
    c1: { startShiftWd: 15, cause: "Документацию выпустили на 15 рабочих дней позже графика." },
    c2: { supplierSlipCal: 24, cause: "Завод задержал металл на 24 календарных дня. Веха октября уже сорвана." },
    c3: { advanceUnpaid: true, cause: "Аванс за опоры не ушёл, отгрузка стоит." },
    c4: { siteLagCal: 12, cause: "Фронт монтажа открыли на 12 дней позже." },
    c5: { blockedWd: 10, cause: "Заявку держат: нет согласованного ТЗ." },
  };

  function rowsOf(lines, hits, today) {
    return lines.map((line) => analyze(line, hits[line.id] || null, today));
  }

  function totals(rows) {
    const budget = rows.reduce((s, r) => s + r.amount, 0);
    const committed = rows.reduce((s, r) => s + r.committed, 0);
    const paid = rows.reduce((s, r) => s + r.paid, 0);
    const delivered = rows.reduce((s, r) => s + r.delivered, 0);
    const late = rows.filter((r) => r.slip > 0);
    return {
      budget,
      committed,
      paid,
      delivered,
      execution: budget ? delivered / budget : 0,
      lateCount: late.length,
      lateAmount: late.reduce((s, r) => s + r.amount, 0),
      noDate: rows.filter((r) => r.undated).length,
      dated: rows.filter((r) => !r.undated).length,
    };
  }

  function sortRows(rows) {
    const band = (r) => {
      if (r.inScenario && r.tone !== "ok") return 0;
      if (r.inScenario) return 1;
      if (r.undated) return 5;
      if (String(r.stage).startsWith("Принято")) return 4;
      if (String(r.stage).startsWith("Запас")) return 3;
      return 2;
    };
    const when = (r) => r.needIso || "9999-99-99";
    return rows.slice().sort((a, b) => band(a) - band(b) || (b.slip || 0) - (a.slip || 0) || when(a).localeCompare(when(b)));
  }

  const PROJECTS = [
    {
      id: "a14",
      code: "А-14",
      title: "Участок А-14",
      subtitle: "12 перекрёстков",
      commissioning: "2027-04-30",
      interactive: true,
      lines: LINES_A,
    },
    {
      id: "b02",
      code: "Б-02",
      title: "Обход Б-02",
      subtitle: "6 перекрёстков",
      commissioning: "2026-12-18",
      interactive: false,
      lines: LINES_B,
      note: "Снимок участка, где даты держатся. Сценарии переключаются на А-14.",
    },
    {
      id: "c07",
      code: "В-07",
      title: "Магистраль В-07",
      subtitle: "4 развязки",
      commissioning: "2026-12-20",
      interactive: false,
      lines: LINES_C,
      lock: LOCK_C,
      note: "Снимок участка, где веха октября уже сорвана. Причины разные и уже сложились.",
    },
  ];

  function view(projectId, scenarioId, today) {
    const project = PROJECTS.find((p) => p.id === projectId) || PROJECTS[0];
    const scenario = project.interactive
      ? SCENARIOS.find((s) => s.id === scenarioId) || SCENARIOS[0]
      : {
          id: "lock",
          group: "Снимок",
          title: project.code,
          lead: project.note,
          hits: project.lock || {},
        };
    const rows = sortRows(rowsOf(project.lines, scenario.hits, today || TODAY));
    const planRows = rowsOf(project.lines, project.lock || {}, today || TODAY);
    const planTotals = project.interactive ? totals(rowsOf(project.lines, {}, today || TODAY)) : totals(planRows);
    const now = totals(rows);
    const pastCommission = rows.filter((r) => r.forecastIso && r.forecastIso > project.commissioning);
    return {
      today: today || TODAY,
      project,
      scenario,
      rows,
      totals: now,
      planTotals,
      pastCommission: pastCommission.length,
      pastCommissionAmount: pastCommission.reduce((s, r) => s + r.amount, 0),
    };
  }

  function money(n) {
    if (n >= 1000000) {
      const v = n / 1000000;
      const digits = v >= 10 ? 1 : 2;
      return v.toLocaleString("ru-RU", { minimumFractionDigits: digits, maximumFractionDigits: digits }) + " млн";
    }
    return Math.round(n / 1000).toLocaleString("ru-RU") + " тыс.";
  }
  const MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
  function longDate(isoText) {
    if (!isoText) return "—";
    const d = parse(isoText);
    return d.getDate() + " " + MONTHS[d.getMonth()] + " " + d.getFullYear();
  }
  function shortDate(isoText) {
    if (!isoText) return "—";
    const d = parse(isoText);
    const z = (n) => String(n).padStart(2, "0");
    return z(d.getDate()) + "." + z(d.getMonth() + 1) + "." + d.getFullYear();
  }

  return {
    TODAY,
    PAY_LAG,
    METHOD,
    SCENARIOS,
    PROJECTS,
    procedureWd,
    analyze,
    view,
    totals,
    money,
    longDate,
    shortDate,
    parse,
    iso,
    stepsOf,
  };
});
