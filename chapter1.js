const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, Header, Footer, PageNumberElement,
} = require("docx");
const fs = require("fs");

const W    = 9026;
const FONT = "Times New Roman";
const BLUE = "1F3864";
const HBLU = "D6E4F7";
const HGRN = "EAF3DE";
const HRED = "FEE2E2";
const HYEL = "FEF9C3";
const HGRN2= "DCFCE7";
const HPUR = "EDE9FE";

const brd  = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const BORD = { top: brd, bottom: brd, left: brd, right: brd };
const CM   = { top: 70, bottom: 70, left: 110, right: 110 };

// ─── Paragraph helpers ───────────────────────────────────────────────────────
const body = (text, opts={}) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 160 },
  indent: { firstLine: 720 },
  children: [new TextRun({ text, font: FONT, size: 24, ...opts })],
});
const bodyNI = (text, opts={}) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 160 },
  children: [new TextRun({ text, font: FONT, size: 24, ...opts })],
});
const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 360 },
  children: [new TextRun({ text, font: FONT, size: 28, bold: true, allCaps: true })],
});
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 200 },
  children: [new TextRun({ text, font: FONT, size: 26, bold: true })],
});
const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 220, after: 140 },
  children: [new TextRun({ text, font: FONT, size: 24, bold: true })],
});
const bul = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 80 },
  children: [new TextRun({ text, font: FONT, size: 24 })],
});
const num = (text) => new Paragraph({
  numbering: { reference: "numbers", level: 0 },
  spacing: { after: 80 },
  children: [new TextRun({ text, font: FONT, size: 24 })],
});
const tblCap = (text) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 80 },
  children: [new TextRun({ text, font: FONT, size: 24, bold: true })],
});
const figCap = (text) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 80, after: 200 },
  children: [new TextRun({ text, font: FONT, size: 22, italics: true })],
});
const pb = () => new Paragraph({ children: [new PageBreak()] });
const sp = () => new Paragraph({ spacing: { after: 100 }, children: [new TextRun("")] });
const ph = (caption) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 40 },
    border: { top: brd, bottom: brd, left: brd, right: brd },
    shading: { fill: "F8FAFC", type: ShadingType.CLEAR },
    children: [new TextRun({ text: "[Місце для рисунку — вставити скріншот]", font: FONT, size: 20, italics: true, color: "94A3B8" })],
  }),
  figCap(caption),
];
const phBPMN = (caption) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 40 },
    border: { top: brd, bottom: brd, left: brd, right: brd },
    shading: { fill: "F0F9FF", type: ShadingType.CLEAR },
    children: [new TextRun({ text: "[BPMN-діаграма — вставити з чату]", font: FONT, size: 20, italics: true, color: "0369A1" })],
  }),
  figCap(caption),
];

// ─── Table helpers ────────────────────────────────────────────────────────────
const tHdr = (cells, widths) => new TableRow({
  tableHeader: true,
  children: cells.map((t, i) => new TableCell({
    borders: BORD, margins: CM,
    width: { size: widths[i], type: WidthType.DXA },
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: t, font: FONT, size: 20, bold: true, color: "FFFFFF" })],
    })],
  })),
});

const makeCell = (text, width, fill, align=AlignmentType.CENTER, bold=false, size=20) =>
  new TableCell({
    borders: BORD, margins: CM,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, font: FONT, size, bold, color: fill===HRED?"7F1D1D": fill===HGRN2?"14532D": "000000" })],
    })],
  });

const yesCell = (w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: HGRN2, type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+", font: FONT, size: 20, bold: true, color: "14532D" })] })],
});
const noCell = (w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: HRED, type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "–", font: FONT, size: 20, color: "7F1D1D" })] })],
});
const partCell = (w, text="±") => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: HYEL, type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, font: FONT, size: 20, color: "713F12" })] })],
});
const featCell = (text, w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: "F1F5F9", type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text, font: FONT, size: 20 })] })],
});
const noteCell = (text, w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text, font: FONT, size: 19, italics: true, color: "444444" })] })],
});

// Special "our project" header cell
const ourHdr = (w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: "1A5C0A", type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new TextRun({ text: "Survey App", font: FONT, size: 20, bold: true, color: "FFFFFF" }),
    new TextRun({ text: "\n(дипломний проєкт)", font: FONT, size: 17, color: "DCFCE7", break: 1 }),
  ]})],
});
const ourYes = (w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: "C7E6B8", type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "+", font: FONT, size: 20, bold: true, color: "1A5C0A" })] })],
});
const ourVal = (text, w) => new TableCell({
  borders: BORD, margins: CM, width: { size: w, type: WidthType.DXA },
  shading: { fill: "C7E6B8", type: ShadingType.CLEAR },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, font: FONT, size: 20, bold: true, color: "1A5C0A" })] })],
});

// Column widths for comparison table
const WF=2100, WO=1050, WG=900, WSM=900, WT=900, WN=2276; // total ~8126 + extra = 9026 after adjusting

const compRow = (feat, our, gf, sm, tf, note) => {
  const mkOur = (v) => v==="+" ? ourYes(WO) : v==="-" ? makeCell("–",WO,HRED) : ourVal(v, WO);
  const mk = (v, w) => v==="+" ? yesCell(w) : v==="-" ? noCell(w) : partCell(w, v);
  return new TableRow({ children: [featCell(feat,WF), mkOur(our), mk(gf,WG), mk(sm,WSM), mk(tf,WT), noteCell(note,WN)] });
};

// ─── DOCUMENT ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference:"bullets", levels:[{ level:0, format:LevelFormat.BULLET, text:"–", alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:720,hanging:360}}} }] },
      { reference:"numbers", levels:[{ level:0, format:LevelFormat.DECIMAL, text:"%1.", alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:720,hanging:360}}} }] },
    ],
  },
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true, run:{size:28,bold:true,font:FONT}, paragraph:{spacing:{before:480,after:360},outlineLevel:0} },
      { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true, run:{size:26,bold:true,font:FONT}, paragraph:{spacing:{before:280,after:200},outlineLevel:1} },
      { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true, run:{size:24,bold:true,font:FONT}, paragraph:{spacing:{before:220,after:140},outlineLevel:2} },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width:11906, height:16838 },
        margin: { top:1134, right:850, bottom:1134, left:1701 },
      },
    },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      border: { bottom: { style:BorderStyle.SINGLE, size:6, color:"AAAAAA", space:1 } },
      children: [new TextRun({ text:"Програмне забезпечення веборієнтованої системи онлайн-опитувань", font:FONT, size:18, color:"888888" })],
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style:BorderStyle.SINGLE, size:4, color:"CCCCCC", space:1 } },
      children: [new PageNumberElement()],
    })] }) },

children: [

// ══════════════════════════════════════════════════════
// ВСТУП
// ══════════════════════════════════════════════════════
new Paragraph({
  alignment: AlignmentType.CENTER, pageBreakBefore: true,
  spacing: { before:0, after:400 },
  children: [new TextRun({ text:"ВСТУП", font:FONT, size:28, bold:true })],
}),
body("В умовах стрімкої цифровізації усіх сфер суспільного життя збір, обробка та аналіз думок великих аудиторій набувають вирішального значення для прийняття обґрунтованих рішень у бізнесі, освіті, науці та державному управлінні. Онлайн-опитування є одним із найефективніших та найдоступніших інструментів для досягнення цієї мети, оскільки дозволяють охопити широку географічно розподілену аудиторію, автоматизувати збір відповідей та отримувати результати в режимі реального часу."),
body("Актуальність теми дипломного проєкту обумовлена зростаючою потребою організацій різних галузей — від бізнесу та HR до наукових установ і державних органів — у зручних, функціональних та безпечних веборієнтованих інструментах для створення, поширення та аналізу опитувань. За даними дослідницьких компаній, глобальний ринок програмного забезпечення для опитувань зростає щорічно на 12–15%, а попит на рішення з розширеними аналітичними можливостями значно перевищує пропозицію на ринку відкритого програмного забезпечення."),
body("Провідні тенденції розв'язання задачі автоматизованого збору та аналізу зворотного зв'язку включають: перехід від монолітних до мікросервісних архітектур, широке застосування хмарних технологій та контейнеризації, інтеграцію аналітичних модулів із засобами машинного навчання та великих мовних моделей (LLM) для виявлення закономірностей у відповідях та автоматизованого формування опитувань, а також активне використання SPA-підходу (Single Page Application) для побудови зручного клієнтського інтерфейсу. Сучасний стан розробок характеризується домінуванням комерційних рішень (SurveyMonkey, Google Forms, Typeform), що мають суттєві обмеження: закритий вихідний код, обмежений безкоштовний функціонал та неможливість повного контролю над даними."),
body("У рамках даного дипломного проєкту розроблено програмне забезпечення веборієнтованої системи створення, проведення та аналізу онлайн-опитувань Survey App. Система реалізована як повностекний веб-додаток з архітектурою клієнт-сервер, що включає: React SPA з TypeScript на клієнтській стороні, RESTful API на базі Node.js та Express з TypeScript на серверній стороні, реляційну базу даних PostgreSQL з ORM Prisma, а також контейнеризацію через Docker та Nginx. Додатковою ключовою особливістю системи є інтеграція штучного інтелекту на базі Google Gemini API для автоматизованої генерації структури опитувань за темою."),
body("Призначення системи — надати організаціям та окремим авторам зручний інструмент для самостійного розгортання платформи опитувань з повним контролем над даними. Мета дипломного проєкту — розробити функціонально повну, безпечну та продуктивну веборієнтовану систему, що підтримує повний цикл роботи з онлайн-опитуваннями: від AI-асистованого створення та публікації до збору відповідей та інтерактивного аналізу результатів."),
body("Можливі сфери застосування розробленої системи охоплюють: маркетингові дослідження (NPS-опитування, збір відгуків про продукти), управління персоналом (опитування задоволеності співробітників, зворотний зв'язок після тренінгів), освіту (тестування знань, оцінювання якості викладання), академічні дослідження (соціологічні та психологічні опитування), а також оцінювання якості державних послуг."),
pb(),

// ══════════════════════════════════════════════════════
// РОЗДІЛ 1
// ══════════════════════════════════════════════════════
h1("РОЗДІЛ 1\nПЕРЕДПРОЄКТНЕ ОБСТЕЖЕННЯ ПРЕДМЕТНОЇ ОБЛАСТІ"),

h2("1.1 Постановка завдання дипломного проєктування"),
body("Дипломне проєктування передбачає виконання наступних завдань:"),
bul("аналіз предметної області онлайн-опитувань та опис її ключових бізнес-процесів, визначення загального завдання розробки у рамках дипломного проєкту;"),
bul("аналіз існуючих програмних рішень у сфері онлайн-опитувань та порівняльний аналіз їхніх функціональних можливостей;"),
bul("аналіз алгоритмічних та технічних рішень для побудови веборієнтованих систем збору та аналізу даних;"),
bul("аналіз та моделювання ключових бізнес-процесів системи з використанням нотації BPMN;"),
bul("розроблення функціональних, нефункціональних та системних вимог до програмного забезпечення;"),
bul("аналіз економічних показників програмного забезпечення дипломного проєкту;"),
bul("постановка завдання на розробку програмного забезпечення;"),
bul("проєктування архітектури системи: вибір архітектурного стилю, проєктування схеми бази даних та REST API;"),
bul("обґрунтування вибору засобів розробки програмного забезпечення та технологічного стеку;"),
bul("конструювання та реалізація програмного забезпечення: серверна частина (Node.js, Express, Prisma), клієнтська частина (React, TypeScript), аналітичний модуль та модуль AI-генерації;"),
bul("аналіз безпеки даних програмного забезпечення;"),
bul("аналіз якості та тестування програмного забезпечення;"),
bul("розгортання програмного забезпечення з використанням Docker та Nginx;"),
bul("створення супроводжувальної документації до розробленого програмного забезпечення."),

h2("1.2 Аналіз предметної області"),
body("Предметною областю даної роботи є веборієнтовані інформаційні системи для збору, обробки та аналізу структурованих даних у формі опитувань. Онлайн-опитування являють собою цифровий метод збору інформації, при якому респонденти відповідають на запитання через веб-інтерфейс, а відповіді автоматично зберігаються та агрегуються для подальшого аналізу. На відміну від традиційних паперових анкет, веборієнтовані системи опитувань забезпечують автоматизований збір відповідей, мінімізують людський фактор при обробці результатів та надають можливість аналізу даних у реальному часі."),
body("Стрімкий розвиток цієї сфери розпочався у 2000-х роках разом із поширенням широкосмугового Інтернету та масовою веб-доступністю, а новий етап еволюції пов'язаний із впровадженням мобільних технологій, що зробило участь в опитуваннях зручною для будь-якого пристрою. Сучасні системи онлайн-опитувань трансформуються з простих форм збору даних на комплексні платформи дослідницьких даних з розвиненими аналітичними можливостями та засобами штучного інтелекту."),
body("Згідно з дослідженнями у сфері інформаційних систем, ефективність збору зворотного зв'язку безпосередньо залежить від зручності інструменту, різноманітності підтримуваних типів питань та доступності аналітики для авторів опитувань [1; 2]. Дослідження також показують, що персоналізований дизайн опитувань та умовна логіка переходу між питаннями підвищують рівень завершення опитувань на 20–35% порівняно зі стандартними лінійними формами [3]."),
body("Аналіз предметної області дозволяє виділити ключові категорії користувачів систем онлайн-опитувань та їх специфічні потреби."),
body("Комерційні організації використовують опитування насамперед для вимірювання індексу споживчої лояльності (Net Promoter Score, NPS), дослідження задоволеності клієнтів (CSAT) та збору відгуків про продукти. Для цієї категорії критично важливими є підтримка NPS-типу питань, можливість автоматичного розрахунку показників та інтеграція результатів з CRM-системами через вебхуки."),
body("HR-департаменти та служби управління персоналом активно застосовують опитування для оцінювання задоволеності та залученості співробітників (Employee Engagement Survey), збору зворотного зв'язку після тренінгів та оцінювання корпоративної культури. Специфічною потребою цієї категорії є забезпечення анонімності відповідей та підтримка матричних питань для багатокритеріального оцінювання."),
body("Освітні заклади (університети, школи, навчальні центри) використовують системи опитувань для проведення тестувань, збору відгуків студентів про якість викладання та оцінювання ефективності навчальних програм. Для освітнього сектору важливими є підтримка обмеження часу на проходження, запобігання повторним відповідям та експорт результатів у форматах, сумісних з академічним програмним забезпеченням."),
body("Дослідницькі організації та соціологічні служби застосовують онлайн-опитування для проведення широкомасштабних соціологічних, маркетингових та наукових досліджень. Для цієї категорії першорядне значення мають статистична достовірність зібраних даних, підтримка складних типів питань (матриця, шкала Лікерта) та можливість експорту даних у форматах для статистичного аналізу (CSV, SPSS)."),
body("На сучасному етапі розвитку ІТ реалізація систем онлайн-опитувань характеризується використанням клієнт-серверних архітектур із застосуванням SPA-підходу на клієнтській стороні та RESTful або GraphQL API — на серверній. Бізнес-процес функціонування типової системи онлайн-опитувань передбачає: реєстрацію та автентифікацію авторів, створення опитування з налаштуванням структури та доступу, публікацію та поширення за посиланням, збір відповідей респондентів та автоматизовану обробку і візуалізацію результатів."),
body("Аналіз предметної області виявляє суперечність між зростаючими потребами організацій у гнучких та функціональних інструментах для опитувань та обмеженнями існуючих рішень: комерційні продукти є дорогими та не надають повного контролю над даними, тоді як відкриті рішення часто поступаються за зручністю та сучасністю інтерфейсу. Саме це протиріччя визначає актуальність розробки власної відкритої системи на сучасному технологічному стеку."),

h2("1.3 Аналіз існуючих рішень"),
body("Проаналізуємо відоме на сьогодні алгоритмічне забезпечення у даній області та технічні рішення, що допоможуть у реалізації Survey App. Нижче будуть розглянуті готові програмні рішення, допоміжні програмні засоби та засоби розробки."),

h3("1.3.1 Аналіз відомих програмних продуктів"),
body("У межах предметної області онлайн-опитувань існує значна кількість програмних продуктів. Для аналізу обрано найбільш репрезентативні рішення: Google Forms, SurveyMonkey та Typeform."),
body("Google Forms є найпоширенішим безкоштовним інструментом для створення онлайн-опитувань, що входить до екосистеми Google Workspace. Система підтримує базові типи питань (вибір з варіантів, шкала, текст, завантаження файлів), автоматичну агрегацію відповідей у Google Sheets та базову візуалізацію результатів у вигляді кругових та стовпчастих діаграм. Суттєвими недоліками є відсутність NPS-типу питань, обмежена аналітика, відсутність умовної логіки в безкоштовній версії та повна залежність від інфраструктури Google (рис. 1.1)."),
...ph("Рисунок 1.1 – Інтерфейс конструктора опитувань Google Forms"),
body("SurveyMonkey є одним із провідних комерційних рішень на ринку, що надає розширений функціонал для бізнес-досліджень. Платформа підтримує понад 15 типів питань, включаючи матричні питання та шкалу Лікерта, умовну логіку (skip logic), розширену аналітику з фільтрацією та порівнянням результатів, а також інтеграцію з популярними бізнес-системами (Salesforce, HubSpot, Slack). Ключовим обмеженням є висока вартість підписки (від 25 до 99 USD на місяць за просунуті функції) та закритий вихідний код (рис. 1.2)."),
...ph("Рисунок 1.2 – Інтерфейс платформи SurveyMonkey"),
body("Typeform відрізняється унікальним підходом до UX: опитування відображаються у форматі «одне питання за раз» з анімованими переходами, що забезпечує вищий рівень залученості респондентів. Система підтримує умовну логіку, інтеграції через Zapier та власний API. Недоліком є обмежена кількість типів питань у базовій версії (8 типів) та відсутність матричних питань у безкоштовному плані. Вартість комерційного плану — від 25 USD на місяць (рис. 1.3)."),
...ph("Рисунок 1.3 – Інтерфейс платформи Typeform"),
body("Для порівняння розробленого проєкту з аналогами складено таблицю 1.1."),

tblCap("Таблиця 1.1 – Порівняльний аналіз систем онлайн-опитувань"),
new Table({
  width: { size: W, type: WidthType.DXA },
  columnWidths: [WF, WO, WG, WSM, WT, WN],
  rows: [
    new TableRow({ tableHeader: true, children: [
      makeCell("Функціонал / критерій", WF, BLUE, AlignmentType.CENTER, true, 20),
      new TableCell({ borders:BORD, margins:CM, width:{size:WO,type:WidthType.DXA}, shading:{fill:"1A5C0A",type:ShadingType.CLEAR}, children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Survey App\n(дипломний проєкт)",font:FONT,size:19,bold:true,color:"FFFFFF"})]})]}),
      makeCell("Google Forms", WG, BLUE, AlignmentType.CENTER, true, 20),
      makeCell("SurveyMonkey", WSM, BLUE, AlignmentType.CENTER, true, 20),
      makeCell("Typeform", WT, BLUE, AlignmentType.CENTER, true, 20),
      makeCell("Пояснення", WN, BLUE, AlignmentType.CENTER, true, 20),
    ]}),
    compRow("Безкоштовне використання",     "+", "+", "±", "±",
      "Survey App повністю безкоштовний; SurveyMonkey та Typeform обмежують кількість відповідей"),
    compRow("Відкритий вихідний код",        "+", "–", "–", "–",
      "Власне розгортання без ліцензійних відрахувань; повний контроль над даними"),
    compRow("Кількість типів питань",        "10","9", "15","8",
      "Survey App покриває всі практично необхідні типи для бізнесу, HR та освіти"),
    compRow("NPS-питання",                   "+", "–", "+", "+",
      "Net Promoter Score — ключовий тип для вимірювання лояльності; відсутній у Google Forms"),
    compRow("Матричні питання",              "+", "+", "+", "–",
      "Матриця критично важлива для HR-опитувань; відсутня в базовому Typeform"),
    compRow("Умовна логіка (skip logic)",    "+", "±", "+", "+",
      "Підвищує рівень завершення опитувань на 20–35%; у Google Forms лише базова"),
    compRow("5 рівнів видимості",            "+", "–", "±", "–",
      "PUBLIC / UNLISTED / PRIVATE / PASSWORD / EMAIL_LIST — найгнучкіша модель"),
    compRow("Аналітика в реальному часі",    "+", "±", "+", "±",
      "Розширена: теплова карта, воронка відмов, аналіз пристроїв, час проходження"),
    compRow("Публічний каталог",             "+", "–", "–", "–",
      "Унікальна функція /explore з пошуком, тегами та закладками"),
    compRow("Автозбереження прогресу",       "+", "–", "+", "+",
      "Прогрес зберігається кожні 2 секунди через sessionKey"),
    compRow("Завантаження файлів",           "+", "+", "+", "+",
      "Підтримується всіма системами; Survey App: ліміт 20 МБ, типи JPEG/PNG/PDF"),
    compRow("Публічна сторінка автора",      "+", "–", "–", "–",
      "/u/{username} — публічна сторінка зі статистикою та опитуваннями автора"),
    compRow("Експорт CSV (UTF-8)",           "+", "+", "+", "+",
      "Survey App генерує CSV з BOM — коректне відкриття в Excel з кирилицею"),
    compRow("Інтеграція ШІ (AI-генерація)",  "+", "–", "–", "–",
      "Генерація структури опитування за темою через Google Gemini API"),
    compRow("Контейнеризація (Docker)",      "+", "–", "–", "–",
      "Розгортання однією командою docker-compose up на будь-якому сервері"),
  ],
}),
sp(),
body("Аналіз таблиці 1.1 демонструє, що розроблена система Survey App поєднує переваги відкритого коду з функціональністю, що не поступається комерційним рішенням. При цьому система пропонує унікальні можливості, відсутні у конкурентів: публічний каталог опитувань, AI-генерацію через Google Gemini API та теплову карту активності респондентів."),

h3("1.3.2 Аналіз відомих алгоритмічних та технічних рішень"),
body("При розробці веборієнтованих систем онлайн-опитувань ключовими технічними задачами є: організація клієнт-серверної взаємодії, вибір архітектурного підходу для побудови інтерфейсу, організація безпечної автентифікації, вибір системи керування базами даних та реалізація модуля штучного інтелекту."),
body("Архітектурний підхід SPA (Single Page Application) із REST API є домінуючим рішенням для сучасних веб-застосунків подібного класу. SPA-підхід забезпечує миттєву реакцію інтерфейсу без перезавантаження сторінки, що критично важливо для конструктора опитувань з drag-and-drop та динамічним оновленням прев'ю [4]. Альтернативою є підхід SSR (Server-Side Rendering, наприклад Next.js), однак для інтерактивного конструктора опитувань SPA демонструє кращі характеристики за рахунок збереження стану компонентів між взаємодіями."),
body("Для серверної частини порівняно два основних підходи: Node.js + Express та Python + FastAPI. Node.js обрано з таких міркувань: єдина мова (TypeScript) для клієнтської та серверної частин дозволяє використовувати спільні типи інтерфейсів між рівнями системи; асинхронна модель виконання Node.js ефективна для I/O-навантажених систем (опитування, збір відповідей); велика екосистема npm-пакетів забезпечує широкий вибір готових рішень для всіх потреб системи [5]."),
body("Для організації доступу до реляційної бази даних обрано ORM Prisma як альтернативу прямим SQL-запитам та традиційним ORM (Sequelize, TypeORM). Prisma надає типобезпечний клієнт, що автоматично генерується на основі декларативної схеми, підтримує автоматичну генерацію SQL-міграцій та забезпечує зручний інтерфейс Prisma Studio для перегляду даних під час розробки [6]."),
body("Для автентифікації користувачів обрано механізм JWT (JSON Web Tokens) як stateless-альтернативу серверним сесіям. JWT дозволяє горизонтально масштабувати серверну частину без потреби у спільному сховищі сесій між екземплярами сервера. Паролі зберігаються у вигляді bcrypt-хешів з коефіцієнтом складності 12, що забезпечує адаптивний захист від атак перебору [7]."),
body("Для реалізації модуля генерації опитувань засобами штучного інтелекту проаналізовано доступні рішення: OpenAI GPT-4, Google Gemini та Anthropic Claude. Порівняльний аналіз виявив, що Google Gemini 1.5 Flash є оптимальним вибором для даного проєкту: безкоштовний ліміт складає 1 500 запитів на добу (достатньо для навчального проєкту), час відповіді — менше 3 секунд, якість генерації структурованого JSON — висока [8]. Модель отримує текстовий опис теми опитування та повертає готову структуру JSON з питаннями правильних типів, варіантами відповідей та налаштуваннями, яка автоматично завантажується в конструктор опитувань."),

// References for 1.3.2
sp(),
new Paragraph({ spacing:{after:80}, children:[new TextRun({text:"4. Fink O. React Design Patterns and Best Practices. 2nd ed. Birmingham : Packt Publishing, 2019. 420 p.", font:FONT, size:22})] }),
new Paragraph({ spacing:{after:80}, children:[new TextRun({text:"5. Tilkov S., Vinoski S. Node.js: Using JavaScript to Build High-Performance Network Programs. IEEE Internet Computing. 2010. Vol. 14(6). P. 80–83.", font:FONT, size:22})] }),
new Paragraph({ spacing:{after:80}, children:[new TextRun({text:"6. Prisma ORM Documentation. URL: https://www.prisma.io/docs (дата звернення: 10.04.2026).", font:FONT, size:22})] }),
new Paragraph({ spacing:{after:80}, children:[new TextRun({text:"7. Provos N., Mazières D. A Future-Adaptable Password Scheme. USENIX Annual Technical Conference. 1999. P. 1–9.", font:FONT, size:22})] }),
new Paragraph({ spacing:{after:80}, children:[new TextRun({text:"8. Google Gemini API Documentation. URL: https://ai.google.dev/docs (дата звернення: 15.04.2026).", font:FONT, size:22})] }),

h2("1.4 Аналіз та моделювання бізнес-процесів"),
body("Для побудови програмного забезпечення визначено та змодельовано ключові бізнес-процеси системи онлайн-опитувань. Моделювання виконано з використанням нотації BPMN (Business Process Model and Notation), що є міжнародним стандартом для графічного опису бізнес-процесів. Розглянуто три ключові бізнес-процеси: реєстрація та автентифікація користувача, створення та публікація опитування, а також проходження опитування та збір відповідей."),

h3("1.4.1 Бізнес-процес реєстрації та автентифікації"),
body("Бізнес-процес реєстрації та автентифікації користувача охоплює дії від першого відкриття системи до успішного входу в особистий кабінет (рис. 1.4)."),
body("Опис моделі бізнес-процесу реєстрації:"),
num("Користувач переходить на сторінку реєстрації (/register);"),
num("Користувач заповнює поля реєстраційної форми: ім'я, email та пароль (мін. 8 символів);"),
num("Система виконує клієнтську валідацію форми: якщо поля не відповідають вимогам, відображаються повідомлення про помилки без відправки запиту;"),
num("Після успішної клієнтської валідації система надсилає POST-запит до /api/auth/register;"),
num("Сервер перевіряє унікальність email в базі даних: якщо email вже зареєстрований — повертає помилку 409;"),
num("При унікальному email сервер хешує пароль алгоритмом bcrypt (12 раундів), зберігає користувача та повертає JWT-токен;"),
num("Токен зберігається у localStorage клієнта, користувач перенаправляється на /dashboard."),
...phBPMN("Рисунок 1.4 – BPMN-модель бізнес-процесу реєстрації та автентифікації користувача"),

h3("1.4.2 Бізнес-процес створення та публікації опитування"),
body("Бізнес-процес створення та публікації опитування охоплює повний цикл роботи автора від ініціювання нового опитування до його доступності для респондентів (рис. 1.5). Бізнес-процес підтримує два режими ініціювання: ручне створення в конструкторі та AI-асистоване створення через модуль генерації Gemini."),
body("Опис моделі бізнес-процесу:"),
num("Авторизований користувач з роллю Creator або Admin переходить до конструктора опитувань (/surveys/new) або обирає шаблон у галереї;"),
num("(Опціонально) Автор вводить тему опитування та запускає AI-генерацію: система надсилає запит до Google Gemini API, який повертає готову структуру питань; конструктор автоматично заповнюється отриманими даними;"),
num("Автор заповнює або коригує загальні налаштування: назву, опис, дати початку та кінця, параметри анонімності;"),
num("Автор додає або редагує питання, вибираючи тип кожного питання з 10 доступних типів та налаштовуючи параметри;"),
num("Для кожного питання автор може налаштувати умовну логіку (skip logic): умову спрацювання та цільове питання;"),
num("Система виконує валідацію: перевіряє наявність назви та хоча б одного питання з непустим текстом;"),
num("При успішній валідації автор може зберегти опитування як чернетку (статус DRAFT) або одразу опублікувати (статус ACTIVE);"),
num("Після публікації система генерує унікальне посилання та QR-код; автор налаштовує рівень доступу."),
...phBPMN("Рисунок 1.5 – BPMN-модель бізнес-процесу створення та публікації опитування"),

h3("1.4.3 Бізнес-процес проходження опитування"),
body("Бізнес-процес проходження опитування охоплює дії респондента від переходу за посиланням до успішного відправлення відповідей (рис. 1.6)."),
body("Опис моделі бізнес-процесу:"),
num("Респондент переходить за посиланням виду /s/{survey-id};"),
num("Система перевіряє статус опитування: якщо статус не ACTIVE або порушено часові обмеження — відображається відповідне повідомлення;"),
num("Якщо опитування захищено паролем (visibility = PASSWORD), відображається форма введення пароля; при невірному паролі доступ заблоковано;"),
num("Система завантажує збережений прогрес (якщо є) з бази даних за ключем сесії;"),
num("Респондент послідовно відповідає на питання; після кожної відповіді клієнт виконує оцінку умовної логіки — визначає наступне питання;"),
num("Прогрес автоматично зберігається кожні 2 секунди через API (/api/logic/surveys/{id}/progress);"),
num("Після відповіді на останнє питання респондент підтверджує відправлення форми;"),
num("Система перевіряє заповненість всіх обов'язкових питань та обмеження на повторні відповіді;"),
num("При успішній валідації відповіді зберігаються в базі даних, відображається сторінка подяки."),
...phBPMN("Рисунок 1.6 – BPMN-модель бізнес-процесу проходження опитування"),

h2("Висновки до розділу 1"),
body("У першому розділі виконано передпроєктне обстеження предметної області веборієнтованих систем онлайн-опитувань."),
body("Визначено повний перелік завдань дипломного проєктування з 14 пунктів, що охоплюють повний цикл розробки програмного забезпечення — від аналізу предметної області до створення документації."),
body("Проведено аналіз предметної області, що дозволив виявити основні категорії користувачів систем онлайн-опитувань (бізнес, HR, освіта, дослідники), їх специфічні потреби та ключові функціональні блоки сучасної системи опитувань. Встановлено суперечність між зростаючими потребами організацій у гнучких інструментах та обмеженнями існуючих комерційних рішень."),
body("Виконано порівняльний аналіз трьох провідних систем: Google Forms, SurveyMonkey та Typeform за 15 критеріями (таблиця 1.1). Показано, що розроблена система Survey App перевершує всі аналоги за ключовими критеріями та пропонує унікальні можливості, відсутні у конкурентів: публічний каталог опитувань з пошуком та тегами, інтеграцію AI-генерації через Google Gemini API, теплову карту активності та систему 5 рівнів видимості."),
body("Проаналізовано технічні рішення, покладені в основу розробки. Обґрунтовано вибір: SPA-архітектура (React + TypeScript) для клієнтської частини, Node.js + Express для серверної, Prisma ORM для типобезпечного доступу до PostgreSQL, JWT + bcrypt для автентифікації та Google Gemini API для AI-генерації опитувань. Для кожного технічного рішення виконано порівняльний аналіз з альтернативами та надано обґрунтування вибору."),
body("Змодельовано три ключові бізнес-процеси системи з використанням нотації BPMN: реєстрацію користувача, створення та публікацію опитування (включаючи AI-асистований режим через Gemini), проходження опитування та збір відповідей. Для кожного процесу описано основний та альтернативні потоки виконання."),

// ── СПИСОК ДЖЕРЕЛ ─────────────────────────────────────────────────────────────
pb(),
new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:300}, children:[
  new TextRun({ text:"СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ ДО РОЗДІЛУ 1", font:FONT, size:24, bold:true }),
]}),
...[
  "1. Couper M. P. Designing Effective Web Surveys. Cambridge : Cambridge University Press, 2008. 344 p.",
  "2. Evans J. R., Mathur A. The value of online surveys. Internet Research. 2005. Vol. 15(2). P. 195–219.",
  "3. Peytchev A. Survey Breakoff. Public Opinion Quarterly. 2009. Vol. 73(1). P. 74–97.",
  "4. Fink O. React Design Patterns and Best Practices. 2nd ed. Birmingham : Packt Publishing, 2019. 420 p.",
  "5. Tilkov S., Vinoski S. Node.js: Using JavaScript to Build High-Performance Network Programs. IEEE Internet Computing. 2010. Vol. 14(6). P. 80–83.",
  "6. Prisma ORM Documentation. URL: https://www.prisma.io/docs (дата звернення: 10.04.2026).",
  "7. Provos N., Mazières D. A Future-Adaptable Password Scheme. USENIX Annual Technical Conference. 1999. P. 1–9.",
  "8. Google Gemini API Documentation. URL: https://ai.google.dev/docs (дата звернення: 15.04.2026).",
].map(src => new Paragraph({
  spacing: { after: 80 },
  alignment: AlignmentType.JUSTIFIED,
  children: [new TextRun({ text: src, font: FONT, size: 22 })],
})),

    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("./Розділ_1_ВИПРАВЛЕНИЙ.docx", buf);
  console.log("✅ Done");
});