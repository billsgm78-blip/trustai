const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/6378ebcf-5099-4f66-82d1-128988cf6763_4c448735-4c0f-40ca-948d-8dbb6a601d9d/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "AniqlaAI_taqdimot.pptx";
const THEME = {
  name: "AniqlaAI",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "0F172A", lt1: "FFFFFF", dk2: "17133D", lt2: "F2F1FE",
    accent1: "6D5DF6", accent2: "0EA5C6", accent3: "E5484D", accent4: "F59E0B", accent5: "16A34A", accent6: "64748B",
    hlink: "6D5DF6", folHlink: "4C3FD1",
  },
};
const HEX = THEME.colors;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "AniqlaAI — sunʼiy intellekt yaratgan kontentni aniqlash";
  pres.author = "AniqlaAI";
  const C = pres.SchemeColor;

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "DARK",
    background: { color: C.text2 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.4, 0.5, 0.45, 0.5],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.7, fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "AniqlaAI · “Raqamli avlod qizlari” · Tuman bosqichi", options: { x: 0.5, y: 5.2, w: 6.5, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
  });

  const I = {
    image: await icon(fa.FaImage, HEX.lt1), video: await icon(fa.FaVideo, HEX.lt1), voice: await icon(fa.FaMicrophone, HEX.lt1), text: await icon(fa.FaFileAlt, HEX.lt1),
    news: await icon(fa.FaNewspaper, HEX.lt1), phone: await icon(fa.FaPhoneAlt, HEX.lt1), school: await icon(fa.FaGraduationCap, HEX.lt1), lang: await icon(fa.FaLanguage, HEX.lt1),
    upload: await icon(fa.FaCloudUploadAlt, HEX.lt1), chip: await icon(fa.FaMicrochip, HEX.lt1), brain: await icon(fa.FaBrain, HEX.lt1), check: await icon(fa.FaCheckCircle, HEX.lt1),
    lock: await icon(fa.FaLock, HEX.lt1), star: await icon(fa.FaStar, HEX.lt1), layers: await icon(fa.FaLayerGroup, HEX.lt1), eye: await icon(fa.FaSearch, HEX.lt1),
    users: await icon(fa.FaUsers, HEX.lt1), road: await icon(fa.FaRoad, HEX.lt1), coins: await icon(fa.FaCoins, HEX.lt1), shield: await icon(fa.FaShieldAlt, HEX.accent1),
    shieldW: await icon(fa.FaShieldAlt, HEX.lt1), question: await icon(fa.FaQuestionCircle, HEX.accent1),
  };
  // icon in a coloured circle (the deck's motif)
  const badge = (s, img, x, y, d, fill, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: name + " circle" });
    const p = d * 0.25;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name + " icon", altText: name });
  };
  const card = (s, x, y, w, h, name, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 }, objectName: name });
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top" }, o));

  // ---------- 1. Title ----------
  pres.addSection({ title: "Kirish" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Kirish" });
  badge(s, I.shieldW, 0.6, 0.55, 0.9, C.accent1, "Logo");
  T(s, "“Raqamli avlod qizlari” tanlovi · Tuman bosqichi · 2026", { x: 1.7, y: 0.75, w: 7.5, h: 0.5, fontSize: 14, color: "C9C4FF" });
  T(s, "AniqlaAI", { x: 0.6, y: 1.65, w: 8.8, h: 1.0, fontSize: 54, bold: true, color: C.background1 });
  T(s, "Sunʼiy intellekt yaratgan rasm, video, ovoz va matnni aniqlovchi platforma", { x: 0.6, y: 2.65, w: 8.4, h: 0.9, fontSize: 22, color: "E6E3FF" });
  card(s, 0.6, 3.85, 2.9, 0.5, "Direction chip", C.accent1);
  T(s, "Yoʻnalish: Sunʼiy intellekt", { x: 0.6, y: 3.85, w: 2.9, h: 0.5, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  T(s, "Muallif: [Ism Familiya] · [Maktab] · [Tuman]", { x: 0.6, y: 4.65, w: 8.8, h: 0.4, fontSize: 14, color: "C9C4FF" });
  s.addNotes("Assalomu alaykum. Mening loyiham — AniqlaAI. Bu sunʼiy intellekt yaratgan rasm, video, ovoz va matnni aniqlab beradigan platforma. Loyiha “Sunʼiy intellekt” yoʻnalishiga tegishli. (≈15 soniya)");

  // ---------- 2. Problem ----------
  pres.addSection({ title: "Muammo va maqsad" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Muammo va maqsad" });
  s.addText("Muammo: haqiqatni soxtadan ajratish qiyin", { placeholder: "title" });
  const probs = [
    [I.news, "Soxta rasm va videolar", "AI yaratgan suratlar yangilik sifatida tarqalmoqda", C.accent3],
    [I.phone, "Ovoz klonlash firibgarligi", "Tanish odam ovozida pul soʻrab qoʻngʻiroq qilinadi", C.accent4],
    [I.school, "AI yozgan insholar", "Oʻqituvchi matnni kim yozganini bila olmaydi", C.accent1],
    [I.lang, "Oʻzbek tilida vosita yoʻq", "Mavjud tekshiruvchilar asosan ingliz tilida va pullik", C.accent2],
  ];
  probs.forEach(([img, h, d, col], k) => {
    const x = 0.5 + (k % 2) * 4.6, y = 1.25 + Math.floor(k / 2) * 1.75;
    card(s, x, y, 4.4, 1.55, "Problem card " + (k + 1));
    badge(s, img, x + 0.25, y + 0.35, 0.8, col, h);
    T(s, h, { x: x + 1.25, y: y + 0.25, w: 3.0, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
    T(s, d, { x: x + 1.25, y: y + 0.72, w: 3.0, h: 0.7, fontSize: 14, color: C.text1 });
  });
  s.addNotes("Bugun sunʼiy intellekt bir necha soniyada haqiqiydek rasm, video, ovoz va matn yarata oladi. Natijada soxta yangiliklar, ovoz klonlash orqali firibgarlik va AI yozgan insholar paydo boʻldi. Oʻzbek tilida bularni tekshiradigan qulay vosita esa yoʻq. (≈25 soniya)");

  // ---------- 3. Goal & fit ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Muammo va maqsad" });
  s.addText("Maqsad va tanlov yoʻnalishiga mosligi", { placeholder: "title" });
  card(s, 0.5, 1.2, 4.4, 3.7, "Goal card", C.accent1);
  badge(s, I.question, 0.8, 1.45, 0.75, C.background1, "Goal");
  T(s, "Maqsad", { x: 1.75, y: 1.55, w: 2.9, h: 0.5, fontSize: 20, bold: true, color: C.background1 });
  T(s, "Har kim bepul va oson tekshira oladigan vosita yaratish: “Bu rasm, video, ovoz yoki matnni sunʼiy intellekt yaratganmi?”", { x: 0.8, y: 2.4, w: 3.85, h: 2.2, fontSize: 16, color: C.background1 });
  T(s, "Yoʻnalishga mosligi", { x: 5.2, y: 1.2, w: 4.3, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
  T(s, [
    { text: "Sunʼiy intellekt: loyiha neyron tarmoqlar asosida ishlaydi va AIning oʻzini aniqlaydi", options: { bullet: true, breakLine: true } },
    { text: "Taʼlim: oʻqituvchilar va media savodxonlik darslari uchun", options: { bullet: true } },
  ], { x: 5.2, y: 1.7, w: 4.3, h: 1.5, fontSize: 14, paraSpaceAfter: 6 });
  T(s, "Kimlar uchun", { x: 5.2, y: 3.3, w: 4.3, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  ["Oʻquvchi", "Oʻqituvchi", "Ota-ona", "Jurnalist"].forEach((t, k) => {
    const x = 5.2 + (k % 2) * 2.2, y = 3.8 + Math.floor(k / 2) * 0.6;
    card(s, x, y, 2.05, 0.48, "Audience " + t);
    T(s, t, { x, y, w: 2.05, h: 0.48, fontSize: 14, bold: true, color: C.accent1, align: "center", valign: "middle" });
  });
  s.addNotes("Maqsadim — har kim bepul va oson foydalana oladigan tekshiruv vositasi yaratish. Loyiha “Sunʼiy intellekt” yoʻnalishiga toʻliq mos: u neyron tarmoqlar yordamida sunʼiy intellekt izlarini topadi. Taʼlim sohasida ham foydali. (≈20 soniya)");

  // ---------- 4. Solution ----------
  pres.addSection({ title: "Yechim" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Yechim" });
  s.addText("Yechim: AniqlaAI — 4 ta tekshiruv bitta joyda", { placeholder: "title" });
  const mods = [
    [I.image, "Rasm", "AI yaratgan rasm, tahrir va montaj izlari", C.accent1],
    [I.video, "Video", "Kadrlar va video ichidagi ovoz", C.accent2],
    [I.voice, "Ovoz", "Sintez qilingan va klonlangan ovoz", C.accent4],
    [I.text, "Matn", "ChatGPT kabi AI yozgan matn", C.accent3],
  ];
  mods.forEach(([img, h, d, col], k) => {
    const x = 0.5 + k * 2.3;
    card(s, x, 1.25, 2.1, 2.45, "Module " + h);
    badge(s, img, x + 0.6, 1.45, 0.9, col, h);
    T(s, h, { x: x + 0.15, y: 2.5, w: 1.8, h: 0.45, fontSize: 18, bold: true, color: C.text2, align: "center" });
    T(s, d, { x: x + 0.15, y: 2.95, w: 1.8, h: 0.7, fontSize: 14, align: "center" });
  });
  T(s, "Natija har doim tushunarli:", { x: 0.5, y: 3.95, w: 9, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
  [["AI yaratgan", C.accent3], ["Ehtimol AI", C.accent4], ["Noaniq", C.accent6], ["AI belgisi topilmadi", C.accent5]].forEach(([t, col], k) => {
    const x = 0.5 + k * 2.3;
    card(s, x, 4.35, 2.1, 0.55, "Verdict " + t, col);
    T(s, t, { x, y: 4.35, w: 2.1, h: 0.55, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  });
  s.addNotes("AniqlaAI to‘rtta bo‘limdan iborat: rasm, video, ovoz va matn. Har biri alohida tekshiriladi. Natija foiz va oddiy xulosa bilan beriladi: “AI yaratgan”, “Ehtimol AI”, “Noaniq” yoki “AI belgisi topilmadi”. (≈20 soniya)");

  // ---------- 5. How it works ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Yechim" });
  s.addText("Qanday ishlaydi: 4 qadam", { placeholder: "title" });
  const steps = [
    [I.upload, "1. Yuklash", "Fayl yoki matn saytga qoʻyiladi"],
    [I.lock, "2. Brauzerda tahlil", "Fayl internetga yuborilmaydi"],
    [I.brain, "3. Neyron tarmoq", "AI modellari va kriminalistik tekshiruvlar"],
    [I.check, "4. Xulosa", "Foiz, xulosa va har bir belgining izohi"],
  ];
  steps.forEach(([img, h, d], k) => {
    const x = 0.5 + k * 2.35;
    badge(s, img, x + 0.55, 1.35, 1.0, k === 3 ? C.accent5 : C.accent1, h);
    if (k < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 1.75, y: 1.7, w: 0.45, h: 0.3, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "Arrow " + (k + 1) });
    T(s, h, { x, y: 2.55, w: 2.1, h: 0.45, fontSize: 16, bold: true, color: C.text2, align: "center" });
    T(s, d, { x, y: 3.0, w: 2.1, h: 0.8, fontSize: 14, align: "center" });
  });
  card(s, 0.5, 4.0, 9.0, 0.9, "Privacy note");
  badge(s, I.shield, 0.7, 4.1, 0.7, C.background1, "Privacy");
  T(s, "Maxfiylik: barcha tahlil foydalanuvchining oʻz qurilmasida bajariladi. Shaxsiy rasm, ovoz yoki matn hech qayerga yuborilmaydi.", { x: 1.6, y: 4.05, w: 7.7, h: 0.8, fontSize: 14, valign: "middle" });
  s.addNotes("Foydalanuvchi faylni yuklaydi. Tahlil to‘g‘ridan-to‘g‘ri brauzerda bajariladi, shuning uchun fayl hech qayerga yuborilmaydi — bu shaxsiy maʼlumotlar uchun muhim. Neyron tarmoqlar faylni tekshiradi va natija tushuntirish bilan chiqadi. (≈20 soniya)");

  // ---------- 6. Technology ----------
  pres.addSection({ title: "Texnik qism" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Texnik qism" });
  s.addText("Texnik xususiyatlar", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 } } });
  const rows = [
    [hdr("Boʻlim"), hdr("Neyron tarmoq (AI modeli)"), hdr("Qoʻshimcha tekshiruvlar")],
    ["Rasm", "SwinV2 — rasm 6 xil koʻrinishda tekshiriladi", "Metamaʼlumot, ELA, shovqin, klon, spektr"],
    ["Video", "SwinV2 — kadrlar + WavLM — ovozi", "Kadrlararo izchillik, ovozni ajratish"],
    ["Ovoz", "WavLM + oʻzim oʻqitgan klassifikator (2 520 ta ovoz)", "Spektrogramma, siqish va kesish izlari"],
    ["Matn", "RoBERTa + oʻzbek tili uslub tahlili", "Til, gap uzunligi, koʻrinmas belgilar"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.2, w: 9.0, colW: [1.3, 4.1, 3.6], fontSize: 13, fontFace: "Arial", color: C.text1, border: { type: "solid", pt: 0.75, color: "D9D6FB" }, fill: { color: C.background1 }, rowH: 0.5, valign: "middle", margin: 0.08 });
  T(s, "Platforma: HTML + JavaScript · ONNX Runtime Web (AI brauzerda ishlaydi) · bepul Vercel hosting · kompyuter va telefonda ishlaydi", { x: 0.5, y: 4.35, w: 9.0, h: 0.6, fontSize: 14, color: C.accent1, bold: true });
  s.addNotes("Har bir bo‘limda alohida neyron tarmoq bor: rasm uchun SwinV2, ovoz uchun WavLM asosidagi o‘zim o‘qitgan model, matn uchun RoBERTa va o‘zbek tili uslub tahlili. Modellar brauzerda ONNX Runtime orqali ishlaydi, sayt kompyuter va telefonda ochiladi. (≈30 soniya)");

  // ---------- 7. Innovation ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Texnik qism" });
  s.addText("Innovatsiya va yangilik", { placeholder: "title" });
  const inn = [
    [I.lang, "Oʻzbek tiliga moslangan", "Interfeys oʻzbekcha, matn detektori oʻzbek tili uslubini ham hisobga oladi"],
    [I.chip, "Oʻzim oʻqitgan ovoz detektori", "70 dan ortiq TTS tizimi va 4 tildagi haqiqiy ovozlarda oʻqitilgan"],
    [I.lock, "Maxfiylik", "AI modellari foydalanuvchi qurilmasida ishlaydi, fayl yuborilmaydi"],
    [I.eye, "Shaffof natija", "Faqat foiz emas: qaysi belgi topilgani oddiy tilda tushuntiriladi"],
  ];
  inn.forEach(([img, h, d], k) => {
    const x = 0.5 + (k % 2) * 4.6, y = 1.2 + Math.floor(k / 2) * 1.85;
    badge(s, img, x, y + 0.1, 0.85, k % 2 ? C.accent2 : C.accent1, h);
    T(s, h, { x: x + 1.05, y: y + 0.05, w: 3.35, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
    T(s, d, { x: x + 1.05, y: y + 0.5, w: 3.35, h: 1.0, fontSize: 14 });
  });
  s.addNotes("Loyihaning yangiligi: birinchidan, o‘zbek tiliga moslangan. Ikkinchidan, ovoz detektorini o‘zim o‘qitdim — tayyor modellar sinovdan o‘tmadi. Uchinchidan, AI to‘liq brauzerda ishlaydi va fayl yuborilmaydi. To‘rtinchidan, natija tushuntirish bilan beriladi. (≈25 soniya)");

  // ---------- 8. Results ----------
  pres.addSection({ title: "Natijalar" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Natijalar" });
  s.addText("Sinov natijalari", { placeholder: "title" });
  s.addChart(pres.charts.BAR, [{ name: "Xato ‘AI’ deyilgan haqiqiy rasmlar, %", labels: ["Birinchi versiya", "Hozirgi versiya"], values: [11.6, 0.9] }], {
    x: 0.5, y: 1.15, w: 4.6, h: 3.8, barDir: "col", chartColors: [HEX.accent3, HEX.accent5],
    showTitle: true, title: "Haqiqiy rasmni xato “AI” deyish, %", titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0.0",
    catAxisLabelColor: HEX.accent6, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 14,
  });
  const stats = [["0,91", "Ovoz detektori aniqligi", "AUC, yangi TTS tizimlarida"], ["5 / 8", "Oʻzbekcha AI matnlar topildi", "avvalgi versiya: 1 / 8"], ["~1%", "Haqiqiy ovozni xato “AI” deyish", "Telegram siqishidan keyin ham"]];
  stats.forEach(([n, l, sub], k) => {
    const y = 1.2 + k * 1.25;
    card(s, 5.4, y, 4.1, 1.1, "Stat " + (k + 1));
    T(s, n, { x: 5.6, y: y + 0.15, w: 1.5, h: 0.8, fontSize: 30, bold: true, color: C.accent1, valign: "middle" });
    T(s, l, { x: 7.15, y: y + 0.15, w: 2.25, h: 0.5, fontSize: 14, bold: true, color: C.text2 });
    T(s, sub, { x: 7.15, y: y + 0.62, w: 2.25, h: 0.4, fontSize: 11, color: C.accent6 });
  });
  s.addNotes("Loyihani minglab namunalarda sinadim. Haqiqiy rasmni xato “AI” deyish 11,6 foizdan 0,9 foizgacha kamaydi. Ovoz detektori o‘qitishda ko‘rmagan tizimlarda ham 0,91 aniqlik (AUC) ko‘rsatdi. O‘zbekcha AI matnlarni topish 8 tadan 1 tadan 5 taga oshdi. (≈30 soniya)");

  // ---------- 9. Comparison ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Natijalar" });
  s.addText("Mavjud analoglardan afzalligi", { placeholder: "title" });
  const yes = (t) => ({ text: t, options: { color: C.accent5, bold: true } });
  const no = (t) => ({ text: t, options: { color: C.accent3 } });
  const cmp = [
    [hdr("Xususiyat"), hdr("Xorijiy xizmatlar (koʻpchiligi)"), hdr("AniqlaAI")],
    ["Tekshiriladigan kontent", no("Odatda bitta tur (faqat matn yoki rasm)"), yes("Rasm, video, ovoz, matn")],
    ["Oʻzbek tili", no("Yoʻq"), yes("Interfeys va matn tahlili oʻzbekcha")],
    ["Fayl qayerda tekshiriladi", no("Serverga yuboriladi"), yes("Foydalanuvchi qurilmasida")],
    ["Narxi", no("Koʻpincha pullik obuna"), yes("Bepul")],
    ["Natija", no("Koʻpincha faqat foiz"), yes("Foiz + belgilar izohi")],
  ];
  s.addTable(cmp, { x: 0.5, y: 1.2, w: 9.0, colW: [2.6, 3.4, 3.0], fontSize: 14, fontFace: "Arial", color: C.text1, border: { type: "solid", pt: 0.75, color: "D9D6FB" }, fill: { color: C.background1 }, rowH: 0.55, valign: "middle", margin: 0.08 });
  T(s, "Masalan: GPTZero — faqat matn, Hive — asosan rasm va video; ikkalasi ham ingliz tilida va fayl serverga yuboriladi.", { x: 0.5, y: 4.82, w: 9.0, h: 0.3, fontSize: 11, color: C.accent6 });
  s.addNotes("Xorijda o‘xshash xizmatlar bor, masalan GPTZero va Hive. Lekin ular odatda bitta turdagi kontentni tekshiradi, o‘zbek tilini bilmaydi, faylni serverga yuboradi va ko‘pincha pullik. AniqlaAI to‘rt turni bepul, o‘zbek tilida va qurilmaning o‘zida tekshiradi. (≈25 soniya)");

  // ---------- 10. Practical use & expected results ----------
  pres.addSection({ title: "Amaliyot" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Amaliyot" });
  s.addText("Amaliyotga tatbiq va kutilayotgan natijalar", { placeholder: "title" });
  const uses = [
    [I.school, "Maktablarda", "Oʻqituvchi insho va referatlarni tekshiradi; media savodxonlik darslarida foydalaniladi"],
    [I.users, "Oilalarda", "Shubhali ovozli xabar yoki qoʻngʻiroq yozuvini tekshirib, firibgarlikdan saqlanish"],
    [I.news, "OAV va ijtimoiy tarmoqlarda", "Jurnalist va blogerlar rasm/videoni tarqatishdan oldin tekshiradi"],
  ];
  uses.forEach(([img, h, d], k) => {
    const y = 1.2 + k * 1.2;
    badge(s, img, 0.5, y + 0.1, 0.8, C.accent1, h);
    T(s, h, { x: 1.5, y: y + 0.05, w: 3.8, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    T(s, d, { x: 1.5, y: y + 0.45, w: 3.8, h: 0.7, fontSize: 14 });
  });
  card(s, 5.6, 1.2, 3.9, 3.7, "Expected results card", C.accent1);
  T(s, "Kutilayotgan natijalar", { x: 5.85, y: 1.4, w: 3.4, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
  T(s, [
    { text: "Soxta kontentga ishonish kamayadi", options: { bullet: true, breakLine: true } },
    { text: "Ovoz klonlash firibgarligidan himoya", options: { bullet: true, breakLine: true } },
    { text: "Oʻquvchilarning mustaqil ishi adolatli baholanadi", options: { bullet: true, breakLine: true } },
    { text: "Yoshlarning media savodxonligi oshadi", options: { bullet: true } },
  ], { x: 5.85, y: 1.95, w: 3.45, h: 2.8, fontSize: 14, color: C.background1, paraSpaceAfter: 8 });
  s.addNotes("AniqlaAI’dan maktablarda o‘qituvchilar, oilalarda ota-onalar va OAVda jurnalistlar foydalana oladi. Natijada odamlar soxta kontentga kamroq ishonadi, firibgarlikdan himoyalanadi va media savodxonligi oshadi. (≈25 soniya)");

  // ---------- 11. Economics & roadmap ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Amaliyot" });
  s.addText("Iqtisodiy jihat va rivojlanish rejasi", { placeholder: "title" });
  card(s, 0.5, 1.2, 4.2, 3.7, "Economics card");
  badge(s, I.coins, 0.75, 1.4, 0.75, C.accent4, "Economics");
  T(s, "Arzon va oʻzini oqlaydi", { x: 1.7, y: 1.5, w: 2.9, h: 0.5, fontSize: 16, bold: true, color: C.text2 });
  T(s, [
    { text: "Hosting bepul, server xarajati yoʻq — AI foydalanuvchi qurilmasida ishlaydi", options: { bullet: true, breakLine: true } },
    { text: "Foydalanuvchilar uchun bepul", options: { bullet: true, breakLine: true } },
    { text: "Kelajakda: maktab va tashkilotlar uchun kengaytirilgan versiya", options: { bullet: true } },
  ], { x: 0.8, y: 2.3, w: 3.7, h: 2.5, fontSize: 14, paraSpaceAfter: 8 });
  const plan = [["Hozir", "Sayt ishlaydi: rasm, video, ovoz, matn"], ["Viloyat bosqichi", "Telegram-bot, koʻproq oʻzbekcha AI namunalari"], ["Respublika bosqichi", "Mobil ilova va maktablar uchun versiya"]];
  badge(s, I.road, 5.0, 1.25, 0.6, C.accent2, "Roadmap");
  T(s, "Rivojlanish rejasi", { x: 5.75, y: 1.3, w: 3.7, h: 0.5, fontSize: 16, bold: true, color: C.text2 });
  plan.forEach(([h, d], k) => {
    const y = 2.05 + k * 0.95;
    s.addShape(pres.shapes.OVAL, { x: 5.1, y: y + 0.1, w: 0.4, h: 0.4, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Plan dot " + (k + 1) });
    T(s, String(k + 1), { x: 5.1, y: y + 0.1, w: 0.4, h: 0.4, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, h, { x: 5.75, y, w: 3.75, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
    T(s, d, { x: 5.75, y: y + 0.35, w: 3.75, h: 0.5, fontSize: 14 });
  });
  s.addNotes("Loyiha juda arzon: hosting bepul, AI foydalanuvchi qurilmasida ishlagani uchun server xarajati yo‘q. Rejam: viloyat bosqichigacha Telegram-bot va ko‘proq o‘zbekcha namunalar, keyin mobil ilova va maktablar uchun versiya. (≈25 soniya)");

  // ---------- 12. Conclusion ----------
  pres.addSection({ title: "Xulosa" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Xulosa" });
  T(s, "Xulosa", { x: 0.6, y: 0.5, w: 8.8, h: 0.7, fontSize: 36, bold: true, color: C.background1 });
  const fin = [["4 ta tekshiruv", "rasm, video, ovoz va matn"], ["Oʻzbek tilida", "bepul va maxfiy"], ["Sinovdan oʻtgan", "minglab namunalarda"]];
  fin.forEach(([h, d], k) => {
    const x = 0.6 + k * 3.0;
    card(s, x, 1.5, 2.75, 1.4, "Summary " + (k + 1), "2B2566");
    T(s, h, { x: x + 0.2, y: 1.65, w: 2.35, h: 0.5, fontSize: 18, bold: true, color: "C9C4FF" });
    T(s, d, { x: x + 0.2, y: 2.15, w: 2.35, h: 0.6, fontSize: 14, color: C.background1 });
  });
  T(s, "AniqlaAI — haqiqatni soxtadan ajratishga yordam beradi", { x: 0.6, y: 3.2, w: 8.8, h: 0.6, fontSize: 22, bold: true, color: C.background1 });
  T(s, "Sayt: [sayt manzili]   ·   Muallif: [Ism Familiya]", { x: 0.6, y: 3.95, w: 8.8, h: 0.4, fontSize: 14, color: "C9C4FF" });
  T(s, "E’tiboringiz uchun rahmat! Savollaringizga javob beraman.", { x: 0.6, y: 4.6, w: 8.8, h: 0.5, fontSize: 18, color: C.background1 });
  s.addNotes("Xulosa qilib aytganda, AniqlaAI to‘rt turdagi kontentni o‘zbek tilida, bepul va maxfiy tekshiradi va minglab namunalarda sinovdan o‘tgan. E’tiboringiz uchun rahmat, savollaringizga javob beraman. (≈15 soniya)");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
