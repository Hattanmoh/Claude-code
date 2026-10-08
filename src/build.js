// Language Champion Program Toolkit
// Builds a 23-slide, fully editable PowerPoint in the AMP-2 Lucid house style.
// Run: npm install && npm run build   ->   output/Language_Champion_Program_Toolkit.pptx

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("./apply-theme");

const ROOT = path.join(__dirname, "..");
const OUT = path.join(ROOT, "output", "Language_Champion_Program_Toolkit.pptx");
const LOGO_NAVY = path.join(ROOT, "assets", "lucid-navy.png");
const LOGO_WHITE = path.join(ROOT, "assets", "lucid-white.png");
const EMAIL = fs.readFileSync(path.join(ROOT, "content", "announcement-email.txt"), "utf8").trim();

// ---------- design tokens (see docs/design-spec.md)
const THEME = {
  name: "Language Champion AMP-2 Lucid",
  headFontFace: "Georgia",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1C2541", // ink: body text
    lt1: "FFFFFF", // white: main background
    dk2: "22335A", // navy: dark slides, titles
    lt2: "F4EEE5", // light sand: alternate background, cards
    accent1: "D9772B", // orange: icon circles, hero numbers
    accent2: "E5D7C1", // sand: text on navy, borders
    accent3: "4A5F86", // slate: secondary fills
    accent4: "A4511A", // burnt orange: eyebrows, labels, buttons
    accent5: "5B6478", // muted: footer, captions
    accent6: "C9B79C", // taupe: connectors, bars
    hlink: "A4511A",
    folHlink: "4A5F86",
  },
};
const HEX = THEME.colors;
const SERIF = "Georgia";
const FOOT = "LANGUAGE CHAMPION  ·  PROGRAM TOOLKIT  ·  AMP-2 LUCID KSA";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Language Champion Program Toolkit";
pres.author = "AMP-2 Learning and Development";
pres.company = "AMP-2 Lucid KSA";
const C = pres.SchemeColor;

// ---------- icons: Font Awesome rendered to PNG
const iconCache = {};
async function icon(name, hex) {
  const key = name + hex;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}
const shadow = () => ({ type: "outer", color: HEX.dk1, opacity: 0.12, blur: 10, offset: 3, angle: 90 });
const NOLINE = { type: "none" };

// ---------- slide layouts
const ph = (name, type, o) => ({ placeholder: { options: Object.assign({ name, type, margin: 0, valign: "top" }, o), text: "" } });
function contentLayout(title, bg) {
  pres.defineSlideMaster({
    title,
    background: { color: bg },
    objects: [
      ph("eyebrow", "body", { x: 0.6, y: 0.42, w: 9.6, h: 0.3, fontFace: "Arial", fontSize: 11, bold: true, color: C.accent4 }),
      ph("title", "title", { x: 0.6, y: 0.74, w: 10.4, h: 0.72, fontFace: SERIF, fontSize: 30, color: C.text2, align: "left" }),
      { image: { x: 11.33, y: 0.5, w: 1.4, h: 0.089, path: LOGO_NAVY } },
      { text: { text: FOOT, options: { x: 0.6, y: 7.0, w: 7.5, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.accent5, margin: 0, charSpacing: 1 } } },
    ],
    slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.accent5, align: "right" },
  });
}
contentLayout("CONTENT", C.background1);
contentLayout("CONTENT_SAND", C.background2);
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { image: { x: 0.8, y: 0.75, w: 1.6, h: 0.102, path: LOGO_WHITE } },
    ph("eyebrow", "body", { x: 0.8, y: 1.55, w: 6.4, h: 0.35, fontFace: "Arial", fontSize: 12, bold: true, color: C.accent2 }),
    ph("title", "title", { x: 0.8, y: 1.95, w: 6.6, h: 1.95, fontFace: SERIF, fontSize: 56, color: C.background1, align: "left" }),
    ph("tagline", "body", { x: 0.8, y: 4.0, w: 6.4, h: 0.55, fontFace: SERIF, fontSize: 26, italic: true, color: C.accent1 }),
    ph("sub", "body", { x: 0.8, y: 4.8, w: 5.9, h: 0.9, fontFace: "Arial", fontSize: 16, color: C.accent2 }),
  ],
});
pres.defineSlideMaster({
  title: "DIVIDER",
  background: { color: C.text2 },
  objects: [
    { image: { x: 11.33, y: 0.5, w: 1.4, h: 0.089, path: LOGO_WHITE } },
    ph("num", "body", { x: 0.8, y: 1.55, w: 4, h: 1.5, fontFace: SERIF, fontSize: 96, color: C.accent1 }),
    ph("title", "title", { x: 0.8, y: 3.15, w: 11, h: 0.95, fontFace: SERIF, fontSize: 44, color: C.background1, align: "left" }),
    ph("sub", "body", { x: 0.8, y: 4.25, w: 10, h: 0.7, fontFace: "Arial", fontSize: 18, color: C.accent2 }),
    { text: { text: FOOT, options: { x: 0.6, y: 7.0, w: 7.5, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.accent2, margin: 0, charSpacing: 1 } } },
  ],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.accent2, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING_DARK",
  background: { color: C.text2 },
  objects: [
    ph("eyebrow", "body", { x: 0.6, y: 0.42, w: 9.6, h: 0.3, fontFace: "Arial", fontSize: 11, bold: true, color: C.accent2 }),
    ph("title", "title", { x: 0.6, y: 0.74, w: 10.4, h: 0.72, fontFace: SERIF, fontSize: 30, color: C.background1, align: "left" }),
    { image: { x: 11.33, y: 0.5, w: 1.4, h: 0.089, path: LOGO_WHITE } },
    { text: { text: FOOT, options: { x: 0.6, y: 7.0, w: 7.5, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.accent2, margin: 0, charSpacing: 1 } } },
  ],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.25, fontFace: "Arial", fontSize: 9, color: C.accent2, align: "right" },
});

// ---------- helpers
let SECTION = "";
function section(t) { SECTION = t; pres.addSection({ title: t }); }
function newSlide(master, eyebrow, title) {
  const s = pres.addSlide({ masterName: master, sectionTitle: SECTION });
  if (eyebrow !== undefined) s.addText(eyebrow, { placeholder: "eyebrow" });
  if (title !== undefined) s.addText(title, { placeholder: "title" });
  return s;
}
function txt(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontFace: "Arial", color: C.text1, valign: "top" }, o)); }
function card(s, x, y, w, h, fill, opts = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign(
    { x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: NOLINE },
    opts.shadow ? { shadow: shadow() } : {},
    opts.line ? { line: opts.line } : {},
    opts.name ? { objectName: opts.name } : {}
  ));
}
async function iconCircle(s, name, x, y, d, fill = C.accent1, iconHex = "FFFFFF") {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: NOLINE });
  const i = d * 0.5;
  s.addImage({ data: await icon(name, iconHex), x: x + (d - i) / 2, y: y + (d - i) / 2, w: i, h: i, altText: name.replace(/^Fa/, "") + " icon" });
}
function label(s, text, x, y, w, color = C.accent4) {
  txt(s, text.toUpperCase(), { x, y, w, h: 0.28, fontSize: 11, bold: true, color, charSpacing: 1.5 });
}
function numCircle(s, n, x, y, d, fill = C.accent1, size = 16, font = SERIF, color = C.background1) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: NOLINE });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: size, fontFace: font, bold: font !== SERIF, color, align: "center", valign: "middle" });
}

(async () => {
  // ======================================================== OVERVIEW
  section("Overview");

  // 1 · Title
  {
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: SECTION });
    s.addText("AMP-2 LUCID KSA  ·  LEARNING & DEVELOPMENT", { placeholder: "eyebrow" });
    s.addText("Language Champion", { placeholder: "title" });
    s.addText("Learn. Connect. Belong.", { placeholder: "tagline" });
    s.addText("Program toolkit: pitch deck, announcement email and committee selection criteria", { placeholder: "sub" });
    s.addShape(pres.shapes.OVAL, { x: 7.55, y: 0.95, w: 5.5, h: 5.5, fill: { color: C.accent3, transparency: 55 }, line: NOLINE, objectName: "Backdrop circle" });
    const bubble = (x, y, w, word, lang, fill, rtl) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.05, rectRadius: 0.22, fill: { color: fill }, line: NOLINE, shadow: shadow(), objectName: "Greeting " + lang });
      s.addText([
        { text: word, options: { fontFace: rtl ? "Arial" : SERIF, fontSize: 26, color: C.text1, bold: !!rtl, rtlMode: !!rtl, breakLine: true } },
        { text: lang, options: { fontFace: "Arial", fontSize: 9, bold: true, color: C.text1, charSpacing: 2 } },
      ], { x, y: y + 0.12, w, h: 0.85, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    };
    bubble(7.95, 1.4, 2.3, "مرحبا", "ARABIC", C.background1, true);
    bubble(10.5, 1.8, 2.3, "Merhaba", "TURKISH", C.accent1);
    bubble(7.6, 2.8, 2.4, "Bonjour", "FRENCH", C.accent2);
    bubble(10.2, 3.2, 2.75, "خوش آمدید", "URDU", C.background1, true);
    bubble(8.3, 4.3, 2.0, "Hola", "SPANISH", C.accent1);
    s.addShape(pres.shapes.OVAL, { x: 10.95, y: 4.65, w: 1.4, h: 1.4, fill: { color: C.text2 }, line: { color: C.accent2, width: 1.5, dashType: "dash" }, rotate: -8, objectName: "Seal" });
    s.addText([
      { text: "10", options: { fontFace: SERIF, fontSize: 30, color: C.background1, breakLine: true } },
      { text: "CHAMPIONS", options: { fontFace: "Arial", fontSize: 9, bold: true, color: C.accent2, charSpacing: 1.5 } },
    ], { x: 10.95, y: 4.75, w: 1.4, h: 1.2, isTextBox: true, margin: 0, align: "center", valign: "middle" });
    s.addNotes("This toolkit has three parts: the pitch deck that introduces the program, the company-wide announcement email, and the committee's selection criteria. Language Champions are the fluent colleagues who teach; learners are the colleagues who join their classes. There are 10 Champion places: two for each of the five languages.");
  }

  // 2 · Inside this toolkit
  {
    const s = newSlide("CONTENT", "OVERVIEW", "Inside This Toolkit");
    const items = [
      ["01", "FaChalkboardTeacher", "The Pitch Deck", "Ten slides on the program, the five languages, the Champion role, selection and the timeline", "Slides 4–13"],
      ["02", "FaEnvelopeOpenText", "The Announcement Email", "A company-wide call for Champions with the registration link, benefits and selection steps", "Slides 15–16"],
      ["03", "FaClipboardCheck", "The Selection Criteria", "Eligibility gates, a 100-point scorecard, a 10-minute conversation and clear decision rules", "Slides 18–22"],
    ];
    const w = (12.133 - 0.7) / 3;
    for (let i = 0; i < 3; i++) {
      const [n, ic, t, d, ref] = items[i];
      const x = 0.6 + i * (w + 0.35), y = 1.85;
      card(s, x, y, w, 3.85, C.background2, { name: "Toolkit card " + n });
      txt(s, n, { x: x + 0.35, y: y + 0.32, w: 1.4, h: 0.7, fontFace: SERIF, fontSize: 40, color: C.accent4 });
      await iconCircle(s, ic, x + w - 1.15, y + 0.32, 0.8);
      txt(s, t, { x: x + 0.35, y: y + 1.25, w: w - 0.7, h: 0.8, fontFace: SERIF, fontSize: 20, color: C.text2 });
      txt(s, d, { x: x + 0.35, y: y + 2.15, w: w - 0.7, h: 1.05, fontSize: 14, color: C.text1 });
      card(s, x + 0.35, y + 3.28, 1.75, 0.4, C.background1);
      txt(s, ref, { x: x + 0.35, y: y + 3.28, w: 1.75, h: 0.4, fontSize: 12, bold: true, color: C.accent4, align: "center", valign: "middle" });
    }
    card(s, 0.6, 6.0, 12.133, 0.7, C.text2, { name: "Naming strip" });
    s.addText([
      { text: "Language Champions", options: { bold: true, color: C.accent2 } },
      { text: "  fluent colleagues who teach", options: { color: C.background1 } },
      { text: "        ·        ", options: { color: C.accent2 } },
      { text: "Learners", options: { bold: true, color: C.accent2 } },
      { text: "  colleagues who join their classes", options: { color: C.background1 } },
    ], { x: 0.6, y: 6.0, w: 12.133, h: 0.7, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 15, align: "center", valign: "middle" });
    s.addNotes("Use this slide to orient the audience. The naming strip matters: Champions teach, learners learn. Keep the terms consistent across the email, the Passport and every communication.");
  }

  // ======================================================== PART 1 · THE PITCH
  section("Part 1 · The Pitch");
  {
    const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: SECTION });
    s.addText("01", { placeholder: "num" });
    s.addText("The Pitch Deck", { placeholder: "title" });
    s.addText("Introduce the program and recruit our first 10 Language Champions", { placeholder: "sub" });
    s.addNotes("Part 1 is the pitch itself: present slides 4 to 13 at town halls, shift huddles and team meetings.");
  }

  // 4 · One team, many voices
  {
    const s = newSlide("CONTENT", "WHY NOW", "One Team, Many Voices");
    s.addShape(pres.shapes.OVAL, { x: 0.75, y: 1.85, w: 3.7, h: 3.7, fill: { color: C.background2 }, line: NOLINE, objectName: "Stat backdrop" });
    txt(s, "31+", { x: 0.75, y: 2.75, w: 3.7, h: 1.6, fontFace: SERIF, fontSize: 100, color: C.accent1, align: "center", valign: "middle" });
    txt(s, "NATIONALITIES WORK SIDE BY SIDE AT AMP-2", { x: 0.6, y: 5.75, w: 4.0, h: 0.6, fontSize: 13, bold: true, color: C.text2, charSpacing: 1, align: "center" });
    const rows = [
      ["FaGlobeAmericas", "Our diversity is a strength we use every day"],
      ["FaHandshake", "Shared language turns that strength into real connection across shops, departments and shifts"],
      ["FaUsers", "The teachers we need already work here: our own people"],
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.0 + i * 1.4;
      await iconCircle(s, rows[i][0], 5.9, y, 0.85, i === 1 ? C.text2 : C.accent1);
      txt(s, rows[i][1], { x: 7.05, y, w: 5.68, h: 0.85, fontSize: 19, color: C.text1, valign: "middle" });
    }
    s.addNotes("Open with the people, not the program. More than 31 nationalities work side by side at AMP-2. Optional: add a photo of a mixed team on the shop floor.");
  }

  // 5 · Five target languages
  {
    const s = newSlide("CONTENT_SAND", "THE FOCUS", "Our Five Target Languages");
    const L = [["مرحبا", "ARABIC", "marhaba · hello", true], ["Merhaba", "TURKISH", "merhaba · hello"], ["Bonjour", "FRENCH", "bonjour · good day"], ["خوش آمدید", "URDU", "khush amadeed · welcome", true], ["Hola", "SPANISH", "hola · hi"]];
    const w = (12.133 - 4 * 0.25) / 5;
    for (let i = 0; i < 5; i++) {
      const [g, l, m, rtl] = L[i];
      const x = 0.6 + i * (w + 0.25), y = 1.8;
      card(s, x, y, w, 2.5, C.background1, { shadow: true, name: "Language " + l });
      txt(s, g, { x, y: y + 0.35, w, h: 0.85, fontFace: rtl ? "Arial" : SERIF, fontSize: 30, bold: !!rtl, rtlMode: !!rtl, color: C.text2, align: "center", valign: "middle" });
      txt(s, l, { x, y: y + 1.35, w, h: 0.3, fontSize: 12, bold: true, color: C.accent4, align: "center", charSpacing: 2 });
      txt(s, m, { x: x + 0.15, y: y + 1.7, w: w - 0.3, h: 0.6, fontSize: 13, color: C.accent5, align: "center" });
    }
    card(s, 0.6, 4.55, 12.133, 1.8, C.background1, { name: "Shared words panel" });
    s.addText([
      { text: "WORDS WE ALREADY SHARE", options: { bold: true, color: C.accent4, charSpacing: 1.5 } },
      { text: "     Arabic · Turkish · Urdu · French · Spanish", options: { color: C.accent5 } },
    ], { x: 0.95, y: 4.75, w: 11.4, h: 0.3, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 11 });
    const words = [["Sugar", "sukkar · şeker · shakar · sucre · azúcar"], ["Coffee", "qahwa · kahve · kaafi · café · café"], ["Tea", "shay · çay · chai · thé · té"]];
    const cw = (12.133 - 0.7 - 0.8) / 3;
    words.forEach(([a, b], i) => {
      const x = 0.95 + i * (cw + 0.4);
      txt(s, a, { x, y: 5.2, w: cw, h: 0.42, fontFace: SERIF, fontSize: 20, color: C.text2 });
      txt(s, b, { x, y: 5.65, w: cw, h: 0.55, fontSize: 13, color: C.text1 });
    });
    txt(s, "Why these five: [add rationale, e.g., most-spoken languages on site]   ·   Each learner chooses one language for the full program", { x: 0.6, y: 6.5, w: 12.133, h: 0.3, fontSize: 11, italic: true, color: C.accent5 });
    s.addNotes("Every learner picks one of the five languages. Avoid national flags on this slide: each language is spoken across many countries. The shared-words strip is a quick confidence boost: the languages already overlap. Fill in the rationale for the five before presenting.");
  }

  // 6 · Peer-to-peer learning
  {
    const s = newSlide("CONTENT", "THE VALUE", "The Power of Peer-to-Peer Learning");
    const steps = [["Learn", "About 10 practical words per session from a fluent colleague"], ["Practice", "Games, role-play and Native Buddy chats"], ["Use", "Real conversations at work, in the canteen and on the line"], ["Record progress", "Every word logged and signed in the Language Passport"]];
    const bw = 2.7, gap = 0.444;
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * (bw + gap);
      numCircle(s, i + 1, x, 1.85, 0.8, C.accent1, 24);
      txt(s, steps[i][0], { x: x + 0.95, y: 1.85, w: bw - 0.95, h: 0.8, fontFace: SERIF, fontSize: 20, color: C.text2, valign: "middle" });
      txt(s, steps[i][1], { x, y: 2.85, w: bw, h: 0.8, fontSize: 13, color: C.text1 });
      if (i < 3) s.addShape(pres.shapes.CHEVRON, { x: x + bw + 0.1, y: 2.07, w: 0.24, h: 0.36, fill: { color: C.accent6 }, line: NOLINE });
    }
    label(s, "Why it works", 0.6, 3.9, 6);
    const vals = [["FaComments", "Authentic", "Real, everyday language, not textbook phrases"], ["FaBriefcase", "Practical", "Workplace-ready words from day one"], ["FaGlobeAmericas", "Cultural", "Culture shared with every word"], ["FaHandshake", "Connected", "Bonds across teams that rarely meet"], ["FaSeedling", "Home-grown", "Delivered by our own talent, with no external trainers"]];
    const w = (12.133 - 4 * 0.25) / 5;
    for (let i = 0; i < 5; i++) {
      const x = 0.6 + i * (w + 0.25), y = 4.25;
      card(s, x, y, w, 2.35, C.background2, { name: "Value " + vals[i][1] });
      await iconCircle(s, vals[i][0], x + 0.3, y + 0.3, 0.7, C.text2);
      txt(s, vals[i][1], { x: x + 0.3, y: y + 1.15, w: w - 0.6, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
      txt(s, vals[i][2], { x: x + 0.3, y: y + 1.5, w: w - 0.6, h: 0.8, fontSize: 13, color: C.text1 });
    }
    s.addNotes("Peer learning works because the teacher lives the language and knows the workplace. Each session follows the same loop: learn, practice, use, record.");
  }

  // 7 · How the program works
  {
    const s = newSlide("CONTENT", "THE PROGRAM", "How the Program Works");
    const stats = [["100", "WORDS PER LEARNER"], ["10", "SESSIONS"], ["2", "MONTHS, NOVEMBER TO DECEMBER"]];
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * 4.1;
      txt(s, stats[i][0], { x, y: 1.75, w: 1.75, h: 1.0, fontFace: SERIF, fontSize: 58, color: C.accent1, valign: "middle" });
      txt(s, stats[i][1], { x: x + 1.8, y: 1.75, w: 2.1, h: 1.0, fontSize: 13, bold: true, color: C.text2, valign: "middle", charSpacing: 1 });
    }
    const cats = ["Greetings", "Numbers", "Time", "Requests", "Directions", "Workplace", "Feelings", "Questions", "Culture", "Bonus"];
    const x0 = 0.95, step = (12.1 - 0.95) / 9, ly = 4.3;
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: 12.1 - x0, h: 0, line: { color: C.text2, width: 2 } });
    for (let i = 0; i < 10; i++) {
      const cx = x0 + i * step, even = (i + 1) % 2 === 0;
      numCircle(s, i + 1, cx - 0.26, ly - 0.26, 0.52, even ? C.accent1 : C.text2, 14, "Arial");
      txt(s, cats[i], { x: cx - 0.62, y: ly + 0.4, w: 1.24, h: 0.3, fontSize: 13, bold: true, color: C.text2, align: "center" });
      if (even) {
        card(s, cx - 0.6, ly - 0.95, 1.2, 0.38, C.accent4);
        txt(s, "Checkpoint " + (i + 1) / 2, { x: cx - 0.6, y: ly - 0.95, w: 1.2, h: 0.38, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
      }
    }
    const xEnd5 = x0 + 4 * step + 0.6;
    card(s, 0.6, 5.3, xEnd5 - 0.6, 0.45, C.background2);
    txt(s, "NOVEMBER  ·  SESSIONS 1–5", { x: 0.6, y: 5.3, w: xEnd5 - 0.6, h: 0.45, fontSize: 12, bold: true, color: C.text2, align: "center", valign: "middle", charSpacing: 1 });
    card(s, xEnd5 + 0.1, 5.3, 12.733 - xEnd5 - 0.1, 0.45, C.background2);
    txt(s, "DECEMBER  ·  SESSIONS 6–10", { x: xEnd5 + 0.1, y: 5.3, w: 12.733 - xEnd5 - 0.1, h: 0.45, fontSize: 12, bold: true, color: C.text2, align: "center", valign: "middle", charSpacing: 1 });
    txt(s, "About 10 words per session   ·   Progress tracked in each learner's Language Passport   ·   Recognition in December", { x: 0.6, y: 6.15, w: 12.133, h: 0.35, fontSize: 13, color: C.accent5 });
    s.addNotes("Ten sessions, about ten words each, grouped by practical theme. Champions sign a Passport checkpoint every two sessions: five checkpoints of 20 words make 100.");
  }

  // 8 · Meet the Champion
  {
    const s = newSlide("CONTENT", "THE ROLE", "Meet the Language Champion");
    card(s, 0.6, 1.75, 4.7, 4.85, C.text2, { name: "Champion profile panel" });
    await iconCircle(s, "FaChalkboardTeacher", 0.95, 2.1, 1.1, C.accent1);
    txt(s, "LANGUAGE CHAMPION", { x: 0.95, y: 3.5, w: 4.0, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 2 });
    txt(s, "Lead · Share · Support", { x: 0.95, y: 3.85, w: 4.1, h: 1.0, fontFace: SERIF, fontSize: 28, color: C.background1 });
    txt(s, "A fluent colleague who teaches their language, and the culture behind it, to fellow employees.", { x: 0.95, y: 4.95, w: 4.0, h: 1.3, fontSize: 15, color: C.accent2 });
    const rows = [["FaUserCheck", "WHO", "Native or fluent speakers of Arabic, Turkish, French, Urdu or Spanish"], ["FaUsers", "HOW MANY", "Two Champions per language: 10 in total"], ["FaBookOpen", "CLASS SIZE", "Each pair leads one class of 10–15 learners"], ["FaGraduationCap", "SUPPORT", "No teaching experience needed: L&D provides training and a session kit"]];
    for (let i = 0; i < 4; i++) {
      const y = 1.85 + i * 1.2;
      await iconCircle(s, rows[i][0], 5.75, y, 0.75);
      label(s, rows[i][1], 6.75, y, 5.9);
      txt(s, rows[i][2], { x: 6.75, y: y + 0.32, w: 5.98, h: 0.7, fontSize: 16, color: C.text1 });
    }
    s.addNotes("The Champion is a peer, not a professional teacher. Stress that L&D trains every Champion and provides the session kit, so fluency and goodwill are what matter.");
  }

  // 9 · Responsibilities
  {
    const s = newSlide("CONTENT_SAND", "THE ROLE", "Champion Roles and Responsibilities");
    const R = [["FaPencilRuler", "Co-design", "Plan the 10 sessions with your co-Champion and L&D"], ["FaChalkboardTeacher", "Lead", "Teach words, pronunciation, a game and a culture moment in each session"], ["FaPassport", "Track", "Sign off learners' Passport checkpoints every two sessions"], ["FaComments", "Encourage", "Give warm, constructive feedback, with no public ranking"], ["FaGlobeAmericas", "Share", "Bring the stories and customs behind the words"], ["FaCalendarCheck", "Commit", "Attend a 90-min orientation, a 2-hour planning workshop and 10 sessions of [session length]"]];
    const w = (12.133 - 0.6) / 3;
    for (let i = 0; i < 6; i++) {
      const x = 0.6 + (i % 3) * (w + 0.3), y = 1.8 + Math.floor(i / 3) * 2.5;
      card(s, x, y, w, 2.25, C.background1, { shadow: true, name: "Duty " + R[i][1] });
      await iconCircle(s, R[i][0], x + 0.35, y + 0.35, 0.75);
      txt(s, R[i][1], { x: x + 1.3, y: y + 0.35, w: w - 1.6, h: 0.75, fontFace: SERIF, fontSize: 22, color: C.text2, valign: "middle" });
      txt(s, R[i][2], { x: x + 0.35, y: y + 1.3, w: w - 0.7, h: 0.85, fontSize: 14, color: C.text1 });
    }
    s.addNotes("Six responsibilities, all supported: L&D provides the session template, and each language has two Champions who share the load. Confirm the session length before presenting.");
  }

  // 10 · What Champions gain
  {
    const s = newSlide("CONTENT", "THE BENEFITS", "What Champions Gain");
    const P = [["FaGraduationCap", "Skill Development", ["Facilitation and coaching training from L&D", "Public speaking and group leadership", "A ready-made session kit"]], ["FaAward", "Recognition", ["Official Champion pin and certificate", "Leadership appreciation", "Celebrated at the December event"]], ["FaNetworkWired", "Networking", ["Connections across every shop and department", "Friendships across 31+ nationalities", "Visibility with site leadership"]]];
    const w = (12.133 - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (w + 0.3), y = 1.8, dark = i === 1;
      card(s, x, y, w, 4.8, dark ? C.text2 : C.background2, { name: "Benefit " + P[i][1] });
      await iconCircle(s, P[i][0], x + 0.4, y + 0.4, 1.0, C.accent1);
      txt(s, P[i][1], { x: x + 0.4, y: y + 1.6, w: w - 0.8, h: 0.9, fontFace: SERIF, fontSize: 24, color: dark ? C.background1 : C.text2 });
      s.addText(P[i][2].map((t, j) => ({ text: t, options: { bullet: { indent: 16 }, breakLine: j < 2 } })), { x: x + 0.4, y: y + 2.65, w: w - 0.8, h: 2.0, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 15, color: dark ? C.accent2 : C.text1, paraSpaceAfter: 10, valign: "top" });
    }
    s.addNotes("Rewards are growth and recognition, not cash. The three pillars match the email: skill development, recognition and networking.");
  }

  // 11 · Selection process
  {
    const s = newSlide("CONTENT", "HOW TO JOIN", "The Selection Process, Step by Step");
    const S5 = [["Apply", "Submit the online application through the registration link", "By Mon 12 Oct"], ["Eligibility check", "HR confirms the five eligibility gates", "Tue 13 Oct"], ["Committee review", "HR, L&D and site leadership score each application", "Tue 13 Oct"], ["Conversation", "A friendly 10-minute chat and teach-back for the shortlist", "Tue 13 Oct"], ["Decision", "Two Champions and one reserve per language; everyone notified", "End of Tue 13 Oct"]];
    const cw = 12.133 / 5;
    s.addShape(pres.shapes.LINE, { x: 0.6 + cw / 2, y: 2.25, w: cw * 4, h: 0, line: { color: C.accent6, width: 2 } });
    for (let i = 0; i < 5; i++) {
      const cx = 0.6 + i * cw + cw / 2;
      numCircle(s, i + 1, cx - 0.4, 1.85, 0.8, C.accent1, 24);
      txt(s, S5[i][0], { x: cx - 1.1, y: 2.85, w: 2.2, h: 0.35, fontSize: 16, bold: true, color: C.text2, align: "center" });
      txt(s, S5[i][1], { x: cx - 1.1, y: 3.25, w: 2.2, h: 1.0, fontSize: 13, color: C.text1, align: "center" });
      card(s, cx - 0.9, 4.4, 1.8, 0.4, C.background2);
      txt(s, S5[i][2], { x: cx - 0.9, y: 4.4, w: 1.8, h: 0.4, fontSize: 12, bold: true, color: C.accent4, align: "center", valign: "middle" });
    }
    card(s, 0.6, 5.3, 12.133, 1.15, C.background2, { name: "Native Buddy note" });
    await iconCircle(s, "FaHandsHelping", 0.95, 5.47, 0.8, C.text2);
    txt(s, "Not selected this round? You'll be invited to be a Native Buddy, helping learners practice between sessions.", { x: 2.0, y: 5.3, w: 10.4, h: 1.15, fontSize: 17, color: C.text1, valign: "middle" });
    s.addNotes("The process is short and transparent: the criteria are published in the email and used by the committee. Unselected applicants stay involved as Native Buddies.");
  }

  // 12 · Timeline
  {
    const s = newSlide("CONTENT", "THE TIMELINE", "Program Timeline and Milestones");
    const M = [["SUN 11 OCT", "Applications open", "Announcement email and shift-huddle reveal"], ["MON 12 OCT", "Applications close", "End-of-shift deadline"], ["TUE 13 OCT", "Champions selected", "Committee review, conversations and notifications"], ["WED 14 OCT", "Orientation", "Champions trained by L&D"], ["SUN 18 OCT", "Co-design workshop", "10 sessions planned; kickoff date locked"], ["NOVEMBER", "Program kickoff", "Sessions 1–5; checkpoints 1–2"], ["DECEMBER", "Complete and celebrate", "Sessions 6–10, evaluation and recognition"]];
    const cw = 2.6, x0 = 0.6 + cw / 2, x1 = 12.733 - cw / 2, ly = 4.0, step = (x1 - x0) / 6;
    s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: C.text2, width: 2.5 } });
    for (let i = 0; i < 7; i++) {
      const cx = x0 + i * step, above = i % 2 === 0, prog = i >= 5;
      const x = Math.min(Math.max(cx - cw / 2, 0.6), 12.733 - cw);
      const y = above ? 1.75 : 4.45, h = 1.8;
      s.addShape(pres.shapes.LINE, { x: cx, y: above ? y + h : ly + 0.16, w: 0, h: above ? ly - 0.16 - (y + h) : y - (ly + 0.16), line: { color: C.accent6, width: 1.25 } });
      s.addShape(pres.shapes.OVAL, { x: cx - 0.16, y: ly - 0.16, w: 0.32, h: 0.32, fill: { color: prog ? C.accent1 : C.text2 }, line: { color: C.background1, width: 2 } });
      card(s, x, y, cw, h, prog ? C.text2 : C.background2, { name: "Milestone " + M[i][0] });
      txt(s, M[i][0], { x: x + 0.25, y: y + 0.2, w: cw - 0.5, h: 0.28, fontSize: 11, bold: true, color: prog ? C.accent2 : C.accent4, charSpacing: 1.5 });
      txt(s, M[i][1], { x: x + 0.25, y: y + 0.45, w: cw - 0.5, h: 0.6, fontFace: SERIF, fontSize: 16, color: prog ? C.background1 : C.text2 });
      txt(s, M[i][2], { x: x + 0.25, y: y + 1.1, w: cw - 0.5, h: 0.6, fontSize: 12, color: prog ? C.accent2 : C.text1 });
    }
    txt(s, "October: recruit and prepare   ·   November–December: the program runs", { x: 0.6, y: 6.4, w: 12.133, h: 0.3, fontSize: 12, color: C.accent5 });
    s.addNotes("October is recruitment and preparation; the program itself runs November to December. Add the kickoff date once it is locked at the co-design workshop on 18 October.");
  }

  // 13 · Call to action
  {
    const s = newSlide("CLOSING_DARK", "CALL TO ACTION", "Step Up. Your Language Matters.");
    const rows = [["FaLink", "APPLY", "[INSERT REGISTRATION LINK]"], ["FaClock", "DEADLINE", "Monday 12 October, end of shift"], ["FaComments", "QUESTIONS", "[HR contact name and email]"]];
    for (let i = 0; i < 3; i++) {
      const y = 2.0 + i * 1.15;
      await iconCircle(s, rows[i][0], 0.6, y, 0.8, C.accent1);
      txt(s, rows[i][1], { x: 1.65, y: y + 0.02, w: 6, h: 0.3, fontSize: 11, bold: true, color: C.accent2, charSpacing: 2 });
      txt(s, rows[i][2], { x: 1.65, y: y + 0.32, w: 6.5, h: 0.45, fontSize: 20, bold: i === 0, color: C.background1 });
    }
    txt(s, "“Every new word creates a new connection.”", { x: 0.6, y: 5.75, w: 12.133, h: 0.6, fontFace: SERIF, fontSize: 26, italic: true, color: C.accent1 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.3, y: 1.95, w: 3.3, h: 3.3, rectRadius: 0.15, fill: { color: C.text2 }, line: { color: C.accent2, width: 2, dashType: "dash" }, objectName: "QR placeholder" });
    s.addImage({ data: await icon("FaQrcode", HEX.accent2), x: 10.35, y: 2.75, w: 1.2, h: 1.2, altText: "QR code placeholder" });
    txt(s, "QR CODE", { x: 9.3, y: 4.15, w: 3.3, h: 0.3, fontSize: 12, bold: true, color: C.accent2, align: "center", charSpacing: 2 });
    txt(s, "Replace with the code for the registration link", { x: 9.55, y: 4.5, w: 2.8, h: 0.5, fontSize: 11, color: C.accent2, align: "center" });
    s.addNotes("Close the pitch with one action. Replace the QR placeholder with the code for the registration link before presenting.");
  }

  // ======================================================== PART 2 · EMAIL
  section("Part 2 · Announcement Email");
  {
    const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: SECTION });
    s.addText("02", { placeholder: "num" });
    s.addText("The Announcement Email", { placeholder: "title" });
    s.addText("A motivating, company-wide call for Language Champions", { placeholder: "sub" });
    s.addNotes("Part 2 is the announcement email. The full text is in the speaker notes of the next slide, ready to copy.");
  }

  // 15 · Email at a glance
  {
    const s = newSlide("CONTENT", "THE EMAIL", "The Email at a Glance");
    const plan = [["WHEN", "Sunday 11 October, right after the shift-huddle reveal"], ["FROM", "[Sender name and title]"], ["TO", "All AMP-2 employees"], ["ALSO", "Huddles, QR posters and screens for the shop floor"], ["SUBJECT", "Speak Arabic, Turkish, French, Urdu or Spanish? Become a Language Champion"], ["ALTERNATIVE SUBJECT", "Your language can connect our team. Apply by Monday 12 October"]];
    const planY = [1.75, 2.5, 3.25, 4.0, 4.75, 5.7]; // the two subject lines wrap, so they get more room
    for (let i = 0; i < plan.length; i++) {
      const y = planY[i];
      label(s, plan[i][0], 0.6, y, 4.3);
      txt(s, plan[i][1], { x: 0.6, y: y + 0.27, w: 4.3, h: 0.5, fontSize: 13, color: C.text1 });
    }
    const ex = 5.3, ey = 1.8, ew = 7.433, eh = 4.85;
    card(s, ex, ey, ew, eh, C.background1, { shadow: true, line: { color: C.accent2, width: 1 }, name: "Email mockup" });
    card(s, ex + 0.2, ey + 0.2, ew - 0.4, 0.85, C.background2);
    s.addText([
      { text: "Subject: ", options: { bold: true } },
      { text: "Speak Arabic, Turkish, French, Urdu or Spanish? Become a Language Champion", options: { breakLine: true } },
      { text: "From: ", options: { bold: true, color: C.accent5 } },
      { text: "[Sender name]", options: { color: C.accent5 } },
    ], { x: ex + 0.4, y: ey + 0.27, w: ew - 0.8, h: 0.72, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 12, color: C.text1, valign: "middle" });
    txt(s, "Dear colleagues,", { x: ex + 0.4, y: ey + 1.3, w: ew - 0.8, h: 0.3, fontSize: 13 });
    s.addText([
      { text: "Every day, more than 31 nationalities work side by side at AMP-2. Today we're launching " },
      { text: "Language Champion", options: { bold: true } },
      { text: ", a peer-to-peer program where colleagues teach colleagues: 100 words, 10 sessions, one language at a time.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: "We're looking for " },
      { text: "10 Language Champions", options: { bold: true } },
      { text: ", two for each of our five languages." },
    ], { x: ex + 0.4, y: ey + 1.65, w: ew - 0.8, h: 1.55, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 13, color: C.text1, valign: "top" });
    card(s, ex + 0.4, ey + 3.35, 5.7, 0.6, C.accent4, { name: "Apply button" });
    txt(s, "APPLY HERE  →  [INSERT REGISTRATION LINK]", { x: ex + 0.4, y: ey + 3.35, w: 5.7, h: 0.6, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, "Deadline: Monday 12 October, end of shift", { x: ex + 0.4, y: ey + 4.1, w: ew - 0.8, h: 0.35, fontSize: 13, bold: true, color: C.accent4 });
    s.addNotes("FULL EMAIL TEXT, ready to copy:\n\n" + EMAIL);
  }

  // 16 · Inside the email
  {
    const s = newSlide("CONTENT_SAND", "THE EMAIL", "Inside the Email");
    const w = (12.133 - 0.3) / 2, h = 2.3;
    const pos = [[0.6, 1.8], [0.6 + w + 0.3, 1.8], [0.6, 4.3], [0.6 + w + 0.3, 4.3]];
    const heads = [["FaLanguage", "Our Five Languages"], ["FaUserCheck", "Who Should Apply"], ["FaAward", "What You'll Gain"], ["FaClipboardCheck", "How Selection Works"]];
    for (let i = 0; i < 4; i++) {
      const [x, y] = pos[i];
      card(s, x, y, w, h, C.background1, { shadow: true, name: heads[i][1] });
      await iconCircle(s, heads[i][0], x + 0.3, y + 0.25, 0.55, C.text2);
      txt(s, heads[i][1], { x: x + 1.0, y: y + 0.25, w: w - 1.3, h: 0.55, fontFace: SERIF, fontSize: 19, color: C.text2, valign: "middle" });
    }
    {
      const [x, y] = pos[0];
      let px = x + 0.3;
      for (const l of ["Arabic", "Turkish", "French", "Urdu", "Spanish"]) {
        const pw = 0.25 + l.length * 0.115;
        card(s, px, y + 0.98, pw, 0.38, C.background2);
        txt(s, l, { x: px, y: y + 0.98, w: pw, h: 0.38, fontSize: 13, bold: true, color: C.text2, align: "center", valign: "middle" });
        px += pw + 0.15;
      }
      txt(s, "Greetings: marhaba · merhaba · bonjour · khush amadeed · hola", { x: x + 0.3, y: y + 1.48, w: w - 0.6, h: 0.3, fontSize: 12, color: C.accent5 });
      txt(s, "Want to learn instead? Learners register at [INSERT LEARNER REGISTRATION LINK]", { x: x + 0.3, y: y + 1.8, w: w - 0.6, h: 0.4, fontSize: 12, color: C.text1 });
    }
    const bl = (arr, x, y, numbered) => s.addText(
      arr.map((t, j) => ({ text: t, options: { bullet: numbered ? { type: "number" } : { indent: 14 }, breakLine: j < arr.length - 1 } })),
      { x: x + 0.3, y: y + 0.95, w: w - 0.6, h: 1.3, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 13, color: C.text1, paraSpaceAfter: 4, valign: "top" }
    );
    bl(["Native or fluent in one of the five languages", "Patient, encouraging and clear", "Free for all 10 sessions, plus 14 and 18 October", "No teaching experience needed: L&D trains you"], pos[1][0], pos[1][1]);
    bl(["Skill development: facilitation and coaching training", "Recognition: Champion pin, certificate, December celebration", "Networking: connections across every shop and nationality"], pos[2][0], pos[2][1]);
    bl(["Apply by Monday 12 October", "Committee scores against published criteria", "10-minute conversation on Tuesday 13 October", "Everyone hears back by end of Tuesday"], pos[3][0], pos[3][1], true);
    s.addNotes("The four blocks below the call to action. Keep the email scannable: bold mini-headings, short bullets, and the registration link near the top. Saying clearly that no teaching experience is needed widens the pool.");
  }

  // ======================================================== PART 3 · SELECTION
  section("Part 3 · Selection Criteria");
  {
    const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: SECTION });
    s.addText("03", { placeholder: "num" });
    s.addText("Committee Selection Criteria", { placeholder: "title" });
    s.addText("A fair, transparent framework for choosing our Language Champions", { placeholder: "sub" });
    s.addNotes("Part 3 is the committee's framework: gates, scorecard, conversation, decision rules and a scoring sheet.");
  }

  // 18 · Committee and gates
  {
    const s = newSlide("CONTENT", "STEP 1", "The Committee and the Eligibility Gates");
    label(s, "The selection committee", 0.6, 1.75, 6.6);
    const roles = [["FaUserTie", "Chair · HR", "Runs the process and records decisions"], ["FaGraduationCap", "L&D representative", "Assesses teaching and facilitation potential"], ["FaBriefcase", "Site leadership delegate", "Sponsor voice; supports manager release"], ["FaLanguage", "Native-speaker assessors", "One per language; rate proficiency only"]];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 3.4, y = 2.1 + Math.floor(i / 2) * 2.05;
      card(s, x, y, 3.2, 1.85, C.background2, { name: "Role " + roles[i][1] });
      await iconCircle(s, roles[i][0], x + 0.25, y + 0.25, 0.6, C.text2);
      txt(s, roles[i][1], { x: x + 1.0, y: y + 0.25, w: 2.05, h: 0.6, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      txt(s, "[Name]", { x: x + 0.25, y: y + 0.95, w: 2.7, h: 0.25, fontSize: 11, color: C.accent5 });
      txt(s, roles[i][2], { x: x + 0.25, y: y + 1.2, w: 2.75, h: 0.55, fontSize: 12, color: C.text1 });
    }
    txt(s, "Conflict rule: no member scores their own direct report.", { x: 0.6, y: 6.3, w: 6.6, h: 0.3, fontSize: 12, italic: true, color: C.accent5 });
    card(s, 7.55, 1.75, 5.183, 4.85, C.text2, { name: "Eligibility gates panel" });
    txt(s, "ELIGIBILITY GATES", { x: 7.9, y: 2.0, w: 4.5, h: 0.28, fontSize: 11, bold: true, color: C.accent2, charSpacing: 2 });
    txt(s, "Pass all five before scoring", { x: 7.9, y: 2.3, w: 4.6, h: 0.45, fontFace: SERIF, fontSize: 20, color: C.background1 });
    const gates = ["Native or fluent in one of the five languages", "Available for all 10 sessions, orientation and the workshop", "Line manager supports release", "In good standing, with no open disciplinary case", "Application submitted by the deadline"];
    for (let i = 0; i < 5; i++) {
      const y = 3.0 + i * 0.7;
      numCircle(s, i + 1, 7.9, y, 0.45, C.accent1, 13, "Arial", C.text1);
      txt(s, gates[i], { x: 8.55, y: y - 0.05, w: 3.95, h: 0.6, fontSize: 14, color: C.background1, valign: "middle" });
    }
    s.addNotes("Gates are pass or fail and come before any scoring. The native-speaker assessors solve a practical problem: most committee members will not speak all five languages.");
  }

  // 19 · Scorecard
  {
    const s = newSlide("CONTENT", "STEP 2", "The 100-Point Scorecard");
    const segs = [["Language proficiency · 30", 30, C.text2, C.background1], ["Teaching and mentoring · 25", 25, C.accent3, C.background1], ["Communication · 25", 25, C.accent4, C.background1], ["Availability · 20", 20, C.accent2, C.text1]];
    let x = 0.6;
    for (const [t, v, f, tc] of segs) {
      const w = 12.133 * v / 100;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.75, w, h: 0.55, fill: { color: f }, line: NOLINE, objectName: "Weight " + t });
      txt(s, t, { x, y: 1.75, w, h: 0.55, fontSize: 12, bold: true, color: tc, align: "center", valign: "middle" });
      x += w;
    }
    const hd = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 12 } });
    const rows = [
      [hd("Criterion"), hd("Weight"), hd("Strong applicant"), hd("Red flags"), hd("How to assess")],
      ["Language proficiency", "30", "Native or near-native; explains meaning, pronunciation and usage with ease", "Conversational only; unsure why words are used", "Native-speaker assessor's 3-minute check"],
      ["Willingness and ability to teach or mentor", "25", "Clear personal motivation; has trained, coached or onboarded others; patient with mistakes", "Motivated mainly by recognition; impatient or corrective tone", "Application answers and teach-back"],
      ["Communication and interpersonal skills", "25", "Explains simply in a shared language; warm and inclusive; checks understanding; works well with a co-Champion", "Hard to follow; talks over others", "Teach-back and scenario question"],
      ["Availability and commitment", "20", "Shift pattern fits session slots; manager supportive; known for following through", "Likely shift clashes; history of dropped commitments", "Application and manager input"],
    ].map((r, ri) => ri === 0 ? r : r.map((c, ci) => ({ text: c, options: { bold: ci === 0, color: ci === 1 ? C.accent4 : C.text1, fill: { color: ri % 2 ? C.background1 : C.background2 }, fontSize: ci === 1 ? 15 : 13 } })));
    s.addTable(rows, { x: 0.6, y: 2.55, w: 12.133, colW: [2.3, 0.95, 3.6, 2.75, 2.533], rowH: [0.45, 0.78, 0.78, 0.78, 0.78], fontFace: "Arial", fontSize: 13, color: C.text1, valign: "middle", margin: [6, 10, 6, 10], border: { type: "solid", pt: 0.75, color: HEX.accent2 }, objectName: "Scorecard table" });
    txt(s, "Rate each criterion 1–4 (see next slide).   Points = Weight × Rating ÷ 4.   Total = 100.", { x: 0.6, y: 6.55, w: 12.133, h: 0.3, fontSize: 12, color: C.accent5 });
    s.addNotes("Proficiency carries the most weight, but teaching willingness and communication together count for 50 points: a fluent speaker who cannot explain simply will struggle as a Champion.");
  }

  // 20 · Rating scale and conversation
  {
    const s = newSlide("CONTENT_SAND", "STEP 3", "Rating Scale and the 10-Minute Conversation");
    card(s, 0.6, 1.75, 5.6, 4.85, C.background1, { shadow: true, name: "Rating scale panel" });
    label(s, "Rating scale", 0.95, 2.0, 4.8);
    const R = [["4", "Excellent", 100], ["3", "Good", 75], ["2", "Developing", 50], ["1", "Limited", 25]];
    for (let i = 0; i < 4; i++) {
      const y = 2.45 + i * 0.68;
      txt(s, R[i][0], { x: 0.95, y, w: 0.5, h: 0.55, fontFace: SERIF, fontSize: 30, color: C.accent1, valign: "middle" });
      txt(s, R[i][1], { x: 1.5, y, w: 1.55, h: 0.55, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: 3.1, y: y + 0.15, w: 2.0 * R[i][2] / 100, h: 0.25, fill: { color: i === 0 ? C.text2 : C.accent6 }, line: NOLINE });
      txt(s, R[i][2] + "%", { x: 5.2, y, w: 0.8, h: 0.55, fontSize: 12, bold: true, color: C.accent5, valign: "middle" });
    }
    card(s, 0.95, 5.25, 4.9, 1.1, C.background2);
    txt(s, "Points = Weight × Rating ÷ 4", { x: 1.15, y: 5.35, w: 4.5, h: 0.4, fontFace: SERIF, fontSize: 17, color: C.text2 });
    txt(s, "Example: Communication rated 3 → 25 × 3 ÷ 4 = 18.75 points", { x: 1.15, y: 5.8, w: 4.6, h: 0.45, fontSize: 12, color: C.text1 });
    label(s, "The 10-minute conversation", 6.5, 1.75, 6.2);
    const segs = [[2, "2 min", C.text2], [3, "3 min", C.accent3], [3, "3 min", C.accent4], [2, "2 min", C.accent5]];
    let x = 6.5;
    for (const [m, t, f] of segs) {
      const w = 6.233 * m / 10;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 2.1, w, h: 0.5, fill: { color: f }, line: { color: C.background2, width: 1.5 } });
      txt(s, t, { x, y: 2.1, w, h: 0.5, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
      x += w;
    }
    const Q = [["1 · Motivation", "“Why do you want to be a Language Champion?”"], ["2 · Teach-back", "“Teach us three greetings as if we were brand-new learners.” The native-speaker assessor joins."], ["3 · Scenario", "“A learner is shy and keeps repeating the same mistake. What do you do?”"], ["4 · Practicalities", "Shift pattern, availability and the applicant's questions"]];
    for (let i = 0; i < 4; i++) {
      const y = 2.85 + i * 0.95;
      card(s, 6.5, y, 6.233, 0.85, C.background1, { name: "Conversation " + (i + 1) });
      txt(s, Q[i][0], { x: 6.75, y: y + 0.1, w: 5.8, h: 0.28, fontSize: 13, bold: true, color: C.text2 });
      txt(s, Q[i][1], { x: 6.75, y: y + 0.38, w: 5.8, h: 0.45, fontSize: 12, italic: i < 3, color: C.text1 });
    }
    s.addNotes("Score each criterion straight after the conversation, while it is fresh. The teach-back is the best single signal for both communication and teaching ability.");
  }

  // 21 · Decision workflow
  {
    const s = newSlide("CONTENT", "STEP 4", "The Decision Workflow");
    const steps = [["Score independently", "Each member rates every applicant alone first"], ["Calibrate", "Discuss scores together, one language at a time"], ["Apply the threshold", "Minimum 60 / 100 to be selected (suggested)"], ["Select", "Top two per language become Champions; the third is the reserve"], ["Record", "The chair logs each decision with a one-line rationale"], ["Notify", "Everyone hears back by end of Tue 13 Oct; others invited as Native Buddies"]];
    s.addShape(pres.shapes.LINE, { x: 0.85, y: 2.1, w: 0, h: 5 * 0.78, line: { color: C.accent6, width: 2 } });
    for (let i = 0; i < 6; i++) {
      const y = 1.85 + i * 0.78;
      numCircle(s, i + 1, 0.6, y, 0.5, C.accent1, 14, "Arial");
      s.addText([
        { text: steps[i][0], options: { bold: true, color: C.text2, breakLine: true } },
        { text: steps[i][1], options: { color: C.text1, fontSize: 13 } },
      ], { x: 1.35, y: y - 0.05, w: 6.3, h: 0.7, isTextBox: true, margin: 0, fontFace: "Arial", fontSize: 15, valign: "top" });
    }
    card(s, 8.2, 1.8, 4.533, 2.55, C.text2, { name: "Tie-breakers" });
    txt(s, "TIE-BREAKERS, IN ORDER", { x: 8.5, y: 2.05, w: 4.0, h: 0.28, fontSize: 11, bold: true, color: C.accent2, charSpacing: 1.5 });
    const tb = ["Higher Communication score", "Better spread across shops and shifts", "Prior teaching or mentoring experience"];
    for (let i = 0; i < 3; i++) {
      numCircle(s, i + 1, 8.5, 2.55 + i * 0.55, 0.38, C.accent1, 12, "Arial", C.text1);
      txt(s, tb[i], { x: 9.05, y: 2.55 + i * 0.55, w: 3.5, h: 0.38, fontSize: 14, color: C.background1, valign: "middle" });
    }
    card(s, 8.2, 4.55, 4.533, 1.95, C.background2, { name: "Contingency" });
    txt(s, "IF A LANGUAGE IS SHORT", { x: 8.5, y: 4.8, w: 4.0, h: 0.28, fontSize: 11, bold: true, color: C.accent4, charSpacing: 1.5 });
    txt(s, "Fewer than two qualified applicants? Invite directly through managers and extend that language's window by one day.", { x: 8.5, y: 5.15, w: 3.95, h: 1.2, fontSize: 14, color: C.text1 });
    s.addNotes("The 60-point minimum is a suggestion: set the final threshold before scoring starts, not after seeing the results.");
  }

  // 22 · Scoring sheet
  {
    const s = newSlide("CONTENT", "STEP 5", "Committee Scoring Sheet");
    const H = ["Applicant", "Language", "Proficiency (30)", "Teaching (25)", "Communication (25)", "Availability (20)", "Total (100)", "Decision"];
    const head = H.map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 11 } }));
    const blank = () => H.map(() => ({ text: "", options: { fill: { color: C.background1 } } }));
    const rows = [head];
    for (let i = 0; i < 8; i++) rows.push(blank());
    s.addTable(rows, { x: 0.6, y: 1.8, w: 12.133, colW: [2.6, 1.4, 1.35, 1.35, 1.5, 1.4, 1.1, 1.433], rowH: [0.5, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42], fontFace: "Arial", fontSize: 11, color: C.text1, valign: "middle", margin: [4, 8, 4, 8], border: { type: "solid", pt: 0.75, color: HEX.accent6 }, objectName: "Scoring sheet" });
    txt(s, "DECISION CODES", { x: 0.6, y: 6.05, w: 1.8, h: 0.4, fontSize: 11, bold: true, color: C.accent4, valign: "middle", charSpacing: 1.5 });
    const codes = [["Selected", C.text2, C.background1], ["Reserve", C.accent3, C.background1], ["Native Buddy", C.accent4, C.background1], ["Not selected", C.accent2, C.text1]];
    let px = 2.45;
    for (const [t, f, tc] of codes) {
      const pw = 0.4 + t.length * 0.1;
      card(s, px, 6.05, pw, 0.4, f);
      txt(s, t, { x: px, y: 6.05, w: pw, h: 0.4, fontSize: 12, bold: true, color: tc, align: "center", valign: "middle" });
      px += pw + 0.2;
    }
    s.addNotes("Print one sheet per committee member, or copy the table into Excel. Each member fills it independently before calibration.");
  }

  // ======================================================== LAUNCH
  section("Launch");
  {
    const s = newSlide("CLOSING_DARK", "BEFORE WE LAUNCH  ·  COMPLETE BY THURSDAY 8 OCTOBER", "Final Checklist");
    const items = ["Registration link and QR code", "Learner registration link", "Sender name and title", "HR contact for questions", "Committee members and native-speaker assessors", "Session length and kickoff date", "Rationale for the five languages", "Manager briefing before Sunday's email"];
    for (let i = 0; i < 8; i++) {
      const x = i < 4 ? 0.6 : 6.9, y = 1.95 + (i % 4) * 0.78;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.04, w: 0.42, h: 0.42, rectRadius: 0.06, fill: { color: C.text2 }, line: { color: C.accent2, width: 1.5 } });
      txt(s, items[i], { x: x + 0.65, y, w: 5.3, h: 0.5, fontSize: 17, color: C.background1, valign: "middle" });
    }
    txt(s, "“Every new word creates a new connection.”", { x: 0.6, y: 5.35, w: 12.133, h: 0.65, fontFace: SERIF, fontSize: 28, italic: true, color: C.accent1 });
    txt(s, "LEARN.   CONNECT.   BELONG.", { x: 0.6, y: 6.1, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 3 });
    s.addNotes("Every bracketed placeholder in the deck and the email is listed here. Brief line managers before the email goes out so release conversations are easy.");
  }

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Built " + path.relative(ROOT, OUT));
})().catch((e) => { console.error(e); process.exit(1); });
