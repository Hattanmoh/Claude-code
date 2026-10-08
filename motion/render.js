#!/usr/bin/env node
// Language Champion teaser: builds motion/teaser.html from the template and renders it to MP4.
//
//   node motion/render.js                 build the HTML, render output/Language_Champion_Teaser.mp4,
//                                         the poster frame and the README preview strip
//   node motion/render.js --html-only     only rebuild motion/teaser.html (open it in a browser to preview)
//   node motion/render.js --still 3 12.5  save single frames as PNG (for checking a moment)
//   node motion/render.js --preview-only  refresh only the poster and the README preview strip
//
// Settings live in motion/config.json. Add registrationUrl to put a scannable QR code on the end card.

const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const { spawn } = require("child_process");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa");

const DIR = __dirname;
const ROOT = path.join(DIR, "..");
const OUT = path.join(DIR, "output");
const TEMPLATE = path.join(DIR, "teaser.template.html");
const HTML = path.join(DIR, "teaser.html");
const MP4 = path.join(OUT, "Language_Champion_Teaser.mp4");
const POSTER = path.join(OUT, "poster.jpg");
const PREVIEW = path.join(ROOT, "docs", "img", "teaser-frames.jpg");
const PREVIEW_TIMES = [4.4, 7.8, 12.8, 17.2, 26.8, 31.2];

const cfg = JSON.parse(fs.readFileSync(path.join(DIR, "config.json"), "utf8"));
const FPS = cfg.fps || 30;
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function iconSvg(name) {
  if (!fa[name]) throw new Error(`Unknown icon ${name}`);
  return ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { size: "1em", "aria-hidden": true }));
}

// End-card right panel: a scannable QR code when a registration link is set, otherwise the greeting cluster.
const BUBBLES = `
    <div class="pin" id="s7-disc" style="left: 1460px; top: 540px;"><div class="c disc" style="width: 660px; height: 660px; background: var(--sand);"></div></div>
    <div class="pin" id="r1" style="left: 1300px; top: 340px;"><div class="c bubble sm" style="background: var(--white); border-radius: 40px 40px 40px 10px;"><span class="ar" style="font-size: 52px; color: var(--ink);">مرحبا</span></div></div>
    <div class="pin" id="r2" style="left: 1615px; top: 375px;"><div class="c bubble sm serif" style="background: var(--orange); border-radius: 40px 40px 10px 40px; font-size: 58px; color: var(--ink);">Merhaba</div></div>
    <div class="pin" id="r3" style="left: 1290px; top: 560px;"><div class="c bubble sm serif" style="background: var(--navy); border-radius: 40px 40px 40px 10px; font-size: 58px; color: var(--white);">Bonjour</div></div>
    <div class="pin" id="r4" style="left: 1625px; top: 610px;"><div class="c bubble sm" style="background: var(--white); border-radius: 40px 40px 10px 40px;"><span class="ar" style="font-size: 52px; color: var(--ink);">خوش آمدید</span></div></div>
    <div class="pin" id="r5" style="left: 1460px; top: 770px;"><div class="c bubble sm serif" style="background: var(--orange); border-radius: 40px 40px 40px 10px; font-size: 58px; color: var(--ink);">Hola</div></div>`;

async function buildHtml() {
  let right, applyLine;
  if (cfg.registrationUrl) {
    let QRCode;
    try { QRCode = require("qrcode"); } catch { throw new Error("A registrationUrl is set, so the qrcode package is needed: run npm install"); }
    const svg = (await QRCode.toString(cfg.registrationUrl, { type: "svg", margin: 0, errorCorrectionLevel: "M", color: { dark: "#1C2541", light: "#FFFFFF" } }))
      .replace("<svg ", '<svg width="360" height="360" ');
    right = `
    <div class="pin" id="s7-qr" style="left: 1460px; top: 520px;">
      <div class="c" style="width: 460px; background: var(--white); border-radius: 32px; padding: 50px 50px 40px; box-shadow: 0 24px 60px rgba(28, 37, 65, 0.14); display: flex; flex-direction: column; align-items: center; gap: 26px;">
        ${svg}
        <div class="eyebrow" style="font-size: 26px; letter-spacing: 9px; color: var(--burnt);">Scan to apply</div>
      </div>
    </div>`;
    applyLine = "Or visit " + esc(cfg.displayUrl || cfg.registrationUrl);
  } else {
    right = BUBBLES;
    applyLine = "Apply through the link in your email";
  }
  const html = fs.readFileSync(TEMPLATE, "utf8")
    .replace(/\{\{ICON:(\w+)\}\}/g, (_, n) => iconSvg(n))
    .replace("{{S7_RIGHT}}", right)
    .replace("{{APPLY_LINE}}", applyLine)
    .replace("{{DEADLINE}}", esc(cfg.deadline));
  const left = html.match(/\{\{[A-Z0-9_:]+\}\}/);
  if (left) throw new Error("Unfilled template token " + left[0]);
  fs.writeFileSync(HTML, html);
  console.log("Built " + path.relative(ROOT, HTML));
}

async function openPage() {
  const { chromium } = require("playwright");
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });
  await page.goto(pathToFileURL(HTML).href + "?render=1");
  await page.evaluate(() => window.fontsReady);
  if (errors.length) { await browser.close(); throw new Error("Page errors:\n" + errors.join("\n")); }
  const duration = await page.evaluate(() => window.DURATION);
  const shot = async (t, type = "png") => {
    await page.evaluate((x) => window.renderFrame(x), t);
    return page.screenshot({ type, quality: type === "jpeg" ? 92 : undefined, clip: { x: 0, y: 0, width: 1920, height: 1080 } });
  };
  return { browser, duration, shot };
}

async function renderVideo({ duration, shot }) {
  fs.mkdirSync(OUT, { recursive: true });
  const frames = Math.round(duration * FPS);
  const ff = spawn("ffmpeg", ["-y", "-loglevel", "error", "-f", "image2pipe", "-framerate", String(FPS), "-c:v", "png", "-i", "-",
    "-c:v", "libx264", "-preset", "slow", "-crf", "18", "-pix_fmt", "yuv420p", "-r", String(FPS), "-movflags", "+faststart", MP4], { stdio: ["pipe", "inherit", "inherit"] });
  const done = new Promise((res, rej) => { ff.on("error", rej); ff.on("close", (c) => (c === 0 ? res() : rej(new Error("ffmpeg exited with " + c)))); });
  for (let i = 0; i < frames; i++) {
    const buf = await shot(i / FPS);
    if (!ff.stdin.write(buf)) await new Promise((r) => ff.stdin.once("drain", r));
    if (i % 150 === 0) process.stdout.write(`  frame ${i}/${frames}\n`);
  }
  ff.stdin.end();
  await done;
  console.log(`Rendered ${path.relative(ROOT, MP4)} (${frames} frames, ${duration}s at ${FPS} fps)`);
}

async function renderPreview({ shot }) {
  const sharp = require("sharp");
  fs.writeFileSync(POSTER, await shot(cfg.posterTime || 31.2, "jpeg"));
  const tw = 640, th = 360, gap = 16, pad = 24, cols = 3, rows = Math.ceil(PREVIEW_TIMES.length / cols);
  const tiles = [];
  for (let i = 0; i < PREVIEW_TIMES.length; i++) {
    // a thin taupe frame keeps the light-sand scenes distinct from the sheet background
    const img = await sharp(await shot(PREVIEW_TIMES[i])).resize(tw - 4, th - 4)
      .extend({ top: 2, bottom: 2, left: 2, right: 2, background: "#C9B79C" }).jpeg({ quality: 88 }).toBuffer();
    tiles.push({ input: img, left: pad + (i % cols) * (tw + gap), top: pad + Math.floor(i / cols) * (th + gap) });
  }
  fs.mkdirSync(path.dirname(PREVIEW), { recursive: true });
  await sharp({ create: { width: pad * 2 + cols * tw + (cols - 1) * gap, height: pad * 2 + rows * th + (rows - 1) * gap, channels: 3, background: "#FFFFFF" } })
    .composite(tiles).jpeg({ quality: 85 }).toFile(PREVIEW);
  console.log(`Saved ${path.relative(ROOT, POSTER)} and ${path.relative(ROOT, PREVIEW)}`);
}

(async () => {
  const args = process.argv.slice(2);
  await buildHtml();
  if (args.includes("--html-only")) return;
  const page = await openPage();
  try {
    const si = args.indexOf("--still");
    if (si >= 0) {
      const dir = process.env.STILLS_DIR || OUT;
      fs.mkdirSync(dir, { recursive: true });
      for (const a of args.slice(si + 1)) {
        const f = path.join(dir, `still-${a}.png`);
        fs.writeFileSync(f, await page.shot(parseFloat(a)));
        console.log("Saved " + f);
      }
      return;
    }
    if (!args.includes("--preview-only")) await renderVideo(page);
    await renderPreview(page);
  } finally {
    await page.browser.close();
  }
})().catch((e) => { console.error(e.message || e); process.exit(1); });
