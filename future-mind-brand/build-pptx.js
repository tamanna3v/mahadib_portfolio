const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const { applyTheme } = require("/root/.claude/skills/synced/f04f7271-28d2-4942-9144-06a195dec267_f7c9d2fc-b597-415b-84f9-1db0adbda9c9/pptx/scripts/apply_theme.js");
const OUT = require("path").join(__dirname, "Future-Mind-Brand-Guidelines.pptx");

const THEME = {
  name: "Future Mind",
  headFontFace: "Inter", bodyFontFace: "Inter",
  colors: { dk1:"1E1630", lt1:"FFFFFF", dk2:"6E3B96", lt2:"E8F8F3",
    accent1:"9257BB", accent2:"A4E5D4", accent3:"6E3B96", accent4:"1E1630", accent5:"5E5670", accent6:"E3DEEA",
    hlink:"9257BB", folHlink:"6E3B96" }
};
const H = { purple:"9257BB", mint:"A4E5D4", ink:"1E1630", white:"FFFFFF", deep:"6E3B96" };

async function markPng(bg, fg, opts = {}) {
  const rect = bg ? `<rect width="100" height="100" fill="#${bg}"/>` : "";
  const sw = opts.sw || 4.2;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="600" height="600" viewBox="0 0 100 100">${rect}<path d="M57 18 L30 55 H49 L43 82 L70 45 H51 Z" fill="none" stroke="#${fg}" stroke-width="${sw}" stroke-linejoin="round"/></svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const IMG = {
    P: await markPng(H.purple, H.white), M: await markPng(H.mint, H.purple), W: await markPng(H.white, H.purple),
    K: await markPng(H.ink, H.white), X: await markPng("F2A33A", "1C8C3C"), bolt: await markPng(null, H.white, { sw: 1.2 }),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "Future Mind Consulting Group — Brand Guidelines";
  pres.author = "Future Mind Consulting Group";
  pres.company = "Future Mind Consulting Group";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: "FM_COVER", background: { color: C.accent1 },
    objects: [
      { image: { x: 8.3, y: 1.45, w: 4.6, h: 4.6, data: IMG.M, altText: "Future Mind Spark mark" } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.7, y: 2.55, w: 7.2, h: 1.3, fontSize: 54, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", align: "left", x: 0.7, y: 3.95, w: 7.2, h: 0.6, fontSize: 24, color: C.accent2, valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "meta", type: "body", align: "left", x: 0.7, y: 6.35, w: 7.2, h: 0.4, fontSize: 12, bold: true, charSpacing: 3, color: C.accent2, margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "FM_CONTENT", background: { color: C.background1 },
    margin: [0.5, 0.6, 0.7, 0.6],
    objects: [
      { image: { x: 12.23, y: 0.5, w: 0.5, h: 0.5, data: IMG.P, altText: "Future Mind mark" } },
      { text: { text: "Future Mind Consulting Group  ·  Brand Guidelines", options: { x: 0.6, y: 6.95, w: 7, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
      { placeholder: { options: { name: "eyebrow", type: "body", align: "left", x: 0.6, y: 0.45, w: 10, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.accent1, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.6, y: 0.78, w: 11, h: 0.75, fontSize: 36, bold: true, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 12.23, y: 6.95, w: 0.5, h: 0.3, fontSize: 10, color: "5E5670", align: "right" },
  });
  pres.defineSlideMaster({
    title: "FM_CLOSING", background: { color: C.accent1 },
    objects: [
      { image: { x: 9.2, y: 0.9, w: 5.6, h: 5.6, data: IMG.bolt, transparency: 75, altText: "" } },
      { placeholder: { options: { name: "eyebrow", type: "body", align: "left", x: 0.7, y: 0.9, w: 8, h: 0.35, fontSize: 12, bold: true, charSpacing: 3, color: C.accent2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.7, y: 1.3, w: 8.5, h: 0.9, fontSize: 40, bold: true, color: C.background1, valign: "top", margin: 0 }, text: "" } },
    ],
  });

  // ---------- helpers ----------
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 14 }, o));
  const panel = (s, x, y, w, h, color, name, extra = {}) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.08, fill: { color }, line: { color, width: 0 }, objectName: name }, extra));
  function lockup(s, x, y, size, img, nameColor, subColor, tag) {
    s.addImage({ x, y, w: size, h: size, data: IMG[img], altText: "Future Mind Spark mark", objectName: `${tag} mark` });
    s.addText([
      { text: "Future Mind", options: { fontSize: Math.round(size * 34), bold: true, color: nameColor, breakLine: true } },
      { text: "CONSULTING GROUP", options: { fontSize: Math.max(7, Math.round(size * 10.5)), color: subColor, charSpacing: Math.max(1, Math.round(size * 4)), bold: true } },
    ], { isTextBox: true, x: x + size * 1.25, y, w: size * 4.4, h: size, valign: "middle", margin: 0, paraSpaceAfter: 2, objectName: `${tag} wordmark` });
  }
  const content = (section) => pres.addSlide({ masterName: "FM_CONTENT", sectionTitle: section });

  // ===== 1. Cover =====
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "FM_COVER", sectionTitle: "Introduction" });
  lockup(s, 0.7, 0.7, 0.6, "M", C.background1, C.accent2, "Cover");
  s.addText("Brand Guidelines", { placeholder: "title" });
  s.addText("How we look, how we speak.", { placeholder: "body" });
  s.addText("VERSION 1.0  ·  NOVEMBER 2026", { placeholder: "meta" });
  s.addNotes("Brand guidelines for Future Mind Consulting Group. Covers the logo, colour palette, typography and applications.");

  // ===== 2. Inside this guide =====
  s = content("Introduction");
  s.addText("INTRODUCTION", { placeholder: "eyebrow" });
  s.addText("One brand, used the same way everywhere", { placeholder: "title" });
  T(s, "This guide is for everyone who creates material for Future Mind Consulting Group — staff, partners and agencies. Follow it to keep every touchpoint clear, consistent and recognisably ours.",
    { x: 0.6, y: 1.75, w: 8.6, h: 0.9, fontSize: 16, color: C.accent5, objectName: "Intro text" });
  const toc = [["01", "Logo", "The Spark mark, clear space, sizes and versions"], ["02", "Colour", "Future Purple, Mind Mint and supporting neutrals"],
    ["03", "Typography", "Inter for English, Hind Siliguri for Bangla"], ["04", "Applications", "Business cards, social posts and a checklist"]];
  toc.forEach(([n, t, d], i) => {
    const x = 0.6 + i * 3.08;
    panel(s, x, 3.05, 2.83, 3.4, i % 2 ? C.background2 : C.accent2, `Card ${n}`);
    T(s, n, { x: x + 0.3, y: 3.3, w: 2.2, h: 0.9, fontSize: 44, bold: true, color: C.accent1, objectName: `Card ${n} number` });
    T(s, t, { x: x + 0.3, y: 4.55, w: 2.3, h: 0.45, fontSize: 20, bold: true, objectName: `Card ${n} title` });
    T(s, d, { x: x + 0.3, y: 5.05, w: 2.3, h: 1.1, fontSize: 14, color: C.accent5, valign: "top", objectName: `Card ${n} text` });
  });

  // ===== 3. Primary logo =====
  pres.addSection({ title: "Logo" });
  s = content("Logo");
  s.addText("01  ·  LOGO", { placeholder: "eyebrow" });
  s.addText("The primary logo", { placeholder: "title" });
  panel(s, 0.6, 1.8, 7.4, 4.75, C.background2, "Logo panel");
  lockup(s, 1.35, 3.53, 1.3, "P", C.text1, C.accent1, "Primary");
  T(s, "The Spark mark", { x: 8.5, y: 1.85, w: 4.2, h: 0.5, fontSize: 22, bold: true, objectName: "Meaning heading" });
  T(s, "An outlined lightning bolt inside a solid Future Purple square. It stands for the moment of insight we bring to clients — quick, clear thinking that turns into action.",
    { x: 8.5, y: 2.45, w: 4.2, h: 1.7, fontSize: 15, color: C.accent5, valign: "top", objectName: "Meaning text" });
  T(s, [
    { text: "Mark + wordmark together is the default.", options: { bullet: true, breakLine: true } },
    { text: "Use the mark alone only for icons and avatars.", options: { bullet: true, breakLine: true } },
    { text: "Always use the master artwork — never redraw or retype it.", options: { bullet: true } },
  ], { x: 8.5, y: 4.3, w: 4.2, h: 2.2, fontSize: 14, valign: "top", paraSpaceAfter: 8, objectName: "Logo rules" });

  // ===== 4. Clear space & minimum size =====
  s = content("Logo");
  s.addText("01  ·  LOGO", { placeholder: "eyebrow" });
  s.addText("Clear space and minimum size", { placeholder: "title" });
  panel(s, 0.6, 1.8, 5.9, 3.7, C.background2, "Clear space panel");
  s.addShape(pres.shapes.RECTANGLE, { x: 2.3, y: 2.15, w: 2.5, h: 2.5, fill: { color: C.background2 }, line: { color: C.accent1, width: 1, dashType: "dash" }, objectName: "Clear space boundary" });
  s.addImage({ x: 2.925, y: 2.775, w: 1.25, h: 1.25, data: IMG.P, altText: "Spark mark", objectName: "Clear space mark" });
  [[3.4, 2.25], [3.4, 4.17], [2.4, 3.21], [4.4, 3.21]].forEach(([x, y], i) =>
    T(s, "X", { x, y, w: 0.3, h: 0.35, fontSize: 12, bold: true, color: C.accent1, align: "center", valign: "middle", objectName: `Clear space X ${i + 1}` }));
  T(s, "Keep a margin of X on every side, where X = ½ the width of the mark. No text, image or edge may enter this space.",
    { x: 0.6, y: 5.65, w: 5.9, h: 0.9, fontSize: 14, color: C.accent5, valign: "top", objectName: "Clear space text" });

  panel(s, 6.83, 1.8, 5.9, 3.7, C.background2, "Minimum size panel");
  lockup(s, 7.35, 3.3, 0.55, "P", C.text1, C.accent1, "Min");
  T(s, [{ text: "Lockup", options: { bold: true, breakLine: true } }, { text: "35 mm / 140 px wide", options: { color: C.accent5 } }],
    { x: 7.35, y: 4.15, w: 2.8, h: 0.7, fontSize: 12, valign: "top", objectName: "Lockup min label" });
  s.addImage({ x: 11.1, y: 3.37, w: 0.4, h: 0.4, data: IMG.P, altText: "Spark mark at minimum size", objectName: "Min mark" });
  T(s, [{ text: "Mark only", options: { bold: true, breakLine: true } }, { text: "8 mm / 24 px", options: { color: C.accent5 } }],
    { x: 10.7, y: 4.15, w: 1.8, h: 0.7, fontSize: 12, valign: "top", objectName: "Mark min label" });
  T(s, "Below these sizes the bolt outline stops being readable. Use the mark alone for favicons, app icons and social avatars.",
    { x: 6.83, y: 5.65, w: 5.9, h: 0.9, fontSize: 14, color: C.accent5, valign: "top", objectName: "Minimum size text" });

  // ===== 5. Versions & misuse =====
  s = content("Logo");
  s.addText("01  ·  LOGO", { placeholder: "eyebrow" });
  s.addText("Colour versions and misuse", { placeholder: "title" });
  const vers = [["Primary", "On white or light backgrounds", C.background1, "P", C.text1, C.accent1, true],
    ["On Mind Mint", "Brand panels and covers", C.accent2, "P", C.text1, C.accent3, false],
    ["Reversed", "On Future Purple or dark photos", C.accent1, "W", C.background1, C.accent2, false],
    ["Mono", "Single-colour print and stamps", C.background1, "K", C.text1, C.text1, true]];
  vers.forEach(([t, d, bg, img, nc, sc, border], i) => {
    const x = 0.6 + i * 3.08;
    panel(s, x, 1.75, 2.83, 1.45, bg, `Version ${t} tile`, border ? { line: { color: C.accent6, width: 1 } } : {});
    lockup(s, x + 0.32, 2.21, 0.52, img, nc, sc, `Version ${t}`);
    T(s, [{ text: t, options: { bold: true, breakLine: true } }, { text: d, options: { color: C.accent5 } }],
      { x, y: 3.3, w: 2.83, h: 0.65, fontSize: 12, valign: "top", objectName: `Version ${t} caption` });
  });
  T(s, "Don’t", { x: 0.6, y: 4.1, w: 4, h: 0.4, fontSize: 20, bold: true, objectName: "Dont heading" });
  const donts = ["Rotate or tilt the mark", "Recolour with off-brand colours", "Stretch or squash", "Add outlines, borders or shadows"];
  donts.forEach((d, i) => {
    const x = 0.6 + i * 3.08, y = 4.6, w = 2.83, h = 1.45;
    panel(s, x, y, w, h, C.background2, `Dont ${i + 1} tile`);
    const cx = x + w / 2, cy = y + h / 2, m = 0.85;
    if (i === 0) s.addImage({ x: cx - m / 2, y: cy - m / 2, w: m, h: m, data: IMG.P, rotate: 340, altText: "Rotated mark — incorrect", objectName: "Dont rotated" });
    if (i === 1) s.addImage({ x: cx - m / 2, y: cy - m / 2, w: m, h: m, data: IMG.X, altText: "Recoloured mark — incorrect", objectName: "Dont recolour" });
    if (i === 2) s.addImage({ x: cx - 0.75, y: cy - 0.32, w: 1.5, h: 0.64, data: IMG.P, altText: "Stretched mark — incorrect", objectName: "Dont stretch" });
    if (i === 3) {
      s.addShape(pres.shapes.RECTANGLE, { x: cx - m / 2 - 0.06, y: cy - m / 2 - 0.06, w: m + 0.12, h: m + 0.12, fill: { color: "0066DB" }, line: { color: "0066DB", width: 0 },
        shadow: { type: "outer", color: "000000", opacity: 0.4, blur: 6, offset: 4, angle: 45 }, objectName: "Dont border" });
      s.addImage({ x: cx - m / 2, y: cy - m / 2, w: m, h: m, data: IMG.P, altText: "Mark with border — incorrect", objectName: "Dont border mark" });
    }
    s.addShape(pres.shapes.LINE, { x: x + 0.15, y: y + 0.15, w: w - 0.3, h: h - 0.3, flipV: true, line: { color: "D64545", width: 1.5 }, objectName: `Dont ${i + 1} strike` });
    T(s, d, { x, y: 6.12, w: 2.83, h: 0.4, fontSize: 12, color: C.accent5, objectName: `Dont ${i + 1} caption` });
  });

  // ===== 6. Primary colours =====
  pres.addSection({ title: "Colour" });
  s = content("Colour");
  s.addText("02  ·  COLOUR", { placeholder: "eyebrow" });
  s.addText("Two colours carry the brand", { placeholder: "title" });
  const sw = [["Future Purple", "Our signature — confident and considered", C.accent1, C.background1, C.accent2, "#9257BB", "146 · 87 · 187", "22 · 53 · 0 · 27", "Logo mark, headlines, key buttons"],
    ["Mind Mint", "Calm, fresh and open", C.accent2, C.text1, C.accent3, "#A4E5D4", "164 · 229 · 212", "28 · 0 · 7 · 10", "Backgrounds, panels, highlights"]];
  sw.forEach(([n, d, bg, fg, ey, hex, rgb, cmyk, use], i) => {
    const x = 0.6 + i * 6.23;
    panel(s, x, 1.8, 5.9, 4.75, bg, `${n} swatch`);
    T(s, "PRIMARY", { x: x + 0.45, y: 2.15, w: 4, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: ey, objectName: `${n} label` });
    T(s, n, { x: x + 0.45, y: 2.5, w: 5, h: 0.7, fontSize: 32, bold: true, color: fg, objectName: `${n} name` });
    T(s, d, { x: x + 0.45, y: 3.2, w: 5, h: 0.4, fontSize: 15, color: fg, objectName: `${n} description` });
    const rows = [["HEX", hex], ["RGB", rgb], ["CMYK", cmyk], ["Use", use]];
    T(s, rows.map(([k, v], j) => ({ text: `${k}\t${v}`, options: { breakLine: j < 3 } })),
      { x: x + 0.45, y: 4.45, w: 5.1, h: 1.8, fontSize: 15, color: fg, valign: "bottom", paraSpaceAfter: 4, tabStops: [{ position: 1.0 }], objectName: `${n} specs` });
  });

  // ===== 7. Neutrals, balance, accessibility =====
  s = content("Colour");
  s.addText("02  ·  COLOUR", { placeholder: "eyebrow" });
  s.addText("Neutrals, balance and contrast", { placeholder: "title" });
  T(s, "Supporting neutrals", { x: 0.6, y: 1.8, w: 5.6, h: 0.4, fontSize: 18, bold: true, objectName: "Neutrals heading" });
  const neu = [["Ink", "#1E1630", "Body text, mono logo", C.text1, C.background1], ["Deep Purple", "#6E3B96", "Purple text on light/mint", C.accent3, C.background1],
    ["Mint Tint", "#E8F8F3", "Soft section backgrounds", C.background2, C.text1], ["White", "#FFFFFF", "Page, breathing space", C.background1, C.text1]];
  neu.forEach(([n, hx, u, bg, fg], i) => {
    const x = 0.6 + (i % 2) * 2.9, y = 2.3 + Math.floor(i / 2) * 1.45;
    panel(s, x, y, 2.7, 1.25, bg, `${n} chip`, { line: { color: i === 3 ? C.accent6 : bg, width: 1 } });
    T(s, [{ text: n, options: { bold: true, breakLine: true } }, { text: `${hx}  ·  ${u}`, options: { fontSize: 10 } }],
      { x: x + 0.2, y: y + 0.2, w: 2.35, h: 0.9, fontSize: 14, color: fg, valign: "bottom", objectName: `${n} chip text` });
  });
  T(s, "Balance", { x: 0.6, y: 5.3, w: 5.6, h: 0.4, fontSize: 18, bold: true, objectName: "Balance heading" });
  const bal = [["White 50%", 2.8, C.background1, C.text1], ["Mint 22%", 1.25, C.accent2, C.text1], ["Purple 18%", 1.05, C.accent1, C.background1], ["Ink", 0.5, C.text1, C.background1]];
  let bx = 0.6;
  bal.forEach(([t, w, bg, fg]) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: 5.8, w, h: 0.55, fill: { color: bg }, line: { color: C.accent6, width: 0.75 }, objectName: `Balance ${t}` });
    T(s, t, { x: bx + 0.1, y: 5.8, w: w - 0.1, h: 0.55, fontSize: 11, bold: true, color: fg, valign: "middle", objectName: `Balance ${t} label` });
    bx += w;
  });
  T(s, "Lead with white space; purple is the focal point.", { x: 0.6, y: 6.4, w: 5.6, h: 0.35, fontSize: 11, color: C.accent5, objectName: "Balance note" });

  T(s, "Accessible pairings", { x: 6.83, y: 1.8, w: 5.9, h: 0.4, fontSize: 18, bold: true, objectName: "Contrast heading" });
  const hdr = { bold: true, fontSize: 11, color: C.accent5, fill: { color: C.background1 } };
  const rowsT = [[{ text: "SAMPLE", options: hdr }, { text: "CONTRAST", options: hdr }, { text: "USE FOR", options: hdr }]];
  [[C.accent1, C.background1, "4.9 : 1", "All text — white on purple", true], [C.accent2, C.text1, "12.2 : 1", "All text — ink on mint", true],
   [C.accent2, C.accent3, "5.4 : 1", "All text — deep purple on mint", true], [C.accent2, C.accent1, "3.4 : 1", "Large text (18pt+) and graphics only", false]]
    .forEach(([bg, fg, r, u, ok]) => rowsT.push([
      { text: "Aa", options: { fill: { color: bg }, color: fg, bold: true, align: "center", fontSize: 16 } },
      { text: r, options: { fontSize: 14 } },
      { text: (ok ? "✓  " : "!  ") + u, options: { fontSize: 14, bold: true, color: ok ? "1F7A55" : "A0661A" } }]));
  s.addTable(rowsT, { x: 6.83, y: 2.3, w: 5.9, colW: [1.1, 1.3, 3.5], rowH: [0.45, 0.75, 0.75, 0.75, 0.75], valign: "middle", color: C.text1,
    border: { type: "solid", pt: 0.75, color: "E3DEEA" }, margin: [0, 0.12, 0, 0.12], objectName: "Contrast table" });

  // ===== 8. Typography =====
  pres.addSection({ title: "Typography" });
  s = content("Typography");
  s.addText("03  ·  TYPOGRAPHY", { placeholder: "eyebrow" });
  s.addText("Our typefaces", { placeholder: "title" });
  panel(s, 0.6, 1.8, 5.9, 4.75, C.background2, "Inter card");
  T(s, "PRIMARY  ·  ENGLISH", { x: 1.0, y: 2.1, w: 5, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.accent1, objectName: "Inter label" });
  T(s, "Aa", { x: 1.0, y: 2.45, w: 2.2, h: 1.5, fontSize: 88, bold: true, color: C.accent1, objectName: "Inter specimen" });
  T(s, [{ text: "Inter", options: { fontSize: 30, bold: true, breakLine: true } }, { text: "Free on Google Fonts · all platforms", options: { fontSize: 12, color: C.accent5 } }],
    { x: 3.3, y: 2.75, w: 3.0, h: 1.0, valign: "bottom", objectName: "Inter name" });
  T(s, "Regular · Medium · SemiBold · Bold", { x: 1.0, y: 4.25, w: 5.2, h: 0.4, fontSize: 15, objectName: "Inter weights" });
  T(s, "ABCDEFGHIJKLMNOPQRSTUVWXYZ\nabcdefghijklmnopqrstuvwxyz 0123456789", { x: 1.0, y: 4.75, w: 5.2, h: 0.75, fontSize: 13, color: C.accent5, objectName: "Inter alphabet" });
  T(s, "Fallback when Inter isn’t installed: Arial.", { x: 1.0, y: 5.75, w: 5.2, h: 0.4, fontSize: 12, color: C.accent5, objectName: "Inter fallback" });

  panel(s, 6.83, 1.8, 5.9, 4.75, C.accent2, "Hind card");
  T(s, "COMPANION  ·  BANGLA", { x: 7.23, y: 2.1, w: 5, h: 0.3, fontSize: 12, bold: true, charSpacing: 3, color: C.accent3, objectName: "Hind label" });
  T(s, "অআ", { x: 7.23, y: 2.45, w: 2.5, h: 1.5, fontSize: 80, bold: true, color: C.accent3, fontFace: "Hind Siliguri", objectName: "Hind specimen" });
  T(s, [{ text: "Hind Siliguri", options: { fontSize: 26, bold: true, breakLine: true } }, { text: "For all Bangla text", options: { fontSize: 12, color: C.text1 } }],
    { x: 9.75, y: 2.75, w: 2.8, h: 1.0, valign: "bottom", objectName: "Hind name" });
  T(s, "ভবিষ্যতের চিন্তা, আজকের সমাধান।", { x: 7.23, y: 4.25, w: 5.2, h: 0.55, fontSize: 20, fontFace: "Hind Siliguri", objectName: "Hind sample" });
  T(s, "অ আ ই ঈ উ ক খ গ ঘ ০১২৩৪৫৬৭৮৯", { x: 7.23, y: 4.85, w: 5.2, h: 0.5, fontSize: 15, fontFace: "Hind Siliguri", color: C.accent3, objectName: "Hind alphabet" });
  T(s, "Never mix in decorative or script fonts.", { x: 7.23, y: 5.75, w: 5.2, h: 0.4, fontSize: 12, objectName: "Hind rule" });

  // ===== 9. Type scale =====
  s = content("Typography");
  s.addText("03  ·  TYPOGRAPHY", { placeholder: "eyebrow" });
  s.addText("Type scale and hierarchy", { placeholder: "title" });
  const scale = [["Display", "Bold · 44 pt", "Ideas that move", 40, true, C.text1], ["Heading", "Bold · 24 pt", "Strategy, people and growth", 24, true, C.text1],
    ["Body", "Regular · 14–16 pt", "We help organisations think clearly about what comes next — and act on it with confidence.", 15, false, C.accent5],
    ["Label", "Bold · 12 pt · spaced", "CASE STUDY  ·  2026", 12, true, C.accent1]];
  const ys = [1.85, 2.95, 3.85, 4.85];
  scale.forEach(([n, spec, sample, fs, b, col], i) => {
    const y = ys[i];
    T(s, [{ text: n, options: { bold: true, breakLine: true } }, { text: spec, options: { color: C.accent5, fontSize: 12 } }],
      { x: 0.6, y, w: 2.6, h: 0.8, fontSize: 14, valign: "middle", objectName: `${n} spec` });
    T(s, sample, { x: 3.5, y, w: 9.2, h: 0.8, fontSize: fs, bold: b, color: col, valign: "middle", charSpacing: n === "Label" ? 3 : 0, objectName: `${n} sample` });
    s.addShape(pres.shapes.LINE, { x: 0.6, y: y + 0.9, w: 12.13, h: 0, line: { color: C.accent6, width: 0.75 }, objectName: `${n} divider` });
  });
  const rules = [["Sentence case", "for headings and titles"], ["Left-align", "text; avoid centred paragraphs"], ["Ink for body;", "purple for headlines and accents"]];
  rules.forEach(([a, b], i) => {
    const x = 0.6 + i * 4.1;
    panel(s, x, 5.95, 3.85, 0.75, C.background2, `Rule ${i + 1}`);
    T(s, [{ text: a + " ", options: { bold: true, color: C.accent3 } }, { text: b }], { x: x + 0.25, y: 5.95, w: 3.45, h: 0.75, fontSize: 13, valign: "middle", objectName: `Rule ${i + 1} text` });
  });

  // ===== 10. Applications =====
  pres.addSection({ title: "Applications" });
  s = content("Applications");
  s.addText("04  ·  APPLICATIONS", { placeholder: "eyebrow" });
  s.addText("The brand in use", { placeholder: "title" });
  const shadow = () => ({ type: "outer", color: "1E1630", opacity: 0.18, blur: 10, offset: 3, angle: 90 });
  // card front 85x55 ratio
  panel(s, 0.6, 2.0, 4.3, 2.78, C.background1, "Card front", { shadow: shadow(), line: { color: C.accent6, width: 0.5 } });
  lockup(s, 0.95, 2.35, 0.5, "P", C.text1, C.accent1, "Card");
  T(s, [{ text: "Full Name", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "Designation", options: { color: C.accent3, bold: true, breakLine: true } },
    { text: "+880 1XXX-XXXXXX  ·  email@domain.com", options: { color: C.accent5 } }],
    { x: 0.95, y: 3.6, w: 3.8, h: 0.95, fontSize: 11, valign: "bottom", objectName: "Card details" });
  T(s, [{ text: "Business card — front", options: { bold: true } }, { text: "  ·  85 × 55 mm, white stock", options: { color: C.accent5 } }],
    { x: 0.6, y: 5.0, w: 4.3, h: 0.35, fontSize: 12, objectName: "Card front caption" });
  panel(s, 5.25, 2.0, 4.3, 2.78, C.accent1, "Card back", { shadow: shadow() });
  s.addImage({ x: 6.8, y: 2.79, w: 1.2, h: 1.2, data: IMG.M, altText: "Mint Spark mark", objectName: "Card back mark" });
  T(s, [{ text: "Business card — back", options: { bold: true } }, { text: "  ·  mint mark on purple", options: { color: C.accent5 } }],
    { x: 5.25, y: 5.0, w: 4.3, h: 0.35, fontSize: 12, objectName: "Card back caption" });
  panel(s, 9.95, 2.0, 2.78, 2.78, C.accent2, "Social post", { shadow: shadow() });
  s.addImage({ x: 10.25, y: 2.3, w: 0.55, h: 0.55, data: IMG.P, altText: "Spark mark", objectName: "Social mark" });
  T(s, [{ text: "Clear thinking", options: { breakLine: true } }, { text: "for " }, { text: "what’s next.", options: { color: C.accent3 } }],
    { x: 10.25, y: 3.4, w: 2.3, h: 1.1, fontSize: 18, bold: true, valign: "bottom", objectName: "Social headline" });
  T(s, [{ text: "Social post", options: { bold: true } }, { text: "  ·  1080 × 1080 px", options: { color: C.accent5 } }],
    { x: 9.95, y: 5.0, w: 2.78, h: 0.35, fontSize: 12, objectName: "Social caption" });
  panel(s, 0.6, 5.85, 12.13, 0.85, C.background2, "Placeholder note");
  T(s, "Replace the placeholder name, designation, phone and email with real details before printing.",
    { x: 0.9, y: 5.85, w: 11.5, h: 0.85, fontSize: 14, color: C.accent5, valign: "middle", objectName: "Placeholder note text" });

  // ===== 11. Closing checklist =====
  s = pres.addSlide({ masterName: "FM_CLOSING", sectionTitle: "Applications" });
  s.addText("BEFORE YOU PUBLISH", { placeholder: "eyebrow" });
  s.addText("Quick brand checklist", { placeholder: "title" });
  const checks = ["Logo from the master file, with clear space respected", "Only Future Purple, Mind Mint and the neutrals",
    "Inter for English, Hind Siliguri for Bangla", "Plenty of white space — one purple focal point per layout", "Text contrast checked against the pairings table"];
  checks.forEach((c, i) => {
    const y = 2.5 + i * 0.78;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.04, w: 0.46, h: 0.46, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: `Check ${i + 1} dot` });
    T(s, "✓", { x: 0.7, y: y + 0.04, w: 0.46, h: 0.46, fontSize: 14, bold: true, color: C.accent3, align: "center", valign: "middle", objectName: `Check ${i + 1} tick` });
    T(s, c, { x: 1.45, y, w: 7.5, h: 0.55, fontSize: 18, color: C.background1, valign: "middle", objectName: `Check ${i + 1} text` });
  });
  lockup(s, 0.7, 6.55, 0.42, "M", C.background1, C.accent2, "Closing");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written");
})();
