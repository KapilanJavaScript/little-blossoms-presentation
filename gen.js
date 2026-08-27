const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
p.layout = "WIDE";

const S = pptxgen.ShapeType;
const INK = "3A2E29";
const GRAPE = "6F2DBD";
const GRAPE_D = "4A1D87";
const CREAM = "FFF6E0";
const WHITE = "FFFFFF";
const SUN = "FFC53D";
const TEAL = "2EC4B6";
const LEAF = "4C9A56";
const LAV = "F3EDFB";
const MUT = "725A4A";
const FONT = "Trebuchet MS";
const BODY = "Calibri";

function bg(color) { return { color }; }
function body(o = {}) { return Object.assign({ fontFace: BODY, color: MUT, size: 13, breakLine: false }, o); }

// ---------- Slide 1 — Title (dark) ----------
const s1 = p.addSlide();
s1.background = bg(GRAPE_D);
s1.addShape("ellipse", { x: -1.2, y: -1.2, w: 3, h: 3, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
s1.addShape("ellipse", { x: 11.3, y: 5.7, w: 3, h: 3, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
s1.addShape("roundRect", { x: 0.9, y: 0.8, w: 5.4, h: 1.0, rectRadius: 0.5, fill: { color: SUN }, line: { color: SUN } });
s1.addText("LITTLE BLOSSOMS \u2022 WEBSITE \u2022 2026", { x: 0.9, y: 0.8, w: 5.4, h: 1.0, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 13, valign: "middle" });
s1.addText("Play today.\nBlossom tomorrow.", { x: 1.0, y: 1.9, w: 11.3, h: 2.6, fontFace: FONT, bold: true, color: WHITE, fontSize: 52, lineSpacingMultiple: 1.02 });
s1.addText("A warm, colourful playschool website for little ones from 8 months to 6 years", { x: 1.0, y: 4.5, w: 10.5, h: 0.8, fontFace: BODY, color: "E8DDF7", fontSize: 19 });
s1.addText("Designed with real photos, WhatsApp-first booking, an age quiz, and a garden of playful details", { x: 1.0, y: 5.2, w: 10.5, h: 0.7, fontFace: BODY, color: "D9C9F3", fontSize: 16 });
s1.addText("Little Blossoms Play School", { x: 1.0, y: 6.6, w: 8, h: 0.5, fontFace: BODY, color: "C6B3EE", fontSize: 14 });

// ---------- Slide 2 — Our website (light) ----------
const s2 = p.addSlide();
s2.background = bg(CREAM);
s2.addText("Our website today", { x: 0.8, y: 0.5, w: 8, h: 0.9, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s2.addText("A single-page site that feels like the school: warm, playful and easy for parents", { x: 0.8, y: 1.35, w: 8.4, h: 0.5, ...body() });
s2.addImage({ path: "img/our-hero.png", x: 9.0, y: 0.6, w: 3.9, h: 6.3, rounding: false, shadow: { color: "6F2DBD", blur: 8, offset: 3, angle: 45, opacity: 0.25 } });
s2.addText("What makes it special", { x: 0.8, y: 2.0, w: 6, h: 0.5, fontFace: FONT, bold: true, fontSize: 18, color: GRAPE });
const feats = [
  ["\u2713", "WhatsApp-first booking", "All enquiries land in one WhatsApp chat with details pre-filled"],
  ["\u2713", "Age quiz \u2192 auto-fill", "Parents answer 2 questions; the form pre-fills program + age"],
  ["\u2713", "Real photos", "Your playschool photos across hero, About, day cards and gallery"],
  ["\u2713", "Interactive scroll", "Scroll progress, spinning corner flowers, gentle reveals"],
  ["\u2713", "Address + Google Map", "Find-us bubble with directions link below the form"],
  ["\u2713", "Fully responsive", "Tested on a 360px mobile viewport \u2014 clean on every screen"],
];
feats.forEach((f, i) => {
  const y = 2.6 + i * 0.72;
  s2.addShape("oval", { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s2.addText(f[0], { x: 0.8, y, w: 0.5, h: 0.5, align: "center", bold: true, color: INK, fontSize: 16, fontFace: FONT, valign: "middle" });
  s2.addText(f[1], { x: 1.5, y: y - 0.02, w: 4.6, h: 0.34, fontFace: FONT, bold: true, fontSize: 14, color: INK });
  s2.addText(f[2], { x: 1.5, y: y + 0.3, w: 6.4, h: 0.4, ...body({ size: 12 }) });
});

// ---------- Slide 3 — Reference: Roughout ----------
const s3 = p.addSlide();
s3.background = bg(LAV);
s3.addText("Reference 1 \u2014 roughout.in (agency demo)", { x: 0.8, y: 0.5, w: 9, h: 0.7, fontFace: FONT, bold: true, fontSize: 22, color: GRAPE });
s3.addText("A playful, modular, rounded landing page in the same spirit as ours", { x: 0.8, y: 1.18, w: 9, h: 0.5, ...body() });
s3.addImage({ path: "img/ref-roughout.png", x: 0.8, y: 1.8, w: 4.6, h: 5.3, rounding: false, shadow: { color: "6F2DBD", blur: 8, offset: 3, angle: 45, opacity: 0.25 } });
s3.addText("What to borrow", { x: 5.8, y: 1.8, w: 6, h: 0.5, fontFace: FONT, bold: true, fontSize: 16, color: GRAPE });
["Big friendly headline inside a bold rounded card", "Custom icons to explain trust points simply", "Interactive age-selector personalises the journey", "Warm cream + lavender + sunny yellow palette", "Floating flower / organic shapes for charm"].forEach((t, i) => {
  s3.addShape("oval", { x: 5.8, y: 2.4 + i * 0.8, w: 0.32, h: 0.32, fill: { color: GRAPE }, line: { color: GRAPE } });
  s3.addText(t, { x: 6.3, y: 2.34 + i * 0.8, w: 5.6, h: 0.6, ...body({ size: 14 }) });
});

// ---------- Slide 4 — Reference: ZENZ ----------
const s4 = p.addSlide();
s4.background = bg(WHITE);
s4.addText("Reference 2 \u2014 ZENZ (fashion / editorial)", { x: 0.8, y: 0.5, w: 9, h: 0.7, fontFace: FONT, bold: true, fontSize: 22, color: INK });
s4.addText("A minimalist, editorial look built on white space and strong photography", { x: 0.8, y: 1.18, w: 9, h: 0.5, ...body() });
s4.addImage({ path: "img/ref-zenz.png", x: 0.8, y: 1.8, w: 4.6, h: 5.3, rounding: false, shadow: { color: "6F2DBD", blur: 8, offset: 3, angle: 45, opacity: 0.25 } });
s4.addText("What to borrow", { x: 5.8, y: 1.8, w: 6, h: 0.5, fontFace: FONT, bold: true, fontSize: 16, color: GRAPE });
["Generous white space \u2014 let content breathe", "High-quality real photos carry the page", "Asymmetric, editorial image placement", "One confident accent against a neutral base", "Craft markers that build trust step-by-step"].forEach((t, i) => {
  s4.addShape("oval", { x: 5.8, y: 2.4 + i * 0.8, w: 0.32, h: 0.32, fill: { color: GRAPE }, line: { color: GRAPE } });
  s4.addText(t, { x: 6.3, y: 2.34 + i * 0.8, w: 5.6, h: 0.6, ...body({ size: 14 }) });
});

// ---------- Slide 5 — Reference: Boutique hotel ----------
const s5 = p.addSlide();
s5.background = bg("F2F7F3");
s5.addText("Reference 3 \u2014 Boutique Hotel (hospitality)", { x: 0.8, y: 0.5, w: 9, h: 0.7, fontFace: FONT, bold: true, fontSize: 22, color: "2C5F2D" });
s5.addText("Full-bleed atmosphere + a conversion bar that is always in reach", { x: 0.8, y: 1.18, w: 9, h: 0.5, ...body() });
s5.addImage({ path: "img/ref-boutique.png", x: 0.8, y: 1.8, w: 4.6, h: 5.3, rounding: false, shadow: { color: "6F2DBD", blur: 8, offset: 3, angle: 45, opacity: 0.25 } });
s5.addText("What to borrow", { x: 5.8, y: 1.8, w: 6, h: 0.5, fontFace: FONT, bold: true, fontSize: 16, color: "2C5F2D" });
["Immersive full-bleed hero sets the mood", "A sticky booking bar keeps the key action visible", "Elegant serif + calm nature palette", "Alternate big images with clean info blocks", "A stylised map / journey graphic to guide visitors"].forEach((t, i) => {
  s5.addShape("oval", { x: 5.8, y: 2.4 + i * 0.8, w: 0.32, h: 0.32, fill: { color: "2C5F2D" }, line: { color: "2C5F2D" } });
  s5.addText(t, { x: 6.3, y: 2.34 + i * 0.8, w: 5.6, h: 0.6, ...body({ size: 14 }) });
});

// ---------- Slide 6 — Parent journey ----------
const s6 = p.addSlide();
s6.background = bg(CREAM);
s6.addText("The parent\u2019s journey", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 28, color: GRAPE });
const steps = [
  ["1", "Discover", "Scroll a playful hero, see real kids and the outdoor playground"],
  ["2", "Match", "Answer 2 quick questions \u2014 the quiz recommends the right program"],
  ["3", "Book", "Tap Book a visit \u2014 program and age are already filled in"],
  ["4", "Connect", "Add name, phone, centre and consent \u2014 one tap opens WhatsApp"],
  ["5", "Visit", "Find the address and map, then get directions \u2014 done"],
];
steps.forEach((st, i) => {
  const x = 0.8 + (i % 3) * 4.05;
  const y = 1.7 + Math.floor(i / 3) * 2.7;
  s6.addShape("roundRect", { x, y, w: 3.7, h: 2.3, rectRadius: 0.3, fill: { color: i === 2 ? LAV : WHITE }, line: { color: i === 2 ? GRAPE : "E4D9C0", width: 1.5 } });
  s6.addShape("oval", { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: SUN }, line: { color: SUN } });
  s6.addText(st[0], { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, align: "center", fontFace: FONT, bold: true, fontSize: 16, color: INK, valign: "middle" });
  s6.addText(st[1], { x: x + 1.05, y: y + 0.32, w: 2.4, h: 0.5, fontFace: FONT, bold: true, fontSize: 16, color: INK });
  s6.addText(st[2], { x: x + 0.3, y: y + 1.0, w: 3.1, h: 1.1, ...body({ size: 12 }) });
});

// ---------- Slide 7 — Design system ----------
const s7 = p.addSlide();
s7.background = bg(WHITE);
s7.addText("Design system", { x: 0.8, y: 0.4, w: 6, h: 0.7, fontFace: FONT, bold: true, fontSize: 26, color: GRAPE });
s7.addText("Brand palette", { x: 0.8, y: 1.2, w: 5, h: 0.5, fontFace: FONT, bold: true, fontSize: 15, color: GRAPE });
const cols = [
  ["6F2DBD", "Grape \u2014 primary"],
  ["FFC53D", "Sun \u2014 accent"],
  ["2EC4B6", "Aqua \u2014 pop"],
  ["4C9A56", "Leaf \u2014 pop"],
  ["FFF6E0", "Cream \u2014 ground"],
  ["3A2E29", "Ink \u2014 text"],
];
cols.forEach((c, i) => {
  const x = 0.8 + (i % 3) * 1.75;
  const y = 1.8 + Math.floor(i / 3) * 1.4;
  s7.addShape("rect", { x, y, w: 1.5, h: 0.9, fill: { color: c[0] }, line: { color: "E4D9C0", width: 1 } });
  s7.addText(c[1], { x, y: y + 0.95, w: 1.5, h: 0.5, align: "center", fontFace: BODY, fontSize: 10, color: MUT });
});
s7.addText("Typography", { x: 0.8, y: 4.5, w: 5, h: 0.5, fontFace: FONT, bold: true, fontSize: 15, color: GRAPE });
s7.addText("Quicksand", { x: 0.8, y: 5.0, w: 4, h: 0.6, fontFace: FONT, bold: true, fontSize: 30, color: INK });
s7.addText("Friendly rounded display for headlines", { x: 0.8, y: 5.7, w: 5, h: 0.4, ...body({ size: 12 }) });
s7.addText("Outfit", { x: 0.8, y: 6.1, w: 4, h: 0.5, fontFace: BODY, fontSize: 20, color: GRAPE });
s7.addText("Clean legible body for parent-facing text", { x: 0.8, y: 6.6, w: 5, h: 0.4, ...body({ size: 12 }) });
s7.addText("Motifs", { x: 6.2, y: 1.2, w: 5, h: 0.5, fontFace: FONT, bold: true, fontSize: 15, color: GRAPE });
[
  "White daisy with a sunny centre \u2014 in the corners, spinning gently on scroll",
  "Rounded cards, pill buttons and soft cloud-like shapes",
  "Flower-divider language replaced by clean rounded sections",
  "Photos framed with a cream border, like polaroids",
].forEach((t, i) => {
  s7.addShape("oval", { x: 6.2, y: 1.8 + i * 1.1, w: 0.3, h: 0.3, fill: { color: SUN }, line: { color: SUN } });
  s7.addText(t, { x: 6.7, y: 1.74 + i * 1.1, w: 5.4, h: 0.95, ...body({ size: 12 }) });
});

// ---------- Slide 8 — Why this works ----------
const s8 = p.addSlide();
s8.background = bg(GRAPE_D);
s8.addText("Why this design works", { x: 0.8, y: 0.5, w: 9, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: WHITE });
const why = [
  ["Trust first", "Real photos + honest fees + concrete ratios (4:1, 7:1) answer parent worries"],
  ["Zero-friction enquiries", "WhatsApp is the inbox: pre-filled details, one tap to send"],
  ["Parents finish fast", "The quiz pre-fills the form, so booking takes seconds, not retyping"],
  ["Looks alive", "Scroll progress, spinning flowers and reveals keep the scroll delightful"],
  ["Mobile-perfect", "Verified on a 360px viewport \u2014 no overlaps, no horizontal scroll"],
];
why.forEach((w, i) => {
  const y = 1.5 + i * 1.05;
  s8.addShape("oval", { x: 0.9, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s8.addText(String(i + 1), { x: 0.9, y, w: 0.5, h: 0.5, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 16, valign: "middle" });
  s8.addText(w[0], { x: 1.6, y: y - 0.05, w: 4.2, h: 0.4, fontFace: FONT, bold: true, fontSize: 16, color: WHITE });
  s8.addText(w[1], { x: 1.6, y: y + 0.35, w: 10.6, h: 0.55, fontFace: BODY, fontSize: 13, color: "D9C9F3" });
});

// ---------- Slide 9 — Rollout / next steps (dark) ----------
const s9 = p.addSlide();
s9.background = bg(GRAPE);
s9.addText("Where it lives & what\u2019s next", { x: 0.8, y: 0.5, w: 9, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: WHITE });
const live = [
  ["Live now", "little-blossoms.higgsfield.app — the main site"],
  ["Live now", "Website repository — deployed and version-controlled"],
  ["Next", "Custom domain — littleblossoms.in"],
  ["Next", "Final content: real address, fees and WhatsApp number"],
  ["Next", "Optional cover video for the feed card"],
];
live.forEach((w, i) => {
  const y = 1.5 + i * 1.0;
  s9.addShape("roundRect", { x: 0.9, y, w: 2.2, h: 0.55, rectRadius: 0.3, fill: { color: w[0] === "Live now" ? SUN : "FFFFFF" } });
  s9.addText(w[0], { x: 0.9, y, w: 2.2, h: 0.55, align: "center", fontFace: FONT, bold: true, fontSize: 13, color: INK, valign: "middle" });
  s9.addText(w[1], { x: 3.4, y: y + 0.02, w: 9, h: 0.5, fontFace: BODY, fontSize: 14, color: "E8DDF7" });
});

// ---------- Slide 10 — Thank you (dark) ----------
const s10 = p.addSlide();
s10.background = bg(GRAPE_D);
s10.addShape("ellipse", { x: 10.9, y: -1.4, w: 3.6, h: 3.6, fill: { color: WHITE, transparency: 90 }, line: { color: WHITE, transparency: 90 } });
s10.addShape("ellipse", { x: -1.5, y: 5.6, w: 3.4, h: 3.4, fill: { color: WHITE, transparency: 90 }, line: { color: WHITE, transparency: 90 } });
s10.addText("Thank you!", { x: 1.0, y: 2.2, w: 11.3, h: 1.4, fontFace: FONT, bold: true, fontSize: 54, color: WHITE });
s10.addText("Play today. Blossom tomorrow.", { x: 1.0, y: 3.7, w: 11.3, h: 0.8, fontFace: FONT, fontSize: 22, color: SUN, italic: true });
s10.addText("Little Blossoms Play School — 12 Blossom Lane, Green Park, Bengaluru", { x: 1.0, y: 5.4, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 14, color: "D9C9F3" });
s10.addText("WhatsApp +91 90000 00000  |  hello@littleblossoms.in", { x: 1.0, y: 6.0, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 14, color: "C6B3EE" });

p.writeFile({ fileName: "Little-Blossoms-Website-Presentation.pptx" }).then((f) => {
  console.log("saved", f);
});