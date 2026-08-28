const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
p.layout = "WIDE";

const INK = "3A2E29", GRAPE = "6F2DBD", GRAPE_D = "4A1D87", CREAM = "FFF6E0",
  WHITE = "FFFFFF", SUN = "FFC53D", TEAL = "2EC4B6", LEAF = "4C9A56", LAV = "F3EDFB",
  MUT = "725A4A", FONT = "Trebuchet MS", BODY = "Calibri";

const body = (o = {}) => Object.assign({ fontFace: BODY, color: MUT, size: 13, breakLine: false }, o);

// ---------- S1 Title ----------
const s1 = p.addSlide();
s1.background = { color: GRAPE_D };
s1.addShape("ellipse", { x: -1.2, y: -1.2, w: 3, h: 3, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
s1.addShape("ellipse", { x: 11.4, y: 5.8, w: 2.9, h: 2.9, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
s1.addShape("roundRect", { x: 0.9, y: 0.8, w: 4.6, h: 1.0, rectRadius: 0.5, fill: { color: SUN }, line: { color: SUN } });
s1.addText("TIER 1  \u2022  STARTER", { x: 0.9, y: 0.8, w: 4.6, h: 1.0, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 14, valign: "middle" });
s1.addText("Your school\u2019s website,\nready to welcome parents", { x: 1.0, y: 1.9, w: 11.3, h: 2.5, fontFace: FONT, bold: true, color: WHITE, fontSize: 44, lineSpacingMultiple: 1.05 });
s1.addText("A complete, mobile-friendly website with WhatsApp booking \u2014 delivered, hosted and supported", { x: 1.0, y: 4.6, w: 10.6, h: 0.8, fontFace: BODY, color: "E8DDF7", fontSize: 17 });
s1.addText("Rs 50,000  + domain (billed directly by the registrar)", { x: 1.0, y: 5.6, w: 10, h: 0.9, fontFace: FONT, bold: true, color: SUN, fontSize: 26 });
s1.addText("Little Blossoms Web Studio", { x: 1.0, y: 6.7, w: 8, h: 0.4, fontFace: BODY, color: "C6B3EE", fontSize: 13 });

// ---------- S2 What you get ----------
const s2 = p.addSlide();
s2.background = { color: CREAM };
s2.addText("What\u2019s included", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s2.addText("Everything a school needs to look professional online and start receiving enquiries", { x: 0.8, y: 1.3, w: 11.6, h: 0.5, ...body() });
const inc = [
  ["Website", "Single-page public site, mobile responsive, fast loading"],
  ["Design", "Custom playful design in your brand colours + your real photos"],
  ["WhatsApp", "Booking form that opens WhatsApp with details pre-filled"],
  ["Map", "Address + Google Map + get-directions block"],
  ["Pages", "Programs, fees, gallery, parent testimonials, contact"],
  ["Extras", "Age quiz, scroll effects, white-daisy corner accents"],
];
inc.forEach((f, i) => {
  const y = 2.0 + i * 0.78;
  s2.addShape("oval", { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s2.addText("\u2713", { x: 0.8, y, w: 0.5, h: 0.5, align: "center", bold: true, color: INK, fontSize: 17, fontFace: FONT, valign: "middle" });
  s2.addText(f[0], { x: 1.5, y: y - 0.02, w: 2.6, h: 0.34, fontFace: FONT, bold: true, fontSize: 15, color: INK });
  s2.addText(f[1], { x: 4.3, y: y + 0.02, w: 8, h: 0.5, ...body({ size: 13 }) });
});
s2.addShape("roundRect", { x: 0.8, y: 6.6, w: 11.6, h: 0.7, rectRadius: 0.35, fill: { color: LAV }, line: { color: GRAPE, width: 1 } });
s2.addText("One-time build fee. Your site is deployed on a live link and is yours to keep.", { x: 0.8, y: 6.6, w: 11.6, h: 0.7, align: "center", fontFace: BODY, fontSize: 14, color: GRAPE, valign: "middle" });

// ---------- S3 Pricing & terms ----------
const s3 = p.addSlide();
s3.background = { color: WHITE };
s3.addText("Pricing & terms", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
const rows = [
  ["Website build", "Rs 50,000", "Design, development, testing, deploy"],
  ["Domain", "Actual cost", "Billed directly by the registrar (e.g. .in or .com)"],
  ["Hosting", "Included first year", "Fast, secure hosting on our platform"],
  ["Support", "3 months", "Issues resolved free for 3 months after launch"],
  ["Timeline", "4\u20136 weeks", "From content handover to live site"],
];
rows.forEach((r, i) => {
  const y = 1.5 + i * 0.95;
  s2; s3.addShape("rect", { x: 0.9, y, w: 3.4, h: 0.75, fill: { color: LAV } });
  s3.addText(r[0], { x: 0.9, y, w: 3.4, h: 0.75, align: "center", fontFace: FONT, bold: true, fontSize: 14, color: GRAPE, valign: "middle" });
  s3.addText(r[1], { x: 4.5, y, w: 3.2, h: 0.75, fontFace: FONT, bold: true, fontSize: 15, color: INK, valign: "middle" });
  s3.addText(r[2], { x: 7.9, y, w: 4.6, h: 0.75, ...body({ size: 12.5, valign: "middle" }) });
});
s3.addShape("roundRect", { x: 0.9, y: 6.4, w: 11.5, h: 0.75, rectRadius: 0.35, fill: { color: SUN }, line: { color: SUN } });
s3.addText("Domain charge is not included in the Rs 50,000 \u2014 you pay the registrar directly, no markup.", { x: 0.9, y: 6.4, w: 11.5, h: 0.75, align: "center", fontFace: BODY, bold: true, fontSize: 14, color: INK, valign: "middle" });

// ---------- S4 What's next / support ----------
const s4 = p.addSlide();
s4.background = { color: GRAPE_D };
s4.addText("What happens next", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: WHITE });
const steps = [
  ["1", "Kickoff", "You share your logo, photos and basic info"],
  ["2", "Design", "We design your site and confirm with you"],
  ["3", "Build & test", "We develop, test and fix anything broken"],
  ["4", "Go live", "Your site is deployed on a live link"],
  ["5", "Support", "3 months of free issue resolution"],
];
steps.forEach((st, i) => {
  const x = 0.8 + (i % 3) * 4.1;
  const y = 1.7 + Math.floor(i / 3) * 2.8;
  s4.addShape("roundRect", { x, y, w: 3.8, h: 2.4, rectRadius: 0.3, fill: { color: WHITE }, line: { color: GRAPE, transparency: 35, width: 1.5 } });
  s4.addShape("oval", { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: SUN }, line: { color: SUN } });
  s4.addText(st[0], { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, align: "center", fontFace: FONT, bold: true, fontSize: 17, color: INK, valign: "middle" });
  s4.addText(st[1], { x: x + 1.05, y: y + 0.34, w: 2.5, h: 0.5, fontFace: FONT, bold: true, fontSize: 17, color: GRAPE });
  s4.addText(st[2], { x: x + 0.3, y: y + 1.05, w: 3.2, h: 1.1, fontFace: BODY, fontSize: 12.5, color: MUT });
});

// ---------- S5 Why ----------
const s5 = p.addSlide();
s5.background = { color: WHITE };
s5.addText("Why this is a smart start", { x: 0.8, y: 0.5, w: 9, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
const why = [
  ["Looks professional", "A polished online presence parents can trust before they visit"],
  ["Brings enquiries", "WhatsApp-first booking means every visitor can reach you in one tap"],
  ["Works everywhere", "Phones, tablets, desktops \u2014 tested on a 360px mobile viewport"],
  ["Fast and affordable", "One fixed fee, no surprises, live within weeks"],
  ["Room to grow", "Ready to upgrade to Tier 2 with add-ons whenever you need"],
];
why.forEach((w, i) => {
  const y = 1.5 + i * 1.02;
  s5.addShape("oval", { x: 0.9, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s5.addText(String(i + 1), { x: 0.9, y, w: 0.5, h: 0.5, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 16, valign: "middle" });
  s5.addText(w[0], { x: 1.6, y: y - 0.05, w: 4.2, h: 0.4, fontFace: FONT, bold: true, fontSize: 16, color: INK });
  s5.addText(w[1], { x: 1.6, y: y + 0.35, w: 10.8, h: 0.55, fontFace: BODY, fontSize: 13, color: MUT });
});

// ---------- S6 Close ----------
const s6 = p.addSlide();
s6.background = { color: GRAPE_D };
s6.addShape("ellipse", { x: 11.0, y: -1.3, w: 3.4, h: 3.4, fill: { color: WHITE, transparency: 90 }, line: { color: WHITE, transparency: 90 } });
s6.addText("Ready when you are", { x: 1.0, y: 2.0, w: 11.3, h: 1.2, fontFace: FONT, bold: true, fontSize: 46, color: WHITE });
s6.addText("Starter Tier \u2014 Rs 50,000 + domain charge", { x: 1.0, y: 3.5, w: 11.3, h: 0.8, fontFace: FONT, fontSize: 22, color: SUN, italic: true });
s6.addText("Let\u2019s plan your school\u2019s website together.", { x: 1.0, y: 5.2, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 15, color: "D9C9F3" });
s6.addText("WhatsApp +91 90000 00000  |  hello@littleblossoms.in", { x: 1.0, y: 6.0, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 14, color: "C6B3EE" });

p.writeFile({ fileName: "Little-Blossoms-Tier1-Starter-Pitch.pptx" }).then((f) => console.log("saved", f));