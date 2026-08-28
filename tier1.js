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
s1.addShape("ellipse", { x: 11.4, y: -1.2, w: 2.9, h: 2.9, fill: { color: WHITE, transparency: 88 }, line: { color: WHITE, transparency: 88 } });
s1.addShape("roundRect", { x: 0.9, y: 0.8, w: 4.6, h: 1.0, rectRadius: 0.5, fill: { color: SUN }, line: { color: SUN } });
s1.addText("TIER 1  \u2022  STARTER", { x: 0.9, y: 0.8, w: 4.6, h: 1.0, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 14, valign: "middle" });
s1.addText("Your school\u2019s website,\nready to welcome parents", { x: 1.0, y: 1.9, w: 11.3, h: 2.5, fontFace: FONT, bold: true, color: WHITE, fontSize: 44, lineSpacingMultiple: 1.05 });
s1.addText("A complete, mobile-friendly website with WhatsApp booking \u2014 plus the designs you need to start", { x: 1.0, y: 4.55, w: 11.2, h: 0.8, fontFace: BODY, color: "E8DDF7", fontSize: 17 });
s1.addText("Rs 50,000  + domain (billed directly by the registrar)", { x: 1.0, y: 5.6, w: 10, h: 0.9, fontFace: FONT, bold: true, color: SUN, fontSize: 26 });
s1.addText("Little Blossoms Web Studio", { x: 1.0, y: 6.7, w: 8, h: 0.4, fontFace: BODY, color: "C6B3EE", fontSize: 13 });

// ---------- S2 What's included ----------
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
  const y = 1.95 + i * 0.72;
  s2.addShape("oval", { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s2.addText("\u2713", { x: 0.8, y, w: 0.5, h: 0.5, align: "center", bold: true, color: INK, fontSize: 17, fontFace: FONT, valign: "middle" });
  s2.addText(f[0], { x: 1.5, y: y - 0.02, w: 2.6, h: 0.34, fontFace: FONT, bold: true, fontSize: 15, color: INK });
  s2.addText(f[1], { x: 4.3, y: y + 0.02, w: 8, h: 0.5, ...body({ size: 13 }) });
});
s2.addShape("roundRect", { x: 0.8, y: 6.55, w: 11.6, h: 0.7, rectRadius: 0.35, fill: { color: LAV }, line: { color: GRAPE, width: 1 } });
s2.addText("One-time build fee. Your site is deployed on a live link and is yours to keep.", { x: 0.8, y: 6.55, w: 11.6, h: 0.7, align: "center", fontFace: BODY, fontSize: 14, color: GRAPE, valign: "middle" });

// ---------- S3 Designs & start ----------
const s3b = p.addSlide();
s3b.background = { color: WHITE };
s3b.addText("Designs & everything needed to start", { x: 0.8, y: 0.5, w: 11, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s3b.addText("We don\u2019t just build the site \u2014 we give you the designs it needs and everything required to go live", { x: 0.8, y: 1.3, w: 11.4, h: 0.5, ...body() });
const dsg = [
  ["Additional designs", "Logos, banners, icons, poster/brochure graphics, social-media creatives \u2014 designed for you in your brand style"],
  ["Photos & content", "Use your photos or we guide you \u2014 we advise and prepare what the site needs"],
  ["Domain & hosting", "We suggest and set up your domain (e.g\. littleblossoms\.in) and hosting"],
  ["Start checklist", "Name, logo, photos, basic text, contact details, preferred domain \u2014 we walk you through it"],
];
dsg.forEach((f, i) => {
  const y = 2.1 + i * 1.05;
  s3b.addShape("oval", { x: 0.9, y, w: 0.5, h: 0.5, fill: { color: TEAL }, line: { color: TEAL } });
  s3b.addText(String(i + 1), { x: 0.9, y, w: 0.5, h: 0.5, align: "center", bold: true, color: WHITE, fontSize: 16, fontFace: FONT, valign: "middle" });
  s3b.addText(f[0], { x: 1.6, y: y - 0.03, w: 3.6, h: 0.4, fontFace: FONT, bold: true, fontSize: 16, color: INK });
  s3b.addText(f[1], { x: 5.3, y: y + 0.02, w: 7.2, h: 0.85, ...body({ size: 13 }) });
});
s3b.addShape("roundRect", { x: 0.9, y: 6.55, w: 11.5, h: 0.7, rectRadius: 0.35, fill: { color: SUN }, line: { color: SUN } });
s3b.addText("Main promise: you hand us your school, we deliver a ready-to-use website with its complete visual kit.", { x: 0.9, y: 6.55, w: 11.5, h: 0.7, align: "center", fontFace: BODY, bold: true, fontSize: 14, color: INK, valign: "middle" });

// ---------- S4 Pricing & terms ----------
const s4 = p.addSlide();
s4.background = { color: WHITE };
s4.addText("Pricing & terms", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
const rows = [
  ["Website build", "Rs 50,000", "Design, development, testing, deploy"],
  ["Additional designs", "Included", "Logo, banners, icons, social graphics for the site"],
  ["Domain", "Actual cost", "Billed directly by the registrar (e.g. .in or .com)"],
  ["Hosting", "Included first year", "Fast, secure hosting on our platform"],
  ["Support", "3 months", "Issues resolved free for 3 months after launch"],
  ["Timeline", "4\u20136 weeks", "From content handover to live site"],
];
rows.forEach((r, i) => {
  const y = 1.4 + i * 0.82;
  s4.addShape("rect", { x: 0.9, y, w: 3.4, h: 0.7, fill: { color: LAV } });
  s4.addText(r[0], { x: 0.9, y, w: 3.4, h: 0.7, align: "center", fontFace: FONT, bold: true, fontSize: 13.5, color: GRAPE, valign: "middle" });
  s4.addText(r[1], { x: 4.5, y, w: 3.2, h: 0.7, fontFace: FONT, bold: true, fontSize: 14.5, color: INK, valign: "middle" });
  s4.addText(r[2], { x: 7.9, y, w: 4.6, h: 0.7, ...body({ size: 12, valign: "middle" }) });
});
s4.addShape("roundRect", { x: 0.9, y: 6.4, w: 11.5, h: 0.75, rectRadius: 0.35, fill: { color: SUN }, line: { color: SUN } });
s4.addText("Domain charge is not included in the Rs 50,000 \u2014 you pay the registrar directly, no markup.", { x: 0.9, y: 6.4, w: 11.5, h: 0.75, align: "center", fontFace: BODY, bold: true, fontSize: 14, color: INK, valign: "middle" });

// ---------- S5 What happens next ----------
const s5 = p.addSlide();
s5.background = { color: GRAPE_D };
s5.addText("What happens next", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: WHITE });
const steps = [
  ["1", "Kickoff", "You share your logo, photos and basic info"],
  ["2", "Designs", "We design logos, banners and site visuals for you"],
  ["3", "Build & test", "We develop, test and fix anything broken"],
  ["4", "Go live", "Your site is deployed on a live link"],
  ["5", "Support", "3 months of free issue resolution"],
];
steps.forEach((st, i) => {
  const x = 0.8 + (i % 3) * 4.1;
  const y = 1.7 + Math.floor(i / 3) * 2.8;
  s5.addShape("roundRect", { x, y, w: 3.8, h: 2.4, rectRadius: 0.3, fill: { color: WHITE }, line: { color: GRAPE, transparency: 35, width: 1.5 } });
  s5.addShape("oval", { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: SUN }, line: { color: SUN } });
  s5.addText(st[0], { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, align: "center", fontFace: FONT, bold: true, fontSize: 17, color: INK, valign: "middle" });
  s5.addText(st[1], { x: x + 1.05, y: y + 0.34, w: 2.8, h: 0.5, fontFace: FONT, bold: true, fontSize: 17, color: GRAPE });
  s5.addText(st[2], { x: x + 0.3, y: y + 1.05, w: 3.2, h: 1.1, fontFace: BODY, fontSize: 12.5, color: MUT });
});

// ---------- S6 Self-managed ----------
const s6 = p.addSlide();
s6.background = { color: CREAM };
s6.addText("Self-managed option (full control)", { x: 0.8, y: 0.5, w: 11, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s6.addText("Prefer to run everything yourself? Choose the self-managed upgrade \u2014 available on Tier 1 as well", { x: 0.8, y: 1.3, w: 11.4, h: 0.5, ...body() });
const boxx = [
  ["2 websites", "One public site + one private user/admin site"],
  ["1 backend", "An admin store where you manage everything"],
  ["1 database", "Your school data stored safely in one place"],
  ["2GB storage", "Space for photos, documents and uploads"],
];
boxx.forEach((f, i) => {
  const y = 2.0 + i * 1.0;
  s6.addShape("rect", { x: 0.9, y, w: 3.3, h: 0.8, fill: { color: LAV } });
  s6.addText(f[0], { x: 0.9, y, w: 3.3, h: 0.8, align: "center", fontFace: FONT, bold: true, fontSize: 15, color: GRAPE, valign: "middle" });
  s6.addText(f[1], { x: 4.5, y, w: 8, h: 0.8, ...body({ size: 14, valign: "middle" }) });
});
s6.addShape("roundRect", { x: 0.9, y: 6.1, w: 11.5, h: 0.9, rectRadius: 0.4, fill: { color: GRAPE }, line: { color: GRAPE } });
s6.addText("Rs 1,80,000 \u2013 Rs 2,00,000", { x: 0.9, y: 6.1, w: 11.5, h: 0.9, align: "center", fontFace: FONT, bold: true, fontSize: 26, color: WHITE, valign: "middle" });

// ---------- S7 Why ----------
const s7 = p.addSlide();
s7.background = { color: WHITE };
s7.addText("Why this is a smart start", { x: 0.8, y: 0.5, w: 9, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
const why = [
  ["Looks professional", "A polished online presence parents can trust before they visit"],
  ["Brings enquiries", "WhatsApp-first booking means every visitor can reach you in one tap"],
  ["Designs included", "Logos, banners and graphics you don\u2019t have to commission separately"],
  ["Works everywhere", "Phones, tablets, desktops \u2014 tested on a 360px mobile viewport"],
  ["Room to grow", "Upgrade to Tier 2 or the Self-managed option whenever you need"],
];
why.forEach((w, i) => {
  const y = 1.5 + i * 1.02;
  s7.addShape("oval", { x: 0.9, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s7.addText(String(i + 1), { x: 0.9, y, w: 0.5, h: 0.5, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 16, valign: "middle" });
  s7.addText(w[0], { x: 1.6, y: y - 0.05, w: 4.2, h: 0.4, fontFace: FONT, bold: true, fontSize: 16, color: INK });
  s7.addText(w[1], { x: 1.6, y: y + 0.35, w: 10.8, h: 0.55, fontFace: BODY, fontSize: 13, color: MUT });
});

// ---------- S8 Close ----------
const s8 = p.addSlide();
s8.background = { color: GRAPE_D };
s8.addShape("ellipse", { x: 11.0, y: -1.3, w: 3.4, h: 3.4, fill: { color: WHITE, transparency: 90 }, line: { color: WHITE, transparency: 90 } });
s8.addText("Ready when you are", { x: 1.0, y: 2.0, w: 11.3, h: 1.2, fontFace: FONT, bold: true, fontSize: 46, color: WHITE });
s8.addText("Starter Tier \u2014 Rs 50,000 + domain  |  Self-managed \u2014 Rs 1.8\u20132.0L", { x: 1.0, y: 3.5, w: 11.3, h: 0.8, fontFace: FONT, fontSize: 20, color: SUN, italic: true });
s8.addText("Let\u2019s plan your school\u2019s website together.", { x: 1.0, y: 5.2, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 15, color: "D9C9F3" });
s8.addText("WhatsApp +91 90000 00000  |  hello@littleblossoms.in", { x: 1.0, y: 6.0, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 14, color: "C6B3EE" });

p.writeFile({ fileName: "Little-Blossoms-Tier1-Starter-Pitch.pptx" }).then((f) => console.log("saved", f));