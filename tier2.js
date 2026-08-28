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
s1.addText("TIER 2  \u2022  PROFESSIONAL", { x: 0.9, y: 0.8, w: 4.6, h: 1.0, align: "center", fontFace: FONT, bold: true, color: INK, fontSize: 14, valign: "middle" });
s1.addText("A complete digital presence,\nmanaged for you", { x: 1.0, y: 1.8, w: 11.3, h: 2.5, fontFace: FONT, bold: true, color: WHITE, fontSize: 44, lineSpacingMultiple: 1.05 });
s1.addText("Website + domain + server + 2 months free support, with optional add-ons that make your school run smoother", { x: 1.0, y: 4.5, w: 10.8, h: 0.8, fontFace: BODY, color: "E8DDF7", fontSize: 16 });
s1.addText("Rs 80,000  \u2014 full package", { x: 1.0, y: 5.5, w: 8, h: 0.8, fontFace: FONT, bold: true, color: SUN, fontSize: 26 });
s1.addText("Little Blossoms Web Studio", { x: 1.0, y: 6.7, w: 8, h: 0.4, fontFace: BODY, color: "C6B3EE", fontSize: 13 });

// ---------- S2 What's included ----------
const s2 = p.addSlide();
s2.background = { color: CREAM };
s2.addText("What\u2019s included in Rs 80,000", { x: 0.8, y: 0.5, w: 9, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s2.addText("Everything in Tier 1, plus the things that make a website actually work for the school", { x: 0.8, y: 1.3, w: 9.6, h: 0.5, ...body() });
const inc = [
  ["Website", "Full design + development, mobile responsive"],
  ["Issues resolved", "Any functional problems fixed before handover \u2014 working from day one"],
  ["2 months free", "Free adjustments and support for the first 2 months"],
  ["Domain", "Your domain (e.g. littleblossoms.in) \u2014 set up for you"],
  ["Server", "Hosting arranged and configured \u2014 handled for you"],
  ["Extras from Tier 1", "WhatsApp booking, map, quiz, gallery, testimonies"],
];
inc.forEach((f, i) => {
  const y = 1.95 + i * 0.78;
  s2.addShape("oval", { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: SUN }, line: { color: SUN } });
  s2.addText("\u2713", { x: 0.8, y, w: 0.5, h: 0.5, align: "center", bold: true, color: INK, fontSize: 17, fontFace: FONT, valign: "middle" });
  s2.addText(f[0], { x: 1.5, y: y - 0.02, w: 3.0, h: 0.34, fontFace: FONT, bold: true, fontSize: 14, color: INK });
  s2.addText(f[1], { x: 4.6, y: y + 0.02, w: 8, h: 0.55, ...body({ size: 13 }) });
});
s2.addShape("roundRect", { x: 0.8, y: 6.6, w: 11.6, h: 0.7, rectRadius: 0.35, fill: { color: LAV }, line: { color: GRAPE, width: 1 } });
s2.addText("A working, worry-free site: fixed, hosted, domain-ready and free to adjust for 2 months.", { x: 0.8, y: 6.6, w: 11.6, h: 0.7, align: "center", fontFace: BODY, fontSize: 14, color: GRAPE, valign: "middle" });

// ---------- S3 Add-ons ----------
const s3 = p.addSlide();
s3.background = { color: WHITE };
s3.addText("Optional add-ons", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s3.addText("Make your website grow with your school \u2014 add what you need, when you need it", { x: 0.8, y: 1.3, w: 9.4, h: 0.5, ...body() });
const adds = [
  ["Google Ads", "Managed monthly campaign", "Ad management (setup + running)",
   "4.0", "Monthly"],
  ["WhatsApp automation", "Auto-replies, follow-ups, broadcasts to parents", "Systematised communication",
   "2.5", "Monthly"],
  ["Admission portal", "Online application + enquiry tracking", "Digital admissions",
   "1.5", "Monthly"],
  ["School database", "Student records, staff, attendance base", "Data organised",
   "1.5", "Monthly"],
];
adds.forEach((a, i) => {
  const x = 0.8 + (i % 2) * 6.0;
  const y = 1.95 + Math.floor(i / 2) * 2.5;
  s3.addShape("roundRect", { x, y, w: 5.7, h: 2.2, rectRadius: 0.25, fill: { color: LAV }, line: { color: GRAPE, width: 1 } });
  s3.addText(a[0], { x: x + 0.3, y: y + 0.25, w: 4.6, h: 0.5, fontFace: FONT, bold: true, fontSize: 17, color: GRAPE });
  s3.addText(a[1], { x: x + 0.3, y: y + 0.8, w: 4.9, h: 0.7, ...body({ size: 12.5 }) });
  s3.addText(a[2], { x: x + 0.3, y: y + 1.5, w: 4.9, h: 0.5, fontFace: BODY, italic: true, size: 11.5, color: MUT });
});
s3.addShape("roundRect", { x: 0.8, y: 6.4, w: 11.7, h: 0.7, rectRadius: 0.35, fill: { color: SUN }, line: { color: SUN } });
s3.addText("Add-on pricing is simple and low: every add-on is Rs 1,000 or below per month.", { x: 0.8, y: 6.4, w: 11.7, h: 0.7, align: "center", fontFace: BODY, bold: true, fontSize: 14, color: INK, valign: "middle" });

// ---------- S4 Monthly service ----------
const s4 = p.addSlide();
s4.background = { color: CREAM };
s4.addText("Monthly service plan", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s4.addText("Keep the site fresh while you run the school \u2014 we handle the updates", { x: 0.8, y: 1.3, w: 9.6, h: 0.5, ...body() });
const svc = [
  ["Content swaps", "Change photos, fees, programs, notices and offers"],
  ["Seasonal updates", "Admission-open banners, events, celebrations"],
  ["Text & page edits", "Reword sections, add news, update timings"],
  ["Small fixes", "Anything broken or outdated gets corrected"],
];
svc.forEach((f, i) => {
  const y = 2.0 + i * 0.85;
  s4.addShape("oval", { x: 0.8, y, w: 0.45, h: 0.45, fill: { color: GRAPE }, line: { color: GRAPE } });
  s4.addText(f[0], { x: 1.5, y: y - 0.05, w: 3.4, h: 0.4, fontFace: FONT, bold: true, fontSize: 15, color: INK });
  s4.addText(f[1], { x: 5.0, y: y, w: 7.4, h: 0.5, ...body({ size: 13 }) });
});
s4.addShape("roundRect", { x: 0.8, y: 5.6, w: 11.7, h: 0.9, rectRadius: 0.4, fill: { color: GRAPE }, line: { color: GRAPE } });
s4.addText("Rs 1,000 or below per month", { x: 0.8, y: 5.6, w: 11.7, h: 0.9, align: "center", fontFace: FONT, bold: true, fontSize: 24, color: WHITE, valign: "middle" });
s4.addText("Unlimited update requests within the plan \u2014 content refreshes handled every month.", { x: 0.8, y: 6.6, w: 11.7, h: 0.5, align: "center", ...body({ size: 12.5 }) });

// ---------- S5 Self-managed option ----------
const s5 = p.addSlide();
s5.background = { color: WHITE };
s5.addText("Self-managed option", { x: 0.8, y: 0.5, w: 8, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s5.addText("For schools that want to upload and maintain their own content \u2014 full control, no dependency", { x: 0.8, y: 1.3, w: 10.4, h: 0.5, ...body() });
const box = [
  ["2 websites", "One public site + one private user/admin site"],
  ["1 backend", "An admin store where you manage everything"],
  ["1 database", "Your school data stored safely in one place"],
  ["2GB storage", "Space for photos, documents and uploads"],
];
box.forEach((f, i) => {
  const y = 2.0 + i * 1.0;
  s5.addShape("rect", { x: 0.9, y, w: 3.3, h: 0.8, fill: { color: LAV } });
  s5.addText(f[0], { x: 0.9, y, w: 3.3, h: 0.8, align: "center", fontFace: FONT, bold: true, fontSize: 15, color: GRAPE, valign: "middle" });
  s5.addText(f[1], { x: 4.5, y, w: 8, h: 0.8, ...body({ size: 14, valign: "middle" }) });
});
s5.addShape("roundRect", { x: 0.9, y: 6.1, w: 11.5, h: 0.9, rectRadius: 0.4, fill: { color: GRAPE }, line: { color: GRAPE } });
s5.addText("Rs 1,80,000 \u2013 Rs 2,00,000", { x: 0.9, y: 6.1, w: 11.5, h: 0.9, align: "center", fontFace: FONT, bold: true, fontSize: 26, color: WHITE, valign: "middle" });

// ---------- S6 Compare tiers ----------
const s6 = p.addSlide();
s6.background = { color: CREAM };
s6.addText("Tier 1 vs Tier 2 at a glance", { x: 0.8, y: 0.5, w: 9, h: 0.8, fontFace: FONT, bold: true, fontSize: 30, color: GRAPE });
s6.addShape("rect", { x: 1.0, y: 1.5, w: 5.4, h: 0.7, fill: { color: LAV } });
s6.addText("Tier 1 \u2022 Starter \u2014 Rs 50,000 + domain", { x: 1.0, y: 1.5, w: 5.4, h: 0.7, align: "center", fontFace: FONT, bold: true, fontSize: 15, color: GRAPE, valign: "middle" });
s6.addShape("rect", { x: 6.9, y: 1.5, w: 5.4, h: 0.7, fill: { color: GRAPE } });
s6.addText("Tier 2 \u2022 Professional \u2014 Rs 80,000", { x: 6.9, y: 1.5, w: 5.4, h: 0.7, align: "center", fontFace: FONT, bold: true, fontSize: 15, color: WHITE, valign: "middle" });
const cmp = [
  ["Website + WhatsApp + map", "\u2713", "\u2713"],
  ["Issues resolved before handover", "\u2013", "\u2713"],
  ["2 months free support", "\u2013", "\u2713"],
  ["Domain & server handled", "\u2013", "\u2713"],
  ["Add-ons (Ads, automation, portal, DB)", "\u2013", "\u2713"],
  ["Monthly content service Rs 1,000", "\u2013", "\u2713"],
];
cmp.forEach((c, i) => {
  const y = 2.4 + i * 0.7;
  s6.addText(c[0], { x: 1.0, y, w: 5.3, h: 0.5, fontFace: BODY, fontSize: 13.5, color: INK, valign: "middle" });
  s6.addText(c[1], { x: 6.5, y, w: 0.7, h: 0.5, align: "center", fontFace: FONT, bold: true, fontSize: 15, color: c[1] === "\u2713" ? LEAF : "C9B8A8", valign: "middle" });
  s6.addText(c[2], { x: 12.0, y, w: 0.7, h: 0.5, align: "center", fontFace: FONT, bold: true, fontSize: 15, color: c[2] === "\u2713" ? LEAF : "C9B8A8", valign: "middle" });
});
s6.addShape("roundRect", { x: 0.8, y: 6.6, w: 11.7, h: 0.6, rectRadius: 0.3, fill: { color: SUN }, line: { color: SUN } });
s6.addText("Need autonomy? Choose the Self-managed option (Rs 1.8\u20132.0L) with 2 sites + backend + database.", { x: 0.8, y: 6.6, w: 11.7, h: 0.6, align: "center", fontFace: BODY, bold: true, fontSize: 13, color: INK, valign: "middle" });

// ---------- S7 Close ----------
const s7 = p.addSlide();
s7.background = { color: GRAPE_D };
s7.addShape("ellipse", { x: 11.0, y: -1.3, w: 3.4, h: 3.4, fill: { color: WHITE, transparency: 90 }, line: { color: WHITE, transparency: 90 } });
s7.addText("Grow with confidence", { x: 1.0, y: 2.0, w: 11.3, h: 1.2, fontFace: FONT, bold: true, fontSize: 46, color: WHITE });
s7.addText("Professional Tier \u2014 Rs 80,000  |  Self-managed \u2014 Rs 1.8\u20132.0L", { x: 1.0, y: 3.5, w: 11.3, h: 0.8, fontFace: FONT, fontSize: 21, color: SUN, italic: true });
s7.addText("Let\u2019s plan the right package for your school.", { x: 1.0, y: 5.2, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 15, color: "D9C9F3" });
s7.addText("WhatsApp +91 90000 00000  |  hello@littleblossoms.in", { x: 1.0, y: 6.0, w: 11.3, h: 0.5, fontFace: BODY, fontSize: 14, color: "C6B3EE" });

p.writeFile({ fileName: "Little-Blossoms-Tier2-Professional-Pitch.pptx" }).then((f) => console.log("saved", f));