"use strict";
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_4x3";   // 10" × 7.5"
pres.title  = "VeriSource – AI-Powered Fake News Detection Platform";

// ── Paths ──────────────────────────────────────────────────────────────────
const LM = "C:/Users/Admin/OneDrive/Desktop/lawyer_unpacked/ppt/media";
const RM = "C:/Users/Admin/OneDrive/Desktop/report_unpacked/word/media";

const IUL_L = LM + "/image1.png";
const WAVE  = LM + "/image2.png";
const IUL_R = LM + "/image3.png";
const REACT = LM + "/image4.png";

// Report images (re-numbered after DOCX regeneration)
const IMG = {
  usecase : { p: RM+"/image2.png",   pw:1385, ph:659  },
  dfd0    : { p: RM+"/image3.png",   pw:1476, ph:802  },
  dfd1    : { p: RM+"/image4.png",   pw:1463, ph:988  },
  erd     : { p: RM+"/image7.png",   pw:1536, ph:1024 },
  seq_u   : { p: RM+"/image5.png",   pw:1379, ph:920  },
  seq_p   : { p: RM+"/image6.png",   pw:1385, ph:923  },
  welcome : { p: RM+"/image8.png",   pw:1385, ph:659  },
  login   : { p: RM+"/image9.png",   pw:898,  ph:823  },
  anlyz   : { p: RM+"/image10.png",  pw:893,  ph:817  },
  dash    : { p: RM+"/image11.png",  pw:1918, ph:911  },
  mob_en  : { p: RM+"/image12.png",  pw:1022, ph:906  },
  press1  : { p: RM+"/image13.jpeg", pw:738,  ph:1600 },
  chat_w  : { p: RM+"/image14.jpeg", pw:738,  ph:1600 },
  chat_m  : { p: RM+"/image15.png",  pw:1331, ph:811  },
  press2  : { p: RM+"/image16.jpeg", pw:738,  ph:1600 },
  mob_home: { p: RM+"/image17.png",  pw:1286, ph:624  },
};

// ── Colours ────────────────────────────────────────────────────────────────
const BLUE   = "1A5F9E";
const DBLUE  = "154D82";
const LBLUE  = "D6E8F7";
const LLBLUE = "EAF4FC";
const BLACK  = "1C2833";
const DGRAY  = "5D6D7E";
const WHITE  = "FFFFFF";
const GREEN  = "1E8449";
const TEAL   = "117A65";
const PURP   = "6C3483";
const ORNG   = "BA4A00";
const NAVY   = "1B2631";
const RED    = "922B21";

// ── Slide counter ──────────────────────────────────────────────────────────
let SN = 0;
const TOTAL = 21;   // numbered slides (title excluded)

// ── Helpers ────────────────────────────────────────────────────────────────
function fit(pw, ph, maxW, maxH) {
  const r = pw / ph;
  if (r > maxW / maxH) return { w: +maxW.toFixed(3), h: +(maxW / r).toFixed(3) };
  return { w: +(maxH * r).toFixed(3), h: +maxH.toFixed(3) };
}

function base(sl, numbered) {
  if (numbered !== false) SN++;
  sl.background = { color: WHITE };
  sl.addImage({ path: IUL_L, x: 0.08, y: 0.06, w: 0.82, h: 0.82 });
  sl.addImage({ path: IUL_R, x: 9.10, y: 0.06, w: 0.82, h: 0.82 });
  sl.addImage({ path: WAVE,  x: 0,    y: 5.27,  w: 10.0, h: 2.25 });
  if (numbered !== false) {
    sl.addText(`${SN} / ${TOTAL}`, {
      x: 8.75, y: 7.08, w: 1.1, h: 0.25,
      fontFace: "Arial", fontSize: 10, color: WHITE,
      align: "right", margin: 0
    });
  }
}

function hdr(sl, title) {
  sl.addText(title, {
    x: 0.95, y: 0.02, w: 8.1, h: 0.88,
    fontFace: "Times New Roman", fontSize: 22, bold: true,
    color: BLACK, valign: "middle", align: "center", margin: 0
  });
  sl.addShape(pres.ShapeType.line, {
    x: 0.95, y: 0.93, w: 8.1, h: 0,
    line: { color: BLUE, width: 1.5 }
  });
}

function cap(sl, txt, x, y, w) {
  sl.addText(txt, {
    x, y, w: w || 9.4, h: 0.3,
    fontFace: "Arial", fontSize: 10.5, color: DGRAY,
    align: "center", italic: true, margin: 0
  });
}

// Filled colored box with text
function box(sl, txt, x, y, w, h, fill, txtColor, fs, bold) {
  sl.addText(txt, {
    x, y, w, h,
    fontFace: "Arial", fontSize: fs || 11, bold: bold !== false,
    color: txtColor || WHITE, align: "center", valign: "middle",
    fill: { color: fill }, line: { color: fill, width: 0 }, margin: 3,
  });
}

// Arrow — uses flipH/flipV to keep all dimensions positive (required by pptxgenjs)
function arrow(sl, x1, y1, x2, y2, color, lineW) {
  const dx = x2 - x1, dy = y2 - y1;
  const goLeft = dx < 0, goUp = dy < 0;
  sl.addShape(pres.ShapeType.line, {
    x: goLeft ? x2 : x1, y: goUp ? y2 : y1,
    w: Math.max(Math.abs(dx), 0.005), h: Math.max(Math.abs(dy), 0.005),
    flipH: goLeft, flipV: goUp,
    line: { color: color || DGRAY, width: lineW || 1.5, endArrowType: "arrow" }
  });
}

// Concise bullet block
function bullets(sl, items, x, y, w, h, fs) {
  const runs = items.map((it, i) => ({
    text: typeof it === "string" ? it : it.t,
    options: {
      bold:      typeof it !== "string" && !!it.b,
      bullet:    typeof it === "string" ? { code: "25CF" } : (it.bul !== false ? { code: "25CF" } : false),
      color:     typeof it !== "string" && it.c ? it.c : BLACK,
      fontSize:  typeof it !== "string" && it.fs ? it.fs : (fs || 14),
      breakLine: i < items.length - 1,
    }
  }));
  sl.addText(runs, {
    x, y, w, h, fontFace: "Arial", fontSize: fs || 14, color: BLACK, valign: "top", margin: 0
  });
}

// Card with colored header bar + bullet list
function card(sl, header, items, x, y, w, h, fill) {
  const fc = fill || BLUE;
  sl.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: LLBLUE }, line: { color: fc, width: 1.5 } });
  sl.addShape(pres.ShapeType.rect, { x, y, w, h: 0.38, fill: { color: fc }, line: { color: fc, width: 0 } });
  sl.addText(header, {
    x, y, w, h: 0.38, fontFace: "Arial", fontSize: 11, bold: true,
    color: WHITE, align: "center", valign: "middle", margin: 0
  });
  const runs = items.map((it, i) => ({
    text: it, options: { bullet: { code: "25CF" }, breakLine: i < items.length - 1, fontSize: 10.5, color: BLACK }
  }));
  sl.addText(runs, {
    x: x + 0.12, y: y + 0.42, w: w - 0.24, h: h - 0.48,
    fontFace: "Arial", fontSize: 10.5, color: BLACK, valign: "top", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 1 ─ TITLE
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl, false);

  // Larger corner logos for title slide
  sl.addImage({ path: IUL_L, x: 0.05, y: 0.05, w: 1.1, h: 1.1 });
  sl.addImage({ path: IUL_R, x: 8.85, y: 0.05, w: 1.1, h: 1.1 });

  sl.addText("Islamic University of Lebanon", {
    x: 1.2, y: 0.06, w: 7.6, h: 0.52,
    fontFace: "Times New Roman", fontSize: 17, bold: true,
    color: DBLUE, align: "center", valign: "middle", margin: 0
  });
  sl.addText("Faculty of Science and Arts", {
    x: 1.2, y: 0.56, w: 7.6, h: 0.36,
    fontFace: "Times New Roman", fontSize: 14,
    color: BLACK, align: "center", valign: "middle", margin: 0
  });

  sl.addShape(pres.ShapeType.line, {
    x: 1.0, y: 0.98, w: 8.0, h: 0, line: { color: BLUE, width: 1.2 }
  });

  sl.addText("Android & Web Application", {
    x: 1.5, y: 1.04, w: 7.0, h: 0.42,
    fontFace: "Arial", fontSize: 15, color: DGRAY, align: "center", valign: "middle", margin: 0
  });

  sl.addText("VeriSource", {
    x: 0.8, y: 1.44, w: 8.4, h: 1.05,
    fontFace: "Times New Roman", fontSize: 54, bold: true,
    color: BLUE, align: "center", valign: "middle", margin: 0
  });
  sl.addText("AI-Powered Fake News Detection Platform", {
    x: 0.8, y: 2.46, w: 8.4, h: 0.5,
    fontFace: "Arial", fontSize: 18, color: BLACK, align: "center", valign: "middle", margin: 0
  });

  sl.addImage({ path: REACT, x: 4.38, y: 3.06, w: 1.25, h: 1.25 });

  // Wave info — student on one line, supervisor on the next
  sl.addText("Presented by:  Hassan Zaarour", {
    x: 0.3, y: 5.42, w: 9.4, h: 0.38,
    fontFace: "Arial", fontSize: 13, bold: false, color: WHITE,
    align: "center", valign: "middle", margin: 0
  });
  sl.addText("Supervisor:  Dr. Bassell Dhaini", {
    x: 0.3, y: 5.8, w: 9.4, h: 0.36,
    fontFace: "Arial", fontSize: 13, bold: false, color: WHITE,
    align: "center", valign: "middle", margin: 0
  });
  sl.addText("Department of Computer and Communications Engineering  ·  2025", {
    x: 0.3, y: 6.17, w: 9.4, h: 0.3,
    fontFace: "Arial", fontSize: 11, color: "AACCEE", align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 2 ─ PROBLEM STATEMENT
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Problem Statement");

  const problems = [
    {
      num: "01",
      title: "Misinformation Spreads Rapidly",
      desc: "Fake news travels 6× faster than real news on social media, reaching millions before corrections are issued.",
      c: RED,
    },
    {
      num: "02",
      title: "Manual Verification is Impractical",
      desc: "Users lack the time, resources, and expertise to independently verify the credibility of every article they encounter.",
      c: ORNG,
    },
    {
      num: "03",
      title: "No Accessible Automated Solution",
      desc: "Existing fact-checking tools are slow, require expert review, or are not available for everyday users in real-time.",
      c: BLUE,
    },
  ];

  const pW = 9.1, pH = 1.2, pGap = 0.2, startX = 0.45, startY = 1.05;

  problems.forEach((p, i) => {
    const y = startY + i * (pH + pGap);
    // Background card
    sl.addShape(pres.ShapeType.rect, {
      x: startX, y, w: pW, h: pH,
      fill: { color: LLBLUE }, line: { color: p.c, width: 2 }
    });
    // Number badge (left strip)
    sl.addShape(pres.ShapeType.rect, {
      x: startX, y, w: 0.65, h: pH,
      fill: { color: p.c }, line: { color: p.c, width: 0 }
    });
    sl.addText(p.num, {
      x: startX, y, w: 0.65, h: pH,
      fontFace: "Arial", fontSize: 20, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    // Title
    sl.addText(p.title, {
      x: startX + 0.75, y: y + 0.08, w: pW - 0.85, h: 0.38,
      fontFace: "Arial", fontSize: 13.5, bold: true, color: p.c,
      valign: "middle", margin: 0
    });
    // Description
    sl.addText(p.desc, {
      x: startX + 0.75, y: y + 0.46, w: pW - 0.85, h: 0.64,
      fontFace: "Arial", fontSize: 12, color: BLACK, valign: "top", margin: 0
    });
  });

  // Solution teaser
  sl.addText("VeriSource solves this with instant AI-powered classification available to anyone.", {
    x: 0.45, y: 4.75, w: 9.1, h: 0.38,
    fontFace: "Arial", fontSize: 13, bold: true, color: BLUE,
    align: "center", valign: "middle", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 3 ─ PROJECT OVERVIEW
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Project Overview");

  bullets(sl, [
    { t: "What is VeriSource?", b: true, bul: false, fs: 15.5 },
    { t: "An AI platform that classifies news articles as Real or Fake", bul: true },
    { t: " ", bul: false },
    { t: "Three input methods:", b: true, bul: false },
    { t: "Plain text — paste or type any article", bul: true },
    { t: "URL — automatic web scraping", bul: true },
    { t: "Image — OCR text extraction", bul: true },
    { t: " ", bul: false },
    { t: "Available on:", b: true, bul: false },
    { t: "React 18 Web Application", bul: true },
    { t: "Expo React Native Mobile App", bul: true },
  ], 0.35, 1.0, 5.0, 4.1);

  const sz = fit(893, 817, 4.2, 3.8);
  const ix = 10 - sz.w - 0.2;
  const iy = 1.0 + (4.1 - sz.h) / 2;
  sl.addImage({ path: IMG.anlyz.p, x: ix, y: iy, w: sz.w, h: sz.h });
  cap(sl, "Article Analysis Screen", ix, iy + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 4 ─ PROJECT OBJECTIVES
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Project Objectives");

  const objs = [
    { t: "AI Classification",       d: "Real or Fake verdict with confidence score and keyword signals",   c: BLUE  },
    { t: "Multi-Format Input",      d: "Plain text, URL scraping, and image OCR in one platform",          c: TEAL  },
    { t: "Multilingual Support",    d: "Auto-detect & translate non-English articles before analysis",     c: GREEN },
    { t: "Dual-Platform",           d: "Full-featured web app and Android mobile app",                     c: DBLUE },
    { t: "Transparency",            d: "Confidence %, keyword signals, and source trust scores",           c: PURP  },
    { t: "Press Ecosystem",         d: "Publishers post news; users receive live notifications",           c: ORNG  },
  ];

  const bW = 2.95, bH = 1.7, gapX = 0.2, gapY = 0.15;
  const sx = (10 - 3*bW - 2*gapX) / 2;

  objs.forEach((o, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = sx + col*(bW+gapX), y = 1.08 + row*(bH+gapY);
    sl.addShape(pres.ShapeType.rect, { x, y, w: bW, h: bH, fill: { color: LLBLUE }, line: { color: o.c, width: 2 } });
    sl.addShape(pres.ShapeType.rect, { x, y, w: bW, h: 0.38, fill: { color: o.c }, line: { color: o.c, width: 0 } });
    sl.addText(o.t, {
      x, y, w: bW, h: 0.38, fontFace: "Arial", fontSize: 11.5, bold: true,
      color: WHITE, align: "center", valign: "middle", margin: 0
    });
    sl.addText(o.d, {
      x: x+0.12, y: y+0.44, w: bW-0.24, h: bH-0.5,
      fontFace: "Arial", fontSize: 12, color: BLACK, valign: "top", margin: 0
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 5 ─ USER ROLES
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "User Roles");

  const roles = [
    { label: "Regular User", fill: TEAL, items: [
      "Register and log in securely (JWT)",
      "Submit articles for AI analysis",
      "View verdict, confidence & keywords",
      "Browse history, analytics & source map",
      "Read press news and use AI Chat",
    ]},
    { label: "Press User  (extends Regular)", fill: BLUE, items: [
      "All Regular User capabilities",
      "Signup triggers admin approval email",
      "Publish, edit and delete press posts",
      "Auto-notify all users on new post",
      "Personal author channel name",
    ]},
    { label: "Admin  (Planned)", fill: DGRAY, items: [
      "Manage all user accounts and roles",
      "Approve or revoke press status",
      "Platform-wide analytics and oversight",
    ]},
  ];

  const cW = 2.95, cH = 3.85, gap = 0.2;
  const sx = (10 - 3*cW - 2*gap) / 2;

  roles.forEach((r, i) => {
    const x = sx + i*(cW+gap);
    card(sl, r.label, r.items, x, 1.05, cW, cH, r.fill);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 6 ─ SYSTEM ARCHITECTURE
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "System Architecture");

  // ── Layout: top-to-bottom hierarchy ──────────────────────────────────────
  // Row 1: Users (centered)
  // Row 2: React Web  |  React Native  (two boxes)
  // Row 3: FastAPI Backend (center, wider)
  // Row 4: PostgreSQL | RoBERTa | OCR | Gemini | NewsAPI

  const bH = 0.44;

  // Row 1 – Users
  const r1y = 1.0, uW = 3.2;
  box(sl, "Users (Web & Mobile)", (10-uW)/2, r1y, uW, bH, GREEN, WHITE, 11);

  // Row 2 – Frontend (y=1.65)
  const r2y = 1.62;
  box(sl, "React 18  (Web App)",        1.0,  r2y, 3.5, bH, TEAL, WHITE, 10.5);
  box(sl, "React Native  (Mobile App)", 5.5,  r2y, 3.5, bH, TEAL, WHITE, 10.5);

  // Row 3 – Backend (y=2.3, taller)
  const r3y = 2.28, beW = 5.2, beH = 0.88;
  sl.addShape(pres.ShapeType.rect, {
    x: (10-beW)/2, y: r3y, w: beW, h: beH,
    fill: { color: DBLUE }, line: { color: LBLUE, width: 2 }
  });
  sl.addText([
    { text: "FastAPI Backend", options: { bold: true, fontSize: 13.5, color: WHITE } },
    { text: "\nREST API  ·  JWT Auth  ·  SQLAlchemy ORM", options: { fontSize: 9.5, color: LBLUE } },
  ], {
    x: (10-beW)/2, y: r3y, w: beW, h: beH,
    fontFace: "Arial", align: "center", valign: "middle", margin: 4
  });

  // Row 4 – Services (y=3.35)
  const r4y = 3.33;
  const svcs = [
    { t: "PostgreSQL",  c: NAVY },
    { t: "RoBERTa AI",  c: PURP },
    { t: "OCR Module",  c: ORNG },
    { t: "Gemini API",  c: TEAL },
    { t: "NewsAPI",     c: GREEN },
  ];
  const sW = 1.7, sGap = 0.175;
  const sStartX = (10 - 5*sW - 4*sGap) / 2;
  svcs.forEach((s, i) => {
    box(sl, s.t, sStartX + i*(sW+sGap), r4y, sW, bH, s.c, WHITE, 10);
  });

  // Layer labels (right-side annotations)
  const lblX = 9.35;
  [
    { t: "Clients",   y: r1y + bH/2 - 0.08 },
    { t: "Frontend",  y: r2y + bH/2 - 0.08 },
    { t: "Backend",   y: r3y + beH/2 - 0.08 },
    { t: "Services",  y: r4y + bH/2 - 0.08 },
  ].forEach(l => {
    sl.addText(l.t, {
      x: lblX, y: l.y, w: 0.6, h: 0.28,
      fontFace: "Arial", fontSize: 8.5, bold: true, color: DGRAY,
      align: "left", valign: "middle", margin: 0
    });
  });

  // ── Arrows ────────────────────────────────────────────────────────────────
  // Users → React Web
  arrow(sl, 10/2 - 0.4, r1y + bH, 1.0 + 3.5/2, r2y, "4A9A6B", 1.3);
  // Users → React Native
  arrow(sl, 10/2 + 0.4, r1y + bH, 5.5 + 3.5/2, r2y, "4A9A6B", 1.3);
  // React Web → Backend
  arrow(sl, 1.0 + 3.5/2, r2y + bH, (10-beW)/2 + beW*0.3, r3y, "4A9CD2", 1.4);
  // React Native → Backend
  arrow(sl, 5.5 + 3.5/2, r2y + bH, (10-beW)/2 + beW*0.7, r3y, "4A9CD2", 1.4);
  // Backend → each service (from bottom center of backend)
  const beCx = 5.0, beBot = r3y + beH;
  svcs.forEach((s, i) => {
    const svcCx = sStartX + i*(sW+sGap) + sW/2;
    arrow(sl, beCx, beBot, svcCx, r4y, "9B7FC2", 1.3);
  });

  // REST label on left connection
  sl.addText("REST / JSON", {
    x: 0.25, y: r2y + bH + 0.05, w: 1.6, h: 0.22,
    fontFace: "Arial", fontSize: 8, color: DGRAY, italic: true,
    align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 7 ─ USE CASE DIAGRAM
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "UML Use Case Diagram");

  // Brief bullets on left
  bullets(sl, [
    { t: "2 actors:", b: true, bul: false, fs: 13 },
    "Regular User — submit, view, chat",
    "Press User — extends Regular + publish",
    { t: " ", bul: false },
    { t: "Core use cases:", b: true, bul: false, fs: 13 },
    "Submit article (text / URL / image)",
    "View verdict, history & analytics",
    "Publish & manage press posts",
  ], 0.25, 1.05, 2.9, 4.0, 12.5);

  const sz = fit(1385, 659, 6.5, 3.9);
  const ix = 10 - sz.w - 0.15;
  const iy = 1.0 + (4.0 - sz.h) / 2;
  sl.addImage({ path: IMG.usecase.p, x: ix, y: iy, w: sz.w, h: sz.h });
  cap(sl, "Figure 1 – VeriSource Use Case Diagram", ix, iy + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 8 ─ DFD LEVEL 0
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Data Flow Diagram – Level 0");

  bullets(sl, [
    { t: "VeriSource as a black box", b: true, bul: false, fs: 13 },
    "Regular User: submit & receive verdict",
    "Press User: publish & manage posts",
    { t: " ", bul: false },
    { t: "External services:", b: true, bul: false, fs: 13 },
    "Google Gemini — AI Chat",
    "NewsAPI — bilingual news feed",
  ], 0.25, 1.05, 2.9, 4.0, 12.5);

  const sz = fit(1476, 802, 6.5, 3.9);
  const ix = 10 - sz.w - 0.15;
  const iy = 1.0 + (4.0 - sz.h) / 2;
  sl.addImage({ path: IMG.dfd0.p, x: ix, y: iy, w: sz.w, h: sz.h });
  cap(sl, "Figure 2 – DFD Level 0 (Context Diagram)", ix, iy + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 9 ─ DFD LEVEL 1
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Data Flow Diagram – Level 1");

  bullets(sl, [
    { t: "6 internal processes:", b: true, bul: false, fs: 13 },
    "Authentication & Access Control",
    "Article Classification (RoBERTa)",
    "Press Publication & Notifications",
    "Analytics & Reporting",
    "AI Chat (Gemini)",
    "News Aggregation (NewsAPI)",
  ], 0.25, 1.05, 2.9, 4.0, 12.5);

  const sz = fit(1463, 988, 6.2, 3.9);
  const ix = 10 - sz.w - 0.15;
  const iy = 1.0 + (4.0 - sz.h) / 2;
  sl.addImage({ path: IMG.dfd1.p, x: ix, y: iy, w: sz.w, h: sz.h });
  cap(sl, "Figure 3 – DFD Level 1 (Internal Processes)", ix, iy + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 10 ─ SEQUENCE DIAGRAMS
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Sequence Diagrams");

  const h = 3.35;
  const w1 = +(h * 1379/920).toFixed(3);
  const w2 = +(h * 1385/923).toFixed(3);
  const gap = 0.25;
  const totalW = Math.min(w1, 4.65) + gap + Math.min(w2, 4.65);
  const sx = (10 - totalW) / 2;
  const sw1 = Math.min(w1, 4.65), sw2 = Math.min(w2, 4.65);
  const imgY = 1.5;

  cap(sl, "User: Article Submission Flow", sx, 1.05, sw1);
  sl.addImage({ path: IMG.seq_u.p, x: sx, y: imgY, w: sw1, h });

  cap(sl, "Press User: Publication Flow", sx+sw1+gap, 1.05, sw2);
  sl.addImage({ path: IMG.seq_p.p, x: sx+sw1+gap, y: imgY, w: sw2, h });

  sl.addShape(pres.ShapeType.line, {
    x: sx+sw1+gap/2, y: 1.05, w: 0, h: 4.1,
    line: { color: "DDDDDD", width: 0.8 }
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 11 ─ ERD
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Entity Relationship Diagram (ERD)");

  bullets(sl, [
    { t: "5 tables:", b: true, bul: false, fs: 13 },
    "users",
    "articles",
    "sources",
    "press_posts",
    "notifications",
    { t: " ", bul: false },
    { t: "All FK constraints enforced", b: true, bul: false, fs: 13 },
  ], 0.25, 1.05, 2.9, 4.0, 12.5);

  const sz = fit(1536, 1024, 6.3, 3.9);
  const ix = 10 - sz.w - 0.15;
  const iy = 1.0 + (4.0 - sz.h) / 2;
  sl.addImage({ path: IMG.erd.p, x: ix, y: iy, w: sz.w, h: sz.h });
  cap(sl, "Figure 6 – Entity Relationship Diagram (ERD)", ix, iy + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 12 ─ TECHNOLOGY STACK  (5 grouped cards in a row)
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Technology Stack");

  const groups = [
    { h: "Frontend",          c: TEAL,  items: ["React 18 + Vite", "Expo SDK 56 (RN)", "react-leaflet (Maps)", "Recharts (Charts)", "Framer Motion"] },
    { h: "Backend",           c: DBLUE, items: ["Python FastAPI", "SQLAlchemy ORM", "Alembic Migrations", "PyJWT (RS256)", "httpx (async)"] },
    { h: "Database",          c: NAVY,  items: ["PostgreSQL 15", "5 normalized tables", "Docker container", "Alembic schema mgmt"] },
    { h: "AI / ML",           c: PURP,  items: ["HuggingFace Transformers", "RoBERTa classifier", "Tesseract.js (OCR)", "deep-translator"] },
    { h: "External APIs",     c: ORNG,  items: ["Google Gemini (Chat)", "NewsAPI (EN / AR)", "CARTO tile layers"] },
  ];

  const cW = 1.72, cH = 3.6, gapX = 0.15;
  const sx = (10 - 5*cW - 4*gapX) / 2;

  groups.forEach((g, i) => {
    card(sl, g.h, g.items, sx + i*(cW+gapX), 1.05, cW, cH, g.c);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 13 ─ AI CLASSIFICATION PIPELINE
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "AI Classification Pipeline");

  const steps = [
    { n: "1", t: "Input",        d: "Text / URL\nor Image",        c: TEAL  },
    { n: "2", t: "Translate",    d: "Auto-detect &\ntranslate",    c: BLUE  },
    { n: "3", t: "Truncate",     d: "≤ 2,000 chars\n(512 tokens)", c: DBLUE },
    { n: "4", t: "RoBERTa",      d: "REAL / FAKE\n+ confidence",   c: PURP  },
    { n: "5", t: "Store",        d: "Save to DB\nupdate trust",    c: NAVY  },
    { n: "6", t: "Result",       d: "Verdict &\nkeywords shown",   c: GREEN },
  ];

  const bW = 1.42, bH = 1.55, sx = 0.27;

  steps.forEach((s, i) => {
    const x = sx + i * (bW + 0.14);
    // Number badge
    box(sl, s.n, x + bW/2 - 0.2, 1.5, 0.4, 0.4, s.c, WHITE, 14, true);
    // Step box
    sl.addShape(pres.ShapeType.rect, { x, y: 2.0, w: bW, h: bH, fill: { color: LLBLUE }, line: { color: s.c, width: 2 } });
    sl.addText(s.t, { x, y: 2.0, w: bW, h: 0.4, fontFace: "Arial", fontSize: 12, bold: true, color: s.c, align: "center", valign: "middle", margin: 0 });
    sl.addText(s.d, { x, y: 2.42, w: bW, h: bH-0.44, fontFace: "Arial", fontSize: 10.5, color: BLACK, align: "center", valign: "top", margin: 2 });
    // Arrow
    if (i < steps.length - 1) {
      arrow(sl, x + bW + 0.01, 2.0 + bH/2, x + bW + 0.135, 2.0 + bH/2, DGRAY, 1.5);
    }
  });

  sl.addText("Model loaded once at startup — no repeated 500 MB disk I/O per request.", {
    x: 0.3, y: 3.75, w: 9.4, h: 0.36,
    fontFace: "Arial", fontSize: 11, color: DGRAY, italic: true, align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 14 ─ WELCOME & AUTHENTICATION
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Welcome & Authentication");

  bullets(sl, [
    { t: "JWT tokens (RS256 keys)", bul: true },
    { t: "Email + password registration", bul: true },
    { t: "Persistent login (localStorage)", bul: true },
    { t: "Role-based UI (press / regular)", bul: true },
    { t: "Web and mobile", bul: true },
  ], 0.3, 1.05, 2.65, 3.0, 13);

  const imgH = 3.2;
  const logSz = fit(898,  823, 2.45, imgH);
  const welSz = fit(1385, 659, 99,  imgH);
  const welW  = Math.min(welSz.w, 4.9);
  const logX = 2.95 + 0.1;
  const welX = logX + logSz.w + 0.2;
  const imgY = 1.05 + (4.0 - imgH) / 2;

  cap(sl, "Login / Register",  logX, 1.02, logSz.w);
  sl.addImage({ path: IMG.login.p,   x: logX, y: imgY, w: logSz.w, h: imgH });
  cap(sl, "Welcome Page (Web)", welX, 1.02, welW);
  sl.addImage({ path: IMG.welcome.p, x: welX, y: imgY, w: welW,     h: imgH });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 15 ─ ARTICLE ANALYSIS
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Article Analysis Screen");

  bullets(sl, [
    { t: "Submit modes:", b: true, bul: false, fs: 14 },
    "Plain text",
    "URL (auto-scrape)",
    "Image (OCR)",
    { t: " ", bul: false },
    { t: "Results:", b: true, bul: false, fs: 14 },
    "REAL / FAKE verdict",
    "Confidence percentage",
    "Keyword signals",
  ], 0.3, 1.05, 2.8, 4.0, 13.5);

  const sz = fit(893, 817, 6.5, 3.9);
  const ix = 10 - sz.w - 0.15;
  const iy = 1.0 + (4.0 - sz.h) / 2;
  sl.addImage({ path: IMG.anlyz.p, x: ix, y: iy, w: sz.w, h: sz.h });
  cap(sl, "Figure 9 – Article Analysis Screen", ix, iy + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 16 ─ DASHBOARD & ANALYTICS
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Dashboard & Analytics");

  bullets(sl, [
    "Real vs. Fake counts at a glance",
    "Submission timeline (bar chart)",
    "Search & filter history",
    "Global source map — geocoded domains",
    "Color-coded trust score markers",
  ], 0.3, 1.05, 2.8, 3.0, 13.5);

  const sz = fit(1918, 911, 6.5, 3.0);
  const ix = 10 - sz.w - 0.15;
  sl.addImage({ path: IMG.dash.p, x: ix, y: 1.08, w: sz.w, h: sz.h });
  cap(sl, "Figure 10 – Analytics Dashboard", ix, 1.08 + sz.h + 0.04, sz.w);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 17 ─ MOBILE APPLICATION
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Mobile Application  (Expo SDK 56)");

  const imgH = 3.5;
  const pw = +(imgH * 1022/906).toFixed(3);   // mob_en (landscape-ish)
  const mhW = +(imgH * 1286/624).toFixed(3);   // mob_home (wide)

  // Show 2 images side by side with captions
  const gap = 0.3;
  const w1 = Math.min(pw, 3.9), w2 = Math.min(mhW, 4.5);
  const total = w1 + gap + w2;
  const sx = (10 - total) / 2;
  const imgY = 1.45;

  cap(sl, "News Headlines (English)", sx,       1.07, w1);
  sl.addImage({ path: IMG.mob_en.p,   x: sx,       y: imgY, w: w1, h: imgH });

  cap(sl, "Mobile Home Screen",        sx+w1+gap, 1.07, w2);
  sl.addImage({ path: IMG.mob_home.p,  x: sx+w1+gap, y: imgY, w: w2, h: imgH });

  // Features strip at bottom
  sl.addText("Bilingual (EN / AR)  ·  Article classification  ·  AI Chat  ·  Press feed  ·  Notifications", {
    x: 0.3, y: 5.05, w: 9.4, h: 0.25,
    fontFace: "Arial", fontSize: 11, color: DGRAY, italic: true, align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 18 ─ PRESS FEATURES
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Press User Features");

  const imgH = 3.4;
  const prW  = +(imgH * 738/1600).toFixed(3);   // ~1.57" portrait
  const pr2x = 10 - prW - 0.1;
  const midX = 0.15 + prW + 0.2;
  const midW = pr2x - midX - 0.2;

  cap(sl, "Publish Screen",  0.15, 1.02, prW);
  sl.addImage({ path: IMG.press1.p, x: 0.15, y: 1.4, w: prW, h: imgH });

  bullets(sl, [
    { t: "Account approval:", b: true, bul: false, fs: 14 },
    "Admin receives email on press signup",
    "Approve or Reject via email or Admin Panel",
    { t: " ", bul: false },
    { t: "Publishing:", b: true, bul: false, fs: 14 },
    "Title, body & cover image upload",
    "Edit or delete own posts",
    "Auto-notify all users on publish",
    { t: " ", bul: false },
    { t: "Press News Page — EN / AR filter", b: false, bul: false, fs: 13 },
  ], midX, 1.05, midW, 4.0, 13);

  cap(sl, "Press News Page", pr2x, 1.02, prW);
  sl.addImage({ path: IMG.press2.p, x: pr2x, y: 1.4, w: prW, h: imgH });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 19 ─ AI CHAT ASSISTANT
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "AI Chat Assistant  (Google Gemini)");

  const imgH = 3.3;
  const chW  = +(imgH * 738/1600).toFixed(3);   // portrait ~1.52"
  const ch2  = fit(1331, 811, 4.4, imgH);
  const ch2x = 10 - ch2.w - 0.1;
  const midX = 0.15 + chW + 0.2;
  const midW = ch2x - midX - 0.2;

  cap(sl, "Web Chat",    0.15, 1.02, chW);
  sl.addImage({ path: IMG.chat_w.p, x: 0.15, y: 1.4, w: chW, h: imgH });

  bullets(sl, [
    { t: "Powered by Gemini Flash Lite", b: true, bul: false, fs: 14 },
    { t: " ", bul: false },
    "Context-aware Q&A on news credibility",
    "Persistent conversation per session",
    "Non-blocking async (httpx)",
    { t: " ", bul: false },
    { t: "Web:", b: true, bul: false, fs: 13 },
    { t: "Floating expandable widget", bul: true, fs: 12 },
    { t: "Mobile:", b: true, bul: false, fs: 13 },
    { t: "Slide-up modal interface", bul: true, fs: 12 },
  ], midX, 1.05, midW, 4.0, 13);

  cap(sl, "Mobile Chat", ch2x, 1.02, ch2.w);
  sl.addImage({ path: IMG.chat_m.p, x: ch2x, y: 1.4, w: ch2.w, h: ch2.h });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 20 ─ DEMO
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Live System Demonstration");

  const vx = 1.4, vy = 1.1, vw = 7.2, vh = 3.5;
  sl.addShape(pres.ShapeType.rect, {
    x: vx, y: vy, w: vw, h: vh,
    fill: { color: "1A1A2E" }, line: { color: BLUE, width: 2 }
  });
  sl.addShape(pres.ShapeType.rect, {
    x: vx + vw/2 - 0.55, y: vy + vh/2 - 0.55, w: 1.1, h: 1.1,
    fill: { color: BLUE }, line: { color: "AACCEE", width: 2 }
  });
  sl.addText("▶", {
    x: vx + vw/2 - 0.55, y: vy + vh/2 - 0.55, w: 1.1, h: 1.1,
    fontFace: "Arial", fontSize: 32, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0
  });
  sl.addText("[ Recorded Demo ]", {
    x: vx, y: vy + vh - 0.55, w: vw, h: 0.5,
    fontFace: "Arial", fontSize: 11, color: "AACCEE", align: "center", margin: 0
  });

  sl.addText("Submit article  ·  View verdict  ·  Analytics dashboard  ·  Mobile app  ·  AI Chat", {
    x: 0.3, y: 4.75, w: 9.4, h: 0.36,
    fontFace: "Arial", fontSize: 12, color: DGRAY, italic: true, align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 21 ─ CONCLUSION
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Conclusion");

  const pts = [
    { t: "Instant AI verdict",         d: "RoBERTa classifies any article in seconds with confidence score & keywords",      c: BLUE  },
    { t: "Three input modes",          d: "Plain text, URL scraping, and image OCR — all in one platform",                   c: TEAL  },
    { t: "Dual-platform",              d: "Full-featured React 18 web app + Expo React Native Android app",                  c: GREEN },
    { t: "Press ecosystem",            d: "Publishers post news; readers receive live notifications",                         c: PURP  },
    { t: "AI Chat assistant",          d: "Google Gemini-powered context-aware chat on any slide",                           c: ORNG  },
  ];

  pts.forEach((p, i) => {
    const y = 1.08 + i * 0.8;
    sl.addShape(pres.ShapeType.rect, { x: 0.35, y, w: 0.06, h: 0.62, fill: { color: p.c }, line: { color: p.c, width: 0 } });
    sl.addText(p.t, { x: 0.55, y, w: 2.8, h: 0.3, fontFace: "Arial", fontSize: 14, bold: true, color: p.c, valign: "bottom", margin: 0 });
    sl.addText(p.d, { x: 0.55, y: y+0.3, w: 9.0, h: 0.32, fontFace: "Arial", fontSize: 12.5, color: BLACK, valign: "top", margin: 0 });
  });

  sl.addText("VeriSource makes credibility assessment instant and accessible for everyone.", {
    x: 0.35, y: 5.05, w: 9.3, h: 0.3,
    fontFace: "Arial", fontSize: 13, bold: true, color: BLUE, align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 22 ─ FUTURE WORK
// ═══════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  base(sl);
  hdr(sl, "Future Work");

  const items = [
    { t: "Multilingual AI Model",    d: "Fine-tune XLM-RoBERTa for direct Arabic classification (no translation step)",   c: BLUE  },
    { t: "Admin Panel",              d: "User management, press approvals, and platform-wide moderation dashboard",        c: TEAL  },
    { t: "Browser Extension",        d: "Real-time article verification while browsing any news website",                  c: GREEN },
    { t: "Fact-Check Integration",   d: "Connect to Snopes, FactCheck.org and Google Fact Check Tools API",               c: PURP  },
  ];

  items.forEach((it, i) => {
    const y = 1.1 + i * 0.96;
    sl.addShape(pres.ShapeType.rect, { x: 0.35, y, w: 0.06, h: 0.76, fill: { color: it.c }, line: { color: it.c, width: 0 } });
    sl.addText(it.t, { x: 0.55, y, w: 3.2, h: 0.36, fontFace: "Arial", fontSize: 14.5, bold: true, color: it.c, valign: "bottom", margin: 0 });
    sl.addText(it.d, { x: 0.55, y: y+0.36, w: 9.1, h: 0.4, fontFace: "Arial", fontSize: 12.5, color: BLACK, valign: "top", margin: 0 });
  });

  sl.addText("Thank You  ·  Questions?", {
    x: 0.3, y: 5.45, w: 9.4, h: 0.45,
    fontFace: "Times New Roman", fontSize: 22, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0
  });
  sl.addText("Hassan Zaarour  ·  Supervisor: Dr. Bassell Dhaini  ·  Islamic University of Lebanon", {
    x: 0.3, y: 5.9, w: 9.4, h: 0.3,
    fontFace: "Arial", fontSize: 11, color: WHITE, align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// WRITE
// ═══════════════════════════════════════════════════════════════════════════
const OUT = "C:/Users/Admin/OneDrive/Desktop/fakenewsdetect/VeriSource_Presentation.pptx";
pres.writeFile({ fileName: OUT })
  .then(() => console.log("Done:", OUT, "— Total slides:", SN + 1))
  .catch(e => { console.error("Error:", e.message); process.exit(1); });
