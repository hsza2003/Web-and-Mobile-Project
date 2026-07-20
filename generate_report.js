const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat, TableOfContents
} = require("docx");
const fs = require("fs");
const path = require("path");

// ── Constants ─────────────────────────────────────────────────────────────────
const CONTENT_W = 9360;
const MEDIA = path.join(__dirname, "..", "report_unpacked", "word", "media");
const BORDER  = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER  = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

// ── Text helpers ──────────────────────────────────────────────────────────────
const h1 = t => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 200 },
  children: [new TextRun({ text: t, bold: true, size: 32, font: "Times New Roman" })]
});
const h2 = t => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 160 },
  children: [new TextRun({ text: t, bold: true, size: 28, font: "Times New Roman" })]
});
const h3 = t => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 120 },
  children: [new TextRun({ text: t, bold: true, size: 24, font: "Times New Roman" })]
});
const para = (t, opts = {}) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { before: 80, after: 120, line: 360, lineRule: "auto" },
  children: [new TextRun({ text: t, font: "Times New Roman", size: 24, ...opts })]
});
const bullet = t => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { before: 60, after: 60 },
  children: [new TextRun({ text: t, font: "Times New Roman", size: 24 })]
});
const centered = (t, opts = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 80, after: 80 },
  children: [new TextRun({ text: t, font: "Times New Roman", size: 24, ...opts })]
});
const spacer = () => new Paragraph({ spacing: { before: 120, after: 120 } });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ── Image helpers ─────────────────────────────────────────────────────────────
function imgRun(filename, w, h) {
  const type = filename.endsWith(".jpeg") || filename.endsWith(".jpg") ? "jpeg" : "png";
  const data = fs.readFileSync(path.join(MEDIA, filename));
  return new ImageRun({ type, data, transformation: { width: w, height: h } });
}

function figureBlock(figNum, caption, filename, w, h) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 80 },
      children: [imgRun(filename, w, h)]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 240 },
      children: [new TextRun({ text: `Figure ${figNum}: ${caption}`, font: "Times New Roman", size: 22, bold: true })]
    })
  ];
}

function placeholderFigure(figNum, caption) {
  return [
    spacer(),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [new TableRow({ children: [new TableCell({
        borders: {
          top:    { style: BorderStyle.DASHED, size: 4, color: "888888" },
          bottom: { style: BorderStyle.DASHED, size: 4, color: "888888" },
          left:   { style: BorderStyle.DASHED, size: 4, color: "888888" },
          right:  { style: BorderStyle.DASHED, size: 4, color: "888888" },
        },
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
        margins: { top: 400, bottom: 400, left: 400, right: 400 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [
            new TextRun({ text: "[ INSERT DIAGRAM - " + caption + " ]", font: "Times New Roman", size: 22, color: "666666", italics: true })
          ]})
        ]
      })]})]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 240 },
      children: [new TextRun({ text: `Figure ${figNum}: ${caption}`, font: "Times New Roman", size: 22, bold: true })]
    })
  ];
}

// ── Detailed Use-Case Table (FitFix style) ────────────────────────────────────
// columnWidths: [2800, 1880, 4680] — rows span differently
function detailedUseCaseTable({ name, id, priority, actor, description, precondition, events, postconditions }) {
  const COL = [2800, 1880, 4680]; // sum = CONTENT_W = 9360

  const mkCell = (text, width, span, fill, bold, color) =>
    new TableCell({
      borders: BORDERS,
      width: { size: width, type: WidthType.DXA },
      ...(span > 1 ? { columnSpan: span } : {}),
      shading: { fill: fill || "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({
        text, bold: !!bold, font: "Times New Roman", size: 22,
        color: color || "000000"
      })] })]
    });

  const mkMultiCell = (lines, width, span, fill) =>
    new TableCell({
      borders: BORDERS,
      width: { size: width, type: WidthType.DXA },
      ...(span > 1 ? { columnSpan: span } : {}),
      shading: { fill: fill || "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: lines.map((line, i) => new Paragraph({
        spacing: { before: i === 0 ? 0 : 40 },
        children: [new TextRun({ text: line, font: "Times New Roman", size: 22 })]
      }))
    });

  const labelRow = (label, content) => new TableRow({ children: [
    mkCell(label, 2800, 1, "EAF2FB", true),
    mkCell(content, 6560, 2, "FFFFFF", false)
  ]});

  const labelMultiRow = (label, lines) => new TableRow({ children: [
    mkCell(label, 2800, 1, "EAF2FB", true),
    mkMultiCell(lines, 6560, 2, "FFFFFF")
  ]});

  const coeHeaderRow = new TableRow({ children: [
    mkCell("Typical Course of Events", 9360, 3, "2E5FA3", true, "FFFFFF")
  ]});

  const coeSubHeaderRow = new TableRow({ children: [
    mkCell("Actor Action", 4680, 2, "D6E4F5", true),
    mkCell("System Response", 4680, 1, "D6E4F5", true)
  ]});

  const eventRows = events.map(([actor, sys], i) => new TableRow({ children: [
    new TableCell({
      borders: BORDERS, width: { size: 4680, type: WidthType.DXA }, columnSpan: 2,
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F7F9FC", type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: actor, font: "Times New Roman", size: 22 })] })]
    }),
    new TableCell({
      borders: BORDERS, width: { size: 4680, type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F7F9FC", type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: sys, font: "Times New Roman", size: 22 })] })]
    })
  ]}));

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: COL,
    rows: [
      labelRow("Use-Case Name", name),
      labelRow("Use-Case ID", id),
      labelRow("Priority", priority),
      labelRow("Primary Actor", actor),
      labelRow("Description", description),
      labelMultiRow("Precondition", precondition),
      coeHeaderRow,
      coeSubHeaderRow,
      ...eventRows,
      labelMultiRow("Postconditions", postconditions)
    ]
  });
}

// ── Flow/Testing table ────────────────────────────────────────────────────────
function flowTable(caption, tableNum, rows) {
  const col1 = 2400, col2 = 2000, col3 = CONTENT_W - col1 - col2;
  const headerRow = new TableRow({ children: [
    new TableCell({ borders: BORDERS, width: { size: col1, type: WidthType.DXA },
      shading: { fill: "2E5FA3", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: "Component", bold: true, color: "FFFFFF", font: "Times New Roman", size: 22 })] })] }),
    new TableCell({ borders: BORDERS, width: { size: col2, type: WidthType.DXA },
      shading: { fill: "2E5FA3", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: "Type", bold: true, color: "FFFFFF", font: "Times New Roman", size: 22 })] })] }),
    new TableCell({ borders: BORDERS, width: { size: col3, type: WidthType.DXA },
      shading: { fill: "2E5FA3", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: "Description", bold: true, color: "FFFFFF", font: "Times New Roman", size: 22 })] })] })
  ]});
  const dataRows = rows.map(([comp, type, desc], i) => new TableRow({ children: [
    new TableCell({ borders: BORDERS, width: { size: col1, type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F7F9FC", type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: comp, bold: true, font: "Times New Roman", size: 22 })] })] }),
    new TableCell({ borders: BORDERS, width: { size: col2, type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F7F9FC", type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: type, font: "Times New Roman", size: 22 })] })] }),
    new TableCell({ borders: BORDERS, width: { size: col3, type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F7F9FC", type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: desc, font: "Times New Roman", size: 22 })] })] })
  ]}));
  return [
    new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [col1, col2, col3], rows: [headerRow, ...dataRows] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 240 },
      children: [new TextRun({ text: `Table ${tableNum}: ${caption}`, font: "Times New Roman", size: 22, bold: true })] })
  ];
}

// ── ER entity table ───────────────────────────────────────────────────────────
function erEntityTable(entityName, pkField, attributes, domain) {
  const col1 = 3000, col2 = CONTENT_W - 3000;
  const headerRow = new TableRow({ children: [
    new TableCell({ columnSpan: 2, borders: BORDERS, width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { fill: "1A4F8A", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [new TextRun({ text: `${entityName}  (${domain})`, bold: true, color: "FFFFFF", font: "Times New Roman", size: 24 })] })] })
  ]});
  const colRow = new TableRow({ children: [
    new TableCell({ borders: BORDERS, width: { size: col1, type: WidthType.DXA },
      shading: { fill: "D6E4F5", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: "Attribute", bold: true, font: "Times New Roman", size: 22 })] })] }),
    new TableCell({ borders: BORDERS, width: { size: col2, type: WidthType.DXA },
      shading: { fill: "D6E4F5", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: "Description / Type", bold: true, font: "Times New Roman", size: 22 })] })] })
  ]});
  const attrRows = attributes.map(([attr, desc], i) => new TableRow({ children: [
    new TableCell({ borders: BORDERS, width: { size: col1, type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F2F7FC", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: attr, font: "Courier New", size: 20, bold: attr === pkField })] })] }),
    new TableCell({ borders: BORDERS, width: { size: col2, type: WidthType.DXA },
      shading: { fill: i % 2 === 0 ? "FFFFFF" : "F2F7FC", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: desc, font: "Times New Roman", size: 22 })] })] })
  ]}));
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [col1, col2],
    rows: [headerRow, colRow, ...attrRows] });
}

// ── Use-Case Data ─────────────────────────────────────────────────────────────
const UC_LOGIN = {
  name: "User Login",
  id: "UC-001",
  priority: "High",
  actor: "Regular User, Press User",
  description: "Authenticated users access the VeriSource platform by providing their registered email and password. The system validates credentials, issues a JWT token, and grants role-appropriate access.",
  precondition: [
    "1. User must have a registered account in the system.",
    "2. User must have a valid email address and password.",
    "3. The backend API server must be running and accessible.",
    "4. The PostgreSQL database must be operational."
  ],
  events: [
    ["1. User navigates to the login/register page.", "System displays the login form with email and password fields."],
    ["2. User enters their registered email address.", "System validates the email format in real time."],
    ["3. User enters their password.", "System masks the password field for security."],
    ["4. User clicks the Login button.", "System sends a POST /auth/login request to the backend."],
    ["", "Backend retrieves the user record by email from the database."],
    ["", "Backend verifies the submitted password against the stored bcrypt hash."],
    ["", "Backend generates a JWT access token with a 7-day expiry."],
    ["", "Backend returns the JWT token and user profile (email, is_press, channel_name)."],
    ["", "System stores the JWT in browser local storage (web) or secure storage (mobile)."],
    ["", "System redirects the user to the platform home page based on their role."]
  ],
  postconditions: [
    "1. User is authenticated and logged into the system.",
    "2. JWT token is stored in client-side storage for subsequent requests.",
    "3. User is redirected to the role-appropriate home page.",
    "4. All protected API routes are accessible with the issued token.",
    "5. System logs the login event."
  ]
};

const UC_REGISTER = {
  name: "User Registration",
  id: "UC-002",
  priority: "High",
  actor: "New User (Prospective Regular or Press User)",
  description: "New users create an account on VeriSource by providing an email, password, and optionally enabling a press account with a channel name. The system validates inputs, hashes the password, stores the user record, and immediately authenticates the new user.",
  precondition: [
    "1. User must have a valid email address not already registered.",
    "2. Backend API server must be running.",
    "3. PostgreSQL database must be operational."
  ],
  events: [
    ["1. User navigates to the register tab on the Auth page.", "System displays the registration form."],
    ["2. User enters email and password.", "System validates email format and ensures password is not empty."],
    ["3. User optionally toggles the Press Account switch.", "System reveals the Channel Name input field when toggle is enabled."],
    ["4. User optionally enters a channel name.", "System records the channel name for press identification."],
    ["5. User clicks the Register button.", "System sends a POST /auth/register request with credentials and flags."],
    ["", "Backend checks if the email already exists in the database."],
    ["", "Backend hashes the password using bcrypt with an adaptive cost factor."],
    ["", "Backend creates a new user record in the users table with is_press, press_approved=false, and channel_name."],
    ["", "If is_press is true, backend generates a unique approval token (UUID) and stores it in press_approval_token."],
    ["", "Backend sends an automated email to the system administrator containing the applicant's username, email, channel name, and two action links: Approve and Reject."],
    ["", "Backend generates and returns a JWT access token."],
    ["", "System stores the JWT and navigates user to the home page."]
  ],
  postconditions: [
    "1. New user account is created in the PostgreSQL database.",
    "2. Password is stored as a bcrypt hash; plain-text password is never persisted.",
    "3. User is authenticated immediately after registration.",
    "4. If is_press is true, the account is created with press_approved = false and a notification email is dispatched to the administrator for review.",
    "5. Press features (publishing, channel management) remain locked until the administrator approves the request by clicking the Approve link in the email or via the Admin Panel.",
    "6. User can begin submitting articles and accessing platform features regardless of press approval status."
  ]
};

const UC_ANALYSIS = {
  name: "Article Submission and Analysis",
  id: "UC-003",
  priority: "High",
  actor: "Authenticated User (Regular or Press)",
  description: "Users submit a news article via free text, a URL, or an uploaded image. The system classifies the article as Real or Fake using the RoBERTa ML model and returns a verdict with a confidence score and keyword signals.",
  precondition: [
    "1. User must be authenticated with a valid JWT token.",
    "2. The BERT classifier model must be loaded in the backend.",
    "3. For URL submissions, the target website must be accessible.",
    "4. For image submissions, Tesseract OCR must be installed."
  ],
  events: [
    ["1. User navigates to the Submit / Analyze page.", "System displays the submission form with text area, URL field, and image upload option."],
    ["2. User pastes article text, enters a URL, or uploads an image.", "System enables the Analyze button once input is detected."],
    ["3. User clicks the Analyze button.", "System sends a POST /articles/analyze request with the JWT token."],
    ["", "If URL: backend scrapes article text and title using BeautifulSoup and requests."],
    ["", "If image: backend extracts text using Tesseract OCR (pytesseract)."],
    ["", "Backend detects the language using langdetect."],
    ["", "If non-English: backend translates text to English via deep-translator (Google Translate)."],
    ["", "Backend passes up to 2000 characters to the RoBERTa classifier pipeline (max 512 tokens)."],
    ["", "Model returns REAL or FAKE label with a confidence score between 0 and 1."],
    ["", "Backend scans text for fake and real keyword signals."],
    ["", "Backend saves article record (title, content, url, is_fake, confidence, signals) to database."],
    ["", "If URL: backend updates or creates a source record and recalculates trust score."],
    ["", "Backend returns the classification result to the frontend."],
    ["", "System displays a verdict card with the label, confidence percentage, and keyword lists."]
  ],
  postconditions: [
    "1. Article is saved to the articles table linked to the authenticated user.",
    "2. Source trust score is recalculated using weighted harmonic average of recent articles.",
    "3. Article appears in the user's submission history.",
    "4. Dashboard statistics are updated to reflect the new article.",
    "5. User receives a clear Real or Fake verdict with supporting evidence."
  ]
};

const UC_PRESS = {
  name: "Press Post Management",
  id: "UC-004",
  priority: "High",
  actor: "Press User (is_press = true)",
  description: "Verified press users can create, edit, and delete news posts. When a post is published, the system automatically creates in-app notifications for all regular users. Only the post author may edit or delete their own posts.",
  precondition: [
    "1. User must be authenticated with is_press = true.",
    "2. For post creation: title and content fields must be non-empty.",
    "3. For deletion: notifications referencing the post must exist in the database."
  ],
  events: [
    ["1. Press user navigates to the Publish / My Posts tab.", "System verifies is_press flag; displays the post creation form and My Posts list."],
    ["2. User enters a post title and content body.", "System validates that required fields are not empty."],
    ["3. Optionally, user attaches an image from camera or gallery.", "System encodes image and includes it in the request payload."],
    ["4. User submits the post.", "System sends POST /press/posts with JWT and post data."],
    ["", "Backend verifies the user has press role (returns HTTP 403 if not)."],
    ["", "Backend saves the press post to the press_posts table."],
    ["", "Backend creates a notification record for each regular user in the system."],
    ["", "Backend returns the created post; press news feed is updated."],
    ["5. User taps Edit on one of their posts.", "Backend verifies author_id matches authenticated user; returns post data for editing."],
    ["6. User submits updated content.", "Backend updates the post record via PUT /press/posts/{id}."],
    ["7. User taps Delete on one of their posts.", "Backend deletes all notifications referencing press_post_id first, then deletes the post."]
  ],
  postconditions: [
    "1. New post is visible on the Press News page for all users.",
    "2. Notifications are created and accessible to regular users via GET /notifications.",
    "3. Edited post reflects updated content immediately.",
    "4. Deleted post is removed from the press feed; all related notifications are purged.",
    "5. Only the original author can modify or delete their posts (ownership enforced)."
  ]
};

const UC_CHAT = {
  name: "AI Chat Assistant",
  id: "UC-005",
  priority: "Medium",
  actor: "Authenticated User (Regular or Press), via Web or Mobile",
  description: "Users interact with a Gemini-powered AI assistant that answers questions about the VeriSource platform, helps interpret classification results, and provides general media literacy guidance. The assistant maintains a rolling conversation history of the last 10 messages.",
  precondition: [
    "1. User must be authenticated.",
    "2. GEMINI_API_KEY must be configured in the backend environment.",
    "3. The Gemini API service must be accessible from the backend.",
    "4. Backend API server must be running."
  ],
  events: [
    ["1. User clicks the floating chat button (web) or AI chat icon (mobile).", "System opens the animated chat panel (web) or slide-up modal (mobile)."],
    ["2. User types a question or message.", "System displays the typed text in the chat interface input field."],
    ["3. User clicks Send.", "System sends POST /chat/assist with the message and conversation history (last 10 messages)."],
    ["", "Backend builds a context-aware prompt including the VeriSource system instructions."],
    ["", "Backend calls the Gemini gemini-flash-lite-latest API with the conversation context."],
    ["", "Gemini API processes the request and generates a response."],
    ["", "Backend returns the AI response text to the frontend."],
    ["", "System displays the AI response in the chat panel."],
    ["4. User continues the conversation.", "System maintains rolling history; older messages beyond 10 are dropped."],
    ["5. User closes the chat panel.", "System hides the chat widget; history is retained for the session."]
  ],
  postconditions: [
    "1. AI response is displayed to the user in the chat interface.",
    "2. Conversation history is preserved for the duration of the session.",
    "3. User receives helpful guidance about platform features or misinformation topics.",
    "4. If Gemini API is unavailable, system returns a graceful error message without crashing."
  ]
};

const UC_DASHBOARD = {
  name: "Dashboard and Analytics",
  id: "UC-006",
  priority: "Medium",
  actor: "Authenticated User",
  description: "Users access a live analytics dashboard showing aggregate article statistics, a real-vs-fake bar chart, and a Global Source Map displaying news domain trust scores plotted geographically on an interactive world map.",
  precondition: [
    "1. User must be authenticated.",
    "2. At least some articles must have been submitted to the system.",
    "3. Backend API server must be running.",
    "4. React-leaflet and CARTO tile service must be accessible."
  ],
  events: [
    ["1. User navigates to the Dashboard page.", "System sends parallel API requests for summary and sources data."],
    ["", "System fetches GET /analytics/summary returning total, fake, real counts and fake percentage."],
    ["", "System fetches GET /analytics/sources returning all source domains with trust scores and coordinates."],
    ["", "System renders four animated stat counters using Framer Motion."],
    ["", "System renders a Recharts bar chart comparing real vs fake article counts with color coding."],
    ["", "System geocodes source domains to latitude/longitude using country metadata and TLD lookup table."],
    ["", "System renders the react-leaflet Global Source Map with CARTO light tile layer."],
    ["", "System plots CircleMarker for each source: green for trust score > 0.5, red otherwise."],
    ["2. User hovers over or taps a map marker.", "System displays a Leaflet Popup with domain name, trust percentage, and article count."],
    ["3. User views the source list below the map.", "System displays sources sorted by total article count with trust score percentages."]
  ],
  postconditions: [
    "1. Dashboard displays up-to-date statistics reflecting all submitted articles.",
    "2. Global Source Map shows source distribution with trust-based color coding.",
    "3. User can identify high-risk and trusted news sources visually.",
    "4. All data is fetched fresh on each page load to reflect latest submissions."
  ]
};

// ── Build document ────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Times New Roman", color: "1A3A5C" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Times New Roman", color: "2E5FA3" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman", color: "1A3A5C" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
      { id: "ListParagraph", name: "List Paragraph", basedOn: "Normal", next: "Normal",
        paragraph: { indent: { left: 720, hanging: 360 } } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E5FA3", space: 4 } },
        children: [new TextRun({ text: "VeriSource — Graduation Project Report", font: "Times New Roman", size: 20, color: "2E5FA3" })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Page ", font: "Times New Roman", size: 20, color: "888888" }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 20, color: "888888" })
        ]
      })] })
    },
    children: [

      // ── COVER PAGE ──────────────────────────────────────────────────────────
      spacer(), spacer(), spacer(),
      centered("VeriSource", { bold: true, size: 48, color: "1A3A5C" }),
      centered("Fake News Detection Platform", { bold: true, size: 32, color: "2E5FA3" }),
      spacer(),
      centered("By", { size: 24 }),
      centered("Hassan Zaarour", { bold: true, size: 28 }),
      spacer(),
      centered("Graduation Report Project", { size: 24 }),
      centered("Submitted in Partial Fulfillment of the Requirements for the", { size: 22 }),
      centered("Degree of Bachelor of Science in Computer Science", { size: 22 }),
      spacer(),
      centered("Department of Computer Science", { size: 24 }),
      centered("Faculty of Sciences & Arts", { size: 24 }),
      spacer(),
      centered("Supervised by:  Prof. / Dr: Bassel Dhaini", { size: 24 }),
      spacer(), spacer(),
      centered("2025 - 2026", { bold: true, size: 24 }),
      pageBreak(),

      // ── DEFENSE COMMITTEE ───────────────────────────────────────────────────
      spacer(), spacer(),
      centered("The Report Defense Committee for (Hassan Zaarour)", { bold: true, size: 24 }),
      centered("Certifies that this is the approved version of the following report.", { size: 24 }),
      spacer(),
      centered("VeriSource", { bold: true, size: 32, color: "1A3A5C" }),
      spacer(), spacer(),
      centered("APPROVED BY:", { bold: true, size: 24 }),
      spacer(),
      para("Supervisor Signature:                        "),
      para("(Name typed under the line)"),
      spacer(),
      para("Examiner Signature:                          "),
      para("(Name typed under the line)"),
      spacer(),
      para("Examiner Signature:                          "),
      para("(Name typed under the line)"),
      pageBreak(),

      // ── ACKNOWLEDGEMENT ─────────────────────────────────────────────────────
      h1("Acknowledgement"),
      para("I would like to express my sincere gratitude to my supervisor, Dr. Bassel Dhaini, for their invaluable guidance, continuous support, and expert feedback throughout the development of this graduation project. Their encouragement and deep understanding of the subject matter were instrumental in shaping the direction and quality of this work."),
      spacer(),
      para("I am also grateful to the faculty of the Department of Computer Science for providing the academic environment and resources necessary to complete this project. This work would not have been possible without the knowledge and skills gained through the courses offered by the department."),
      spacer(),
      para("Finally, I extend my deepest thanks to my family and friends for their patience, understanding, and unwavering moral support throughout this challenging and rewarding journey."),
      pageBreak(),

      // ── ABSTRACT ────────────────────────────────────────────────────────────
      h1("Abstract"),
      para("VeriSource is a full-stack fake news detection platform designed to address the growing challenge of misinformation in online media. The system enables users to submit news articles in three formats — plain text, URLs, or images — and classifies them as Real or Fake using a fine-tuned RoBERTa transformer model served through a Python FastAPI backend. Classification results include a confidence score and keyword signals that explain the model's decision in plain language."),
      spacer(),
      para("The system supports two user roles: regular users who can submit articles for analysis, and press users (journalists) who additionally have the ability to publish news posts visible to all platform users. When a press post is published, the system automatically generates in-app notifications for regular users."),
      spacer(),
      para("A key feature of VeriSource is the integration of an AI-powered chat assistant using the Google Gemini API, available on both the web and mobile applications. This assistant is configured with a platform-specific system prompt and helps users interpret results, navigate the platform, and learn about media literacy."),
      spacer(),
      para("The backend leverages SQLAlchemy with a PostgreSQL database deployed in Docker, JWT-based authentication, and a modular FastAPI router architecture. The web frontend is built with React 18, Vite, Tailwind CSS, Framer Motion, and react-leaflet for an interactive global source map. The mobile application is built with Expo SDK 56 and React Native, featuring a bilingual news feed supporting Arabic and English content sourced from the NewsAPI."),
      pageBreak(),

      // ── LIST OF FIGURES (before List of Tables per requirement) ─────────────
      h1("List of Figures"),
      para("Figure 1: VeriSource System Use Case Diagram"),
      para("Figure 2: Data Flow Diagram - Level 0 (Context Diagram)"),
      para("Figure 3: Data Flow Diagram - Level 1"),
      para("Figure 4: User Sequence Diagram - Article Submission Flow"),
      para("Figure 5: Press User Sequence Diagram - Article Publication Flow"),
      para("Figure 6: Entity Relationship Diagram (ERD)"),
      para("Figure 7: Welcome Page (Web)"),
      para("Figure 8: Login and Register Screens"),
      para("Figure 9: Article Analysis Screen"),
      para("Figure 10: Dashboard Page"),
      para("Figure 11: Mobile News Headlines Screen"),
      para("Figure 12: Press Publish Screen"),
      para("Figure 13: AI Chat Assistant (Web)"),
      para("Figure 14: AI Chat Assistant (Mobile)"),
      para("Figure 15: Press News Page (Web)"),
      para("Figure 16: Mobile Home Screen with Arabic News Feed"),
      pageBreak(),

      // ── LIST OF TABLES ───────────────────────────────────────────────────────
      h1("List of Tables"),
      para("Table 1: User Login Use Case"),
      para("Table 2: User Registration Use Case"),
      para("Table 3: Article Submission and Analysis Use Case"),
      para("Table 4: Press Post Management Use Case"),
      para("Table 5: AI Chat Assistant Use Case"),
      para("Table 6: Dashboard and Analytics Use Case"),
      para("Table 7: Welcome / Home Page Components"),
      para("Table 8: Login and Register Flow"),
      para("Table 9: Article Analysis Flow"),
      para("Table 10: Dashboard Components"),
      para("Table 11: Mobile News Feed Components"),
      para("Table 12: Press Publish Components"),
      para("Table 13: AI Chat Assistant Components"),
      pageBreak(),

      // ── TABLE OF CONTENTS ────────────────────────────────────────────────────
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 1: Introduction
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 1: Introduction and Context"),
      para("In the current digital age, the rapid spread of false and misleading information across social media platforms and online news outlets has become one of the most pressing challenges facing modern society. The ability to quickly verify the credibility of a news article before sharing or acting upon it is no longer a convenience but a necessity. Manual fact-checking processes are slow, resource-intensive, and unable to scale to the volume of content published online every second."),
      spacer(),
      para("To address this challenge, the VeriSource platform was developed as an intelligent, multi-platform fake news detection system. VeriSource combines state-of-the-art natural language processing with an accessible web and mobile interface, allowing users of all technical backgrounds to instantly analyze news articles and receive AI-powered credibility assessments."),
      spacer(),
      para("The system is built as a cross-platform solution comprising three interconnected components: a React web application, an Expo React Native mobile application, and a Python FastAPI backend. These components communicate through a RESTful API that handles article classification, user authentication, press publication, analytics, and AI-assisted chat."),
      spacer(),
      para("VeriSource distinguishes itself by supporting two distinct user roles. Regular users can submit news articles for AI-powered analysis and view a dashboard of aggregated statistics and source trust scores. Press users, who are verified journalists or channel operators, can additionally publish news posts to a public press feed, effectively creating a community-curated news channel within the platform."),

      h2("1.1 Objectives"),
      para("The primary objectives of the VeriSource platform are as follows:"),
      bullet("To develop an accurate and reliable AI-based classifier for detecting fake news articles using a pre-trained RoBERTa transformer model served through a production-grade FastAPI backend."),
      bullet("To build a cross-platform system accessible via both web browser and mobile devices, ensuring wide usability for diverse user populations."),
      bullet("To provide real-time article analysis with confidence-scored results that give users meaningful insight into the credibility of submitted content."),
      bullet("To support a dedicated press user role that allows verified journalists to publish and manage news content, contributing to a trusted news ecosystem within the platform."),
      bullet("To implement an interactive AI chat assistant using the Google Gemini API to assist users with platform navigation and media literacy education."),
      bullet("To aggregate and display Arabic and English news headlines from external APIs, enabling multilingual news consumption on the mobile application."),
      bullet("To maintain a statistical dashboard with a global source map, trust score tracking, and real-time analytics on classified articles."),

      h2("1.2 Challenges"),
      para("The development of VeriSource involved several technical and design challenges, including:"),
      bullet("Integrating a pre-trained RoBERTa transformer model into a production FastAPI backend with acceptable inference latency for real-time user requests."),
      bullet("Designing a dual-role user system with role-based access control that clearly distinguishes between regular users and press users while sharing the same authentication flow."),
      bullet("Aggregating real-time news content from external APIs in both Arabic and English while handling API rate limits, response inconsistencies, and latency."),
      bullet("Building a geographically aware source trust system that maps news domains to countries and assigns trust scores based on weighted historical classification data."),
      bullet("Ensuring consistent behavior between the web and mobile platforms while sharing the same backend API, particularly with respect to authentication, file uploads, and LAN connectivity during development."),
      bullet("Integrating the Google Gemini AI API with a free-tier key while navigating model availability and quota restrictions across both platforms."),

      h2("1.3 Contributions"),
      para("This project makes the following contributions to the field of information verification and digital media technology:"),
      bullet("The design and implementation of a full-stack fake news detection system that combines transformer-based NLP classification with a clean, role-aware web and mobile interface."),
      bullet("A novel press publication workflow that integrates verified journalist accounts into a consumer-facing news platform, bridging automated classification and human-authored content."),
      bullet("An intelligent bilingual news aggregation feed for mobile users, supporting seamless switching between Arabic and English headlines from the NewsAPI."),
      bullet("The integration of a conversational AI assistant (powered by Google Gemini) specifically tailored to help users navigate the platform and improve their media literacy."),
      bullet("A transparent source trust scoring system backed by geographic metadata, providing users with visual and statistical insight into the credibility of news domains worldwide."),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 2: Literature Review
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 2: Literature Review"),
      para("This chapter reviews existing research and platforms in the domain of automated fake news detection, natural language processing for credibility assessment, and cross-platform media applications. It situates VeriSource within the broader academic and industrial landscape and highlights the gaps that this project addresses."),

      h2("2.1 The Problem of Online Misinformation"),
      para("The proliferation of digital media has fundamentally transformed how people consume news. While this shift has democratized information access, it has simultaneously created conditions in which false, misleading, and fabricated content can spread rapidly and widely before corrections can be issued. Research by Vosoughi, Roy, and Aral (2018) demonstrated that false news stories on Twitter spread six times faster than true ones, reaching a far greater audience in the same timeframe. The social and political consequences of this asymmetry are well documented, from election interference to public health misinformation."),
      para("Early approaches to automated fake news detection relied on stylometric features such as writing style, grammatical complexity, and lexical diversity to distinguish credible from non-credible content. While these methods showed promise on constrained datasets, they lacked the semantic understanding required to generalize across topics, languages, and domains."),

      h2("2.2 Natural Language Processing Approaches"),
      para("The emergence of deep learning and transformer-based language models has significantly advanced the field of text classification, including fake news detection. BERT (Bidirectional Encoder Representations from Transformers), introduced by Devlin et al. (2019), established a new paradigm by pre-training on massive corpora and fine-tuning on task-specific datasets. RoBERTa, a robustly optimized variant of BERT, achieved further gains through more aggressive training procedures and larger batch sizes."),
      para("Research using BERT-based classifiers on labeled fake news datasets such as LIAR, FakeNewsNet, and GossipCop has consistently demonstrated accuracy improvements over traditional machine learning approaches. VeriSource leverages the hamzab/roberta-fake-news-classification model, a RoBERTa variant fine-tuned specifically for binary fake/real classification of news articles."),

      h2("2.3 Existing Platforms and Tools"),
      para("Several commercial and academic platforms have attempted to address fake news detection:"),
      bullet("Snopes and PolitiFact provide manual fact-checking services operated by human journalists. While these services are highly reliable, they are slow and cannot scale to the volume of online content."),
      bullet("ClaimBuster is an academic tool that uses machine learning to score the check-worthiness of factual claims in political speech. It focuses on claim detection rather than full-article classification."),
      bullet("NewsGuard assigns trust ratings to news domains based on editorial reviews. While comprehensive, it covers only a fraction of the global news ecosystem and does not classify individual articles."),
      bullet("Media Bias/Fact Check offers source-level bias and credibility ratings but does not analyze individual articles for veracity."),
      spacer(),
      para("None of these existing solutions combine automated AI classification, press content publication, multilingual news aggregation, geographic source mapping, and an AI chat assistant in a single integrated platform — the combination that VeriSource uniquely provides."),

      h2("2.4 Cross-Platform Architecture in Media Applications"),
      para("Modern media platforms increasingly adopt cross-platform architectures to reach users across devices. The React Native framework, maintained by Meta, enables developers to build native mobile applications from a shared JavaScript codebase targeting both iOS and Android. Expo extends React Native with a managed workflow, pre-built native modules, and an over-the-air update system, significantly reducing development complexity."),
      para("On the backend side, Python FastAPI has emerged as a leading choice for high-performance machine learning APIs due to its asynchronous support, automatic OpenAPI documentation generation, and native Pydantic integration for data validation. Its performance characteristics, comparable to Node.js and Go in benchmarks, make it well-suited for serving ML inference endpoints in production."),

      h2("2.5 AI-Assisted User Interfaces"),
      para("The integration of conversational AI assistants into consumer applications has grown substantially with the availability of large language model APIs. Google Gemini, OpenAI GPT, and Anthropic Claude provide REST APIs that enable developers to embed sophisticated conversational agents into applications with minimal infrastructure. VeriSource uses the Gemini gemini-flash-lite-latest model due to its favorable free-tier quotas and fast response times, configuring it with a domain-specific system prompt to act as a VeriSource support agent."),

      h2("2.6 Contribution of VeriSource to the Field"),
      para("VeriSource contributes to the existing body of work by:"),
      bullet("Combining AI-based article classification with a live press publication system in a single platform."),
      bullet("Providing multilingual support for Arabic and English news consumption, addressing a geographic gap in existing English-centric tools."),
      bullet("Implementing a geographically distributed source trust map that visualizes credibility at the domain level."),
      bullet("Integrating a Gemini-powered AI assistant trained on platform-specific context to improve media literacy among users."),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 3: Requirements Analysis
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 3: Requirements Analysis"),
      para("This chapter defines the functional and non-functional requirements of the VeriSource platform. It outlines the core system capabilities, the actors who interact with the system, and the use case scenarios that govern those interactions. Data flow diagrams and sequence diagrams are also presented to illustrate how information moves through the system."),

      h2("3.1 Functional Requirements"),
      para("The VeriSource platform supports two primary user roles: Regular Users and Press Users. All users must authenticate before accessing protected features. The following functional requirements define the core capabilities of the system:"),
      bullet("User Authentication and Access Control: The system shall support secure user registration and login using email and password. Passwords shall be stored as bcrypt hashes. JWT tokens with a 7-day expiry shall be issued on successful login. Role-based access control shall distinguish between regular users and press users throughout the platform."),
      bullet("Article Submission and AI Classification: The system shall allow authenticated users to submit news articles as free text, URLs, or images. URL submissions shall trigger automatic web scraping. Image submissions shall trigger OCR text extraction. Non-English text shall be automatically detected and translated before classification. The RoBERTa model shall classify content as Real or Fake with a confidence score and keyword signals."),
      bullet("Press Post Management: Press users shall be able to create, edit, and delete news posts with a title, body text, optional image, and language tag. Post creation shall automatically generate in-app notifications for all regular users. Only the post author shall be permitted to edit or delete their own posts."),
      bullet("Submission History: Authenticated users shall be able to view a chronological history of their previously submitted articles with classification results and timestamps."),
      bullet("Dashboard and Analytics: The system shall maintain aggregate statistics on all analyzed articles and display them on a dashboard including total articles, fake count, real count, and fake percentage. A bar chart and global source map shall provide visual representations of the data."),
      bullet("Global Source Map: The system shall extract domain names from URL-based submissions, geocode them to countries, and display them on an interactive world map with trust-score-based color coding."),
      bullet("Multilingual News Feed: The mobile application shall fetch and display real-time news headlines from the NewsAPI, with a toggle to switch between English and Arabic content."),
      bullet("AI Chat Assistant: Both web and mobile applications shall include a Gemini-powered AI chat widget that assists users with platform navigation and media literacy questions."),
      bullet("Notification System: The system shall deliver in-app notifications to regular users when a press user publishes a new article."),

      h2("3.2 Non-Functional Requirements"),
      bullet("Performance: The RoBERTa classifier endpoint shall respond within 5 seconds for typical article lengths. The REST API shall handle concurrent requests without degradation. Frontend pages shall load within 2 seconds on standard broadband connections."),
      bullet("Security: All API endpoints that return or modify user-specific data shall require a valid JWT bearer token. Passwords shall be stored using bcrypt with an adaptive cost factor. Press-only routes shall return HTTP 403 Forbidden for non-press authenticated users."),
      bullet("Usability: The web interface shall be responsive and accessible on modern desktop browsers. The mobile application shall follow native design patterns for Android and iOS. All error messages shall be clearly communicated to users in plain language."),
      bullet("Scalability: The backend shall be designed to support horizontal scaling, with stateless API routes and a centralized PostgreSQL database accessible to multiple backend instances."),
      bullet("Reliability: The system shall handle external API failures (news aggregation, Gemini API) gracefully, displaying appropriate error states without crashing. Database transactions shall use SQLAlchemy session management to maintain data integrity."),
      bullet("Maintainability: The codebase shall follow modular patterns with clear separation of concerns across API routers, database models, ML services, and frontend components."),

      // ── USE CASE SECTION ─────────────────────────────────────────────────────
      h2("3.3 UML Use Cases and Scenarios"),
      para("This section describes the primary use cases of the VeriSource system using the detailed UML use case table format. Two primary actors interact with the platform: Regular Users and Press Users. Each use case table below follows the standardized format with Actor Actions and System Responses presented in a structured two-column layout."),

      // UC-001: User Login
      spacer(),
      para("Table 1: User Login Use Case", { bold: true }),
      detailedUseCaseTable(UC_LOGIN),
      spacer(),

      // UC-002: User Registration
      para("Table 2: User Registration Use Case", { bold: true }),
      detailedUseCaseTable(UC_REGISTER),
      spacer(),

      // UC-003: Article Submission and Analysis
      para("Table 3: Article Submission and Analysis Use Case", { bold: true }),
      detailedUseCaseTable(UC_ANALYSIS),
      spacer(),

      // UC-004: Press Post Management
      para("Table 4: Press Post Management Use Case", { bold: true }),
      detailedUseCaseTable(UC_PRESS),
      spacer(),

      // UC-005: AI Chat Assistant
      para("Table 5: AI Chat Assistant Use Case", { bold: true }),
      detailedUseCaseTable(UC_CHAT),
      spacer(),

      // UC-006: Dashboard and Analytics
      para("Table 6: Dashboard and Analytics Use Case", { bold: true }),
      detailedUseCaseTable(UC_DASHBOARD),
      spacer(),

      // Use Case Diagram
      h3("3.3.1 VeriSource System UML Use Case Diagram"),
      spacer(),
      ...figureBlock(1, "VeriSource System Use Case Diagram", "image1.png", 600, 285),
      para("The UML Use Case Diagram for VeriSource illustrates the interactions between the platform and its two primary actors: Regular Users and Press Users. The diagram captures the complete set of system functionalities available to each actor and highlights the extension relationship between the Press User role and the Regular User role."),
      para("The Regular User represents the primary consumer of the platform. Their interactions include submitting articles for fake news analysis, viewing classification results, accessing the submission history, browsing the analytics dashboard, reading the global source map, consuming press news, and interacting with the AI chat assistant."),
      para("The Press User extends the Regular User role with additional capabilities for content publication. Press users can create, edit, and delete their own press posts, which are then visible to all platform users. The system automatically generates notifications for regular users when a new press post is published."),

      // ── DATA FLOW DIAGRAMS ───────────────────────────────────────────────────
      h2("3.4 Data Flow Diagram"),
      para("The Data Flow Diagrams illustrate how data moves through the VeriSource system between external actors, internal processes, and data stores. Two levels of DFD are presented: the Level 0 context diagram provides a high-level system overview, and the Level 1 DFD breaks the system into its core functional processes."),

      h3("3.4.1 DFD Level 0 — Context Diagram"),
      spacer(),
      ...figureBlock(2, "Data Flow Diagram - Level 0 (Context Diagram)", "image2.png", 600, 326),
      spacer(),
      para("The Level 0 Data Flow Diagram represents the VeriSource system as a single central process and illustrates its interaction with external entities at a high level. The diagram shows how the system's primary actors exchange data with the system without exposing its internal processing details."),
      para("The Regular User submits article content (text, URL, or image) to the system and receives in return a classification result, a confidence score, submission history, dashboard analytics, press news posts, and AI chat responses. The Press User interacts with the system to publish and manage news posts, and receives publication confirmations and their post management data in return."),
      para("The VeriSource system also integrates with two external services. The Google Gemini API is called by the backend for every AI chat interaction, receiving a conversation context and returning a generated response. The NewsAPI is queried by the mobile frontend to retrieve real-time English and Arabic news headlines for the bilingual news feed. Overall, this DFD Level 0 defines the system boundary of VeriSource and provides a clear overview of its data exchanges at the highest level of abstraction."),

      h3("3.4.2 DFD Level 1 — Internal Processes"),
      spacer(),
      ...figureBlock(3, "Data Flow Diagram - Level 1", "image3.png", 600, 405),
      spacer(),
      para("The Level 1 Data Flow Diagram decomposes the VeriSource system into its key internal processes, showing how data flows between external entities, system processes, and persistent data stores. This level of detail reveals the modular architecture of the platform and the specific responsibilities of each functional component."),
      para("The diagram includes the following core processes: User Authentication and Access Control, which handles registration, login, and JWT token management; the Article Classification Service, which preprocesses submitted content, calls the RoBERTa ML model, and stores results; the Press Publication Module, which manages press post CRUD operations and notification generation; the Analytics and Reporting process, which aggregates statistics and geocodes source domains; the AI Chat Module, which interfaces with the Gemini API and manages conversation history; and the News Aggregation Service, which fetches and serves external headlines on the mobile platform."),
      para("Persistent data is maintained in five PostgreSQL tables: users, articles, sources, press_posts, and notifications. Each process reads from and writes to the appropriate tables through the SQLAlchemy ORM, ensuring referential integrity and transactional consistency. External services — the Gemini API and NewsAPI — are accessed exclusively by internal processes and never directly by data stores."),

      // ── SEQUENCE DIAGRAMS ────────────────────────────────────────────────────
      h2("3.5 User Sequence Diagram"),
      para("The User Sequence Diagram illustrates the chronological flow of interactions between a Regular User, the web or mobile client, and the VeriSource backend during the core article submission and analysis workflow."),
      spacer(),
      ...figureBlock(4, "User Sequence Diagram - Article Submission Flow", "image5.png", 580, 387),
      spacer(),
      para("This sequence diagram illustrates the complete interaction flow between a regular user and the VeriSource system during article submission. The user authenticates via the login endpoint, receiving a JWT token that is stored client-side. The user then navigates to the analysis page and submits an article. The backend receives the authenticated request, preprocesses the content (scraping if URL, OCR if image, translating if non-English), passes the text to the RoBERTa classifier, stores the result in the articles table, updates the source trust score, and returns the classification verdict to the frontend. The diagram demonstrates the stateless nature of the API — every request carries the JWT in the Authorization header — and the sequential nature of the classification pipeline."),

      h2("3.6 Press User Sequence Diagram"),
      para("The Press User Sequence Diagram illustrates the interaction flow when a press user publishes a news article on the platform, including the automatic notification generation that follows successful publication."),
      spacer(),
      ...figureBlock(5, "Press User Sequence Diagram - Article Publication Flow", "image6.png", 580, 387),
      spacer(),
      para("This sequence diagram illustrates the interaction flow of a press user within the VeriSource system. The press user authenticates and navigates to the Publish tab on the web or mobile interface. Upon submitting a post, the backend verifies the is_press flag on the authenticated user's account, saves the post to the press_posts table, and then iterates through all regular user accounts to create individual notification records in the notifications table. The frontend subsequently fetches the updated press post list and notification count. The diagram also shows the press user querying their own posts via the /press/posts/mine endpoint and executing a deletion, during which the backend first removes all associated notification records before deleting the post itself to satisfy the foreign key constraint."),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 4: System Design
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 4: System Design"),

      h2("4.1 Introduction"),
      para("This chapter presents the system design of the VeriSource platform, covering the data architecture, database schema, entity relationships, technology stack, and module decomposition. VeriSource is designed as a three-tier architecture comprising a presentation layer (web and mobile frontends), an application layer (FastAPI backend with ML services), and a data layer (PostgreSQL). Each tier communicates through well-defined interfaces: RESTful HTTP for frontend-backend communication, SQLAlchemy ORM for backend-database interaction, and the Hugging Face Transformers pipeline for backend-ML communication."),
      para("The RoBERTa-based fake news classifier operates as a service within the backend, loaded once at startup and reused for all subsequent classification requests, avoiding the overhead of repeated model initialization. Similarly, the Google Gemini API client is initialized per-request using an async HTTP client, enabling non-blocking AI interactions within the FastAPI event loop."),

      h2("4.2 Data Design"),
      para("The VeriSource data layer is built on PostgreSQL, a robust open-source relational database management system that provides strong consistency, full ACID transaction support, and rich querying capabilities through SQL. The database is hosted in a Docker container for portability and ease of deployment across development and production environments."),
      para("The schema consists of five normalized tables that together capture all persistent state in the platform: user identities and roles, classified articles, news source reputations, press-authored posts, and user notifications. The design adheres to third normal form to eliminate data redundancy and enforce referential integrity through foreign key constraints."),
      spacer(),
      para("Table structure overview:", { bold: true }),
      spacer(),
      erEntityTable("users", "id", [
        ["id", "INTEGER — Primary Key, auto-increment. Uniquely identifies each user account."],
        ["email", "VARCHAR — Unique, indexed. The user's login email address."],
        ["username", "VARCHAR — Unique, indexed. Display name for the user."],
        ["hashed_password", "VARCHAR — bcrypt hash of the user's password. Plain-text is never stored."],
        ["is_active", "BOOLEAN — Default true. Allows soft-disabling accounts without deletion."],
        ["is_press", "BOOLEAN — Default false. Grants access to press publication features when true."],
        ["channel_name", "VARCHAR — Nullable. Display name for press users' publication channel."],
        ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set by PostgreSQL server on insert."]
      ], "User and Authentication"),
      spacer(),
      erEntityTable("articles", "id", [
        ["id", "INTEGER — Primary Key, auto-increment."],
        ["title", "VARCHAR — Nullable. Article title, extracted from scraped URL or provided by user."],
        ["content", "TEXT — Required. Full article text used for classification."],
        ["url", "VARCHAR — Nullable. Source URL for URL-based submissions."],
        ["input_type", "VARCHAR — One of: 'url', 'text', or 'image'. Tracks submission method."],
        ["is_fake", "BOOLEAN — Classification result from the RoBERTa model."],
        ["confidence_score", "FLOAT — Model confidence between 0.0 and 1.0."],
        ["fake_signals", "JSON — Array of suspicious keywords detected in the article text."],
        ["real_signals", "JSON — Array of credibility keywords detected in the article text."],
        ["source_id", "INTEGER FK → sources.id — Nullable. Linked source domain for URL submissions."],
        ["submitted_by", "INTEGER FK → users.id — Nullable. Allows anonymous submissions."],
        ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on insert."]
      ], "Article and Classification"),
      spacer(),
      erEntityTable("sources", "id", [
        ["id", "INTEGER — Primary Key, auto-increment."],
        ["name", "VARCHAR — Unique. Display name of the news source domain."],
        ["domain", "VARCHAR — Unique. The domain string (e.g., bbc.com) extracted from submitted URLs."],
        ["country", "VARCHAR — Nullable. Country of origin used for geographic map placement."],
        ["trust_score", "FLOAT — Value between 0.0 and 1.0. Recalculated after each new article from this domain using a weighted harmonic average where recent articles carry higher weight."],
        ["total_articles", "INTEGER — Count of all articles analyzed from this domain."],
        ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on first encounter of the domain."],
        ["updated_at", "TIMESTAMP WITH TIME ZONE — Auto-updated on every trust score recalculation."]
      ], "Source Trust"),
      spacer(),
      erEntityTable("press_posts", "id", [
        ["id", "INTEGER — Primary Key, auto-increment."],
        ["title", "VARCHAR — Required. Headline of the press post."],
        ["content", "TEXT — Required. Full body text of the press post."],
        ["language", "VARCHAR — Default 'en'. Language tag for filtering (e.g., 'en', 'ar')."],
        ["image_url", "VARCHAR — Nullable. URL path to an optional cover image uploaded with the post."],
        ["author_id", "INTEGER FK → users.id — Required. The press user who authored the post."],
        ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on insert."]
      ], "Press Publication"),
      spacer(),
      erEntityTable("notifications", "id", [
        ["id", "INTEGER — Primary Key, auto-increment."],
        ["type", "VARCHAR — Default 'press_post'. Category of notification."],
        ["title", "VARCHAR — Required. Short notification headline."],
        ["message", "VARCHAR — Required. Full notification body text."],
        ["press_post_id", "INTEGER FK → press_posts.id — Nullable. Links notification to its source press post."],
        ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on insert."]
      ], "Notification"),
      spacer(),
      para("Key Relationships and Referential Integrity:", { bold: true }),
      bullet("users → articles (One-to-Many): A single user can submit many articles. Each article optionally references one user via submitted_by. Cascade not applied — deleting a user does not delete their articles."),
      bullet("users → press_posts (One-to-Many): A press user can author many posts. Each post belongs to exactly one author via author_id. Non-press users cannot create posts (HTTP 403 enforced at API layer)."),
      bullet("articles → sources (Many-to-One): Multiple articles from the same domain reference one source record. The sources table is populated automatically when new domains are encountered in URL submissions."),
      bullet("press_posts → notifications (One-to-Many): Each press post can generate many notifications (one per user). On post deletion, all associated notifications must be removed first to satisfy the foreign key constraint, implemented via SQLAlchemy's delete with synchronize_session=False followed by db.flush()."),

      h2("4.3 ER Diagram"),
      para("The Entity Relationship Diagram below illustrates the complete logical data structure of the VeriSource platform. It shows all five entities, their attributes, data types, primary keys, and the foreign key relationships that enforce referential integrity across the schema."),
      spacer(),
      ...figureBlock(6, "Entity Relationship Diagram (ERD)", "image4.png", 580, 387),
      spacer(),
      para("The VeriSource relational schema is intentionally lean and normalized. The five-table design eliminates data redundancy while maintaining complete representational coverage of all platform features. Users are the central entity from which all other records radiate outward through foreign key references. The sources table serves as a derived analytical store rather than a transactional entity, accumulating credibility evidence from URL-based article submissions and updating its trust_score and total_articles fields after every new classification."),

      h2("4.4 Technologies Used"),
      para("Backend and API:"),
      bullet("Python 3.11: Primary backend language, selected for its extensive NLP and ML ecosystem."),
      bullet("FastAPI: High-performance asynchronous web framework with automatic OpenAPI documentation and Pydantic validation."),
      bullet("SQLAlchemy ORM: Database abstraction layer providing declarative models and connection pooling for PostgreSQL."),
      bullet("PostgreSQL (Docker): Relational database for persistent storage of users, articles, press posts, notifications, and source records."),
      bullet("python-jose: JWT generation and validation for stateless authentication."),
      bullet("passlib / bcrypt: Password hashing with adaptive cost factor."),
      bullet("httpx: Async HTTP client used for Gemini API calls from the backend."),
      bullet("Hugging Face Transformers: Library used to load and run the pre-trained RoBERTa fake news classifier."),
      spacer(),
      para("AI and Intelligence:"),
      bullet("RoBERTa (hamzab/roberta-fake-news-classification): Pre-trained transformer model fine-tuned on a labeled fake news dataset. Returns REAL or FAKE with a confidence score."),
      bullet("Google Gemini API (gemini-flash-lite-latest): Generative AI model used for the VeriSource AI chat assistant on both web and mobile platforms."),
      spacer(),
      para("Mobile Application:"),
      bullet("React Native and Expo SDK 56: Cross-platform mobile framework targeting Android and iOS from a unified JavaScript codebase."),
      bullet("expo-router: File-based routing system for React Native navigation and tab management."),
      bullet("expo-image-picker: Native camera and gallery access for press post image uploads."),
      bullet("NewsAPI: External REST API for fetching real-time English and Arabic news headlines."),
      spacer(),
      para("Web Application:"),
      bullet("React 18 with Vite: Frontend library and build tool for the web application."),
      bullet("Tailwind CSS v4: Utility-first CSS framework for consistent and responsive styling."),
      bullet("Framer Motion: Animation library for page transitions, hover effects, and animated UI components."),
      bullet("Recharts: Composable charting library used for the real-vs-fake bar chart on the analytics dashboard."),
      bullet("react-leaflet with CARTO tiles: Interactive map component library used for the Global Source Map."),
      bullet("Axios: Promise-based HTTP client for API communication between the frontend and backend."),

      h2("4.5 System Modules and Components"),
      bullet("Authentication Module: Handles user registration, login, and JWT token issuance. Enforces role-based access so that press-only routes return HTTP 403 for non-press users."),
      bullet("Article Analysis Module: Receives article text or URL, scrapes URL content if needed, runs RoBERTa inference, stores the result, and updates the source trust score."),
      bullet("Press Publication Module: Full CRUD API for press posts with ownership enforcement. Automatically triggers notification creation for all regular users on post creation."),
      bullet("Dashboard Analytics Module: Aggregates classification statistics and source trust data for display on the dashboard with geographic coordinates for map rendering."),
      bullet("AI Chat Module: Manages conversation history, builds context-aware Gemini prompts including a VeriSource-specific system prompt, and returns AI responses."),
      bullet("Notification Module: Creates per-user notification records when press posts are published, and serves them to the web client via GET /notifications."),
      bullet("Press Approval Module: When a user registers with the press role, the backend generates a single-use UUID token, stores it in the press_approval_token column, and dispatches an HTML email to the administrator containing the applicant's details and two action links — Approve and Reject. Clicking Approve sets press_approved = true and clears the token. Clicking Reject sets is_press = false, clears the channel name, and removes the token, reverting the account to a regular user. An in-app Admin Panel at /admin provides the same approve and reject actions for cases where the email link is inaccessible."),
      bullet("Mobile News Aggregation: Client-side service that fetches top headlines from the NewsAPI in English or Arabic based on the user's language toggle selection."),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 5: Software and Implementation
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 5: Project Software and Implementation"),

      h2("5.1 Introduction"),
      para("This chapter presents an overview of the implementation of the VeriSource platform. It describes how the functional requirements defined in Chapter 3 are realized through the frontend web application, the mobile application, the FastAPI backend, and the AI classification service. The chapter concludes with a summary of the data flow that integrates all components end-to-end, and a breakdown of software functions organized by user role."),

      h2("5.2 Frontend: Web Application (React 18 + Vite)"),
      para("The web frontend is implemented as a Single Page Application (SPA) using React 18 with the Vite build tool. Navigation is handled via React Router v6 with protected routes that check for a valid JWT token in local storage before rendering authenticated pages."),
      spacer(),
      para("Key Pages and Features:"),
      bullet("WelcomePage: The landing page features an animated gradient hero section, a live statistics grid, and a feature card row. Framer Motion provides entry animations on scroll."),
      bullet("SubmitPage: Provides a text area and URL input for article submission. On submission, the page calls the /articles/analyze endpoint and displays the verdict card with confidence score and keyword signals."),
      bullet("DashboardPage: Renders four animated stat counters, a Recharts bar chart comparing real vs fake article counts, a Leaflet map with CircleMarker dots for each source domain, and a sortable source list."),
      bullet("HistoryPage: Displays the authenticated user's past article submissions with classification results and timestamps in a chronological list."),
      bullet("PressNewsPage: Lists all press posts with optional language filtering. Posts include title, content preview, author channel name, and creation date. Press users see edit and delete controls on their own posts."),
      bullet("AuthPage: Combined login and registration form with a role toggle for enabling press accounts and entering a channel name."),
      bullet("ChatWidget: A floating chat button with a pulse animation that opens an animated panel with bilingual AI assistant capabilities, supporting rolling 10-message conversation history."),

      h2("5.3 Frontend: Mobile Application (React Native + Expo)"),
      para("The mobile application is built with Expo SDK 56 and uses expo-router for file-based tab navigation. The four primary tabs are Home (news feed), Analyze, Dashboard, and Publish (for press users)."),
      spacer(),
      para("Key Features:"),
      bullet("Bilingual News Feed: The Home tab fetches headlines from the NewsAPI and allows users to toggle between English (EN) and Arabic (AR) feeds with instant switching."),
      bullet("Article Analysis: The Analyze tab provides a text input and URL field for submitting articles to the backend classifier, displaying the same verdict card as the web application."),
      bullet("Press Post Management: Press users can view their own posts and tap to edit or delete them. The edit form pre-fills existing content for convenient updates."),
      bullet("Image Upload: The Publish tab supports camera and gallery image selection via expo-image-picker for press post cover images."),
      bullet("AI Chat Widget: A floating action button opens a modal with an animated FlatList chat interface. Messages are rendered with distinct styling for user and AI roles."),

      h2("5.4 Backend: API Layer (FastAPI + Python)"),
      para("The backend is structured as a FastAPI application with a modular router architecture. Each functional domain has its own router module registered in main.py. The backend binds to 0.0.0.0:8000 to accept connections from both localhost and LAN-connected mobile devices."),
      spacer(),
      para("API Endpoints:"),
      bullet("POST /auth/register — Creates a new user. Accepts email, password, is_press flag, and optional channel_name. If is_press is true, generates a press_approval_token and dispatches an approval email to the administrator."),
      bullet("POST /auth/login — Validates credentials and returns a JWT access token with a 7-day expiry."),
      bullet("POST /articles/analyze — Runs the RoBERTa classifier on submitted text or URL content. Requires JWT authentication."),
      bullet("GET /articles/history — Returns the authenticated user's submission history ordered by creation date descending."),
      bullet("GET /analytics/summary — Returns aggregate statistics: total articles, fake count, real count, and fake percentage."),
      bullet("GET /analytics/sources — Returns all source records with domain, country, geographic coordinates, and trust score."),
      bullet("GET /press/posts — Returns all published press posts ordered by creation date descending."),
      bullet("POST /press/posts — Creates a new press post. Restricted to press users. Also triggers notification creation."),
      bullet("GET /press/posts/mine — Returns the authenticated press user's own posts."),
      bullet("PUT /press/posts/{id} — Updates a specific press post. Verifies that the requesting user owns the post."),
      bullet("DELETE /press/posts/{id} — Deletes notifications first, then the press post."),
      bullet("GET /admin/pending-press — Returns a list of all users with is_press = true and press_approved = false. Restricted to the administrator account."),
      bullet("POST /admin/approve-press/{id} — Sets press_approved = true and clears the approval token for the specified user. Restricted to the administrator account."),
      bullet("POST /admin/reject-press/{id} — Reverts the user to a regular account by setting is_press = false and clearing the approval token and channel name. Restricted to the administrator account."),
      bullet("GET /admin/approve-press?token={token} — One-click approval link embedded in the admin notification email. Validates the token, approves the press account, and returns an HTML confirmation page."),
      bullet("GET /admin/reject-press?token={token} — One-click rejection link embedded in the admin notification email. Validates the token, reverts the user to a regular account, and returns an HTML confirmation page."),
      bullet("GET /notifications — Returns all notifications for the authenticated user."),
      bullet("POST /chat/assist — Sends a user message and conversation history to the Gemini API. Returns the AI assistant response."),

      h2("5.5 Backend: AI Classification Service"),
      para("The RoBERTa-based classification model is loaded once at application startup using the Hugging Face transformers library pipeline function. This singleton pattern ensures that the 500MB+ model is loaded into memory only once, avoiding repeated disk I/O on every request. The classify_text function accepts raw text, detects its language, translates if necessary, truncates to 2000 characters, and passes the result to the pipeline with truncation=True and max_length=512. The model returns a label (REAL or FAKE) and a confidence score, which are then used to set the is_fake boolean and confidence_score float on the article record."),

      h2("5.6 Summary: Data Flow and Integration"),
      para("The end-to-end data flow of VeriSource operates as follows: A user submits an article on the web or mobile client. The client sends an authenticated HTTP POST request to /articles/analyze with the JWT in the Authorization Bearer header. The FastAPI backend validates the token, extracts the user identity, preprocesses the submitted content (scraping, OCR, or plain text), translates if needed, runs the RoBERTa classifier, and saves the result to PostgreSQL. The response is returned to the client, which renders the verdict card. In parallel, the source trust score is recalculated and the dashboard statistics are updated for all future dashboard loads."),

      // ── SOFTWARE FUNCTIONS PER ROLE (NEW SECTION) ────────────────────────────
      h2("5.7 Software Functions per Role"),
      para("The following subsections summarize the complete set of software functions available to each user role within the VeriSource platform. This role-based breakdown reflects both the frontend interface capabilities and the backend API endpoints accessible to each actor."),

      h3("5.7.1 For Regular Users"),
      bullet("Register a new account using email and password via the Auth page."),
      bullet("Log in and receive a JWT access token for authenticated API access."),
      bullet("Submit a news article as plain text for AI-based fake news classification."),
      bullet("Submit a URL for automatic web scraping and AI classification."),
      bullet("Submit an image for OCR text extraction and AI classification."),
      bullet("View the classification verdict including Real/Fake label, confidence score, and detected keyword signals."),
      bullet("Browse personal article submission history with timestamps and results."),
      bullet("Access the Dashboard showing aggregate statistics: total, fake, real article counts, and fake percentage."),
      bullet("View the Recharts bar chart comparing real vs fake article counts."),
      bullet("Explore the Global Source Map with geographic markers color-coded by source trust score."),
      bullet("View the source list with domain names, countries, trust score percentages, and article counts."),
      bullet("Read press posts published by verified press users on the Press News page."),
      bullet("Filter press posts by language (English or Arabic)."),
      bullet("Receive in-app notifications when a press user publishes a new post."),
      bullet("Interact with the Gemini-powered AI chat assistant on both web and mobile."),
      bullet("Browse real-time English and Arabic news headlines on the mobile app via the bilingual news feed."),

      h3("5.7.2 For Press Users"),
      para("Press users have all the capabilities of Regular Users plus the following exclusive functions:"),
      bullet("Publish a new press post with a title, content body, and optional cover image."),
      bullet("Select an image from the device camera or gallery for press post cover photos (mobile)."),
      bullet("View a personal list of all posts they have authored via the My Posts section."),
      bullet("Edit the title, content, or language tag of any post they authored."),
      bullet("Delete any post they authored — the system automatically removes all associated notifications before deleting the post."),
      bullet("Have their channel name displayed as the author attribution on all published posts."),
      bullet("Trigger automatic notification generation for all regular users upon publishing a new post."),

      h3("5.7.3 For Admins"),
      para("The admin role is planned for a future version of VeriSource. The capabilities listed below represent the intended admin functionality described in the system design."),
      para("In the anticipated implementation, admins will have the following platform-wide management capabilities:"),
      bullet("View a complete list of all registered user accounts including their role, press status, and registration date."),
      bullet("Promote a regular user to press status or revoke press status from an existing press user."),
      bullet("Deactivate or suspend user accounts that violate platform policies without permanently deleting their data."),
      bullet("View platform-wide analytics including all article submissions, fake detection rates by time period, and source trust score distributions."),
      bullet("Delete any press post from the platform regardless of authorship for content moderation purposes."),
      bullet("Access detailed audit logs of authentication events, article submissions, and press post publications."),
      bullet("Manage API key configuration and system environment settings."),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 6: Testing Results
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 6: Testing Results Samples and Overview"),
      para("This chapter presents a detailed overview of the functional testing results across the primary modules of the VeriSource platform. Each section includes a screenshot of the relevant interface accompanied by a component breakdown table and an explanation paragraph describing the functionality demonstrated."),

      h2("6.1 Welcome Page and Navigation"),
      spacer(),
      ...figureBlock(7, "Welcome Page (Web) - Hero section, stats cards, and feature grid", "image7.png", 580, 361),
      spacer(),
      para("Explanation: The Welcome Page serves as the main landing page of VeriSource and is the first screen a visitor sees upon opening the platform. It features an animated gradient hero section built with Framer Motion that displays the platform name and a descriptive tagline. Below the hero section, a live statistics grid shows real-time counts of total articles analyzed, confirmed fake articles, and confirmed real articles fetched from the /analytics/summary endpoint. A feature card row beneath the statistics highlights the three core platform capabilities: AI-powered fake news classification, source trust score tracking, and press user publication. The page is fully responsive and uses Tailwind CSS utility classes for consistent layout and styling across viewport sizes."),
      spacer(),
      ...flowTable("Welcome / Home Page Components", 7, [
        ["Hero Section", "Animated UI", "Framer Motion gradient banner with platform name, tagline, and call-to-action button"],
        ["Stats Grid", "Live Data", "Real-time counters for total, fake, and real articles fetched from /analytics/summary"],
        ["Feature Cards", "Static UI", "Three cards highlighting AI Classification, Source Trust, and Press Publication features"],
        ["Navigation Bar", "Routing", "React Router links to Submit, Dashboard, History, Press News, and Auth pages"],
        ["Language Toggle", "i18n", "Bilingual toggle switching between English and Arabic UI labels using LanguageContext"]
      ]),

      h2("6.2 Authentication Module"),
      spacer(),
      ...figureBlock(8, "Login and Register Screens - Web and Mobile", "image8.png", 480, 440),
      spacer(),
      para("Explanation: The Authentication Module provides a unified login and registration interface accessible from both the web browser and the mobile application. On the web, a single AuthPage component renders either the login or registration form based on a toggled tab selector. The registration form includes an optional Press Account switch that, when enabled, reveals a Channel Name input field for journalists setting up their publication identity. On successful login or registration, the system issues a JWT token with a 7-day expiry which is stored in the browser's local storage (web) or AsyncStorage (mobile) and used for all subsequent authenticated API requests. The mobile login screen follows native design patterns with appropriately sized touch targets and keyboard-aware layout adjustments."),
      spacer(),
      ...flowTable("Login and Register Flow", 8, [
        ["Login Form", "Auth UI", "Email and password fields with validation; submits POST /auth/login and stores JWT"],
        ["Register Form", "Auth UI", "Email, password, optional press toggle, and channel name; submits POST /auth/register"],
        ["Press Toggle", "Role Config", "Boolean switch that sets is_press=true and reveals the channel_name input field"],
        ["JWT Storage", "Security", "Token stored in localStorage (web) or AsyncStorage (mobile) for session persistence"],
        ["Error Handling", "UX", "Displays clear error messages for invalid credentials or duplicate email addresses"]
      ]),

      h2("6.3 Article Analysis Module"),
      spacer(),
      ...figureBlock(9, "Article Analysis Screen - Text and URL input with classification result", "image9.png", 480, 440),
      spacer(),
      para("Explanation: The Article Analysis screen is the core functional feature of VeriSource and the primary reason users visit the platform. The Submit page presents two input modes: a text area for pasting article content directly, and a URL field for submitting a web article link for automatic scraping. Upon clicking the Analyze button, the frontend sends an authenticated POST request to /articles/analyze. The backend processes the content through the preprocessing pipeline — scraping if URL, OCR if image, translating if non-English — and then passes the text to the RoBERTa classification model. The result is displayed as a verdict card showing the REAL or FAKE label in large colored text, the confidence percentage, and two expandable lists of detected fake signals and real signals (keywords from the article text that influenced the classification). The confidence score helps users understand the certainty of the model's prediction."),
      spacer(),
      ...flowTable("Article Analysis Flow", 9, [
        ["Text Input", "Submission", "Free-text area accepting pasted article content for direct classification"],
        ["URL Input", "Submission", "URL field triggering BeautifulSoup scraping of article text and title"],
        ["Verdict Card", "Result UI", "Displays REAL/FAKE label, confidence percentage, and keyword signal lists"],
        ["Fake Signals", "Explainability", "Keywords like 'shocking', 'conspiracy', 'exposed' found in the article text"],
        ["Real Signals", "Explainability", "Keywords like 'according to', 'research', 'official' found in the article text"],
        ["History Save", "Persistence", "Article saved to database and added to user's submission history"]
      ]),

      h2("6.4 Dashboard Page"),
      spacer(),
      ...figureBlock(10, "Dashboard Page - Stats counters, bar chart, and global source map", "image10.png", 580, 275),
      spacer(),
      para("Explanation: The Dashboard Page provides a comprehensive analytics overview of all articles analyzed through the VeriSource platform. It is structured in three vertical sections. The top section displays four animated stat counter cards built with Framer Motion: total articles submitted, number of fake articles detected, number of real articles confirmed, and the overall fake percentage. The middle section renders a Recharts bar chart with two bars comparing real versus fake article totals, using green and red color coding respectively. The bottom section shows the Global Source Map, a react-leaflet interactive map using CARTO light tile layers. Each news source domain that has been encountered through URL submissions is plotted as a CircleMarker at its inferred geographic coordinates: sources with a trust score above 50% appear as green dots, while lower-trust sources appear as red dots. Clicking or hovering over a marker shows a popup with the domain name, trust percentage, and total article count. Below the map, a sortable list of all sources provides the same information in tabular form."),
      spacer(),
      ...flowTable("Dashboard Components", 10, [
        ["Stat Counters", "Analytics", "Animated counters for total, fake, real articles and fake percentage from /analytics/summary"],
        ["Bar Chart", "Visualization", "Recharts BarChart comparing real vs fake article counts with green/red color coding"],
        ["Source Map", "Geo Analytics", "react-leaflet map with CARTO tiles; CircleMarkers color-coded by source trust score"],
        ["Map Popup", "Interaction", "Leaflet Popup showing domain, trust percentage, and article count on marker click"],
        ["Source List", "Data Table", "Scrollable list of all sources with domain, country, trust score, and article count"]
      ]),

      h2("6.5 Mobile Application Testing"),
      spacer(),
      ...figureBlock(11, "Mobile News Headlines Screen - EN/AR language toggle and article list", "image11.png", 350, 310),
      spacer(),
      para("Explanation: The Mobile News Headlines screen demonstrates the bilingual news aggregation feature of the VeriSource mobile application. The Home tab fetches real-time headlines from the NewsAPI and displays them in a scrollable FlatList with article title, source name, and publication timestamp. A language toggle at the top of the screen allows users to instantly switch between English (EN) and Arabic (AR) news feeds without reloading the application. The Arabic feed displays right-to-left text correctly using React Native's built-in RTL support. Each article card includes a clickable link that opens the full article in the device's default browser using Expo's Linking API."),
      spacer(),
      ...figureBlock(12, "Press Publish Screen - Post creation form, image upload, and My Posts list", "image12.png", 340, 737),
      spacer(),
      para("Explanation: The Press Publish screen is exclusive to authenticated press users and provides the full post management interface on the mobile application. The top section of the screen shows the post creation form with fields for title, content body, and an optional image attachment. The image selector uses expo-image-picker to access the device camera or photo gallery, encoding the selected image as a base64 string for upload. Below the form, the My Posts section lists all posts authored by the current press user, with edit and delete action buttons on each card. Tapping Edit pre-fills the creation form with the existing post content for convenient updates. Tapping Delete triggers a confirmation prompt and then sends a DELETE /press/posts/{id} request, which removes associated notifications before deleting the post record."),
      spacer(),
      ...flowTable("Mobile News Feed Components", 11, [
        ["Language Toggle", "i18n", "EN/AR switch that refetches headlines from NewsAPI in the selected language"],
        ["Article FlatList", "News Feed", "Scrollable list of real-time headlines with title, source, and timestamp"],
        ["Article Link", "Navigation", "Opens full article URL in device browser via Expo Linking.openURL"],
        ["RTL Support", "Accessibility", "Arabic content renders right-to-left using React Native's RTL layout system"]
      ]),
      ...flowTable("Press Publish Components", 12, [
        ["Post Form", "Publication", "Title, content, and language fields for creating new press posts"],
        ["Image Picker", "Media Upload", "expo-image-picker for camera/gallery access; image encoded as base64 for upload"],
        ["My Posts List", "Management", "Shows all authored posts with edit and delete action buttons"],
        ["Edit Flow", "CRUD", "Pre-fills form with existing post data; submits PUT /press/posts/{id} on save"],
        ["Delete Flow", "CRUD", "Deletes notifications first via backend FK constraint handling, then removes post"]
      ]),

      h2("6.6 AI Chat Assistant"),
      spacer(),
      ...figureBlock(13, "AI Chat Assistant (Web) - Floating chat button and expanded animated chat panel", "image13.jpeg", 340, 737),
      spacer(),
      para("Explanation: The AI Chat Assistant on the web platform is implemented as a floating widget that persists across all pages of the application. A circular button in the bottom-right corner of the screen displays a pulse animation to draw user attention. When clicked, the button expands into an animated side panel using Framer Motion spring animations. The chat panel displays a scrollable message history with distinct visual styling for user messages (right-aligned, blue background) and AI responses (left-aligned, gray background). The assistant is configured with a VeriSource-specific system prompt that instructs the Gemini model to act as a knowledgeable support agent for the platform, capable of explaining classification results, guiding users through features, and discussing media literacy topics. The rolling conversation history of the last 10 messages is maintained in local component state and sent with each new request to provide conversational context."),
      spacer(),
      ...figureBlock(14, "AI Chat Assistant (Mobile) - Slide-up modal chat interface", "image14.jpeg", 500, 305),
      spacer(),
      para("Explanation: The AI Chat Assistant on the mobile application is accessible via a floating action button positioned at the bottom-right of every screen. Tapping the button triggers a slide-up modal animation revealing the full-screen chat interface. The mobile chat uses a React Native FlatList for message rendering, with inverted layout so the most recent messages appear at the bottom. User messages and AI responses are styled with distinct colors and alignment. The same Gemini API endpoint (/chat/assist) is used as on the web, ensuring consistent AI behavior across platforms. The mobile implementation handles keyboard avoidance automatically through React Native's KeyboardAvoidingView, keeping the input field visible when the device keyboard is open."),
      spacer(),
      ...flowTable("AI Chat Assistant Components", 13, [
        ["Chat Button", "Trigger", "Floating button with pulse animation; expands to chat panel (web) or modal (mobile)"],
        ["Message List", "Chat UI", "Scrollable history of user/AI messages with distinct visual styling per role"],
        ["Input Field", "Interaction", "Text input with send button; clears on submission"],
        ["Gemini API", "AI Backend", "POST /chat/assist sends message + rolling history; returns Gemini response"],
        ["System Prompt", "AI Config", "VeriSource-specific instructions configure Gemini as a platform support agent"],
        ["History Window", "Context", "Last 10 messages sent with each request to maintain conversational coherence"]
      ]),

      h2("6.7 Press News Page and Mobile Home"),
      spacer(),
      ...figureBlock(15, "Press News Page (Web) - Published articles with language filter tabs", "image15.png", 340, 737),
      spacer(),
      para("Explanation: The Press News Page displays all articles published by verified press users in reverse chronological order. Users can filter the displayed posts by language using tab selectors for English and Arabic content at the top of the page. Each press post card shows the post title, a content preview, the author's channel name as the attribution label, and the publication date. Press users who authored a post see Edit and Delete action buttons on their own cards, while regular users see only the read-only view. When a press post is published, the system generates individual notification records for all regular users, which they can access via the notification bell icon in the navigation bar. This design creates a lightweight community news channel experience within the credibility-focused platform."),
      spacer(),
      ...figureBlock(16, "Mobile Home Screen with Arabic News Feed Active", "image16.jpeg", 580, 281),
      spacer(),
      para("Explanation: The Mobile Home screen shown here demonstrates the Arabic news feed mode with the language toggle set to AR. The NewsAPI returns Arabic-language headlines from major regional outlets, which are displayed in right-to-left text alignment. The screen illustrates the seamless bilingual switching capability — users tap the AR button and the feed immediately refetches and re-renders with Arabic content without any page reload. This feature addresses an underserved use case in existing fake news detection tools, most of which provide English-only interfaces. The consistent tab navigation bar at the bottom of the screen provides access to the Analyze, Dashboard, and Publish tabs regardless of the current language setting."),
      pageBreak(),

      // ═══════════════════════════════════════════════════════════════════════
      // CHAPTER 7: Conclusion
      // ═══════════════════════════════════════════════════════════════════════
      h1("Chapter 7: General Conclusion and Future Recommendations"),
      para("In conclusion, VeriSource successfully delivers a comprehensive and technically rigorous solution to the challenge of automated fake news detection. The platform combines a state-of-the-art RoBERTa transformer model with an accessible, multi-platform interface that serves both casual users and verified press professionals. By providing instant, AI-powered classification of news articles submitted as text, URLs, or images, VeriSource makes credibility assessment available to anyone with a web browser or a smartphone."),
      spacer(),
      para("VeriSource addresses a real-world problem with measurable impact: by providing instant, confidence-scored classification results alongside keyword-level explainability, the platform helps users develop critical reading skills rather than simply accepting verdicts at face value. The source trust scoring system, backed by weighted harmonic averaging and geographic metadata, adds a domain-level perspective that complements the article-level analysis."),
      spacer(),
      para("The bilingual support for Arabic and English content, the global source trust map, and the Gemini-powered AI assistant collectively position VeriSource as a platform designed for a globally diverse user base. The press publication workflow bridges the gap between automated detection and human-authored credible content, creating a community ecosystem within the platform."),
      spacer(),
      para("VeriSource establishes a solid technical foundation from which the platform can continue to grow. The modular backend architecture, stateless JWT authentication, and containerized PostgreSQL database all support future scaling and feature expansion without requiring fundamental redesign."),

      h2("7.1 Future Work"),
      para("Several enhancements and extensions are planned for future iterations of VeriSource:"),
      spacer(),
      para("1. Real-Time Collaborative Fact-Checking", { bold: true }),
      para("Extending the platform to allow community-driven fact-checking annotations where registered users can flag potentially misleading articles and contribute to a crowdsourced credibility rating layer."),
      spacer(),
      para("2. Multilingual Classification", { bold: true }),
      para("Fine-tuning a multilingual BERT variant (mBERT or XLM-RoBERTa) to support direct Arabic-language article classification, eliminating the current dependency on translation as a preprocessing step."),
      spacer(),
      para("3. Browser Extension", { bold: true }),
      para("A Chrome and Firefox extension that highlights potentially fake news articles inline while browsing, powered by the same backend classification API, providing real-time credibility indicators without leaving the current page."),
      spacer(),
      para("4. Social Media Integration", { bold: true }),
      para("Enabling users to submit content directly from Twitter/X, Facebook, and WhatsApp shares for instant classification, using platform-specific share APIs and deep links."),
      spacer(),
      para("5. Admin Dashboard", { bold: true }),
      para("Implementing a full admin panel with user management capabilities — including press role promotion, account suspension, and platform-wide analytics — corresponding to the admin role described in Chapter 5."),
      spacer(),
      para("6. Push Notifications", { bold: true }),
      para("Implementing push notification delivery for the mobile application using Expo Notifications so users receive alerts about new press posts even when the application is not open."),
      spacer(),
      para("7. Model Explainability", { bold: true }),
      para("Adding attention visualization and SHAP explainability features that highlight the specific words or phrases in an article that most influenced the classification decision, providing deeper transparency into the AI's reasoning."),
      pageBreak(),

      // ── REFERENCES ───────────────────────────────────────────────────────────
      h1("References"),
      para("Research References:", { bold: true }),
      bullet("[1] Vosoughi, S., Roy, D., & Aral, S. (2018). The spread of true and false news online. Science, 359(6380), 1146-1151."),
      bullet("[2] Devlin, J., Chang, M. W., Lee, K., & Toutanova, K. (2019). BERT: Pre-training of deep bidirectional transformers for language understanding. NAACL-HLT 2019."),
      bullet("[3] Eisenman, B. (2016). Learning React Native: Building Native Mobile Apps with JavaScript. O'Reilly Media."),
      bullet("[4] Ramirez, S. (2021). FastAPI: Modern, Fast Web APIs with Python 3.6+. FastAPI Documentation."),
      bullet("[5] Google DeepMind. (2024). Gemini: A Family of Highly Capable Multimodal Models. Technical Report."),
      spacer(),
      para("Technical References:", { bold: true }),
      bullet("[6] React Documentation. https://react.dev/"),
      bullet("[7] Expo Documentation (SDK 56). https://docs.expo.dev/"),
      bullet("[8] FastAPI Documentation. https://fastapi.tiangolo.com/"),
      bullet("[9] SQLAlchemy Documentation. https://docs.sqlalchemy.org/"),
      bullet("[10] Hugging Face Transformers Documentation. https://huggingface.co/docs/transformers/"),
      bullet("[11] Leaflet.js Documentation. https://leafletjs.com/"),
      bullet("[12] Tailwind CSS Documentation. https://tailwindcss.com/docs/"),
      bullet("[13] Framer Motion Documentation. https://www.framer.com/motion/"),
      bullet("[14] Recharts Documentation. https://recharts.org/"),
      bullet("[15] NewsAPI Documentation. https://newsapi.org/docs/"),

    ] // end children
  }] // end sections
});

// ── Output ────────────────────────────────────────────────────────────────────
const OUT = path.join(__dirname, "VeriSource_Project_Report.docx");
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("Done:", OUT);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
