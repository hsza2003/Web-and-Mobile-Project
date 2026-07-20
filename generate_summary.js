// VeriSource Project Summary Document Generator
// Run: node generate_summary.js

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer, PageBreak,
} = require('docx');
const fs = require('fs');

// ─── Helpers ────────────────────────────────────────────────────────────────

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 100, bottom: 100, left: 140, right: 140 };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 36, color: "1F3864" })],
    spacing: { before: 400, after: 200 },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 28, color: "2E5596" })],
    spacing: { before: 300, after: 160 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, size: 24, color: "2E5596" })],
    spacing: { before: 240, after: 120 },
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold: opts.bold || false, size: 22 })],
    spacing: { before: 80, after: 80 },
    alignment: AlignmentType.JUSTIFIED,
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 22 })],
    spacing: { before: 40, after: 40 },
  });
}

function spacer() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 60, after: 60 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function twoColTable(leftHeader, leftContent, rightHeader, rightContent) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders, margins: cellMargins,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: "1F3864", type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [new TextRun({ text: leftHeader, bold: true, size: 22, color: "FFFFFF" })] })],
          }),
          new TableCell({
            borders, margins: cellMargins,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: "1F3864", type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [new TextRun({ text: rightHeader, bold: true, size: 22, color: "FFFFFF" })] })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            borders, margins: cellMargins,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: "EBF0FA", type: ShadingType.CLEAR },
            children: leftContent.map(t => new Paragraph({ children: [new TextRun({ text: t, size: 20 })], spacing: { before: 30, after: 30 } })),
          }),
          new TableCell({
            borders, margins: cellMargins,
            width: { size: 4680, type: WidthType.DXA },
            shading: { fill: "EBF0FA", type: ShadingType.CLEAR },
            children: rightContent.map(t => new Paragraph({ children: [new TextRun({ text: t, size: 20 })], spacing: { before: 30, after: 30 } })),
          }),
        ],
      }),
    ],
  });
}

function schemaTable(tableName, rows) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders, margins: cellMargins,
        width: { size: 2800, type: WidthType.DXA },
        shading: { fill: "1F3864", type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: "Column", bold: true, size: 20, color: "FFFFFF" })] })],
      }),
      new TableCell({
        borders, margins: cellMargins,
        width: { size: 6560, type: WidthType.DXA },
        shading: { fill: "1F3864", type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: "Type & Description", bold: true, size: 20, color: "FFFFFF" })] })],
      }),
    ],
  });

  const dataRows = rows.map(([col, desc], i) =>
    new TableRow({
      children: [
        new TableCell({
          borders, margins: cellMargins,
          width: { size: 2800, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? "EBF0FA" : "F5F8FF", type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: col, bold: true, size: 19, color: "1F3864" })] })],
        }),
        new TableCell({
          borders, margins: cellMargins,
          width: { size: 6560, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? "EBF0FA" : "F5F8FF", type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: desc, size: 19 })] })],
        }),
      ],
    })
  );

  return [
    new Paragraph({
      children: [new TextRun({ text: `Table: ${tableName}`, bold: true, size: 22, color: "1F3864" })],
      spacing: { before: 200, after: 80 },
    }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2800, 6560],
      rows: [headerRow, ...dataRows],
    }),
    spacer(),
  ];
}

// ─── Document Content ────────────────────────────────────────────────────────

const children = [

  // ── COVER ──────────────────────────────────────────────────────────────────
  new Paragraph({
    children: [new TextRun({ text: "VeriSource", bold: true, size: 72, color: "1F3864" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 800, after: 200 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "AI-Powered Fake News Detection Platform", size: 36, color: "2E5596" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Comprehensive Project Summary", bold: true, size: 28, color: "555555" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
  }),
  new Paragraph({
    children: [new TextRun({ text: "Islamic University of Lebanon  |  2025", size: 22, color: "888888" })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 800 },
  }),
  pageBreak(),

  // ── 1. PROJECT OVERVIEW ────────────────────────────────────────────────────
  h1("1. Project Overview"),
  para("VeriSource is a full-stack, AI-powered fake news detection platform designed to help users verify the credibility of news articles in real time. The system accepts article content in three forms — plain text, URL, or image — and returns an AI-generated verdict (Real or Fake) with a confidence score, detected keyword signals, and a plain-language explanation of why the classifier made that decision."),
  para("The platform is accessible through two interfaces: a React 18 web application and a React Native / Expo mobile app. Both interfaces share a single FastAPI backend that hosts the RoBERTa classification model, manages user accounts, serves press-published news, and provides an AI chat assistant powered by Google Gemini."),
  spacer(),
  para("Key capabilities:", { bold: true }),
  bullet("Real-time fake news classification using a fine-tuned RoBERTa transformer model"),
  bullet("Support for text, URL (with automatic web scraping), and image (with OCR) submission"),
  bullet("Bilingual platform: English and Arabic news feeds"),
  bullet("Press user role: registered journalists can publish news posts to all platform users"),
  bullet("Press approval workflow: admin email approval required before press features unlock"),
  bullet("Dashboard analytics with source trust scores and a global interactive map"),
  bullet("AI chat assistant powered by Google Gemini for general and news-related questions"),
  bullet("Notification system: users receive in-app alerts when new press posts are published"),
  spacer(),

  // ── 2. SYSTEM ARCHITECTURE ─────────────────────────────────────────────────
  h1("2. System Architecture"),
  para("VeriSource follows a three-tier architecture: a presentation layer (web and mobile frontends), an application layer (FastAPI backend with ML services), and a data layer (PostgreSQL). All communication between the frontend and backend uses RESTful HTTP with JSON. Authentication is stateless using JWT RS256 tokens."),
  spacer(),
  h2("2.1 Architecture Overview"),
  twoColTable(
    "Layer", ["Presentation", "Application", "Data"],
    "Components", [
      "Web: React 18 + Vite (SPA, Tailwind CSS, Framer Motion)  |  Mobile: React Native + Expo SDK 56",
      "FastAPI (Python 3.11) — REST API, JWT auth, business logic  |  RoBERTa ML classifier  |  Gemini API (AI chat)",
      "PostgreSQL (Docker) — 5 normalized tables  |  SQLAlchemy ORM",
    ]
  ),
  spacer(),

  h2("2.2 Request Flow"),
  para("The following describes the end-to-end request flow for an article submission:"),
  bullet("User enters text or a URL on the web or mobile interface."),
  bullet("The client sends POST /articles/submit/text (or /url, /image) with a JWT Bearer token."),
  bullet("FastAPI validates the JWT, identifies the user, and routes to the analysis service."),
  bullet("The analysis service preprocesses the content: scrapes URL if needed, runs OCR on image, detects language, translates non-English text."),
  bullet("The preprocessed text is passed to the RoBERTa pipeline with top_k=None to retrieve both FAKE and TRUE scores."),
  bullet("The higher-scoring label determines the verdict; keyword signals are extracted from the text."),
  bullet("The result is stored in PostgreSQL (articles table), source trust score is updated, and the response is returned to the client."),
  bullet("The client renders the verdict card with the label, confidence percentage, keyword signals, and explanation."),
  spacer(),
  pageBreak(),

  // ── 3. BACKEND ─────────────────────────────────────────────────────────────
  h1("3. Backend — FastAPI + Python"),

  h2("3.1 Project Structure"),
  para("The backend is organized as a FastAPI application with the following directory structure:"),
  spacer(),
  new Paragraph({
    children: [new TextRun({
      text:
        "backend/\n" +
        "  app/\n" +
        "    api/        — Route modules: auth, articles, press, admin, chat, news, notifications\n" +
        "    ml/         — classifier.py: RoBERTa inference logic\n" +
        "    models/     — SQLAlchemy ORM model classes\n" +
        "    schemas/    — Pydantic v2 request/response models\n" +
        "    services/   — email.py, news.py (NewsAPI + RSS)\n" +
        "    database.py — SQLAlchemy engine & session factory\n" +
        "    main.py     — App entry, router registration, CORS config",
      size: 18, font: "Courier New",
    })],
    spacing: { before: 80, after: 80 },
  }),
  spacer(),

  h2("3.2 Authentication"),
  para("Authentication uses JWT tokens signed with HS256. Tokens carry the user's ID and email, and expire after 7 days. The get_current_user() FastAPI dependency validates the Bearer token on every protected route. Passwords are hashed with bcrypt via passlib. Role enforcement is done at the API layer: press-only routes call a require_press() dependency that checks the is_press and press_approved flags on the authenticated user."),

  h2("3.3 AI Classification Service"),
  para("The RoBERTa model (hamzab/roberta-fake-news-classification) is loaded once at application startup using the Hugging Face transformers pipeline function. This singleton avoids repeated 500 MB+ model initialization. The classify_text function:"),
  bullet("Detects the article language; translates to English if needed"),
  bullet("Truncates text to 2,000 characters before passing to the model"),
  bullet("Calls the pipeline with top_k=None to retrieve both FAKE and TRUE label scores"),
  bullet("Selects the label with the higher score as the verdict"),
  bullet("Scans the text for a curated list of fake-news signal keywords (e.g., 'breaking', 'shocking', 'cover-up') and credibility keywords (e.g., 'according to', 'study confirms', 'peer-reviewed')"),
  bullet("Generates a plain-language explanation string based on the detected signals and confidence level"),
  bullet("Returns is_fake boolean, confidence_score float, fake_signals list, real_signals list, and explanations list"),

  h2("3.4 Press Approval Workflow"),
  para("When a user registers with is_press=true, the backend:"),
  bullet("Generates a UUID token and stores it in the press_approval_token column"),
  bullet("Sets press_approved=false — press features are locked until approval"),
  bullet("Dispatches an HTML email to the admin (zaarourhassan4@gmail.com) via Gmail SMTP"),
  bullet("The email contains the applicant's name, email, channel name, and two action buttons: green Approve and red Reject"),
  bullet("Approve link: GET /admin/approve-press?token= — sets press_approved=true, clears token"),
  bullet("Reject link: GET /admin/reject-press?token= — sets is_press=false, clears token and channel_name"),
  bullet("The in-app Admin Panel at /admin offers the same approve/reject actions for cases where the email link cannot be used"),

  h2("3.5 News Aggregation Service"),
  para("The backend serves English news via the NewsAPI REST API. Arabic news is fetched in parallel from five RSS feeds (RT Arabic, BBC Arabic, DW Arabic, France24 Arabic, Sky News Arabia) using ThreadPoolExecutor, with a NewsAPI Arabic fallback if fewer than 3 articles are returned. Both feeds support 13 regional filters (Lebanon, Syria, USA, Palestine, etc.) with bilingual keyword matching."),
  spacer(),
  pageBreak(),

  // ── 4. FRONTEND WEB ────────────────────────────────────────────────────────
  h1("4. Frontend — React 18 Web Application"),

  h2("4.1 Technology Stack"),
  bullet("React 18 with Vite — SPA build tool with hot module replacement"),
  bullet("React Router v6 — client-side routing with protected route wrappers"),
  bullet("Tailwind CSS v4 — utility-first responsive styling"),
  bullet("Framer Motion — animations: page transitions, hover effects, animated counters"),
  bullet("Recharts — bar chart for real vs fake article comparison on dashboard"),
  bullet("react-leaflet + CARTO tiles — interactive global source map"),
  bullet("Axios — HTTP client configured with VITE_API_URL (defaults to http://localhost:8000)"),

  h2("4.2 Pages and Features"),
  h3("WelcomePage"),
  para("The public landing page with animated hero section, live stats grid (total articles, fake rate, sources tracked), and a feature card row explaining VeriSource's capabilities. Bilingual news feed with category filters (Politics, Technology, Health, Science, Sports, Business, Entertainment). Supports region filter with 13 geographic options. Category keywords include both English and Arabic script terms for accurate Arabic feed filtering."),

  h3("AuthPage"),
  para("Combined login and registration form. Registration includes an optional 'I am a Press User' toggle that reveals a channel name field. On submit, an account is created with is_press=true; the press features remain locked until the admin approves the account."),

  h3("SubmitPage"),
  para("Article submission interface with three input modes: Text (textarea), URL (input field), Image (file upload). On submission, displays a verdict card with: Real/Fake badge, confidence percentage, fake signal keywords, real signal keywords, and an explanation paragraph."),

  h3("DashboardPage"),
  para("Analytics dashboard with four animated counters (total articles, fake count, real count, fake percentage), a Recharts bar chart, a Leaflet map with CircleMarker pins for each tracked source domain, and a sortable source trust table."),

  h3("HistoryPage"),
  para("Paginated list of the authenticated user's article submissions, ordered newest first, showing verdict badge, confidence score, submission method, and timestamp."),

  h3("PressNewsPage"),
  para("Lists all published press posts with language toggle (EN/AR). Regular users read posts. Press users also see Edit and Delete controls on their own posts. Post creation/editing is done via a modal form."),

  h3("AdminPage (route: /admin)"),
  para("Accessible only to the administrator account (zaarourhassan4@gmail.com). Displays pending press approval requests with applicant name, email, and channel name. Approve and Reject buttons call the respective admin API endpoints."),

  h3("ChatWidget"),
  para("Floating chat button (bottom-right) with pulse animation. Opens a slide-up panel with a 10-message rolling conversation history sent to the Gemini API for contextual responses. The system prompt allows the assistant to answer any question, not limited to VeriSource topics."),
  spacer(),
  pageBreak(),

  // ── 5. MOBILE ──────────────────────────────────────────────────────────────
  h1("5. Mobile Application — React Native + Expo"),

  h2("5.1 Technology Stack"),
  bullet("React Native with Expo SDK 56"),
  bullet("expo-router — file-based tab navigation (Home, Analyze, Dashboard, Press/Publish, Chat)"),
  bullet("expo-image-picker — camera and gallery access for press post images"),
  bullet("AsyncStorage — stores JWT token and user email locally"),
  bullet("Axios — HTTP client pointing to the backend IP (BASE_URL in services/api.js)"),
  bullet("NewsAPI — fetches real-time English and Arabic headlines"),

  h2("5.2 Key Screens and Features"),
  h3("Home Tab — Bilingual News Feed"),
  para("Fetches top headlines via GET /news/headlines (English) or GET /news/headlines/ar (Arabic). Language toggle allows instant EN/AR switching. Each article card shows title, source, date, and a thumbnail image."),

  h3("Analyze Tab — Article Submission"),
  para("Text input and URL field for submitting articles to the backend classifier. Displays the same verdict card as the web app: Real/Fake label, confidence score, keyword signals, and explanation."),

  h3("Dashboard Tab — Analytics"),
  para("Displays aggregate statistics (total articles, fake count, real count) and a source trust list fetched from the backend."),

  h3("Press / Publish Tab"),
  para("Regular users browse press posts from all press publishers. Press users see their own posts and can edit or delete them. The Publish tab provides a form with title, content, language selector, and optional cover image via expo-image-picker."),

  h3("Chat Widget"),
  para("A floating action button (FAB) opens a full-screen modal with an animated FlatList chat interface. User and AI messages are rendered with distinct bubble styling. Supports any question via the Gemini-powered backend endpoint."),

  h2("5.3 Connectivity Note"),
  para("The mobile app communicates with the backend over the local network using a hardcoded LAN IP in verisource-mobile/services/api.js. When the LAN IP changes, BASE_URL must be updated manually. For phone testing on a different WiFi network, use npx expo start --tunnel (requires @expo/ngrok) which creates a stable public tunnel URL."),
  spacer(),
  pageBreak(),

  // ── 6. DIAGRAMS ────────────────────────────────────────────────────────────
  h1("6. System Diagrams"),
  para("The following section describes all diagrams included in the VeriSource technical report. Each diagram is described in detail to support the project presentation."),

  h2("6.1 Use Case Diagram (Figure 1)"),
  para("The UML Use Case Diagram shows the interactions between the platform and its two primary actors:"),
  bullet("Regular User: submits articles (text/URL/image), views results, browses history, uses dashboard, reads press news, uses AI chat assistant."),
  bullet("Press User: extends Regular User — additionally creates, edits, and deletes press posts; posts trigger notifications to all regular users."),
  para("The diagram highlights the extension relationship: the Press User role extends the Regular User role rather than being a separate system actor, meaning every press user retains all regular user capabilities."),

  h2("6.2 Data Flow Diagram Level 0 — Context Diagram (Figure 2)"),
  para("The DFD Level 0 (Context Diagram) represents VeriSource as a single process and shows its data exchanges with four external entities:"),
  bullet("Regular User → submits article content; receives verdict, history, dashboard data, press posts, AI chat responses."),
  bullet("Press User → submits/manages press posts; receives confirmation and post management data."),
  bullet("Google Gemini API → receives conversation context; returns AI-generated responses."),
  bullet("NewsAPI → queried by the mobile frontend; returns English and Arabic headlines."),
  para("This diagram defines the system boundary and provides the highest-level view of VeriSource's data exchanges."),

  h2("6.3 Data Flow Diagram Level 1 — Internal Processes (Figure 3)"),
  para("The DFD Level 1 decomposes VeriSource into six internal processes:"),
  bullet("1. User Authentication and Access Control — registration, login, JWT issuance, press approval token management."),
  bullet("2. Article Classification Service — preprocessing (scraping, OCR, translation), RoBERTa inference, result storage."),
  bullet("3. Press Publication Module — press post CRUD, ownership verification, notification generation."),
  bullet("4. Analytics and Reporting — aggregates statistics, geocodes source domains, serves dashboard data."),
  bullet("5. AI Chat Module — builds Gemini prompts, manages conversation history, returns AI responses."),
  bullet("6. News Aggregation Service — fetches English headlines from NewsAPI and Arabic news from RSS feeds."),
  para("Data flows to and from five persistent PostgreSQL tables: users, articles, sources, press_posts, notifications."),

  h2("6.4 User Sequence Diagram (Figure 4)"),
  para("The User Sequence Diagram shows the chronological flow for a regular user submitting an article:"),
  bullet("User → Client: opens the Submit page."),
  bullet("Client → Backend: POST /auth/login → receives JWT token."),
  bullet("User → Client: enters text or URL and clicks Submit."),
  bullet("Client → Backend: POST /articles/submit/text (or /url) with JWT in Authorization header."),
  bullet("Backend: preprocesses content, calls RoBERTa classifier, stores result in articles table, updates source trust score."),
  bullet("Backend → Client: returns classification result (is_fake, confidence_score, fake_signals, real_signals, explanations)."),
  bullet("Client → User: renders verdict card with full result."),
  para("The diagram illustrates the stateless nature of the API — every request carries the JWT — and the sequential preprocessing pipeline within the backend."),

  h2("6.5 Press User Sequence Diagram (Figure 5)"),
  para("The Press User Sequence Diagram illustrates the press post publication flow:"),
  bullet("Press User → Client: opens Publish tab, fills form, submits."),
  bullet("Client → Backend: POST /press/posts with JWT — backend verifies is_press flag."),
  bullet("Backend: saves post to press_posts table, iterates all users to create notification records."),
  bullet("Client: fetches updated press post list and notification count."),
  bullet("Press User → Client: selects a post to delete."),
  bullet("Client → Backend: DELETE /press/posts/{id} — backend first deletes all associated notifications, then deletes the post (foreign key constraint)."),

  h2("6.6 Entity Relationship Diagram (Figure 6)"),
  para("The ERD illustrates the complete logical data structure with all five tables, their attributes, data types, primary keys, and foreign key relationships. Key relationships:"),
  bullet("users → articles (One-to-Many): a user can submit many articles."),
  bullet("users → press_posts (One-to-Many): a press user can author many posts."),
  bullet("articles → sources (Many-to-One): many articles from the same domain reference one source record."),
  bullet("press_posts → notifications (One-to-Many): each press post generates one notification per user."),
  spacer(),
  pageBreak(),

  // ── 7. DATABASE SCHEMA ─────────────────────────────────────────────────────
  h1("7. Database Schema"),
  para("VeriSource uses PostgreSQL with five normalized tables. The schema follows third normal form to eliminate redundancy and enforce referential integrity."),
  spacer(),

  ...schemaTable("users", [
    ["id", "INTEGER — Primary Key, auto-increment. Uniquely identifies each user account."],
    ["email", "VARCHAR — Unique, indexed. The user's login email address."],
    ["username", "VARCHAR — Unique, indexed. Display name for the user."],
    ["hashed_password", "VARCHAR — bcrypt hash. Plain-text password is never stored."],
    ["is_active", "BOOLEAN — Default true. Allows soft-disabling accounts."],
    ["is_press", "BOOLEAN — Default false. Grants press publication features when true."],
    ["press_approved", "BOOLEAN — Default false. Press features locked until admin approves."],
    ["press_approval_token", "VARCHAR — Nullable UUID. Single-use token for email approval links."],
    ["channel_name", "VARCHAR — Nullable. Display name for press users' publication channel."],
    ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set by PostgreSQL on insert."],
  ]),

  ...schemaTable("articles", [
    ["id", "INTEGER — Primary Key, auto-increment."],
    ["title", "VARCHAR — Nullable. Article title extracted from scraping or provided by user."],
    ["content", "TEXT — Required. Full article text used for classification."],
    ["url", "VARCHAR — Nullable. Source URL for URL-based submissions."],
    ["input_type", "VARCHAR — One of: 'url', 'text', 'image'. Tracks submission method."],
    ["is_fake", "BOOLEAN — Classification result from the RoBERTa model."],
    ["confidence_score", "FLOAT — Model confidence between 0.0 and 1.0."],
    ["fake_signals", "JSON — Array of suspicious keywords detected in the article text."],
    ["real_signals", "JSON — Array of credibility keywords detected in the article text."],
    ["source_id", "INTEGER FK → sources.id — Nullable. Linked source domain for URL submissions."],
    ["submitted_by", "INTEGER FK → users.id — Nullable. Allows anonymous submissions."],
    ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on insert."],
  ]),

  ...schemaTable("sources", [
    ["id", "INTEGER — Primary Key, auto-increment."],
    ["name", "VARCHAR — Unique. Display name of the news source."],
    ["domain", "VARCHAR — Unique. Domain string (e.g. bbc.com) from submitted URLs."],
    ["country", "VARCHAR — Nullable. Country of origin for geographic map placement."],
    ["trust_score", "FLOAT — 0.0–1.0. Recalculated with weighted harmonic average after each new article."],
    ["total_articles", "INTEGER — Count of all articles analyzed from this domain."],
    ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on first encounter of domain."],
    ["updated_at", "TIMESTAMP WITH TIME ZONE — Auto-updated on every trust score recalculation."],
  ]),

  ...schemaTable("press_posts", [
    ["id", "INTEGER — Primary Key, auto-increment."],
    ["title", "VARCHAR — Required. Headline of the press post."],
    ["content", "TEXT — Required. Full body text of the press post."],
    ["language", "VARCHAR — Default 'en'. Language tag: 'en' or 'ar'."],
    ["image_url", "VARCHAR — Nullable. URL path to an optional cover image."],
    ["author_id", "INTEGER FK → users.id — Required. The press user who authored the post."],
    ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on insert."],
  ]),

  ...schemaTable("notifications", [
    ["id", "INTEGER — Primary Key, auto-increment."],
    ["type", "VARCHAR — Default 'press_post'. Category of notification."],
    ["title", "VARCHAR — Required. Short notification headline."],
    ["message", "VARCHAR — Required. Full notification body text."],
    ["press_post_id", "INTEGER FK → press_posts.id — Nullable. Links notification to source press post."],
    ["created_at", "TIMESTAMP WITH TIME ZONE — Auto-set on insert."],
  ]),

  h2("7.1 Relationships Summary"),
  para("users → articles (1:N): cascade not applied; deleting a user does not delete their articles."),
  para("users → press_posts (1:N): non-press users cannot create posts (HTTP 403 enforced at API layer)."),
  para("articles → sources (N:1): sources table is auto-populated when new domains appear in URL submissions."),
  para("press_posts → notifications (1:N): on post deletion, all associated notifications are removed first to satisfy the FK constraint."),
  spacer(),
  pageBreak(),

  // ── 8. API ENDPOINTS ───────────────────────────────────────────────────────
  h1("8. API Endpoint Reference"),
  para("All endpoints are served from the FastAPI backend at http://localhost:8000. Authenticated endpoints require Authorization: Bearer <token> in the request header."),
  spacer(),

  h2("Authentication"),
  bullet("POST /auth/register — Create account (email, password, is_press, channel_name)"),
  bullet("POST /auth/login — Get JWT token (email, password)"),
  bullet("GET /auth/me — Get current authenticated user profile"),

  h2("Article Analysis"),
  bullet("POST /articles/submit/text — Classify plain text article"),
  bullet("POST /articles/submit/url — Classify article from URL (scrapes automatically)"),
  bullet("POST /articles/submit/image — Classify article from image (OCR)"),
  bullet("GET /articles/history — User's past submissions with results"),

  h2("Analytics"),
  bullet("GET /analytics/summary — Total, fake, real counts and fake percentage"),
  bullet("GET /analytics/sources — All source domains with trust scores and coordinates"),

  h2("Press"),
  bullet("GET /press/posts — All press posts (optional ?language=en|ar)"),
  bullet("POST /press/posts — Create press post (press user only)"),
  bullet("GET /press/posts/mine — Authenticated press user's own posts"),
  bullet("PUT /press/posts/{id} — Update press post (owner only)"),
  bullet("DELETE /press/posts/{id} — Delete press post and its notifications"),
  bullet("POST /press/upload-image — Upload cover image, returns image_url"),

  h2("Admin"),
  bullet("GET /admin/pending-press — List pending press approval requests (admin only)"),
  bullet("POST /admin/approve-press/{id} — Approve press user (admin only)"),
  bullet("POST /admin/reject-press/{id} — Reject press user (admin only)"),
  bullet("GET /admin/approve-press?token= — One-click approval from email link"),
  bullet("GET /admin/reject-press?token= — One-click rejection from email link"),

  h2("Notifications"),
  bullet("GET /notifications — All notifications for authenticated user"),

  h2("Chat"),
  bullet("POST /chat/assist — Send message + history to Gemini, receive AI response"),

  h2("News"),
  bullet("GET /news/headlines — English top headlines (optional ?region=all|lebanon|...)"),
  bullet("GET /news/headlines/ar — Arabic headlines from RSS feeds"),
  bullet("GET /news/regions — List of all supported regions with EN/AR labels"),
  spacer(),
  pageBreak(),

  // ── 9. TECH STACK SUMMARY ─────────────────────────────────────────────────
  h1("9. Technology Stack Summary"),
  spacer(),
  twoColTable(
    "Layer", ["Frontend (Web)", "Frontend (Mobile)", "Backend", "AI / ML", "Data & Infrastructure"],
    "Technologies", [
      "React 18 + Vite, React Router v6, Tailwind CSS v4, Framer Motion, Recharts, react-leaflet, Axios",
      "React Native + Expo SDK 56, expo-router, expo-image-picker, AsyncStorage, Axios",
      "Python 3.11 + FastAPI, SQLAlchemy ORM, Pydantic v2, python-jose (JWT), passlib/bcrypt, httpx, feedparser",
      "Hugging Face Transformers, hamzab/roberta-fake-news-classification, Google Gemini API, langdetect + googletrans",
      "PostgreSQL (Docker), Gmail SMTP (approval emails), NewsAPI (English), RSS feeds (Arabic)",
    ]
  ),
  spacer(),
];

// ─── Build & Write ────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22 } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Calibri", color: "1F3864" },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Calibri", color: "2E5596" },
        paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Calibri", color: "2E5596" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "VeriSource — Project Summary  |  Page ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888" }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('VeriSource_Summary.docx', buffer);
  console.log('Done: VeriSource_Summary.docx');
});
