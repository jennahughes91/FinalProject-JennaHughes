const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, Header, Footer, PageNumber, PageBreak,
} = require('/usr/local/lib/node_modules_global/lib/node_modules/docx');
const fs = require('fs');

// ─── Color palette ────────────────────────────────────────────────────────────
const BLUE_DARK   = "1F4E79";
const BLUE_MID    = "2E75B6";
const BLUE_LIGHT  = "D6E4F0";
const BLUE_HEADER = "BDD7EE";
const GRAY_ROW    = "F2F7FB";
const WHITE       = "FFFFFF";
const ACCENT      = "C6EFCE";
const ACCENT_TXT  = "276221";
const ORANGE_BG   = "FFF2CC";
const ORANGE_SIDE = "ED7D31";
const TEAL_BG     = "E2EFDA";
const TEAL_SIDE   = "375623";

// ─── Border helpers ───────────────────────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "BBCFE0" };
const noneBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const tableBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder,
                       insideH: thinBorder, insideV: thinBorder };
const noBorders    = { top: noneBorder, bottom: noneBorder, left: noneBorder, right: noneBorder,
                       insideH: noneBorder, insideV: noneBorder };

// ─── Paragraph helpers ────────────────────────────────────────────────────────
const h1 = text => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 120 },
  children: [new TextRun({ text, font: "Arial", size: 30, bold: true, color: BLUE_DARK })]
});
const h2 = text => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 80 },
  children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: BLUE_MID })]
});
const h3 = text => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 180, after: 60 },
  children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: BLUE_DARK })]
});
const body = (text, opts = {}) => new Paragraph({
  spacing: { before: 60, after: 80 },
  children: [new TextRun({ text, font: "Arial", size: 20, ...opts })]
});
const bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: `bullets-${level}`, level: 0 },
  spacing: { before: 40, after: 40 },
  children: [new TextRun({ text, font: "Arial", size: 20 })]
});
const spacer = () => new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun("")] });
const divider = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE_MID, space: 4 } },
  spacing: { before: 120, after: 120 },
  children: [new TextRun("")]
});

// ─── Callout box ─────────────────────────────────────────────────────────────
const callout = (label, text, sideColor = BLUE_MID, bgColor = BLUE_LIGHT) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [940, 8420],
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 940, type: WidthType.DXA }, borders: noBorders,
        shading: { fill: sideColor, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: label, font: "Arial", size: 17, bold: true, color: WHITE })] })]
      }),
      new TableCell({
        width: { size: 8420, type: WidthType.DXA }, borders: noBorders,
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [new Paragraph({
          children: [new TextRun({ text, font: "Arial", size: 19, italics: true, color: "333333" })]
        })]
      })
    ]})]
  });

// ─── Data table ──────────────────────────────────────────────────────────────
const dataTable = (headers, widths, rows) => {
  const totalWidth = widths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA }, borders: tableBorders,
      shading: { fill: BLUE_HEADER, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 19, bold: true, color: BLUE_DARK })] })]
    }))
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => {
      const isGreen = typeof cell === 'string' && cell.startsWith('✓');
      return new TableCell({
        width: { size: widths[ci], type: WidthType.DXA }, borders: tableBorders,
        shading: { fill: isGreen ? ACCENT : (ri % 2 === 0 ? WHITE : GRAY_ROW), type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({
          text: cell, font: "Arial", size: 19,
          color: isGreen ? ACCENT_TXT : "2C2C2C"
        })] })]
      });
    })
  }));
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA }, columnWidths: widths,
    borders: tableBorders,
    rows: [headerRow, ...dataRows]
  });
};

// ─── Step badge ───────────────────────────────────────────────────────────────
const stepBadge = (num, title, color = BLUE_DARK) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [700, 8660],
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 700, type: WidthType.DXA }, borders: tableBorders,
        shading: { fill: color, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 80, right: 80 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: String(num), font: "Arial", size: 22, bold: true, color: WHITE })] })]
      }),
      new TableCell({
        width: { size: 8660, type: WidthType.DXA }, borders: tableBorders,
        shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: BLUE_DARK })] })]
      })
    ]})]
  });

// ─── Two-column layout helper ─────────────────────────────────────────────────
const twoCol = (leftItems, rightItems, leftW = 4600, rightW = 4600) =>
  new Table({
    width: { size: leftW + rightW + 160, type: WidthType.DXA },
    columnWidths: [leftW, rightW],
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: leftW, type: WidthType.DXA }, borders: noBorders,
        margins: { top: 0, bottom: 0, left: 0, right: 80 },
        children: leftItems
      }),
      new TableCell({
        width: { size: rightW, type: WidthType.DXA }, borders: noBorders,
        margins: { top: 0, bottom: 0, left: 80, right: 0 },
        children: rightItems
      })
    ]})]
  });

// ─── Phase label ─────────────────────────────────────────────────────────────
const phaseBadge = (label, sub, color) =>
  new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    borders: { top: thinBorder, bottom: thinBorder, left: { style: BorderStyle.SINGLE, size: 12, color }, right: thinBorder },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      borders: { top: thinBorder, bottom: thinBorder, left: { style: BorderStyle.SINGLE, size: 12, color }, right: thinBorder },
      shading: { fill: "F8FBFF", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 200, right: 120 },
      children: [
        new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: 22, bold: true, color })] }),
        new Paragraph({ children: [new TextRun({ text: sub, font: "Arial", size: 19, color: "555555" })] })
      ]
    })] })]
  });

// ─── Document ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: { config: [
    { reference: "bullets-0",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 600, hanging: 300 } } } }] },
    { reference: "bullets-1",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 960, hanging: 300 } } } }] },
  ]},
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: BLUE_MID },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } }
    },
    headers: { default: new Header({ children: [new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE_MID, space: 4 } },
      children: [
        new TextRun({ text: "Sprint Backlog Prioritization System  |  System Design", font: "Arial", size: 18, color: "555555" }),
        new TextRun({ text: "     INTERNAL — CONFIDENTIAL", font: "Arial", size: 18, color: "AAAAAA" }),
      ]
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE_MID, space: 4 } },
      children: [
        new TextRun({ text: "Page ", font: "Arial", size: 18, color: "888888" }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "888888" }),
        new TextRun({ text: " of ", font: "Arial", size: 18, color: "888888" }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: "888888" }),
      ]
    })] }) },

    children: [

      // ══════════════════════════════════════════════════════════════════════
      // TITLE PAGE
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({
        spacing: { before: 960, after: 240 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Sprint Backlog Prioritization System", font: "Arial", size: 52, bold: true, color: BLUE_DARK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 120 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "System Design Document", font: "Arial", size: 32, color: BLUE_MID })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 480 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Internal Interactive Application  ·  Powered by Anthropic Claude  ·  April 2026", font: "Arial", size: 22, color: "888888", italics: true })]
      }),
      divider(),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 1. EXECUTIVE SUMMARY
      // ══════════════════════════════════════════════════════════════════════
      h1("1.  Executive Summary"),
      body("This document describes the system design for an internal interactive web application that automates development backlog prioritization. Users upload a backlog and sprint capacity file, and the system returns a live-ranked list of work items — with weak descriptions flagged and improvement suggestions attached."),
      spacer(),
      body("Critically, this is not a batch pipeline with a static output file. Priority rankings update in real time as users adjust business area and product team priority weights through the application interface. Named priority profiles can be saved and reloaded for future sprints."),
      spacer(),
      callout("CORE DESIGN PRINCIPLE", "AI runs once per file upload (description evaluation and suggestions). Ranking reruns instantly in the browser whenever the user changes a priority weight. These two phases are architecturally separate."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 2. ARCHITECTURE OVERVIEW
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("2.  Architecture Overview"),
      body("The system is organized into five layers. The most important architectural decision is the split between the AI Analysis Phase (runs once on upload) and the Interactive Ranking Phase (runs instantly in the browser on every priority change)."),
      spacer(),
      dataTable(
        ["Layer", "Name", "Responsibility", "Key Technology"],
        [800, 2000, 3700, 2860],
        [
          ["1", "Interactive UI", "File upload, live priority controls, ranked list, profile manager, export", "React + Tailwind CSS"],
          ["2", "Ingestion & Parsing", "Excel extraction, schema normalization, validation", "Python, pandas, openpyxl"],
          ["3", "AI Analysis Pipeline", "Description evaluation, suggestion generation — runs once per upload", "FastAPI, Claude API"],
          ["4", "Interactive Ranking Engine", "Formula-based re-scoring on every priority change — no LLM calls", "Python (API) or JavaScript (client-side)"],
          ["5", "Profile & State Store", "Persist priority profiles, job cache, analysis results", "PostgreSQL or SQLite"],
        ]
      ),
      spacer(),
      callout("WHY SPLIT?", "LLM calls take 3–15 seconds. A live re-ranking interaction must update in under 200ms. Separating AI analysis (slow, runs once) from formula ranking (fast, runs continuously) makes both possible."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 3. THE TWO-PHASE ARCHITECTURE
      // ══════════════════════════════════════════════════════════════════════
      h1("3.  The Two-Phase Architecture"),
      body("Every run through the system passes through two distinct phases. Understanding this split is essential to all downstream design decisions."),
      spacer(),
      twoCol(
        [
          phaseBadge("PHASE 1 — Analyze", "Triggered once when the user uploads files", BLUE_DARK),
          spacer(),
          bullet("Parse and normalize Excel files"),
          bullet("Evaluate each description: strong or weak (Claude)"),
          bullet("Generate suggested rewrites for weak items (Claude)"),
          bullet("Compute base effort and priority field values"),
          bullet("Cache all results against a file hash"),
          spacer(),
          body("Duration: 15–45 seconds for a 100-item backlog. Runs asynchronously; user sees a progress indicator.", { color: "555555", italics: true }),
        ],
        [
          phaseBadge("PHASE 2 — Rank", "Triggered on every priority weight change", ORANGE_SIDE),
          spacer(),
          bullet("Apply current business area weights to items"),
          bullet("Apply current product team weights to items"),
          bullet("Compute weighted priority score per item"),
          bullet("Re-sort the ranked list"),
          bullet("Update UI display instantly"),
          spacer(),
          body("Duration: under 50ms. Pure formula math — no LLM, no server round-trip if run client-side.", { color: "555555", italics: true }),
        ],
        4560, 4560
      ),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 4. DATA INGESTION & PARSING
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("4.  Data Ingestion & Parsing Layer"),
      h2("4.1  Input Files"),
      body("The system accepts two Excel uploads per session:"),
      bullet("Backlog file — one row per work item. Required columns: Item ID, Title, Description, Business Area, Business Area Priority, Product Team, Product Team Priority, Effort Estimate."),
      bullet("Sprint capacity file — team velocity and available capacity for the upcoming sprint."),
      spacer(),
      h2("4.2  Parsing & Normalization"),
      body("A Python parser (pandas + openpyxl) normalizes both files into structured JSON. It handles column name aliasing, priority label-to-number conversion, T-shirt size mapping, and missing field detection. This step is entirely rule-based — no LLM needed."),
      spacer(),
      body("Two pieces of data are extracted here that are critical for the interactive ranking engine: the distinct list of business areas (with their initial priority values) and the distinct list of product teams (with their initial priority values). These become the controls in the priority panel."),
      spacer(),
      callout("NOTE", "If business area or product team names differ across rows for the same entity (e.g., 'Fin' vs 'Finance'), a normalization map should be maintained in config. The alternative — asking Claude to deduplicate — is overkill and adds latency at ingestion."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 5. AI ANALYSIS PIPELINE
      // ══════════════════════════════════════════════════════════════════════
      h1("5.  AI Analysis Pipeline  (Phase 1)"),
      body("The AI pipeline runs once per file upload. Results are cached by file hash so re-uploads of the same backlog skip the LLM calls entirely. The pipeline has three steps."),
      spacer(),

      stepBadge("A", "Parse & Normalize  (Rules-Based)"),
      spacer(),
      body("Reads both Excel files. Outputs normalized items list, capacity object, parse warnings, and the distinct business area and product team control sets."),
      spacer(),

      stepBadge("B", "Description Quality Evaluation  (Claude)"),
      spacer(),
      body("For each work item, Claude evaluates whether the description is actionable and sufficiently detailed. Items are batched in groups of 25 to manage latency."),
      spacer(),
      body("The system prompt defines four evaluation criteria: Does the description explain the business purpose? Is the scope clear? Are acceptance criteria implied? Could a developer start work without asking questions? Claude returns a structured JSON response per item with four fields: quality (strong / weak), confidence score (0–1), a one-sentence reason, and a draft improvement for weak items."),
      spacer(),
      callout("AI VALUE", "A length or keyword rule cannot catch a 200-word description that never explains why the work matters. Claude evaluates meaning — flagging vague descriptions while passing concise ones that are genuinely actionable.", BLUE_MID, BLUE_LIGHT),
      spacer(),

      stepBadge("C", "Suggestion Generation  (Claude — Weak Items Only)"),
      spacer(),
      body("For every item flagged as weak, Claude drafts an improved description. It receives the item title, the existing description, the weakness reason, and a template paragraph showing your team's preferred description format (configurable in the prompt template file)."),
      spacer(),
      body("Output is stored in the analysis cache and surfaced in the app UI alongside the original description, not overwriting it. Users choose whether to accept each suggestion."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 6. INTERACTIVE RANKING ENGINE
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("6.  Interactive Ranking Engine  (Phase 2)"),
      h2("6.1  How Live Re-Ranking Works"),
      body("When a user adjusts a business area or product team priority weight, the ranking engine recomputes scores for all items using the updated weights and re-sorts the displayed list. This happens in under 50 milliseconds."),
      spacer(),
      body("Because this must be instant, it runs as pure JavaScript in the browser — no server round-trip. The analysis results from Phase 1 are loaded into the browser once (as a JSON payload) and the ranking formula executes client-side on every change."),
      spacer(),
      h2("6.2  Scoring Formula"),
      body("Each work item receives a composite priority score from three inputs:"),
      spacer(),
      dataTable(
        ["Input", "Default Weight", "Direction", "Source"],
        [2600, 1600, 1600, 3560],
        [
          ["Business Area Priority", "45%", "Higher = more urgent", "User-controlled weight × item's business area value"],
          ["Product Team Priority", "35%", "Higher = more urgent", "User-controlled weight × item's product team value"],
          ["Effort Estimate", "20%", "Lower = higher score", "Inverse of normalized effort; smaller = quicker win"],
        ]
      ),
      spacer(),
      body("The formula weights themselves (45% / 35% / 20%) are fixed at the application level. What users control are the relative priority values assigned to each business area and each product team — for example, setting Finance to priority 5 and Marketing to priority 2."),
      spacer(),
      h2("6.3  What Users Control"),
      body("The priority panel exposes two sets of controls, one per dimension:"),
      spacer(),
      dataTable(
        ["Control", "UI Element", "Effect on Ranking"],
        [2600, 2400, 4360],
        [
          ["Business Area Priority", "Drag-to-rank list or numeric inputs (1–5 per area)", "Items from higher-ranked areas score more heavily on the 45% business weight"],
          ["Product Team Priority", "Drag-to-rank list or numeric inputs (1–5 per team)", "Items from higher-ranked teams score more heavily on the 35% team weight"],
        ]
      ),
      spacer(),
      body("Both control sets are derived automatically from the uploaded backlog — the app reads the distinct business areas and product teams from the parsed data and creates one control per entity. No manual configuration is required."),
      spacer(),
      h2("6.4  Claude Tie-Breaking"),
      body("The formula may produce identical or near-identical scores for competing items. Claude tie-breaking (used in the original batch pipeline) does not run live — it would add latency incompatible with interactive use. Instead, tie-breaking applies in two places:"),
      bullet("On initial load, after Phase 1 completes. Claude reviews any items within 5 points of each other and produces a tie-break order. This is stored with the analysis cache and applied as a stable secondary sort."),
      bullet("On export. When the user downloads the final Excel, the export pipeline re-runs Claude tie-breaking against the current priority weights before generating the file."),
      spacer(),
      callout("DESIGN NOTE", "Users see live formula-ranked results while working. The final exported file applies Claude tie-breaking on top, so the downloaded output may differ slightly from the on-screen order for closely-ranked items. A tooltip in the UI explains this.", ORANGE_SIDE, ORANGE_BG),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 7. PRIORITY PROFILE MANAGEMENT
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("7.  Priority Profile Management"),
      h2("7.1  What a Profile Contains"),
      body("A named priority profile is a saved snapshot of a user's priority weight configuration. It stores everything needed to reproduce a particular ranking view:"),
      spacer(),
      dataTable(
        ["Field", "Type", "Description"],
        [2400, 1400, 5560],
        [
          ["profile_id", "UUID", "Unique identifier"],
          ["name", "string", "User-assigned label, e.g. 'Q2 Planning — Finance-led'"],
          ["created_by", "string", "User ID of creator"],
          ["created_at", "datetime", "Timestamp"],
          ["business_area_weights", "JSON object", "Map of business area name → priority value (1–5)"],
          ["product_team_weights", "JSON object", "Map of product team name → priority value (1–5)"],
          ["notes", "string (optional)", "Free-text annotation for context"],
        ]
      ),
      spacer(),
      h2("7.2  Profile Lifecycle"),
      bullet("Create. After adjusting the priority panel, the user clicks 'Save as Profile', enters a name, and optionally adds a note. The current weight configuration is written to the database."),
      bullet("Load. From a 'Profiles' dropdown or sidebar, the user selects any saved profile. The priority panel updates immediately and the ranked list re-renders with the loaded weights."),
      bullet("Update. The user adjusts weights after loading a profile and clicks 'Update Profile' to overwrite, or 'Save as New' to fork it."),
      bullet("Delete. Profiles can be deleted by their creator. Deletion is soft (flagged, not removed) to preserve audit history."),
      spacer(),
      h2("7.3  Profile Scoping"),
      body("Profiles are personal by default — visible only to the creator. A 'Share with team' toggle promotes a profile to team-visible, allowing colleagues to load it as a starting point. Team-shared profiles are read-only for non-owners; others fork rather than edit."),
      spacer(),
      callout("USE CASE", "A product lead saves a 'Pre-PI Planning' profile before the quarterly planning session. The team loads that shared profile as a baseline, each makes personal adjustments, and a facilitator exports the final ranked list from the agreed-upon profile.", TEAL_SIDE, TEAL_BG),
      spacer(),
      h2("7.4  Profile Storage Schema"),
      body("Stored in a relational database (PostgreSQL recommended for teams; SQLite sufficient for small internal deployments)."),
      spacer(),
      dataTable(
        ["Table", "Key Columns"],
        [2400, 7000],
        [
          ["profiles", "profile_id, name, created_by, created_at, updated_at, deleted_at, shared, notes"],
          ["profile_weights", "profile_id, dimension (business_area | product_team), entity_name, weight_value"],
          ["analysis_cache", "file_hash, created_at, items_json, description_evals_json, suggestions_json"],
          ["jobs", "job_id, file_hash, status, created_at, completed_at, error_json"],
        ]
      ),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 8. INTERACTIVE UI LAYER
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("8.  Interactive UI Layer"),
      h2("8.1  Application Layout"),
      body("The application is a single-page React app with three primary zones: an upload panel (used once per session), a priority control panel (used interactively), and a ranked backlog panel (the live output)."),
      spacer(),
      dataTable(
        ["Zone", "Location", "Contents"],
        [1800, 1800, 5760],
        [
          ["Upload Panel", "Top / modal on first visit", "File pickers for backlog and capacity Excel, upload button, progress indicator during Phase 1"],
          ["Priority Panel", "Left sidebar", "Business area weight controls, product team weight controls, profile save/load, reset-to-default button"],
          ["Ranked Backlog", "Main content area", "Live-sorted table of work items with rank, score, description quality flag, and expandable suggestion"],
          ["Export Bar", "Bottom bar / header", "Download button (triggers Claude tie-breaking then Excel generation), current profile name badge"],
        ]
      ),
      spacer(),
      h2("8.2  Priority Panel Detail"),
      body("Each dimension (business areas and product teams) has its own section in the priority panel. Within each section, entities are displayed as a list where the user can:"),
      bullet("Enter a numeric weight (1–5) for each entity directly, or"),
      bullet("Drag to reorder, with the app assigning weights automatically based on position (top = 5, bottom = 1)."),
      spacer(),
      body("Changes apply instantly — there is no Apply button. A small 'unsaved changes' indicator appears when the current weights differ from the loaded profile, prompting the user to save or discard."),
      spacer(),
      h2("8.3  Ranked Backlog Panel Detail"),
      body("The backlog table updates within 50ms of any priority control change. Each row shows:"),
      spacer(),
      dataTable(
        ["Column", "Description"],
        [2600, 6760],
        [
          ["Rank", "Current rank position (1 = highest priority); updates live"],
          ["Item ID + Title", "Identifier and title from the backlog"],
          ["Priority Score", "Computed composite score; updates live"],
          ["Business Area", "The item's business area with its current weight shown as a badge"],
          ["Product Team", "The item's product team with its current weight shown as a badge"],
          ["Effort", "Normalized effort estimate"],
          ["Description Quality", "Strong (green) or Weak (amber) badge from Phase 1 evaluation"],
          ["Suggestion", "Expandable row: shows Claude's draft improvement for weak items; accept / dismiss controls"],
        ]
      ),
      spacer(),
      h2("8.4  Application States"),
      dataTable(
        ["State", "Trigger", "UI Behavior"],
        [2000, 2800, 4560],
        [
          ["Empty", "First visit", "Upload panel shown prominently; priority panel and backlog are hidden"],
          ["Analyzing", "Files uploaded", "Progress indicator showing Phase 1 steps; priority panel disabled"],
          ["Ready", "Phase 1 complete", "Full UI unlocked; backlog table populated; priority panel active"],
          ["Re-ranking", "User changes a weight", "Backlog table re-sorts in under 50ms; no loading state needed"],
          ["Exporting", "User clicks Download", "Brief loading state while Claude tie-breaking and Excel generation run (5–15s)"],
          ["Profile saved", "User saves profile", "Toast notification; profile added to dropdown"],
        ]
      ),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 9. LLM INTEGRATION DETAILS
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("9.  LLM Integration Details"),
      h2("9.1  Model Selection"),
      spacer(),
      dataTable(
        ["Use Case", "Model", "Rationale"],
        [3000, 2600, 3760],
        [
          ["Description evaluation (batch, Phase 1)", "claude-haiku-4-5", "High volume, structured judgment — fast and cost-effective"],
          ["Suggestion generation (Phase 1)", "claude-sonnet-4-6", "Writing quality matters; Sonnet produces more natural prose"],
          ["Tie-breaking on export", "claude-sonnet-4-6", "Contextual multi-item reasoning; runs once at export time"],
        ]
      ),
      spacer(),
      h2("9.2  Prompt Design Principles"),
      bullet("Use structured output (tool_use / JSON schema) for all calls to guarantee machine-parseable responses without regex."),
      bullet("Batch 20–25 items per call to balance quality and latency. Larger batches dilute attention; smaller batches multiply cost."),
      bullet("Include two or three few-shot examples of weak and strong descriptions in every description evaluation prompt."),
      bullet("Store all prompts in versioned template files (not hardcoded strings). This allows prompt iteration without code deploys."),
      spacer(),
      h2("9.3  Caching Strategy"),
      body("Phase 1 results are cached against an MD5 hash of the uploaded backlog file. On re-upload of the same file, the system returns cached analysis results immediately with no LLM calls. Cache records include the model version and prompt template version, so a cache entry is invalidated when either changes."),
      spacer(),
      h2("9.4  Fallback Strategy"),
      body("If any Claude API call fails after two retries: description evaluation falls back to marking affected items as 'Review Needed'; suggestion generation is skipped with a note in the UI; tie-breaking on export is skipped and the formula sort is used. The pipeline always returns a usable result."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 10. FINE-TUNING vs. PROMPT ENGINEERING
      // ══════════════════════════════════════════════════════════════════════
      h1("10.  Fine-Tuning vs. Prompt Engineering"),
      h2("10.1  Start with Prompting"),
      body("Launch with prompt engineering and few-shot examples. Claude's base models achieve 82–90% agreement with expert human reviewers on description quality judgment when given a well-designed prompt. Fine-tuning is a Phase 2 investment, not a launch requirement."),
      spacer(),
      h2("10.2  When Fine-Tuning Becomes Worth It"),
      spacer(),
      dataTable(
        ["Condition", "Fine-Tune?", "Notes"],
        [3800, 1300, 4260],
        [
          ["Prompting accuracy < 80% on your backlog style", "✓ Yes", "Your team's descriptions have patterns the base model does not know"],
          ["You have 100+ PM-labeled examples (weak / strong)", "✓ Yes", "Enough signal for meaningful training; expect +8–12pp accuracy gain"],
          ["Consistent tone in generated suggestions is critical", "✓ Yes", "Fine-tuning enforces style better than prompt instructions alone"],
          ["Cost at scale is a concern (1,000+ items/week)", "✓ Yes", "Fine-tuned Haiku can match Sonnet accuracy at lower cost"],
          ["Requirements are still changing rapidly", "No", "Fine-tuned models need retraining each time criteria shift"],
          ["Data is sparse (fewer than 50 labeled examples)", "No", "Insufficient signal; fine-tuning on sparse data hurts generalization"],
        ]
      ),
      spacer(),
      h2("10.3  Fine-Tuning Approach"),
      bullet("Data collection: have two or three experienced PMs label 150–200 past backlog items as strong or weak with a one-sentence reason. Resolve disagreements to create a gold-standard dataset."),
      bullet("Train/test split: hold out 20% for evaluation. Never train on these."),
      bullet("Training format: Anthropic fine-tuning JSONL — each record is a prompt/completion pair matching production prompt structure."),
      bullet("Baseline comparison: only deploy fine-tuned model if it outperforms base model on the held-out test set by more than 5 percentage points."),
      bullet("Retraining trigger: every 6 months, or when the team adopts a new ticket template."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 11. TOOL STACK
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("11.  Recommended Tool Stack"),
      spacer(),
      dataTable(
        ["Layer", "Tool / Library", "Purpose"],
        [1800, 2600, 4960],
        [
          ["Frontend", "React + Tailwind CSS", "Single-page application; priority controls, live backlog table, profile manager"],
          ["State Management", "Zustand (or React Context)", "Holds current priority weights and analysis results; drives live re-ranking"],
          ["Ranking (client-side)", "Pure JavaScript", "Formula scoring runs in the browser — no server round-trip on weight changes"],
          ["Backend API", "Python + FastAPI", "Handles file upload, Phase 1 AI pipeline, export, profile CRUD"],
          ["Excel Parsing", "pandas + openpyxl", "Read input files; write annotated output file on export"],
          ["AI Orchestration", "LangGraph (or Python)", "Multi-step Phase 1 pipeline with state, retries, and fallbacks"],
          ["LLM — Evaluation", "Claude Haiku (claude-haiku-4-5)", "Batched description quality evaluation"],
          ["LLM — Generation", "Claude Sonnet (claude-sonnet-4-6)", "Suggestion generation and tie-breaking"],
          ["Structured Output", "Anthropic tool_use / JSON schema", "Enforces machine-parseable LLM responses"],
          ["Job Queue", "Celery + Redis", "Runs Phase 1 asynchronously; frontend polls for job status"],
          ["Database", "PostgreSQL (or SQLite for small teams)", "Priority profiles, analysis cache, job records"],
          ["Auth", "SSO / existing internal auth", "Restrict access; associate profiles with user identity"],
          ["Hosting", "Internal server or AWS (ECS + RDS)", "Internal tool; no multi-tenancy required"],
        ]
      ),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 12. API DESIGN
      // ══════════════════════════════════════════════════════════════════════
      h1("12.  Backend API Design"),
      body("The FastAPI backend exposes four resource groups. All responses are JSON; Excel I/O is multipart form data."),
      spacer(),
      dataTable(
        ["Method", "Endpoint", "Description"],
        [1200, 3400, 4760],
        [
          ["POST", "/jobs", "Upload backlog + capacity files; start Phase 1 pipeline; return job_id"],
          ["GET", "/jobs/{job_id}", "Poll job status and progress; return analysis results when complete"],
          ["GET", "/profiles", "List saved profiles for current user (includes shared team profiles)"],
          ["POST", "/profiles", "Save a new priority profile with weights"],
          ["PUT", "/profiles/{id}", "Update an existing profile"],
          ["DELETE", "/profiles/{id}", "Soft-delete a profile"],
          ["POST", "/export", "Trigger tie-breaking + Excel generation with given weights; return file"],
        ]
      ),
      spacer(),
      body("The ranking formula is not a backend endpoint — it runs in the browser. The backend only needs to be involved when performing LLM operations (Phase 1) or generating the export file."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 13. KEY DESIGN DECISIONS
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("13.  Key Design Decisions"),
      spacer(),
      h3("Client-side ranking over server-side"),
      body("Running the ranking formula in the browser eliminates network latency entirely and makes the 'under 50ms' target achievable on any hardware. The tradeoff is that the formula logic must be duplicated in JavaScript (browser) and Python (export pipeline). A shared formula spec document mitigates drift."),
      spacer(),
      h3("AI runs once; ranking runs continuously"),
      body("Separating Phase 1 (AI, slow) from Phase 2 (formula, instant) is the central architectural decision. Without this split, every weight adjustment would trigger a 15-second LLM call, making the interactive UX unusable."),
      spacer(),
      h3("Analysis cache by file hash"),
      body("Caching Phase 1 results by MD5 file hash means a re-upload of an unchanged backlog costs nothing in LLM calls. Sprint-over-sprint reuse of a backlog with minor additions only processes the new items (delta processing is a Phase 2 enhancement)."),
      spacer(),
      h3("Profiles store weights, not ranked output"),
      body("Profiles save priority weight configurations, not the resulting item order. The same profile applied to a new backlog upload produces a fresh ranking based on the new items. This makes profiles durable across sprints without requiring complex versioning of backlog state."),
      spacer(),
      h3("Export triggers Claude tie-breaking, not live view"),
      body("Claude tie-breaking is reserved for the export step to avoid adding latency to the interactive experience. The live ranked list uses formula order; the downloaded file applies Claude's contextual judgment as a final pass. A tooltip in the UI explains the potential minor difference."),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 14. PHASED ROLLOUT
      // ══════════════════════════════════════════════════════════════════════
      h1("14.  Phased Rollout Recommendation"),
      spacer(),
      dataTable(
        ["Phase", "Scope", "Key Milestone"],
        [1100, 4600, 3660],
        [
          ["Phase 1", "Upload + formula ranking + live priority panel (no AI). Export to Excel.", "Users can upload, adjust weights, see live re-ranking, and export in under 5 seconds."],
          ["Phase 2", "Add Phase 1 AI pipeline: description evaluation + flags + suggestions.", "PM team validates AI flags on 50 real items. Measure agreement rate target: > 80%."],
          ["Phase 3", "Add named profile save/load. Add Claude tie-breaking on export.", "End-to-end interactive application with persistent profiles in production."],
          ["Phase 4 (opt.)", "Team-shared profiles. Delta processing for incremental backlog updates. Fine-tuning if accuracy data justifies.", "Fine-tuned model outperforms base by > 5pp on held-out test set."],
        ]
      ),
      spacer(),

      // ══════════════════════════════════════════════════════════════════════
      // 15. OPEN QUESTIONS
      // ══════════════════════════════════════════════════════════════════════
      h1("15.  Open Questions"),
      bullet("Formula weight defaults. Are the default formula weights (45% business / 35% team / 20% effort) acceptable, or should they be user-configurable per session?"),
      bullet("Priority input UX. Should the priority panel use drag-to-reorder lists, numeric sliders (1–5), or raw numeric inputs? Drag is more intuitive but harder to build; numbers are faster but less tactile."),
      bullet("Column mapping UI. Are backlog column names consistent enough to auto-map, or is a one-time mapping configuration step needed the first time a new team uploads a file?"),
      bullet("Suggestion acceptance workflow. When a user accepts a suggested description in the UI, should it be written back to the original Excel, captured only in the export file, or pushed to the upstream backlog tool (e.g., Jira)?"),
      bullet("Weak description threshold. What confidence score from Claude should trigger a weak flag? 0.70 is a reasonable default; validate against real items before going live."),
      bullet("Audit trail. Is there a compliance requirement to log which model version, prompt version, and priority profile produced each export? If so, this metadata should be written to the output file footer."),
      bullet("Multi-backlog sessions. Should users be able to upload multiple backlog files in a single session and compare rankings across them, or is one backlog per session sufficient?"),
      spacer(),
      divider(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 80 },
        children: [new TextRun({ text: "End of Document", font: "Arial", size: 18, italics: true, color: "888888" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/sessions/ecstatic-intelligent-shannon/mnt/outputs/Sprint_Backlog_Prioritization_System_Design.docx', buffer);
  console.log('Done');
});
