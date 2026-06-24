const express = require("express");
const cors = require("cors");
const Groq = require("groq-sdk");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, TabStopPosition, TabStopType,
  convertInchesToTwip, LineRuleType, TableBorders,
} = require("docx");

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// ─── Inline Frontend ────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>DocuForge</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Instrument+Serif:ital@0;1&family=Inter:wght@300;400;500;600&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg:       #0d0f14;
    --surface:  #13161e;
    --border:   #1f2433;
    --muted:    #2a2f42;
    --accent:   #4f6ef7;
    --accent2:  #7c3aed;
    --fg:       #e8eaf0;
    --fg-dim:   #6b7280;
    --success:  #10b981;
    --radius:   10px;
  }

  html, body {
    height: 100%;
    background: var(--bg);
    color: var(--fg);
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 15px;
    line-height: 1.6;
    -webkit-font-smoothing: antialiased;
  }

  /* ── Layout ── */
  .shell {
    min-height: 100vh;
    display: grid;
    grid-template-rows: auto 1fr auto;
    max-width: 860px;
    margin: 0 auto;
    padding: 0 24px;
  }

  /* ── Header ── */
  header {
    padding: 36px 0 28px;
    display: flex;
    align-items: baseline;
    gap: 14px;
    border-bottom: 1px solid var(--border);
  }
  .logo {
    font-family: 'Instrument Serif', Georgia, serif;
    font-size: 28px;
    font-weight: 400;
    letter-spacing: -0.5px;
    background: linear-gradient(135deg, #a5b4fc 0%, #818cf8 50%, #6366f1 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
  }
  .logo-tag {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: var(--fg-dim);
    letter-spacing: 0.08em;
    text-transform: uppercase;
    border: 1px solid var(--border);
    padding: 2px 8px;
    border-radius: 4px;
    position: relative;
    top: -1px;
  }

  /* ── Main area ── */
  main {
    padding: 32px 0;
    display: flex;
    flex-direction: column;
    gap: 20px;
  }

  .label-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .label {
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--fg-dim);
  }
  .char-count {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    color: var(--fg-dim);
  }

  textarea {
    width: 100%;
    min-height: 380px;
    resize: vertical;
    background: var(--surface);
    color: var(--fg);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 18px 20px;
    font-family: 'DM Mono', monospace;
    font-size: 13.5px;
    line-height: 1.7;
    outline: none;
    transition: border-color 0.18s;
    caret-color: var(--accent);
  }
  textarea:focus {
    border-color: var(--accent);
    box-shadow: 0 0 0 3px rgba(79,110,247,0.12);
  }
  textarea::placeholder { color: var(--fg-dim); opacity: 0.6; }

  /* ── Options row ── */
  .options-row {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    align-items: center;
  }
  .chip {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 12px;
    color: var(--fg-dim);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 4px 12px;
    cursor: default;
    user-select: none;
    background: var(--surface);
  }
  .chip-dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--success);
    flex-shrink: 0;
  }

  /* ── Button ── */
  .btn-wrap { display: flex; gap: 12px; align-items: center; }

  button {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    padding: 13px 28px;
    background: var(--accent);
    color: #fff;
    border: none;
    border-radius: var(--radius);
    font-family: 'Inter', sans-serif;
    font-size: 14px;
    font-weight: 600;
    letter-spacing: 0.01em;
    cursor: pointer;
    transition: background 0.15s, transform 0.1s, box-shadow 0.15s;
    box-shadow: 0 2px 16px rgba(79,110,247,0.28);
    position: relative;
    overflow: hidden;
  }
  button::before {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(135deg, transparent 0%, rgba(255,255,255,0.06) 100%);
  }
  button:hover:not(:disabled) {
    background: #5f7ef8;
    box-shadow: 0 4px 24px rgba(79,110,247,0.4);
    transform: translateY(-1px);
  }
  button:active:not(:disabled) { transform: translateY(0); }
  button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
  }

  /* ── Spinner ── */
  .spinner {
    width: 16px; height: 16px;
    border: 2.5px solid rgba(255,255,255,0.25);
    border-top-color: #fff;
    border-radius: 50%;
    animation: spin 0.75s linear infinite;
    flex-shrink: 0;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ── Status message ── */
  #status {
    display: none;
    align-items: center;
    gap: 10px;
    font-size: 13px;
    padding: 13px 16px;
    border-radius: var(--radius);
    border: 1px solid;
  }
  #status.info {
    display: flex;
    background: rgba(79,110,247,0.08);
    border-color: rgba(79,110,247,0.25);
    color: #a5b4fc;
  }
  #status.success {
    display: flex;
    background: rgba(16,185,129,0.08);
    border-color: rgba(16,185,129,0.25);
    color: #6ee7b7;
  }
  #status.error {
    display: flex;
    background: rgba(239,68,68,0.08);
    border-color: rgba(239,68,68,0.25);
    color: #fca5a5;
  }

  /* ── Info grid ── */
  .info-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
    margin-top: 4px;
  }
  .info-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 16px 18px;
  }
  .info-card-title {
    font-size: 10px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--fg-dim);
    margin-bottom: 6px;
  }
  .info-card-val {
    font-family: 'DM Mono', monospace;
    font-size: 12px;
    color: var(--fg);
    line-height: 1.5;
  }

  /* ── Footer ── */
  footer {
    padding: 20px 0;
    border-top: 1px solid var(--border);
    font-size: 12px;
    color: var(--fg-dim);
    display: flex;
    justify-content: space-between;
    align-items: center;
  }

  @media (max-width: 600px) {
    .info-grid { grid-template-columns: 1fr; }
    .options-row { flex-direction: column; align-items: flex-start; }
  }
</style>
</head>
<body>
<div class="shell">
  <header>
    <span class="logo">DocuForge</span>
    <span class="logo-tag">AI → .docx</span>
  </header>

  <main>
    <div>
      <div class="label-row" style="margin-bottom:10px;">
        <span class="label">Raw Text Input</span>
        <span class="char-count" id="charCount">0 chars</span>
      </div>
      <textarea
        id="rawText"
        placeholder="Paste any raw text here — academic papers, reports, notes, prose, tables in any format. DocuForge will detect structure, convert implicit tables to proper ones, group references, and export a formatted Word document."
        spellcheck="false"
      ></textarea>
    </div>

    <div class="options-row">
      <div class="chip"><span class="chip-dot"></span> Times New Roman · 12pt</div>
      <div class="chip"><span class="chip-dot"></span> 1-inch margins</div>
      <div class="chip"><span class="chip-dot"></span> Auto table detection</div>
      <div class="chip"><span class="chip-dot"></span> Reference grouping</div>
    </div>

    <div class="btn-wrap">
      <button id="formatBtn" onclick="handleFormat()">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5">
          <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
        </svg>
        Format &amp; Export .docx
      </button>
    </div>

    <div id="status"></div>

    <div class="info-grid">
      <div class="info-card">
        <div class="info-card-title">AI Model</div>
        <div class="info-card-val">llama-3.3-70b<br/>via Groq</div>
      </div>
      <div class="info-card">
        <div class="info-card-title">Output Format</div>
        <div class="info-card-val">Word .docx<br/>Times New Roman</div>
      </div>
      <div class="info-card">
        <div class="info-card-title">What Gets Detected</div>
        <div class="info-card-val">Headers · Tables<br/>Lists · References</div>
      </div>
    </div>
  </main>

  <footer>
    <span>DocuForge — paste raw, export professional</span>
    <span>Powered by Groq + docx</span>
  </footer>
</div>

<script>
  const textarea = document.getElementById('rawText');
  const charCount = document.getElementById('charCount');
  const status = document.getElementById('status');
  const btn = document.getElementById('formatBtn');

  textarea.addEventListener('input', () => {
    const n = textarea.value.length;
    charCount.textContent = n.toLocaleString() + ' chars';
  });

  function setStatus(type, html) {
    status.className = type;
    status.style.display = 'flex';
    status.innerHTML = html;
  }
  function clearStatus() {
    status.className = '';
    status.style.display = 'none';
    status.innerHTML = '';
  }

  async function handleFormat() {
    const text = textarea.value.trim();
    if (!text) {
      setStatus('error', '<span>⚠ Paste some text first.</span>');
      return;
    }

    btn.disabled = true;
    btn.innerHTML = \`<span class="spinner"></span> Formatting…\`;
    setStatus('info', '<span class="spinner"></span><span>Sending to AI for structure detection…</span>');

    try {
      const res = await fetch('/api/format', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(err.error || 'Server error');
      }

      setStatus('info', '<span class="spinner"></span><span>Building .docx file…</span>');

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'DocuForge_output.docx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus('success', '<span>✓ Document downloaded successfully!</span>');
    } catch (e) {
      setStatus('error', \`<span>✗ \${e.message}</span>\`);
    } finally {
      btn.disabled = false;
      btn.innerHTML = \`
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5">
          <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
        </svg>
        Format &amp; Export .docx\`;
    }
  }
</script>
</body>
</html>`;

// ─── Groq System Prompt ──────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are a document structure expert. Your ONLY job is to convert raw text into clean, well-structured Markdown. You NEVER summarize, shorten, paraphrase, or remove ANY content. Every word in the input must appear in the output.

CRITICAL RULES — VIOLATING ANY OF THESE IS FAILURE:
1. Return ONLY valid Markdown. No commentary, no preamble, no explanation, no code fences, no backtick blocks. Just raw Markdown text.
2. NEVER remove, summarize, or alter any content. Preserve every sentence, every number, every name.
3. NEVER add new content that wasn't in the original text.

STRUCTURE DETECTION — apply ALL of these:

TITLES & HEADINGS:
- The first major topic or document title → # Title (H1)
- Major section breaks, numbered sections (1., 2., I., II., A., B.) → ## Heading (H2)
- Subsections, sub-topics → ### Subheading (H3)
- If a line is clearly a label or header (short, standalone, contextually a section name) → make it a heading
- Even unlabeled topic shifts should be marked as ## headings using the most logical section name

PARAGRAPHS:
- Continuous prose → plain paragraph text
- Preserve all paragraph breaks

BULLET LISTS:
- Any list of items, even written as "First X, then Y, also Z" → convert to bullet list (- item)
- Enumerated items in prose ("there are three factors: A, B, and C") → bullet list

NUMBERED LISTS:
- Sequential steps, ordered items → 1. 2. 3. numbered list

TABLES — THIS IS CRITICAL:
- ANY data that compares, contrasts, or enumerates attributes across multiple entities MUST become a Markdown table
- This includes data written as prose like: "Product A costs $10 and has 5 features. Product B costs $20 and has 8 features." → TABLE
- This includes sentences like "In 2020 revenue was $1M, in 2021 it was $2M" → TABLE
- This includes any sentence patterns like "X has Y, Z has W" when comparing → TABLE
- When in doubt, make it a table. Never leave comparative or multi-attribute data as prose.
- Markdown table format: | Col1 | Col2 | Col3 |\\n|------|------|------|\\n| val | val | val |
- Always include a header row with logical column names

REFERENCES / BIBLIOGRAPHY:
- Detect ALL citations, references, footnotes, endnotes, bibliography entries
- Even inline citations like [1], (Smith, 2020), or numbered notes
- Group ALL of them into a single ## References section at the VERY END of the document
- Format each as a bullet: - [1] Author. Title. Journal. Year.
- If citations appear mid-text, keep a marker in the text (e.g., [1]) and move the full reference to the References section

FORMATTING DETAILS:
- Bold: **text** for emphasis already present in original
- Italics: *text* for titles of works, technical terms
- Horizontal rule (---) between major sections only if a clear break exists in original

FINAL CHECK before outputting:
- Did you include every single sentence from the input? YES required.
- Did you convert ALL comparative/tabular data to Markdown tables? YES required.
- Did you group all references at the bottom? YES required.
- Is your output pure Markdown with zero commentary? YES required.`;

// ─── Markdown → docx Parser ──────────────────────────────────────────────────
const FONT = "Times New Roman";
const INCH = convertInchesToTwip(1);

function makeRuns(text) {
  const runs = [];
  const parts = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g);
  for (const part of parts) {
    if (part.startsWith("**") && part.endsWith("**")) {
      runs.push(new TextRun({ text: part.slice(2, -2), bold: true, font: FONT, size: 24 }));
    } else if (part.startsWith("*") && part.endsWith("*")) {
      runs.push(new TextRun({ text: part.slice(1, -1), italics: true, font: FONT, size: 24 }));
    } else if (part) {
      runs.push(new TextRun({ text: part, font: FONT, size: 24 }));
    }
  }
  return runs.length ? runs : [new TextRun({ text: "", font: FONT, size: 24 })];
}

function lineSpacing() {
  return { line: Math.round(1.15 * 240), lineRule: LineRuleType.AUTO };
}

function parseMarkdownToDocx(md) {
  const lines = md.split("\n");
  const children = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // ── Table ──
    if (line.trim().startsWith("|")) {
      const tableLines = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) {
        tableLines.push(lines[i]);
        i++;
      }
      // filter out separator rows
      const dataRows = tableLines.filter(l => !/^\|[\s\-:|]+\|/.test(l.trim()));
      if (dataRows.length >= 1) {
        const rows = dataRows.map((rowLine, rowIdx) => {
          const cells = rowLine.split("|").slice(1, -1).map(c => c.trim());
          return new TableRow({
            children: cells.map(cellText =>
              new TableCell({
                shading: rowIdx === 0 ? { fill: "2a2f42", type: ShadingType.CLEAR, color: "auto" } : undefined,
                children: [new Paragraph({
                  children: rowIdx === 0
                    ? [new TextRun({ text: cellText, bold: true, font: FONT, size: 20, color: rowIdx === 0 ? "FFFFFF" : "000000" })]
                    : makeRuns(cellText).map(r => new TextRun({ ...r, size: 20 })),
                  spacing: { before: 60, after: 60 },
                })],
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
              })
            ),
          });
        });

        children.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
            left: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
            right: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" },
            insideH: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
            insideV: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
          },
          rows,
        }));
        children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
      }
      continue;
    }

    // ── Horizontal rule ──
    if (/^---+$/.test(line.trim())) {
      children.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 1 } },
        spacing: { before: 120, after: 120 },
        children: [],
      }));
      i++; continue;
    }

    // ── H1 ──
    if (line.startsWith("# ")) {
      children.push(new Paragraph({
        children: [new TextRun({ text: line.slice(2).trim(), bold: true, font: FONT, size: 36 })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240, ...lineSpacing() },
      }));
      i++; continue;
    }

    // ── H2 ──
    if (line.startsWith("## ")) {
      children.push(new Paragraph({
        children: [new TextRun({ text: line.slice(3).trim(), bold: true, font: FONT, size: 28 })],
        spacing: { before: 280, after: 120, ...lineSpacing() },
      }));
      i++; continue;
    }

    // ── H3 ──
    if (line.startsWith("### ")) {
      children.push(new Paragraph({
        children: [new TextRun({ text: line.slice(4).trim(), bold: true, italics: true, font: FONT, size: 26 })],
        spacing: { before: 200, after: 80, ...lineSpacing() },
      }));
      i++; continue;
    }

    // ── Bullet list ──
    if (/^[-*+] /.test(line.trim())) {
      children.push(new Paragraph({
        children: makeRuns(line.trim().slice(2)),
        bullet: { level: 0 },
        spacing: { before: 40, after: 40, ...lineSpacing() },
        indent: { left: convertInchesToTwip(0.25), hanging: convertInchesToTwip(0.25) },
      }));
      i++; continue;
    }

    // ── Numbered list ──
    if (/^\d+\. /.test(line.trim())) {
      const text = line.trim().replace(/^\d+\. /, "");
      children.push(new Paragraph({
        children: makeRuns(text),
        numbering: { reference: "default-numbering", level: 0 },
        spacing: { before: 40, after: 40, ...lineSpacing() },
      }));
      i++; continue;
    }

    // ── Blank line ──
    if (line.trim() === "") {
      i++; continue;
    }

    // ── Normal paragraph ──
    children.push(new Paragraph({
      children: makeRuns(line.trim()),
      spacing: { before: 80, after: 80, ...lineSpacing() },
    }));
    i++;
  }

  return children;
}

// ─── API Route ───────────────────────────────────────────────────────────────
app.post("/api/format", async (req, res) => {
  const { text } = req.body;
  if (!text || typeof text !== "string" || !text.trim()) {
    return res.status(400).json({ error: "No text provided." });
  }

  const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

  let markdown;
  try {
    const completion = await groq.chat.completions.create({
      model: "llama-3.3-70b-versatile",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: text },
      ],
      temperature: 0.1,
      max_tokens: 8192,
    });
    markdown = completion.choices[0]?.message?.content || "";
  } catch (err) {
    console.error("Groq error:", err.message);
    return res.status(502).json({ error: "Groq API error: " + err.message });
  }

  let docxChildren;
  try {
    docxChildren = parseMarkdownToDocx(markdown);
  } catch (err) {
    console.error("Parse error:", err.message);
    return res.status(500).json({ error: "Failed to parse markdown: " + err.message });
  }

  const doc = new Document({
    numbering: {
      config: [{
        reference: "default-numbering",
        levels: [{
          level: 0,
          format: "decimal",
          text: "%1.",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: convertInchesToTwip(0.25), hanging: convertInchesToTwip(0.25) },
            },
          },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          margin: { top: INCH, bottom: INCH, left: INCH, right: INCH },
        },
      },
      children: docxChildren,
    }],
  });

  try {
    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="DocuForge_output.docx"');
    res.send(buffer);
  } catch (err) {
    console.error("Packer error:", err.message);
    res.status(500).json({ error: "Failed to build .docx: " + err.message });
  }
});

// ─── Serve frontend ──────────────────────────────────────────────────────────
app.get("/", (req, res) => res.send(HTML));

// ─── Start ───────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`DocuForge running on http://localhost:${PORT}`));
