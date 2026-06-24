const express = require("express");
const cors = require("cors");
const { GoogleGenAI } = require("@google/genai");
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
        <div class="info-card-val">gemini-3.1-flash-lite<br/>via Google Gemini</div>
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
    <span>Powered by Gemini + docx</span>
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

  // ── Client-side chunking ──────────────────────────────────────────────────
  // Each chunk is sent in a separate short HTTP request so Render's 90s
  // response timeout is never hit. A short delay between chunks happens
  // in the browser, not on the server.

  function splitChunks(text, n) {
    const paras = text.split(/\\n\\s*\\n/).filter(p => p.trim());
    if (paras.length <= n) return paras;
    const size = Math.ceil(paras.length / n);
    const out = [];
    for (let i = 0; i < paras.length; i += size)
      out.push(paras.slice(i, i + size).join('\\n\\n'));
    return out;
  }

  function mergeMarkdowns(parts) {
    const bodies = [], refs = [];
    const refRe = /^#{1,3}\\s+references\\s*$/i;
    for (const part of parts) {
      const lines = part.split('\\n');
      let inRef = false, body = [];
      for (const line of lines) {
        if (refRe.test(line.trim())) { inRef = true; continue; }
        if (inRef && /^#{1,3}\\s+/.test(line.trim())) { inRef = false; body.push(line); continue; }
        if (inRef) { if (line.trim()) refs.push(line); }
        else body.push(line);
      }
      bodies.push(body.join('\\n').trim());
    }
    let merged = bodies.join('\\n\\n');
    if (refs.length) {
      const seen = new Set();
      const unique = refs.filter(l => { const k = l.trim().toLowerCase(); return seen.has(k) ? false : seen.add(k); });
      merged += '\\n\\n## References\\n\\n' + unique.join('\\n');
    }
    return merged;
  }

  async function handleFormat() {
    const text = textarea.value.trim();
    if (!text) { setStatus('error', '<span>⚠ Paste some text first.</span>'); return; }

    const DELAY_MS = 3000;
    const TARGET_TOKENS = 1200;
    const numChunks = Math.min(8, Math.max(3, Math.ceil(text.length / (4 * TARGET_TOKENS))));
    const chunks = splitChunks(text, numChunks);

    btn.disabled = true;
    btn.innerHTML = \`<span class="spinner"></span> Formatting\u2026\`;

    try {
      const markdownParts = [];

      for (let i = 0; i < chunks.length; i++) {
        if (i > 0) {
          // Countdown so the user knows it's not frozen
          for (let s = Math.ceil(DELAY_MS / 1000); s > 0; s--) {
            setStatus('info', \`<span class=\"spinner\"></span><span>Chunk \${i}/\${chunks.length} done \u2014 waiting \${s}s before next\u2026</span>\`);
            await new Promise(r => setTimeout(r, 1000));
          }
        }

        setStatus('info', \`<span class=\"spinner\"></span><span>Processing chunk \${i + 1} of \${chunks.length}\u2026</span>\`);

        const res = await fetch('/api/format-chunk', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ chunk: chunks[i] }),
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({ error: 'Unknown error' }));
          throw new Error(err.error || 'Server error on chunk ' + (i + 1));
        }
        const data = await res.json();
        markdownParts.push(data.markdown);
      }

      setStatus('info', '<span class="spinner"></span><span>Building .docx file\u2026</span>');
      const merged = mergeMarkdowns(markdownParts);

      const docRes = await fetch('/api/build-docx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ markdown: merged }),
      });
      if (!docRes.ok) {
        const err = await docRes.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(err.error || 'Failed to build document');
      }

      const blob = await docRes.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'DocuForge_output.docx';
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
      setStatus('success', '<span>\u2713 Document downloaded successfully!</span>');

    } catch (e) {
      setStatus('error', \`<span>\u2717 \${e.message}</span>\`);
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

// ─── Gemini System Prompt ────────────────────────────────────────────────────
// Kept short and strict so gemini-3.1-flash-lite stays on-task per chunk.
const SYSTEM_PROMPT = `Convert raw text to Markdown. Rules:
1. Output ONLY Markdown — no commentary, no code fences.
2. Preserve EVERY word and EVERY value exactly as given. Never summarise, omit, regenerate, or add content. If the input already contains a table (rows of data separated by | or aligned in columns), copy every cell's text into the Markdown table EXACTLY as written — never leave a cell blank, never rebuild the table from memory.
3. Headings: ONLY mark a line as a heading (#, ##, ###) if it is clearly a real section/document title in the source (e.g. "Methodology", "Results", "Discussion", "Conclusion", "References"). A normal sentence or citation discussion paragraph is NEVER a heading, no matter how it looks. Never bold/center a paragraph by turning it into a heading.
4. Lists: enumerated items → numbered list; item groups → bullet list.
5. Tables: comparative or multi-attribute data MUST become a Markdown table with a header row AND fully populated data rows, copied verbatim from the input (see rule 2). If you cannot find real values to put in every cell, do NOT create a table — output that content as normal paragraph text instead. Never output a table with empty cells.
6. References: in-text author-year mentions inside normal prose (e.g. "Dutta et al. (2021) introduced...") are NOT citations to collect — leave them exactly where they are, inline, as part of the paragraph. ONLY if this chunk contains an actual References/Bibliography section (a heading literally saying "References" or "Bibliography" followed by a list of full source entries) should you reproduce that section, headed EXACTLY "## References" (two # characters, nothing more, nothing less), with every entry preserved verbatim. If this chunk has no such section, do not output a References heading at all.
7. Bold existing emphasis as **text**; italicise titles/terms as *text*.`;

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

// ─── Helpers ─────────────────────────────────────────────────────────────────
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Dynamically pick chunk count: target ~1200 user tokens/chunk, clamp 3–8 */
function calcChunkCount(text) {
  const est = Math.ceil(text.length / 4);
  return Math.min(8, Math.max(3, Math.ceil(est / 1200)));
}

/** Split text into N roughly-equal chunks at paragraph boundaries */
function splitIntoParagraphChunks(text, n) {
  const paragraphs = text.split(/\n\s*\n/).filter((p) => p.trim());
  if (paragraphs.length <= n) return paragraphs.map((p) => p.trim());

  const chunkSize = Math.ceil(paragraphs.length / n);
  const chunks = [];
  for (let i = 0; i < paragraphs.length; i += chunkSize) {
    chunks.push(paragraphs.slice(i, i + chunkSize).join("\n\n").trim());
  }
  return chunks;
}

/**
 * Merge multiple markdown outputs:
 * - Concatenate body sections in order
 * - Collect all ## References blocks and consolidate into one at the end
 */
function mergeMarkdowns(parts) {
  const bodies = [];
  const refLines = [];
  const refHeadingRe = /^#{1,3}\s+references\s*$/i;

  for (const part of parts) {
    const lines = part.split("\n");
    let inRefs = false;
    const bodyLines = [];

    for (const line of lines) {
      if (refHeadingRe.test(line.trim())) {
        inRefs = true;
        continue;
      }
      if (inRefs && /^#{1,3}\s+/.test(line.trim())) {
        // Another heading after references — stop collecting refs
        inRefs = false;
        bodyLines.push(line);
        continue;
      }
      if (inRefs) {
        if (line.trim()) refLines.push(line);
      } else {
        bodyLines.push(line);
      }
    }
    bodies.push(bodyLines.join("\n").trim());
  }

  let merged = bodies.join("\n\n");
  if (refLines.length) {
    // Deduplicate reference lines
    const seen = new Set();
    const unique = refLines.filter((l) => {
      const key = l.trim().toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    merged += "\n\n## References\n\n" + unique.join("\n");
  }
  return merged;
}

/** Call Gemini for one chunk with a per-chunk timeout */
async function processChunk(ai, chunk, chunkIndex) {
  const CHUNK_TIMEOUT = 55000; // 55s per chunk
  const response = await Promise.race([
    ai.models.generateContent({
      model: "gemini-3.1-flash-lite",
      contents: chunk,
      config: {
        systemInstruction: SYSTEM_PROMPT,
        temperature: 0.1,
        maxOutputTokens: 2500,
      },
    }),
    new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error(`Chunk ${chunkIndex + 1} timed out. Try with shorter text.`)),
        CHUNK_TIMEOUT
      )
    ),
  ]);
  return response.text || "";
}

// ─── API Routes ──────────────────────────────────────────────────────────────

// POST /api/format-chunk  { chunk: string } → { markdown: string }
// Called once per chunk from the frontend. Each request is short (<30s).
app.post("/api/format-chunk", async (req, res) => {
  const { chunk } = req.body;
  if (!chunk || typeof chunk !== "string" || !chunk.trim()) {
    return res.status(400).json({ error: "No chunk provided." });
  }

  const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  try {
    const result = await processChunk(ai, chunk, 0);
    res.json({ markdown: result });
  } catch (err) {
    console.error("Gemini chunk error:", err.message);
    res.status(502).json({ error: "Gemini API error: " + err.message });
  }
});

// POST /api/build-docx  { markdown: string } → .docx binary
// Called once after all chunks are merged by the frontend.
app.post("/api/build-docx", async (req, res) => {
  const { markdown } = req.body;
  if (!markdown || typeof markdown !== "string" || !markdown.trim()) {
    return res.status(400).json({ error: "No markdown provided." });
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
