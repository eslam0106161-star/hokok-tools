'use strict';
/**
 * يحوّل كل ملفات PDF الموجودة في مجلد input/ إلى ملفات Word (شرح + أسئلة أكاديمية)
 * بالكامل من غير تدخل يدوي — مصمم للعمل داخل GitHub Actions (سحابي بالكامل).
 *
 * المراحل:
 *  1) تحويل كل صفحة PDF لصورة (عبر pdftoppm من poppler-utils)
 *  2) استخراج النص من كل صورة عبر Gemini Vision
 *  3) تحليل الهيكل الهرمي للكتاب من النص المستخرج (بدون AI)
 *  4) توليد شرح + أسئلة أكاديمية لكل جزء ورقي عبر Gemini
 *  5) بناء ملف Word (.docx) نهائي وحفظه في مجلد output/
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const JSZip = require('jszip');

const ROOT_DIR = path.join(__dirname, '..');
const INPUT_DIR = path.join(ROOT_DIR, 'input');
const OUTPUT_DIR = path.join(ROOT_DIR, 'output');
const WORK_DIR = path.join(ROOT_DIR, '.work');

const MODEL = process.env.GEMINI_MODEL || 'gemini-3.5-flash-lite';
const WORD_COUNT_OPTION = process.env.WORD_COUNT || 'medium';
const MCQ_COUNT = parseInt(process.env.MCQ_COUNT || '5', 10);
const ESSAY_COUNT = parseInt(process.env.ESSAY_COUNT || '3', 10);
const SKIP_FOOTNOTES = (process.env.SKIP_FOOTNOTES || 'yes') === 'yes';
const PDF_DPI = parseInt(process.env.PDF_DPI || '150', 10);

const API_KEYS = (process.env.GEMINI_API_KEYS || '')
  .split(',').map(s => s.trim()).filter(Boolean);

if (API_KEYS.length === 0) {
  console.error('❌ لا يوجد أي مفتاح Gemini API. أضف GEMINI_API_KEYS كـ Secret في إعدادات الريبو (مفصولة بفاصلة لو أكتر من مفتاح).');
  process.exit(1);
}

let keyIdx = 0;
function currentKey() { return API_KEYS[keyIdx]; }
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function pacingMs() { return MODEL.includes('lite') ? 2200 : 4300; }

function isKeyExhaustedError(status, message) {
  const m = (message || '').toLowerCase();
  if (status === 401 || status === 403) return true;
  if (m.includes('api key not valid') || m.includes('permission')) return true;
  if (status === 429 && (m.includes('per day') || m.includes('perday') || m.includes('daily') || m.includes('requests per day'))) return true;
  return false;
}

/* ================= استدعاءات Gemini ================= */

async function geminiFetch(url, body) {
  let res;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (networkErr) {
    const e = new Error('تعذّر الاتصال بالشبكة');
    e.status = 0;
    throw e;
  }
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try { const j = await res.json(); if (j.error && j.error.message) msg = j.error.message; } catch (e) {}
    const e = new Error(msg);
    e.status = res.status;
    throw e;
  }
  return res.json();
}

async function callGeminiVision(base64Jpeg, apiKey, model, skipFootnotes) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const basePrompt = 'استخرج كل النص الموجود في هذه الصورة بدقة تامة كما هو مكتوب، بنفس اللغة. حافظ على فواصل الفقرات والأسطر قدر الإمكان. لا تضف أي شرح أو تعليق أو مقدمة، أعد النص المستخرج فقط. إذا لم يوجد نص، أعد سطراً فارغاً.';
  const footnotesInstruction = ' تنبيه مهم: قد تحتوي الصفحة على قسم هوامش أو مراجع أسفلها (عادة مفصول عن المتن بخط أفقي، ويبدأ غالباً بأرقام مرجعية مثل (1) أو (2) ويذكر أسماء مؤلفين وكتب ودور نشر وأرقام صفحات). تجاهل هذا القسم تماماً ولا تُدرجه في النص المستخرج، واستخرج فقط نص المتن الأساسي أعلى الخط الفاصل.';
  const promptText = skipFootnotes ? (basePrompt + footnotesInstruction) : basePrompt;
  const body = {
    contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: 'image/jpeg', data: base64Jpeg } }] }],
    generationConfig: { temperature: 0 }
  };
  const data = await geminiFetch(url, body);
  const text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  return text.trim();
}

const wordCountText = { short: 'حوالي 150 كلمة', medium: 'حوالي 300 كلمة', long: 'حوالي 450 كلمة' };

async function callGeminiForUnit(unit, apiKey, model, wcOption, mcqCount, essayCount) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const pathLabel = unit.pathLabel || unit.title;
  const prompt = `أنت أستاذ جامعي متخصص في القانون، مهمتك مساعدة طالب حقوق يذاكر لامتحان فعلي — الأولوية القصوى هي إفادة الطالب أكاديميًا، وليس مجرد تلخيص النص.

عنوان هذا الجزء ضمن الكتاب: "${pathLabel}"
النص الأصلي لهذا الجزء:
"""
${unit.text}
"""

المطلوب أولاً — الشرح:
اكتب شرحًا أكاديميًا واضحًا بالعربية الفصحى السهلة (${wordCountText[wcOption]})، يعيد صياغة الأفكار القانونية بأسلوب مبسط ومنظم (قدّم التعريفات أولاً، ثم الشروط أو الأركان، ثم الأمثلة أو التطبيقات إن وردت)، بدون الإخلال بالدقة القانونية، وبدون إضافة معلومات من خارج النص المعطى.

المطلوب ثانياً — أسئلة الاختيار من متعدد (بالضبط ${mcqCount} أسئلة):
- كل سؤال يجب أن يختبر فهم الطالب للفكرة القانونية (تعريف، شرط، استثناء، تفرقة بين مفهومين) وليس مجرد استرجاع كلمة من النص حرفيًا.
- الخيارات الأربعة يجب أن تكون كلها معقولة ومتقاربة في الصياغة (تجنّب الخيارات الساذجة أو الواضح خطؤها)، بحيث يحتاج الطالب فعلاً للفهم لا للتخمين.
- نوّع بين أسئلة التعريف، وأسئلة "أيهما ليس من..."، وأسئلة تطبيق الفكرة على مثال قصير، وأسئلة التفرقة بين مصطلحين متقاربين إن أمكن.

المطلوب ثالثاً — الأسئلة المقالية (بالضبط ${essayCount} أسئلة):
- يجب أن تكون أسئلة تحليلية أو تطبيقية حقيقية (اشرح، ناقش، قارن، بيّن العلاقة بين، ما رأيك في تطبيق كذا على حالة كذا) — وليست أسئلة استرجاع سطحية مثل "عرّف كذا" فقط.
- كل سؤال يجب أن يكون قابلاً للإجابة بالاعتماد على محتوى هذا الجزء فقط.

أعد النتيجة بصيغة JSON فقط، بدون أي نص أو علامات إضافية قبله أو بعده، وبدون أي إشارة داخل نص الأسئلة نفسها لكونها مأخوذة من "النص" أو "المقطع" (صُغ كل سؤال كسؤال أكاديمي مستقل تمامًا)، بالشكل التالي بالضبط:
{
  "explanation": "الشرح الكامل هنا",
  "mcq": [
    {"question": "نص السؤال", "choices": ["اختيار 1", "اختيار 2", "اختيار 3", "اختيار 4"], "correct_index": 0}
  ],
  "essay": ["سؤال مقالي تحليلي أول", "سؤال مقالي تحليلي ثانٍ"]
}
يجب أن يحتوي "mcq" على ${mcqCount} عنصر بالضبط، و"essay" على ${essayCount} عنصر بالضبط.`;

  const body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.4, responseMimeType: 'application/json' } };
  const data = await geminiFetch(url, body);
  let text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  text = text.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  try { return JSON.parse(text); }
  catch (e) { const err = new Error('تعذّر تفسير استجابة النموذج كـ JSON'); err.status = -1; throw err; }
}

/* ================= تحليل الهيكل الهرمي (بدون AI) ================= */

const ordinalPattern = '(?:الأول|الثاني|الثالث|الرابع|الخامس|السادس|السابع|الثامن|التاسع|العاشر|الحادي عشر|الثاني عشر|الثالث عشر|الرابع عشر|الخامس عشر|السادس عشر|السابع عشر|الثامن عشر|التاسع عشر|العشرون)';
const levelDefs = [
  { key: 'section', label: 'القسم', arName: 'قسم' },
  { key: 'part', label: 'الباب', arName: 'باب' },
  { key: 'chapter', label: 'الفصل', arName: 'فصل' },
  { key: 'topic', label: 'المبحث', arName: 'مبحث' },
  { key: 'subtopic', label: 'المطلب', arName: 'مطلب' },
];
levelDefs.forEach(l => { l.regex = new RegExp(`^${l.label}\\s+${ordinalPattern}\\s*$`); });

function cleanLines(text) {
  return text.split(/\r?\n/).filter(line => {
    const t = line.trim();
    if (/^---\s*الصفحة\s*\d+\s*---$/.test(t)) return false;
    if (/^-\s*\d+\s*-$/.test(t)) return false;
    return true;
  });
}
function isTocMarker(t) { return t === 'فهرس الموضوعات' || t === 'قائمة المراجع' || t === 'رقم الصفحة' || t === 'الموضوع'; }
function isDotLeaderLine(t) { return /\.{2,}\s*\d+\s*$/.test(t) || /^\d+\s*\.{2,}/.test(t); }

function parseBook(rawText) {
  const lines = cleanLines(rawText);
  const n = lines.length;
  const excluded = new Array(n).fill(false);
  {
    let i = 0;
    while (i < n) {
      const t = lines[i].trim();
      if (isTocMarker(t) || isDotLeaderLine(t)) {
        let lastHit = i, j = i;
        while (j < n) {
          const tj = lines[j].trim();
          if (isTocMarker(tj) || isDotLeaderLine(tj)) lastHit = j;
          if (j - lastHit > 20) break;
          j++;
        }
        for (let k = i; k <= lastHit; k++) excluded[k] = true;
        i = lastHit + 1;
      } else { i++; }
    }
    for (let z = 0; z < n; z++) {
      if (excluded[z] && (z === 0 || !excluded[z - 1])) {
        let back = z - 1, steps = 0;
        while (back >= 0 && steps < 8) {
          const t = lines[back].trim();
          const isBareHeading = levelDefs.some(l => l.regex.test(t));
          if (t === '' || isBareHeading) { excluded[back] = true; back--; steps++; }
          else break;
        }
      }
    }
  }

  let matches = [];
  for (let idx = 0; idx < n; idx++) {
    if (excluded[idx]) continue;
    const t = lines[idx].trim();
    if (!t) continue;
    for (let lvl = 0; lvl < levelDefs.length; lvl++) {
      if (levelDefs[lvl].regex.test(t)) { matches.push({ level: lvl, lineIdx: idx }); break; }
    }
  }
  matches.forEach((h, i) => {
    let j = h.lineIdx + 1;
    while (j < n && (excluded[j] || !lines[j].trim())) j++;
    const nextHeadingLine = (i + 1 < matches.length) ? matches[i + 1].lineIdx : n;
    if (j < nextHeadingLine) { h.title = lines[j].trim(); h.titleLineIdx = j; }
    else { h.title = '(بدون عنوان)'; }
  });
  matches.forEach((h, i) => {
    const startLine = (h.titleLineIdx !== undefined ? h.titleLineIdx : h.lineIdx) + 1;
    const endLine = (i + 1 < matches.length) ? matches[i + 1].lineIdx : n;
    const chunk = [];
    for (let k = startLine; k < endLine; k++) { if (!excluded[k]) chunk.push(lines[k]); }
    h.content = chunk.join('\n').trim();
  });

  const root = { level: -1, title: 'root', content: '', children: [] };
  const stack = [root];
  matches.forEach(h => {
    const node = { level: h.level, title: h.title, content: h.content, children: [] };
    while (stack.length > 1 && stack[stack.length - 1].level >= h.level) stack.pop();
    stack[stack.length - 1].children.push(node);
    stack.push(node);
  });

  const leaves = [];
  function collect(node, ancestorIntros) {
    if (node.level >= 0) {
      if (node.children.length === 0) {
        const fullText = ancestorIntros.concat(node.content ? [node.content] : []).join('\n\n').trim();
        node.unitId = leaves.length;
        leaves.push({ id: leaves.length, level: node.level, title: node.title, text: fullText, node });
        return;
      } else {
        const newIntros = node.content && node.content.trim() ? ancestorIntros.concat([node.content.trim()]) : ancestorIntros;
        node.children.forEach(c => collect(c, newIntros));
        return;
      }
    }
    node.children.forEach(c => collect(c, ancestorIntros));
  }
  collect(root, []);

  const stats = { section: 0, part: 0, chapter: 0, topic: 0, subtopic: 0 };
  matches.forEach(h => { stats[levelDefs[h.level].key]++; });

  return { root, leaves, stats };
}

function assignPathLabels(root) {
  function walk(node, pathParts) {
    if (node.level >= 0) {
      const parts = pathParts.concat([node.title]);
      if (node.children.length === 0 && node.unitId !== undefined) node.pathLabel = parts.join(' > ');
      node.children.forEach(c => walk(c, parts));
    } else {
      node.children.forEach(c => walk(c, pathParts));
    }
  }
  walk(root, []);
}

/* ================= بناء ملف Word ================= */

function escXml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function paraXml(text, styleId, opts) {
  opts = opts || {};
  const bold = opts.bold ? '<w:b/>' : '';
  const styleTag = styleId ? `<w:pStyle w:val="${styleId}"/>` : '';
  const jc = opts.center ? '<w:jc w:val="center"/>' : '<w:jc w:val="right"/>';
  const lines = String(text || '').split(/\n+/);
  const runs = lines.map((ln, i) => {
    const br = i > 0 ? '<w:br/>' : '';
    return `<w:r><w:rPr><w:rtl/>${bold}</w:rPr>${br}<w:t xml:space="preserve">${escXml(ln)}</w:t></w:r>`;
  }).join('');
  return `<w:p><w:pPr><w:bidi/>${styleTag}${jc}</w:pPr>${runs}</w:p>`;
}
function pageBreakXml() { return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`; }

function buildDocumentXmlBody(parsedBook, aiResults, titleName) {
  let body = '';
  body += paraXml('ملخص وأسئلة مراجعة للكتاب', 'Heading1', { center: true });
  body += paraXml(titleName, 'Normal', { center: true });
  body += pageBreakXml();

  const levelStyles = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5'];
  function walk(node) {
    if (node.level >= 0) {
      const styleId = levelStyles[Math.min(node.level, 4)];
      body += paraXml(node.title, styleId);
      if (node.children.length === 0) {
        const r = aiResults[node.unitId];
        if (r && !r.fail) {
          if (r.explanation) body += paraXml(r.explanation, 'Normal');
          if (r.mcq && r.mcq.length) {
            body += paraXml('أسئلة اختيار من متعدد', 'Normal', { bold: true });
            const letters = ['أ', 'ب', 'ج', 'د'];
            r.mcq.forEach((q, qi) => {
              body += paraXml(`${qi + 1}. ${q.question || ''}`, 'Normal');
              (q.choices || []).forEach((ch, ci) => { body += paraXml(`   ${letters[ci]}) ${ch}`, 'Normal'); });
              const correctLetter = letters[q.correct_index] || '';
              body += paraXml(`   ✔ الإجابة الصحيحة: ${correctLetter}`, 'Normal', { bold: true });
            });
          }
          if (r.essay && r.essay.length) {
            body += paraXml('أسئلة مقالية للمراجعة', 'Normal', { bold: true });
            r.essay.forEach((q, qi) => { body += paraXml(`${qi + 1}. ${q}`, 'Normal'); });
          }
        } else if (r && r.fail) {
          body += paraXml(`[تعذّر توليد شرح هذا الجزء: ${r.note || ''}]`, 'Normal');
        } else {
          body += paraXml('(لم تتم معالجة هذا الجزء بعد)', 'Normal');
        }
      }
      node.children.forEach(walk);
    } else {
      node.children.forEach(walk);
    }
  }
  walk(parsedBook.root);

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${body}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:bidi/></w:sectPr>
</w:body>
</w:document>`;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
  <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="140"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="260" w:after="140"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/><w:color w:val="2B5F56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="220" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="2B5F56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="200" w:after="100"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="2B5F56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="180" w:after="90"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="C1502E"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="C1502E"/></w:rPr></w:style>
</w:styles>`;

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const DOC_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

function coreXml(titleName) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>${escXml(titleName)}</dc:title>
<dc:creator>خط أنابيب GitHub Actions — من PDF إلى شرح وأسئلة أكاديمية</dc:creator>
</cp:coreProperties>`;
}

async function buildDocx(parsedBook, aiResults, titleName) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML);
  zip.folder('_rels').file('.rels', RELS_XML);
  zip.folder('docProps').file('core.xml', coreXml(titleName));
  const wordFolder = zip.folder('word');
  wordFolder.file('document.xml', buildDocumentXmlBody(parsedBook, aiResults, titleName));
  wordFolder.file('styles.xml', STYLES_XML);
  wordFolder.folder('_rels').file('document.xml.rels', DOC_RELS_XML);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/* ================= تحويل PDF لصور واستخراج النص ================= */

function convertPdfToImages(pdfPath, outDir) {
  fs.mkdirSync(outDir, { recursive: true });
  const prefix = path.join(outDir, 'page');
  execSync(`pdftoppm -jpeg -r ${PDF_DPI} "${pdfPath}" "${prefix}"`, { stdio: 'inherit' });
  const files = fs.readdirSync(outDir)
    .filter(f => f.endsWith('.jpg'))
    .sort((a, b) => {
      const na = parseInt(a.match(/-(\d+)\.jpg$/)?.[1] || '0', 10);
      const nb = parseInt(b.match(/-(\d+)\.jpg$/)?.[1] || '0', 10);
      return na - nb;
    });
  return files.map(f => path.join(outDir, f));
}

async function extractTextFromImages(imagePaths) {
  const pageTexts = [];
  for (let i = 0; i < imagePaths.length; i++) {
    const base64 = fs.readFileSync(imagePaths[i]).toString('base64');
    let text = null;
    let attempts = 0, rateLimitBackoffs = 0;
    while (text === null) {
      if (keyIdx >= API_KEYS.length) throw new Error(`نفدت حدود كل المفاتيح عند الصفحة ${i + 1}`);
      try {
        text = await callGeminiVision(base64, currentKey(), MODEL, SKIP_FOOTNOTES);
      } catch (err) {
        if (isKeyExhaustedError(err.status, err.message)) { keyIdx++; attempts = 0; rateLimitBackoffs = 0; continue; }
        if (err.status === 429) {
          rateLimitBackoffs++;
          if (rateLimitBackoffs > 6) { keyIdx++; rateLimitBackoffs = 0; continue; }
          await sleep(Math.min(4000 * rateLimitBackoffs, 30000));
        } else {
          attempts++;
          if (attempts >= 3) { text = `[تعذّر استخراج هذه الصفحة: ${err.message}]`; }
          else await sleep(1500);
        }
      }
    }
    pageTexts.push(`--- الصفحة ${i + 1} ---\n${text}`);
    console.log(`  📄 صفحة ${i + 1}/${imagePaths.length} تم استخراجها`);
    await sleep(pacingMs());
  }
  return pageTexts.join('\n\n');
}

/* ================= المعالجة الكاملة لملف واحد ================= */

async function processOnePdf(pdfPath) {
  const baseName = path.basename(pdfPath, path.extname(pdfPath));
  console.log(`\n========== معالجة: ${baseName} ==========`);

  console.log('🖼️  جارٍ تحويل الصفحات لصور...');
  const pageDir = path.join(WORK_DIR, baseName);
  const imagePaths = convertPdfToImages(pdfPath, pageDir);
  console.log(`   عدد الصفحات: ${imagePaths.length}`);

  console.log('🔎 جارٍ استخراج النص عبر Gemini...');
  const fullText = await extractTextFromImages(imagePaths);
  fs.writeFileSync(path.join(OUTPUT_DIR, `${baseName}-text.txt`), fullText, 'utf-8');
  console.log('   ✅ تم حفظ النص المستخرج');

  console.log('🌳 جارٍ تحليل الهيكل الهرمي...');
  const parsedBook = parseBook(fullText);
  console.log(`   ${parsedBook.stats.section} قسم، ${parsedBook.stats.part} باب، ${parsedBook.stats.chapter} فصل، ${parsedBook.stats.topic} مبحث، ${parsedBook.leaves.length} جزء سيُشرح`);
  if (parsedBook.leaves.length === 0) {
    console.warn('   ⚠️ لم يتم العثور على أي عناوين بصيغة "الباب الأول" ونحوها — تأكد من صيغة عناوين الكتاب.');
  }
  assignPathLabels(parsedBook.root);

  console.log('✍️  جارٍ توليد الشرح والأسئلة لكل جزء...');
  const aiResults = {};
  for (let i = 0; i < parsedBook.leaves.length; i++) {
    const unit = parsedBook.leaves[i];
    unit.pathLabel = unit.node.pathLabel || unit.title;
    let done = false, attempts = 0, rateLimitBackoffs = 0;
    while (!done) {
      if (keyIdx >= API_KEYS.length) throw new Error(`نفدت حدود كل المفاتيح عند الجزء ${i + 1}`);
      try {
        const parsed = await callGeminiForUnit(unit, currentKey(), MODEL, WORD_COUNT_OPTION, MCQ_COUNT, ESSAY_COUNT);
        aiResults[unit.id] = {
          fail: false,
          explanation: parsed.explanation || '',
          mcq: Array.isArray(parsed.mcq) ? parsed.mcq : [],
          essay: Array.isArray(parsed.essay) ? parsed.essay : []
        };
        done = true;
      } catch (err) {
        if (isKeyExhaustedError(err.status, err.message)) { keyIdx++; attempts = 0; rateLimitBackoffs = 0; continue; }
        if (err.status === 429) {
          rateLimitBackoffs++;
          if (rateLimitBackoffs > 6) { keyIdx++; rateLimitBackoffs = 0; continue; }
          await sleep(Math.min(4000 * rateLimitBackoffs, 30000));
        } else {
          attempts++;
          if (attempts >= 2) { aiResults[unit.id] = { fail: true, note: err.message, explanation: '', mcq: [], essay: [] }; done = true; }
          else await sleep(1200);
        }
      }
    }
    const status = aiResults[unit.id].fail ? '✗ فشل' : '✓';
    console.log(`  ${i + 1}/${parsedBook.leaves.length} ${status} ${unit.pathLabel}`);
    await sleep(pacingMs());
  }

  console.log('📦 جارٍ بناء ملف Word...');
  const docxBuffer = await buildDocx(parsedBook, aiResults, baseName);
  const outPath = path.join(OUTPUT_DIR, `${baseName}-شرح-مبسط.docx`);
  fs.writeFileSync(outPath, docxBuffer);
  console.log(`✅ تم إنشاء: ${outPath}`);
}

async function main() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.mkdirSync(WORK_DIR, { recursive: true });

  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`❌ مجلد ${INPUT_DIR} غير موجود.`);
    process.exit(1);
  }
  const pdfs = fs.readdirSync(INPUT_DIR).filter(f => f.toLowerCase().endsWith('.pdf'));
  if (pdfs.length === 0) {
    console.error(`❌ لا يوجد أي ملف PDF داخل مجلد input/. ارفع ملف PDF هناك أولاً.`);
    process.exit(1);
  }

  console.log(`النموذج: ${MODEL} | طول الشرح: ${WORD_COUNT_OPTION} | أسئلة اختيار: ${MCQ_COUNT} | أسئلة مقالية: ${ESSAY_COUNT}`);
  console.log(`عدد ملفات PDF المطلوب معالجتها: ${pdfs.length}`);

  for (const pdf of pdfs) {
    await processOnePdf(path.join(INPUT_DIR, pdf));
  }

  console.log('\n🎉 اكتملت معالجة كل الملفات بنجاح.');
}

main().catch(err => {
  console.error('\n❌ فشل غير متوقع:', err.message || err);
  process.exit(1);
});
