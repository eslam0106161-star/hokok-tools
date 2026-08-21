'use strict';
/**
 * يقرأ ملف/ملفات .docx (الناتجة من أداة "شرح الكتاب") من مجلد input-docx/
 * ويرفع الهيكل الهرمي + الشرح + الأسئلة مباشرة إلى Firestore — بدون المرور
 * بتطبيق الموبايل خالص. مصمم للعمل داخل GitHub Actions (سحابي بالكامل).
 *
 * بيطابق تمامًا نفس منطق:
 *  - DocxTextExtractor.kt (استخراج نص Markdown من عناوين Word)
 *  - MarkdownOutlineParser.kt (بناء الهيكل الهرمي)
 *  - LeafContentTextParser.kt (استخراج الشرح والأسئلة الجاهزة)
 * ونفس مخطط بيانات Firestore اللي التطبيق نفسه بيستخدمه بالضبط.
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const admin = require('firebase-admin');

const ROOT_DIR = path.join(__dirname, '..');
const INPUT_DIR = path.join(ROOT_DIR, 'input-docx');
const YEAR_LEVEL = parseInt(process.env.YEAR_LEVEL || '1', 10);
const SUBJECT_NAME_OVERRIDE = (process.env.SUBJECT_NAME || '').trim();
const BATCH_LIMIT = 400;

/* ================= تهيئة Firebase Admin ================= */

const serviceAccountJson = process.env.FIREBASE_SERVICE_ACCOUNT;
if (!serviceAccountJson) {
  console.error('❌ لا يوجد FIREBASE_SERVICE_ACCOUNT. أضفه كـ Secret في إعدادات الريبو (محتوى ملف JSON بالكامل).');
  process.exit(1);
}
let serviceAccount;
try {
  serviceAccount = JSON.parse(serviceAccountJson);
} catch (e) {
  console.error('❌ FIREBASE_SERVICE_ACCOUNT ليس JSON صالحًا.');
  process.exit(1);
}
admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

/* ================= استخراج نص Markdown من docx (مطابق لـ DocxTextExtractor.kt) ================= */

async function extractMarkdownFromDocx(filePath) {
  const data = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(data);
  const documentXmlFile = zip.file('word/document.xml');
  if (!documentXmlFile) {
    throw new Error('الملف ليس docx صالحًا — لا يوجد word/document.xml بداخله.');
  }
  const xml = await documentXmlFile.async('string');

  const paragraphs = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  const lines = [];
  for (const p of paragraphs) {
    const styleMatch = p.match(/<w:pStyle\s+w:val="([^"]+)"/);
    let level = null;
    if (styleMatch) {
      const m = /^Heading(\d)$/i.exec(styleMatch[1]);
      if (m) level = parseInt(m[1], 10);
    }
    const textMatches = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)];
    let text = textMatches.map(m => m[1]).join('');
    text = text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
    text = text.trim();
    if (!text) continue;
    if (level) {
      lines.push('#'.repeat(Math.min(Math.max(level, 1), 6)) + ' ' + text);
    } else {
      lines.push(text);
    }
  }
  const markdown = lines.join('\n\n');
  if (!markdown.trim()) {
    throw new Error('لم يتم العثور على أي نص قابل للقراءة داخل ملف الـ docx.');
  }
  return markdown;
}

/* ================= بناء الهيكل الهرمي (مطابق لـ MarkdownOutlineParser.kt) ================= */

const LEVEL_ORDER = ['SECTION', 'PART', 'CHAPTER', 'TOPIC', 'SUBTOPIC'];

function parseMarkdownOutline(markdown) {
  const headingRegex = /^(#{1,6})\s+(.*)$/gm;
  const rawDepths = new Set();
  let m;
  while ((m = headingRegex.exec(markdown))) rawDepths.add(m[1].length);
  const sortedDepths = [...rawDepths].sort((a, b) => a - b);
  const depthToLevel = {};
  sortedDepths.forEach((d, i) => { depthToLevel[d] = LEVEL_ORDER[Math.min(i, LEVEL_ORDER.length - 1)]; });

  const roots = [];
  const stack = [];
  const lineHeadingRegex = /^(#{1,6})\s+(.*)$/;

  for (const line of markdown.split('\n')) {
    const hm = lineHeadingRegex.exec(line);
    if (hm) {
      const rawDepth = hm[1].length;
      const title = hm[2].trim();
      const level = depthToLevel[rawDepth] || 'SUBTOPIC';
      const node = { level, name: title, ownText: '', children: [] };
      while (stack.length && stack[stack.length - 1].rawDepth >= rawDepth) stack.pop();
      if (!stack.length) roots.push(node);
      else stack[stack.length - 1].node.children.push(node);
      stack.push({ rawDepth, node });
    } else if (stack.length && line.trim()) {
      const current = stack[stack.length - 1].node;
      current.ownText = current.ownText ? current.ownText + '\n' + line.trim() : line.trim();
    }
  }
  return roots;
}

function collectLeaves(nodes, out) {
  for (const node of nodes) {
    if (node.children.length === 0) out.push(node);
    else collectLeaves(node.children, out);
  }
}

/* ================= استخراج الشرح والأسئلة الجاهزة (مطابق لـ LeafContentTextParser.kt) ================= */

const MCQ_HEADER = /^أسئلة\s+اختيار\s+من\s+متعدد\s*$/;
const ESSAY_HEADER = /^أسئلة\s+مقالية\s+للمراجعة\s*$/;
const QUESTION_NUM = /^(\d+)[.)]\s*(.+)$/;
const OPTION_LINE = /^([أبجد])[)\.]\s*(.+)$/;
const CORRECT_ANSWER = /^✔?\s*الإجابة\s+الصحيحة\s*:\s*([أبجد])/;
const LETTERS = ['أ', 'ب', 'ج', 'د'];

function parseLeafContent(ownText) {
  const lines = ownText.split('\n').map(l => l.trim()).filter(Boolean);
  const explanationLines = [];
  const mcq = [];
  const essay = [];
  let section = 0;
  let pendingQ = null;
  let pendingOpts = [];

  function flushMcq(correctLetter) {
    if (pendingQ && pendingOpts.length) {
      const idx = LETTERS.indexOf(correctLetter);
      const correctText = (idx >= 0 && idx < pendingOpts.length) ? pendingOpts[idx] : '';
      mcq.push({ text: pendingQ, options: pendingOpts.slice(), correctAnswer: correctText });
    }
    pendingQ = null; pendingOpts = [];
  }

  for (const line of lines) {
    if (MCQ_HEADER.test(line)) { section = 1; continue; }
    if (ESSAY_HEADER.test(line)) { if (section === 1) flushMcq(null); section = 2; continue; }
    if (section === 0) { explanationLines.push(line); continue; }
    if (section === 1) {
      const cm = CORRECT_ANSWER.exec(line);
      const om = OPTION_LINE.exec(line);
      const qm = QUESTION_NUM.exec(line);
      if (cm) flushMcq(cm[1]);
      else if (om) pendingOpts.push(om[2]);
      else if (qm) { flushMcq(null); pendingQ = qm[2]; pendingOpts = []; }
      else if (pendingQ) pendingQ += ' ' + line;
      continue;
    }
    if (section === 2) {
      const qm = QUESTION_NUM.exec(line);
      if (qm) essay.push(qm[2]);
      else if (essay.length) essay[essay.length - 1] += ' ' + line;
    }
  }
  if (section === 1) flushMcq(null);
  return { explanation: explanationLines.join('\n\n'), mcq, essay };
}

/* ================= الرفع لـ Firestore (Batch Writes، نفس مخطط بيانات التطبيق) ================= */

async function findOrCreateSubject(name, yearLevel) {
  const snap = await db.collection('subjects')
    .where('name', '==', name)
    .where('yearLevel', '==', yearLevel)
    .limit(1).get();
  if (!snap.empty) {
    return { id: snap.docs[0].id, created: false };
  }
  const ref = db.collection('subjects').doc();
  await ref.set({ id: ref.id, name, yearLevel, createdAt: Date.now() });
  return { id: ref.id, created: true };
}

async function uploadOutline(subjectId, roots) {
  const leaves = [];
  let batch = db.batch();
  let opCount = 0;

  async function flush(force) {
    if (opCount > 0 && (force || opCount >= BATCH_LIMIT)) {
      await batch.commit();
      batch = db.batch();
      opCount = 0;
    }
  }

  async function walk(nodes, parentId) {
    let order = 0;
    for (const node of nodes) {
      order++;
      const ref = db.collection('subjects').doc(subjectId).collection('curriculumNodes').doc();
      const nodeId = ref.id;
      const hasChildren = node.children.length > 0;
      const data = {
        id: nodeId,
        subjectId,
        parentId: parentId || null,
        level: node.level,
        name: node.name,
        order,
        hasChildren,
        explanation: null
      };
      batch.set(ref, data);
      batch.set(db.collection('curriculumNodes').doc(nodeId), data);
      opCount += 2;
      await flush(false);

      if (hasChildren) {
        await walk(node.children, nodeId);
      } else {
        leaves.push({ nodeId, ownText: node.ownText, name: node.name });
      }
    }
  }

  await walk(roots, null);
  await flush(true);
  return leaves;
}

async function uploadLeavesContent(subjectId, leaves) {
  let batch = db.batch();
  let opCount = 0;
  let filled = 0;
  let totalMcq = 0, totalEssay = 0;

  async function flush(force) {
    if (opCount > 0 && (force || opCount >= BATCH_LIMIT)) {
      await batch.commit();
      batch = db.batch();
      opCount = 0;
    }
  }

  for (const leaf of leaves) {
    const parsed = parseLeafContent(leaf.ownText);
    const explanation = {
      lessonId: leaf.nodeId,
      sections: parsed.explanation ? [{ id: 'sec_1', order: 1, title: '', type: 'text', content: parsed.explanation }] : [],
      updatedAt: Date.now()
    };

    batch.update(db.collection('subjects').doc(subjectId).collection('curriculumNodes').doc(leaf.nodeId), { explanation });
    batch.update(db.collection('curriculumNodes').doc(leaf.nodeId), { explanation });
    opCount += 2;
    await flush(false);

    let qOrder = 0;
    for (const q of parsed.mcq) {
      qOrder++;
      const qRef = db.collection('questions').doc();
      batch.set(qRef, {
        id: qRef.id, lessonId: leaf.nodeId, curriculumNodeId: leaf.nodeId,
        type: 'MCQ', text: q.text, options: q.options, correctAnswer: q.correctAnswer,
        explanation: '', order: qOrder
      });
      opCount++;
      await flush(false);
      totalMcq++;
    }
    for (const eq of parsed.essay) {
      qOrder++;
      const qRef = db.collection('questions').doc();
      batch.set(qRef, {
        id: qRef.id, lessonId: leaf.nodeId, curriculumNodeId: leaf.nodeId,
        type: 'essay', text: eq, options: [], correctAnswer: '',
        explanation: '', order: qOrder
      });
      opCount++;
      await flush(false);
      totalEssay++;
    }
    filled++;
    if (filled % 10 === 0) console.log(`   ${filled}/${leaves.length} جزء...`);
  }

  await flush(true);
  return { filled, totalMcq, totalEssay };
}

/* ================= المعالجة الكاملة لملف واحد ================= */

async function processOneDocx(filePath) {
  const baseName = path.basename(filePath, path.extname(filePath));
  const subjectName = SUBJECT_NAME_OVERRIDE || baseName;
  console.log(`\n========== معالجة: ${baseName} ==========`);
  console.log(`اسم المادة: "${subjectName}" | الفرقة: ${YEAR_LEVEL}`);

  console.log('📄 جارٍ استخراج نص الملف...');
  const markdown = await extractMarkdownFromDocx(filePath);
  console.log(`   طول النص: ${markdown.length} حرف`);

  console.log('🌳 جارٍ تحليل الهيكل الهرمي (بدون AI)...');
  const roots = parseMarkdownOutline(markdown);
  if (roots.length === 0) {
    throw new Error('لم يتم العثور على أي عناوين (Heading) داخل الملف — تأكد من استخدام أنماط العناوين في Word.');
  }
  const leaves = [];
  collectLeaves(roots, leaves);
  console.log(`   ${leaves.length} جزء ورقي جاهز`);

  console.log('🔎 جارٍ البحث عن المادة أو إنشائها في Firestore...');
  const subject = await findOrCreateSubject(subjectName, YEAR_LEVEL);
  console.log(`   subjectId: ${subject.id} (${subject.created ? 'مادة جديدة' : 'مادة موجودة بالفعل — هيُضاف لها هيكل جديد'})`);

  console.log('📤 جارٍ رفع الهيكل الهرمي...');
  const uploadedLeaves = await uploadOutline(subject.id, roots);
  console.log(`   ✅ تم رفع الهيكل بالكامل`);

  console.log('✍️  جارٍ رفع الشرح والأسئلة لكل جزء...');
  const result = await uploadLeavesContent(subject.id, uploadedLeaves);
  console.log(`✅ اكتملت المعالجة: ${result.filled} جزء | ${result.totalMcq} سؤال اختيار | ${result.totalEssay} سؤال مقالي`);
}

async function main() {
  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`❌ مجلد ${INPUT_DIR} غير موجود.`);
    process.exit(1);
  }
  const docxFiles = fs.readdirSync(INPUT_DIR).filter(f => f.toLowerCase().endsWith('.docx'));
  if (docxFiles.length === 0) {
    console.error('❌ لا يوجد أي ملف .docx داخل مجلد input-docx/. ارفع الملف هناك أولاً.');
    process.exit(1);
  }

  for (const f of docxFiles) {
    await processOneDocx(path.join(INPUT_DIR, f));
  }
  console.log('\n🎉 اكتمل رفع كل الملفات مباشرة إلى قاعدة البيانات بنجاح.');
}

main().catch(err => {
  console.error('\n❌ فشل غير متوقع:', err.message || err);
  process.exit(1);
});
