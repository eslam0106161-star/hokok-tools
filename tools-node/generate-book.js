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
const TF_COUNT = parseInt(process.env.TF_COUNT || '5', 10);
const SKIP_FOOTNOTES = (process.env.SKIP_FOOTNOTES || 'yes') === 'yes';
const PDF_DPI = parseInt(process.env.PDF_DPI || '150', 10);

// مراحل التحقق: كل مرحلة بتضيف استدعاء Gemini إضافي (تكلفة/وقت إضافي)، لكنها
// بتزوّد دقة الاستخراج والمحتوى بشكل ملموس. ممكن تتقفل عبر الـ Secrets/Env لو
// الوقت أو حصة الـ API ضيقة.
const VERIFY_EXTRACTION = (process.env.VERIFY_EXTRACTION || 'yes') === 'yes';
const VERIFY_CONTENT = (process.env.VERIFY_CONTENT || 'yes') === 'yes';
// مرحلة تصنيف ذكي للعناوين الغامضة (أولاً/ثانياً، أو عنوان بصري) — استدعاء
// واحد رخيص لكل كتاب كامل (مش لكل صفحة أو جزء)، بيراجع مستوى العناوين دي
// بناءً على فهم النموذج للسياق الكامل للكتاب بدل الاعتماد على تخمين محلي بسيط.
const VERIFY_STRUCTURE = (process.env.VERIFY_STRUCTURE || 'yes') === 'yes';

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
  const basePrompt = `استخرج كل النص الموجود في هذه الصورة بدقة تامة كما هو مكتوب، بنفس اللغة. حافظ على فواصل الفقرات والأسطر قدر الإمكان.

تنبيه مهم بخصوص العناوين: كتب القانون غالبًا فيها عناوين وعناوين فرعية متعددة المستويات (باب، فصل، مبحث، مطلب...)، لكن كتير من الكتب بتكتب عناوين مهمة بصريًا (خط عريض bold، أو حجم خط أكبر من باقي الفقرة، أو سطر متوسط في نص الصفحة) من غير أي كلمة زي "باب" أو "فصل" أو "مبحث" قبلها — أمثلة: عنوان مستقل قائم بذاته، أو عنوان مسبوق بكلمة "تمهيد" أو "مقدمة" أو "فصل تمهيدي"، أو عنوان مسبوق بـ"أولاً"/"ثانياً"/"ثالثاً". أي سطر من دول لازم تحطله العلامة "[HEADING] " في أول السطر (قبل نص العنوان نفسه مباشرة) في النص المستخرج، **فقط لو كان شكله البصري في الصورة فعلاً مختلف عن باقي الفقرة** (خط عريض، أو حجم أكبر، أو متوسط الصفحة، أو سطر مستقل قبل فقرة جديدة تمامًا) — لا تحط العلامة دي على جمل عادية في وسط فقرة حتى لو بدأت بكلمة "أولاً" وهي جزء من جملة عادية (مثال: "أولاً، يجب أن نلاحظ أن..." وسط فقرة مش عنوان، بينما "أولاً: نظرية كذا" في سطر مستقل هو عنوان).

لا تضف أي شرح أو تعليق أو مقدمة غير مطلوبة، أعد النص المستخرج فقط (مع علامات [HEADING] المطلوبة). إذا لم يوجد نص، أعد سطراً فارغاً.`;
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

/**
 * مرحلة تحقق 1: يراجع النموذج النص المستخرج بالرجوع لنفس صورة الصفحة مرة ثانية،
 * بالتركيز تحديداً على الأرقام (أرقام مواد قانونية، تواريخ، نسب، أرقام ترتيبية)
 * وأسماء الأعلام، ويصحح فقط أخطاء الاستخراج دون أي إضافة أو حذف أو إعادة صياغة.
 * يرجع { text, changed } — changed=true لو النص اتغيّر فعلاً بعد المراجعة.
 */
async function callGeminiVerifyExtraction(base64Jpeg, extractedText, apiKey, model) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const promptText = `أنت مدقق نصوص صارم. معك صورة صفحة ونص تم استخراجه منها سابقاً عبر OCR. راجع النص المستخرج بمقارنته حرفياً بما هو مكتوب في الصورة، بالتركيز الخاص على:
- أرقام المواد القانونية والفصول والبنود
- التواريخ والسنوات والنسب المئوية والأرقام الترتيبية
- أسماء الأعلام والمصطلحات القانونية الدقيقة
صحّح فقط أخطاء الاستخراج (رقم أو كلمة أو حرف غلط عن الصورة). ممنوع منعاً باتاً: إضافة أي معلومة غير موجودة بالصورة، حذف أي جزء من النص الأصلي، أو إعادة صياغة أي جملة سليمة بأسلوب مختلف.
النص قد يحتوي على علامات "[HEADING] " في أول بعض الأسطر (تشير إلى عناوين بصرية بارزة اكتُشفت مسبقًا) — حافظ عليها بالضبط كما هي في مكانها، لا تحذفها ولا تضف علامات جديدة ولا تنقلها لسطر آخر، فهي ليست جزءًا من نص الصورة نفسه بل علامة تنظيمية مضافة مسبقًا.
لو النص المستخرج مطابق تماماً لما في الصورة، أعده كما هو دون أي تغيير.
أعد النص الكامل بعد المراجعة فقط، بدون أي تعليق أو شرح أو مقدمة.

النص المستخرج المطلوب مراجعته:
"""
${extractedText}
"""`;
  const body = {
    contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: 'image/jpeg', data: base64Jpeg } }] }],
    generationConfig: { temperature: 0 }
  };
  const data = await geminiFetch(url, body);
  const text = (data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '').trim();
  return { text: text || extractedText, changed: text.trim() !== extractedText.trim() };
}

/**
 * تعقيم دفاعي: النموذج أحيانًا (نادرًا لكن فعليًا) بيرجّع عنصر أسئلة مقالية أو
 * MCQ بشكل مختلف شوية عن الـ schema المطلوب (مثلاً object بدل string، أو حقل
 * بمفتاح مختلف). بدون هذا التعقيم، الفرق ده بيوصل صامتًا لملف Word النهائي
 * كـ "[object Object]" حرفيًا. الدالة دي بتحوّل أي انحراف بسيط لصيغة سليمة،
 * وبتسجّل تحذير في اللوج لو حصل تحويل فعلي عشان تعرف إن فيه مشكلة تستحق المتابعة.
 */
function sanitizeQAContent(raw, unitLabel) {
  const warnings = [];
  const asText = (v, fallbackKeys) => {
    if (typeof v === 'string') return v;
    if (v && typeof v === 'object') {
      for (const k of (fallbackKeys || ['question', 'text', 'content'])) {
        if (typeof v[k] === 'string' && v[k].trim()) return v[k];
      }
    }
    if (v == null) return '';
    return String(v);
  };

  const explanation = asText(raw.explanation);

  const mcq = (Array.isArray(raw.mcq) ? raw.mcq : []).map((q, i) => {
    if (!q || typeof q !== 'object') { warnings.push(`mcq[${i}] ليس object صالحًا`); return null; }
    const question = asText(q.question, ['question', 'text']);
    let choices = Array.isArray(q.choices) ? q.choices.map(c => asText(c)) : [];
    // عدد الخيارات مسموح يتراوح بين 2 و4 (زي الامتحانات الحقيقية بالظبط —
    // مش كل سؤال لازم يبقى له 4 اختيارات)، لكن أقل من 2 أو أكتر من 4 غير منطقي.
    if (choices.length < 2) {
      warnings.push(`mcq[${i}] عدد خياراته ${choices.length} أقل من 2`);
      while (choices.length < 2) choices.push('(خيار مفقود)');
    }
    if (choices.length > 4) {
      warnings.push(`mcq[${i}] عدد خياراته ${choices.length} أكتر من 4`);
      choices = choices.slice(0, 4);
    }
    let correctIndex = Number.isInteger(q.correct_index) ? q.correct_index : parseInt(q.correct_index, 10);
    if (!Number.isInteger(correctIndex) || correctIndex < 0 || correctIndex >= choices.length) {
      warnings.push(`mcq[${i}] correct_index غير صالح (${q.correct_index}) — تم ضبطه على 0`);
      correctIndex = 0;
    }
    return { question, choices, correct_index: correctIndex };
  }).filter(Boolean);

  const essay = (Array.isArray(raw.essay) ? raw.essay : []).map((q, i) => {
    const t = asText(q, ['question', 'text']);
    if (typeof q !== 'string') warnings.push(`essay[${i}] لم يكن نصًا (${typeof q}) — تم تحويله`);
    return t;
  }).filter(t => t.trim());

  const tf = (Array.isArray(raw.tf) ? raw.tf : []).map((q, i) => {
    if (!q || typeof q !== 'object') { warnings.push(`tf[${i}] ليس object صالحًا`); return null; }
    const statement = asText(q.statement, ['statement', 'question', 'text']);
    if (!statement.trim()) { warnings.push(`tf[${i}] بدون نص عبارة`); return null; }
    let answer;
    if (typeof q.answer === 'boolean') answer = q.answer;
    else if (typeof q.answer === 'string') answer = /^(true|صح|صحيح|1)$/i.test(q.answer.trim());
    else { warnings.push(`tf[${i}] قيمة answer غير صالحة (${q.answer}) — تم ضبطها على false`); answer = false; }
    return { statement, answer };
  }).filter(Boolean);

  if (warnings.length) {
    console.warn(`  ⚠️  تعقيم بيانات في "${unitLabel}": ${warnings.join(' | ')}`);
  }

  return { explanation, mcq, essay, tf, hadWarnings: warnings.length > 0 };
}

const wordCountText = { short: 'حوالي 150 كلمة', medium: 'حوالي 300 كلمة', long: 'حوالي 450 كلمة' };

async function callGeminiForUnit(unit, apiKey, model, wcOption, mcqCount, essayCount, tfCount, siblingContext) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const pathLabel = unit.pathLabel || unit.title;
  siblingContext = siblingContext || {};

  let continuityBlock = '';
  if (siblingContext.prevTitle || siblingContext.nextTitle) {
    continuityBlock = '\nسياق مكان هذا الجزء داخل الكتاب (للانتباه لتفادي التكرار أو التناقض — وليس مصدرًا للمعلومات):\n';
    if (siblingContext.prevTitle) {
      continuityBlock += `- الجزء السابق مباشرة كان بعنوان "${siblingContext.prevTitle}"`;
      if (siblingContext.prevExcerpt) continuityBlock += ` وتناول ما يلي (ملخص فقط): "${siblingContext.prevExcerpt}"`;
      continuityBlock += '\n';
    }
    if (siblingContext.nextTitle) {
      continuityBlock += `- الجزء التالي مباشرة بعد هذا الجزء بعنوان "${siblingContext.nextTitle}" (لا تستبق شرحه أو تجب على أسئلة تخصه).\n`;
    }
    continuityBlock += 'استخدم هذا السياق فقط لتفادي تكرار نفس الشرح حرفيًا أو الوقوع في تناقض مع الجزء السابق (مثلاً في تعريف مصطلح مشترك أو ترقيم)، لكن لا تعتمد عليه كمصدر معلومات، ولا تشر إليه صراحة في نص الشرح أو الأسئلة.\n';
  }

  const prompt = `أنت أستاذ جامعي متخصص في القانون، مهمتك مساعدة طالب حقوق يذاكر لامتحان فعلي — الأولوية القصوى هي إفادة الطالب أكاديميًا، وليس مجرد تلخيص النص.

عنوان هذا الجزء ضمن الكتاب: "${pathLabel}"
${continuityBlock}
النص الأصلي لهذا الجزء (المصدر الوحيد المسموح للشرح والأسئلة):
"""
${unit.text}
"""

المطلوب أولاً — الشرح:
اكتب شرحًا أكاديميًا واضحًا بالعربية الفصحى السهلة (${wordCountText[wcOption]})، بدون الإخلال بالدقة القانونية، وبدون إضافة معلومات من خارج النص المعطى. انسخ أي رقم (رقم مادة، تاريخ، نسبة) كما ورد بالنص الأصلي حرفيًا دون تقريب أو تخمين.

قواعد إلزامية لصياغة الشرح:
- **أعد بناء الشرح من الصفر بترتيب وصياغة مختلفين تمامًا عن النص الأصلي** — ممنوع إعادة استخدام نفس الجمل أو نفس ترتيب الأفكار حرفيًا حتى لو بكلمات مرادفة قليلاً؛ اكتبه كأنك بتشرحه لطالب من فهمك أنت للفكرة، مش بتلخّص الفقرة.
- **لو النص الأصلي منظّم بتعداد واضح** (شروط أو خصائص أو أقسام مرقّمة بأ/ب/ج أو أولاً/ثانياً/ثالثاً أو أرقام)، **حافظ على نفس بنية التعداد** في الشرح (نقطة مستقلة لكل عنصر على سطر جديد) بدل ما تذوّبها في فقرة نثرية متصلة — التعداد ده بنية مفيدة للمذاكرة مش مجرد أسلوب كتابة.
- قدّم التعريفات أولاً، ثم الشروط أو الأركان، ثم الأمثلة أو التطبيقات إن وردت في النص الأصلي (لا تخترع أمثلة غير موجودة).
- اكتب الشرح بالعربية فقط، حتى لو ورد في النص الأصلي مصطلح أجنبي (فرنسي أو إنجليزي أو لاتيني) جنب المقابل العربي — لا تُدرج المصطلح الأجنبي في الشرح إطلاقًا، اكتفِ بالمصطلح العربي فقط.

ملاحظات توضيح عامية ("خد بالك"):
بعد أي جملة أو نقطة أكاديمية فيها مصطلح مركّب، أو شرط/استثناء سهل الالتباس (زي الفرق بين "يجوز" و"يجب")، أو فكرة مجردة محتاجة تقريب، **زوّد سطرًا مستقلاً بعدها مباشرة** بالصيغة الحرفية التالية بالظبط:
💡 خد بالك: [توضيح مبسط بالعامية المصرية للفكرة اللي فاتت، بأسلوب قريب وودود، من غير ما تضيف معلومة قانونية جديدة مش موجودة في الجملة الأكاديمية نفسها]
لا تحط ملاحظة "خد بالك" بعد كل جملة — بس عند النقط اللي فعلاً محتاجة تبسيط، وسيب باقي الشرح من غيرها. ممنوع خلط العامية جوه الجملة الأكاديمية نفسها؛ العامية تكون بس في سطر "💡 خد بالك" المستقل ده.

المطلوب ثانياً — أسئلة الاختيار من متعدد (بالضبط ${mcqCount} أسئلة):
عدد الخيارات لكل سؤال **مش لازم يكون 4 دايمًا** — استخدم 2 أو 3 أو 4 خيارات حسب طبيعة السؤال نفسه (بالظبط زي الامتحانات الحقيقية): سؤال بديهي بين احتمالين يكفيه خياران، سؤال فيه بديل ثالث منطقي زوّده لثلاثة، وسؤال محتاج تمييز دقيق بين عدة مفاهيم استخدم أربعة.

وزّع الأسئلة على القوالب الآتية حسب ما يناسب مضمون هذا الجزء تحديدًا (لا تستخدم قالبًا لا ينطبق على النص، ولا يلزم استخدام كل قالب):
- تعريف: يختبر التعريف اللغوي أو الاصطلاحي لمصطلح ورد في النص، أو استكمال تعريف ناقص ("... هو كل عمل مبتكر أدبي أو فني...").
- تفرقة: يختبر الفرق بين مفهومين متقاربين وردا في النص أو يمكن استنتاجهما منه.
- شرط/استثناء دقيق: يختبر حكمًا شرطيًا بصيغة دقيقة زي "يجوز/لا يجوز"، "يشترط/لا يشترط"، "يلزم/لا يلزم" — بحيث يوقع الطالب اللي فاهم الفكرة نص فهم في الفخ، بنفس الدقة اللي بتتاخد بيها أسئلة الصح والخطأ في الامتحانات الحقيقية (استلهم الدقة من صياغات زي "لا يشترط إذن كتابي من المؤلف للقيام بالتحوير" كنموذج للمستوى المطلوب من الدقة، مش كنص تنسخه).
- قصة قانونية قصيرة (حالة افتراضية): اعرض سيناريو واقعي قصير (2-4 أسطر) بأسماء أشخاص افتراضية، ثم اسأل عن الحكم القانوني الصحيح بناءً عليه — ممكن يتبعه سؤال فرعي تاني مبني على نفس القصة (زي "وفقاً للنقطة السابقة، ما المحكمة المختصة؟").
- الرأي الراجح: يختبر أي الآراء أو المذاهب هو الراجح أو المعتمد، إن ورد ترجيح صريح في النص.
- موقف القانون/القضاء المصري: إن ورد في النص إشارة لموقف المشرّع المصري أو القضاء تحديدًا، اسأل عنه.

خيارات كل سؤال يجب أن تكون كلها معقولة ومتقاربة في الصياغة (تجنّب الخيارات الساذجة أو الواضح خطؤها)، بحيث يحتاج الطالب فعلاً للفهم لا للتخمين.

**قاعدة مهمة بخصوص الأرقام (طبّقها في الاختيار من متعدد والمقالي معًا)**:
فرّق بين نوعين من الأرقام:
- **رقم استشهاد** (رقم مادة قانونية، رقم قانون، رقم فقرة) — **قلّل الأسئلة اللي محورها الأساسي حفظ الرقم ده لأقصى درجة**؛ ده نادرًا جدًا ما يُسأل عنه مباشرة في الامتحانات الحقيقية. لو احتجت تذكره، اذكره كجزء من سياق السؤال مش كموضوعه ("نصت المادة على كذا... فما الحكم؟" — السؤال عن الحكم مش عن رقم المادة).
- **رقم حكم قانوني جوهري** (مدة تقادم، مدة حماية، نصاب، نسبة، مهلة قانونية) — ده **رقم من صلب المادة العلمية** ويُسأل عنه بشكل طبيعي زي أي فكرة تانية، لأنه هو نفسه الحكم القانوني المطلوب حفظه وفهمه (زي "تنقضي الحقوق المالية بعد 50 سنة من الوفاة" — ده مش تفصيل هامشي، ده صلب الحكم).
لو مش متأكد رقم معين من أنهي نوع، اعتبره "رقم استشهاد" وقلل الاعتماد عليه.

المطلوب ثالثاً — الأسئلة المقالية (بالضبط ${essayCount} أسئلة):
اكتب الأسئلة بنفس الصياغة والأسلوب المعتمد فعليًا في امتحانات نهاية العام بكليات الحقوق المصرية — وليس بصياغة عامة أو مبسطة. وزّع الأسئلة على القوالب الآتية حسب ما يناسب مضمون هذا الجزء تحديدًا (لا تستخدم قالبًا لا ينطبق على النص، ولا يلزم استخدام كل قالب):

١) تعريف + مذاهب/آراء + أدلة + راجح:
   "عرّف [المصطلح] لغةً واصطلاحاً، ثم اذكر مذاهب الفقهاء (أو آراء الفقه القانوني، حسب طبيعة المادة) وأدلتهم في [المسألة]، وبيّن الرأي الراجح."
٢) تطبيق على حالة افتراضية:
   "بيّن آراء الفقهاء (أو الفقه القانوني) مع ذكر الدليل فيما لو [صف حالة افتراضية واقعية مبنية على مضمون النص]."
٣) مقارنة أو تفرقة:
   "بيّن أوجه الاختلاف (أو الفرق) بين [مفهوم أ] و[مفهوم ب]."
٤) ربط بالعمل القضائي أو موقف القانون المصري:
   "اذكر مذاهب الفقهاء وأدلتهم في [المسألة]، مبيناً ما جرى عليه العمل في المحاكم" أو "...مع بيان موقف القانون المصري من ذلك."
٥) مناقشة عبارة (مناسب لمواد القانون الوضعي غير الفقهية):
   "اشرح مدى صحة العبارة الآتية: '[عبارة قانونية مستخلصة من مضمون النص]'." أو "ناقش [فكرة رئيسية وردت في النص]."
٦) سؤال مركّب من فرعين (أ) و(ب):
   سؤال واحد يضم فرعين مختلفين ضمن نطاق هذا الجزء، كل فرع بصياغة مستقلة وواضحة، بنفس القوالب أعلاه.

قواعد صياغة إلزامية:
- استخدم فقط الأفعال والصيغ الكلاسيكية المعتادة في الامتحانات: عرّف، اذكر، بيّن، وضح، ناقش، قارن، اشرح بالتفصيل — لا صيغًا حوارية أو تبسيطية.
- ممنوع استخدام عبارات مثل "بناءً على النص" أو "كما ورد في النص" أو "في هذا الجزء" — صُغ كل سؤال كأنه سؤال امتحان مستقل تمامًا كما يقابله الطالب فعليًا في ورقة الامتحان.
- إن كانت المادة فقهًا شرعيًا استخدم "مذاهب الفقهاء وأدلتهم"، وإن كانت قانونًا وضعيًا استبدلها بـ"آراء الفقه القانوني" أو بمطالبة تحليل/مناقشة مباشرة حسب طبيعة المادة.
- لا تخترع حالات افتراضية أو تفاصيل قانونية غير مدعومة بمضمون النص المعطى.
- كل سؤال يجب أن يكون قابلاً للإجابة بالاعتماد على محتوى هذا الجزء فقط.

المطلوب رابعاً — أسئلة الصح والخطأ (بالضبط ${tfCount} أسئلة):
اكتب عبارات قصيرة ودقيقة (سطر واحد لكل عبارة)، نصفها تقريبًا صحيح ونصفها الآخر خطأ (وزّع الصح والخطأ بشكل غير متوقع، مش نمط ثابت زي "كل الأزواج صح")، بنفس الدقة الشديدة المعتمدة في امتحانات كليات الحقوق المصرية الفعلية. القوالب:
- عبارة تختبر حكمًا شرطيًا دقيقًا بصيغة "يجوز/لا يجوز"، "يشترط/لا يشترط"، "يجب/لا يجب" — الفخ المطلوب هو دقة الصياغة نفسها (زي "لا يشترط إذن كتابي من المؤلف للقيام بالتحوير") مش الفكرة العامة.
- عبارة تختبر تعريفًا (صحيحًا أو محرَّفًا بتغيير كلمة جوهرية فيه).
- عبارة تنسب حكمًا أو استثناءً لطرف غلط (مثلاً تخلط بين حكمين متقابلين، أو تعمم استثناء على غير محله).
- عبارة تختبر ترتيبًا زمنيًا أو منطقيًا لخطوة إجرائية.
كل عبارة يجب أن تكون قابلة للحكم عليها (صح أو خطأ) بشكل قاطع بالاعتماد على محتوى هذا الجزء وحده فقط، من غير أي التباس أو احتمال لتفسيرين مختلفين. لا تكرر نفس الفكرة في أكتر من عبارة.

أعد النتيجة بصيغة JSON فقط، بدون أي نص أو علامات إضافية قبله أو بعده، وبدون أي إشارة داخل نص الأسئلة نفسها لكونها مأخوذة من "النص" أو "المقطع" (صُغ كل سؤال كسؤال أكاديمي مستقل تمامًا)، بالشكل التالي بالضبط:
{
  "explanation": "الشرح الكامل هنا",
  "mcq": [
    {"question": "نص سؤال بأربعة اختيارات", "choices": ["اختيار 1", "اختيار 2", "اختيار 3", "اختيار 4"], "correct_index": 0},
    {"question": "نص سؤال باختيارين بس (مثلاً بين حكمين متقابلين)", "choices": ["اختيار 1", "اختيار 2"], "correct_index": 1}
  ],
  "essay": ["سؤال مقالي تحليلي أول", "سؤال مقالي تحليلي ثانٍ"],
  "tf": [
    {"statement": "عبارة قانونية دقيقة قابلة للحكم عليها", "answer": true},
    {"statement": "عبارة قانونية أخرى محرَّفة عمدًا في نقطة دقيقة", "answer": false}
  ]
}
عدد عناصر "choices" في كل سؤال مستقل بذاته (2 أو 3 أو 4 حسب طبيعة السؤال — راجع القواعد أعلاه)، والمثال هنا لتوضيح الشكل العام فقط. يجب أن يحتوي "mcq" على ${mcqCount} عنصر بالضبط، و"essay" على ${essayCount} عنصر بالضبط، و"tf" على ${tfCount} عنصر بالضبط.`;

  const body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.4, responseMimeType: 'application/json' } };
  const data = await geminiFetch(url, body);
  let text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  text = text.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  let parsed;
  try { parsed = JSON.parse(text); }
  catch (e) { const err = new Error('تعذّر تفسير استجابة النموذج كـ JSON'); err.status = -1; throw err; }
  return sanitizeQAContent(parsed, pathLabel);
}

/**
 * مرحلة تحقق 2: يراجع النموذج (بحرارة صفر) المحتوى المولَّد مقابل النص الأصلي
 * فقط، ويصحح أي تأليف/معلومة خارجية، أي رقم غير مطابق، أو أي correct_index
 * غير صحيح فعليًا، أو أي سؤال مقالي غير قابل للإجابة من هذا الجزء وحده.
 * يرجع { result, changed, notes } — notes تلخيص مختصر لما تم تصحيحه (أو فارغة).
 */
async function callGeminiVerifyContent(unit, generated, apiKey, model, mcqCount, essayCount, tfCount) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const pathLabel = unit.pathLabel || unit.title;
  const prompt = `أنت مدقق أكاديمي صارم متخصص في القانون. مهمتك مراجعة محتوى (شرح + أسئلة) تم توليده اعتمادًا على نص أصلي واحد فقط، والتحقق من خمسة أمور بدقة:

١) عدم التأليف: هل يوجد في الشرح أو الأسئلة أي معلومة، مثال، أو تفصيل قانوني غير موجود إطلاقًا في النص الأصلي (سواء من معرفة عامة عن القانون أو استنتاج غير مباشر)؟
٢) دقة الأرقام الجوهرية: هل كل رقم يمثل حكمًا قانونيًا جوهريًا (مدة تقادم أو حماية، نصاب، نسبة، مهلة) ورد في الشرح أو الأسئلة مطابق حرفيًا لما ورد في النص الأصلي؟ (أرقام الاستشهاد المجردة زي رقم مادة أو رقم قانون أقل أهمية، فلو مش موجودة أو مش دقيقة ومفيش سؤال محوره الأساسي حفظها تحديدًا، متعتبرهاش خطأ يستأهل تصحيح).
٣) صحة إجابات الاختيار من متعدد: لكل سؤال MCQ (بغض النظر عن عدد خياراته — 2 أو 3 أو 4، مش لازم يبقوا 4 دايمًا)، تحقق فعليًا بالرجوع للنص الأصلي أن "correct_index" هو فعلاً الإجابة الصحيحة، وليس افتراضًا.
٤) قابلية الإجابة: هل كل سؤال مقالي يمكن الإجابة عليه فعلاً بالاعتماد على محتوى النص الأصلي المعطى وحده؟
٥) صحة أسئلة الصح والخطأ: لكل عبارة في "tf"، تحقق فعليًا بالرجوع للنص الأصلي أن قيمة "answer" (true أو false) مطابقة فعلاً لما تقوله العبارة، وأن العبارة نفسها قابلة للحكم عليها بشكل قاطع من غير التباس.

عنوان الجزء: "${pathLabel}"
النص الأصلي (المرجع الوحيد للتحقق):
"""
${unit.text}
"""

المحتوى المطلوب تدقيقه (JSON):
${JSON.stringify({ explanation: generated.explanation, mcq: generated.mcq, essay: generated.essay, tf: generated.tf })}

التعليمات:
- إن كان كل شيء سليمًا تمامًا، أعد نفس المحتوى دون أي تغيير — بما في ذلك عدد خيارات كل سؤال MCQ زي ما هو، لا تضيف أو تحذف خيارات لمجرد توحيد العدد.
- إن وجدت أي خطأ من الأنواع الخمسة أعلاه، صححه مباشرة: احذف/عدّل الجزء الخاطئ من الشرح، أو صحح الرقم الجوهري الخاطئ ليطابق النص الأصلي، أو صحح "correct_index"، أو استبدل السؤال المقالي غير القابل للإجابة بسؤال بديل مبني على النص الأصلي فقط ويتبع نفس أسلوب امتحانات كليات الحقوق (عرّف/اذكر/بيّن/ناقش/قارن)، أو صحح "answer" الخاطئة في أسئلة الصح والخطأ، أو استبدل عبارة صح/خطأ ملتبسة بعبارة بديلة قاطعة الحكم مبنية على النص الأصلي فقط.
- حافظ دائمًا على وجود ${mcqCount} عنصر بالضبط في "mcq" و${essayCount} عنصر بالضبط في "essay" و${tfCount} عنصر بالضبط في "tf" حتى بعد أي تصحيح.
- أضف حقل "verification_notes": مصفوفة نصية قصيرة (سطر واحد لكل ملاحظة) تلخص أي تصحيح فعلي تم إجراؤه (رقم جوهري، معلومة، إجابة MCQ، سؤال، إجابة صح/خطأ). لو ملحوظش أي خطأ اتركها مصفوفة فارغة [].

أعد النتيجة بصيغة JSON فقط بنفس البنية بالضبط (عدد الخيارات داخل "choices" حسب كل سؤال، مش ثابت):
{"explanation":"...","mcq":[{"question":"...","choices":["...","..."],"correct_index":0}],"essay":["...","..."],"tf":[{"statement":"...","answer":true}],"verification_notes":[]}`;

  const body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } };
  const data = await geminiFetch(url, body);
  let text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  text = text.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  let parsed;
  try { parsed = JSON.parse(text); }
  catch (e) { const err = new Error('تعذّر تفسير استجابة التحقق كـ JSON'); err.status = -1; throw err; }

  const notes = Array.isArray(parsed.verification_notes) ? parsed.verification_notes.filter(Boolean) : [];
  const sanitized = sanitizeQAContent(
    {
      explanation: parsed.explanation ?? generated.explanation,
      mcq: (Array.isArray(parsed.mcq) && parsed.mcq.length) ? parsed.mcq : generated.mcq,
      essay: (Array.isArray(parsed.essay) && parsed.essay.length) ? parsed.essay : generated.essay,
      tf: (Array.isArray(parsed.tf) && parsed.tf.length) ? parsed.tf : generated.tf
    },
    pathLabel
  );
  const result = { explanation: sanitized.explanation, mcq: sanitized.mcq, essay: sanitized.essay, tf: sanitized.tf };
  return { result, changed: notes.length > 0, notes };
}

/* ================= تحليل الهيكل الهرمي (بدون AI) ================= */

// ملحوظة: "الثاني"/"الحادي" ليهم تهجئتان شائعتان في الكتب العربية القديمة —
// بالياء (ي) أو بالألف المقصورة (ى، زي "الثانى" و"الحادى") — الاثنان مقبولان هنا.
const ordinalPattern = '(?:الأول|الثان[يى]|الثالث|الرابع|الخامس|السادس|السابع|الثامن|التاسع|العاشر|الحاد[يى] عشر|الثان[يى] عشر|الثالث عشر|الرابع عشر|الخامس عشر|السادس عشر|السابع عشر|الثامن عشر|التاسع عشر|العشرون)';
const levelDefs = [
  { key: 'section', label: 'القسم', arName: 'قسم' },
  { key: 'part', label: 'الباب', arName: 'باب' },
  { key: 'chapter', label: 'الفصل', arName: 'فصل' },
  { key: 'topic', label: 'المبحث', arName: 'مبحث' },
  { key: 'subtopic', label: 'المطلب', arName: 'مطلب' },
];
levelDefs.forEach(l => {
  // الصيغة 1 (الأصلية): العنوان لوحده في سطر منفصل، والعنوان الوصفي في السطر التالي.
  l.regexBare = new RegExp(`^${l.label}\\s+${ordinalPattern}\\s*$`);
  // الصيغة 2 (شائعة جدًا في كتب القانون المصرية): العنوان الوصفي في نفس السطر بعد فاصل.
  l.regexInline = new RegExp(`^${l.label}\\s+${ordinalPattern}\\s*[:：\\-–—]\\s*(\\S.*)$`);
});
function matchHeading(t) {
  for (let lvl = 0; lvl < levelDefs.length; lvl++) {
    const l = levelDefs[lvl];
    if (l.regexBare.test(t)) return { level: lvl, inlineTitle: null };
    const m = t.match(l.regexInline);
    if (m) return { level: lvl, inlineTitle: m[1].trim() };
  }
  return null;
}

// عناوين خاصة ثابتة المستوى: بتظهر في نفس صف "الباب" في فهرس الكتاب غالبًا
// (تمهيد، مقدمة، فصل تمهيدي، باب تمهيدي) لكن من غير ترقيم "الأول/الثاني".
// نديها نفس مستوى "الباب" (index 1 في levelDefs) عشان تبقى بند مستقل موازي للأبواب.
const SPECIAL_TOP_LEVEL_INDEX = 1;
const specialTopLevelRegex = /^(تمهيد|مقدمة|فصل تمهيدي|باب تمهيدي)\s*(?:[:：\-–—]\s*(\S.*))?$/;

// عناوين بصيغة "أولاً/ثانياً/ثالثاً..." — شائعة كتقسيم فرعي غير مسمى داخل فصل
// أو مبحث، من غير كلمة "فصل" أو "مبحث" قبلها. مستواها نسبي: أعمق بواحد من
// آخر عنوان معروف قبلها (مش مستوى ثابت مطلق زي باقي الأنماط).
const wordOrdinalRegex = /^(أولاً|أولا|ثانياً|ثانيا|ثالثاً|ثالثا|رابعاً|رابعا|خامساً|خامسا|سادساً|سادسا|سابعاً|سابعا|ثامناً|ثامنا|تاسعاً|تاسعا|عاشراً|عاشرا)\s*(?:[:：\-–—]\s*(\S.*))?$/;

// علامة عنوان بصري (خط عريض/حجم أكبر) اكتُشفت في مرحلة الاستخراج (OCR) ومحطوطة
// في أول السطر — بتغطي عناوين مالهاش أي كلمة مفتاحية أو ترقيم معروف خالص.
// مستواها نسبي زي "أولاً/ثانياً" بالظبط.
const ocrHeadingMarkerRegex = /^\[HEADING\]\s*(.*)$/;
function stripHeadingMarker(t) { return t.replace(/^\[HEADING\]\s*/, ''); }

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
function isAnyHeadingLine(raw) {
  const ocrMarked = /^\[HEADING\]\s*/.test(raw);
  const t = stripHeadingMarker(raw);
  if (ocrMarked) return true;
  if (matchHeading(t)) return true;
  if (specialTopLevelRegex.test(t)) return true;
  if (wordOrdinalRegex.test(t)) return true;
  return false;
}

function detectHeadingMatches(rawText) {
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
          const isHeadingLine = isAnyHeadingLine(t);
          if (t === '' || isHeadingLine) { excluded[back] = true; back--; steps++; }
          else break;
        }
      }
    }
  }

  let matches = [];
  // آخر مستوى عنوان "مطلق" (كلمة مفتاحية أو نوع خاص ثابت) — العناوين النسبية
  // (أولاً/ثانياً، والعنوان البصري) بتُبنى دايمًا على أساسه هو، مش على أساس
  // بعضها البعض، عشان "أولاً" و"ثانياً" و"ثالثاً" المتتالية تفضل إخوة على نفس
  // المستوى مش متداخلة جوه بعض.
  let lastAbsoluteLevel = null;
  for (let idx = 0; idx < n; idx++) {
    if (excluded[idx]) continue;
    const raw = lines[idx].trim();
    if (!raw) continue;
    const ocrMarked = /^\[HEADING\]\s*/.test(raw);
    const t = stripHeadingMarker(raw);
    if (!t) continue;

    // 1) كلمة مفتاحية + ترقيم (باب/فصل/مبحث/مطلب) — الأعلى موثوقية، مستوى ثابت مطلق
    const hm = matchHeading(t);
    if (hm) { matches.push({ level: hm.level, lineIdx: idx, inlineTitle: hm.inlineTitle, kind: 'keyword' }); lastAbsoluteLevel = hm.level; continue; }

    // 2) عنوان خاص ثابت (تمهيد/مقدمة/فصل تمهيدي) — بنفس مستوى "الباب"، والعنوان
    // معروف من كلمة المطابقة نفسها (مفيش سطر عنوان منفصل بعده زي باقي الأنماط).
    const sm = t.match(specialTopLevelRegex);
    if (sm) { matches.push({ level: SPECIAL_TOP_LEVEL_INDEX, lineIdx: idx, inlineTitle: sm[2] ? sm[2].trim() : sm[1], kind: 'special' }); lastAbsoluteLevel = SPECIAL_TOP_LEVEL_INDEX; continue; }

    // 3) عنوان "أولاً/ثانياً..." — نسبي: أعمق بواحد من آخر عنوان مطلق (مش من
    // بعضها البعض)، فتفضل كل العناوين المتتالية من النوع ده إخوة على نفس المستوى.
    // مستواه "غامض" ومرشّح لمرحلة التصنيف الذكي لاحقًا لو مفعّلة.
    const wm = t.match(wordOrdinalRegex);
    if (wm) {
      const lvl = (lastAbsoluteLevel === null ? 4 : lastAbsoluteLevel + 1);
      matches.push({ level: lvl, lineIdx: idx, inlineTitle: wm[2] ? wm[2].trim() : wm[1], kind: 'wordOrdinal' });
      continue;
    }

    // 4) عنوان بصري مكتشف وقت الاستخراج (علامة [HEADING])، بدون أي نمط نصي معروف — نسبي بنفس منطق (3)، وغامض أيضًا.
    if (ocrMarked) {
      const lvl = (lastAbsoluteLevel === null ? 4 : lastAbsoluteLevel + 1);
      matches.push({ level: lvl, lineIdx: idx, inlineTitle: t || '(عنوان بصري)', kind: 'ocrMarked' });
      continue;
    }
  }
  matches.forEach((h, i) => {
    if (h.inlineTitle) { h.title = h.inlineTitle; h.titleLineIdx = h.lineIdx; return; }
    let j = h.lineIdx + 1;
    while (j < n && (excluded[j] || !lines[j].trim())) j++;
    const nextHeadingLine = (i + 1 < matches.length) ? matches[i + 1].lineIdx : n;
    if (j < nextHeadingLine) { h.title = stripHeadingMarker(lines[j].trim()); h.titleLineIdx = j; }
    else { h.title = '(بدون عنوان)'; }
  });
  matches.forEach((h, i) => {
    const startLine = (h.titleLineIdx !== undefined ? h.titleLineIdx : h.lineIdx) + 1;
    const endLine = (i + 1 < matches.length) ? matches[i + 1].lineIdx : n;
    const chunk = [];
    for (let k = startLine; k < endLine; k++) { if (!excluded[k]) chunk.push(stripHeadingMarker(lines[k])); }
    h.content = chunk.join('\n').trim();
  });

  return { lines, excluded, matches, n };
}

function buildTreeFromMatches(matches) {
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

  const stats = { section: 0, part: 0, chapter: 0, topic: 0, subtopic: 0, other: 0 };
  matches.forEach(h => {
    const def = levelDefs[h.level];
    if (def) stats[def.key]++; else stats.other++;
  });

  return { root, leaves, stats };
}

/**
 * مرحلة تصنيف ذكي (اختيارية): بترجع لـ Gemini قائمة كل العناوين المكتشفة في
 * الكتاب كله بالترتيب (نص العناوين بس، مش المحتوى — استدعاء واحد رخيص جدًا
 * لكل كتاب)، وتسأله يصحح مستوى العناوين "الغامضة" فقط (أولاً/ثانياً، أو
 * العنوان البصري) بناءً على فهمه للسياق الكامل، بينما العناوين "المؤكدة"
 * (باب/فصل/مبحث بكلمة صريحة) بتفضل زي ما هي كمرجع ثابت لا يتغيّر.
 */
async function callGeminiReclassifyHeadings(matches, apiKey, model) {
  const ambiguousCount = matches.filter(h => h.kind === 'wordOrdinal' || h.kind === 'ocrMarked').length;
  if (ambiguousCount === 0) return { matches, changed: 0 };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const outlineLines = matches.map((h, i) => {
    const status = (h.kind === 'wordOrdinal' || h.kind === 'ocrMarked') ? 'غامض' : 'مؤكد';
    return `[${i}] (${status}, مستوى حالي: ${h.level}) ${h.title}`;
  }).join('\n');

  const prompt = `أنت خبير في تحليل الهياكل الهرمية لكتب القانون. معك قائمة مرتبة بكل عناوين كتاب واحد بالترتيب الفعلي كما وردت فيه.

المقياس الرقمي للمستوى: 0=قسم، 1=باب، 2=فصل، 3=مبحث، 4=مطلب، وأي رقم أكبر من 4 يعني مستوى أعمق من مطلب (فرع داخل مطلب).

كل عنوان معلّم بحالة:
- "مؤكد": مستواه صحيح ومعروف بثقة تامة (من كلمة صريحة زي باب/فصل/مبحث/مطلب أو صيغة تمهيد/مقدمة معروفة) — مرجع ثابت، لا تغيّر مستواه إطلاقًا.
- "غامض": مستواه الحالي مجرد تخمين أولي بسيط (عمق واحد أكبر من آخر عنوان مؤكد قبله) ومحتاج تصحيح فعلي منك بناءً على فهمك للسياق الكامل وترتيب الكتاب المنطقي.

قائمة العناوين:
${outlineLines}

المطلوب: لكل عنوان "غامض" فقط، حدد المستوى الهرمي الصحيح (رقم صحيح) بناءً على موضعه بين العناوين المؤكدة المجاورة له وعلاقته المنطقية بالعناوين الغامضة القريبة منه (مثلاً لو "أولاً" و"ثانياً" و"ثالثاً" متتالية تحت نفس المبحث، الأرجح إنهم إخوة على نفس المستوى، إلا لو مضمون عناوينهم نفسه يوحي بخلاف ذلك). لو رأيت أن التخمين الأولي صحيح بالفعل، أعده بنفس الرقم. لا تُدرج أي عنوان "مؤكد" في ردك إطلاقًا.

أعد النتيجة بصيغة JSON فقط: مصفوفة تشمل العناوين "الغامضة" فقط بهذا الشكل بالضبط:
[{"index": 5, "level": 3}, {"index": 6, "level": 3}]`;

  const body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } };
  const data = await geminiFetch(url, body);
  let text = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';
  text = text.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  let corrections;
  try { corrections = JSON.parse(text); }
  catch (e) { const err = new Error('تعذّر تفسير استجابة تصنيف العناوين كـ JSON'); err.status = -1; throw err; }
  if (!Array.isArray(corrections)) { const err = new Error('استجابة تصنيف العناوين ليست مصفوفة'); err.status = -1; throw err; }

  let changed = 0;
  corrections.forEach(c => {
    const idx = Number.isInteger(c.index) ? c.index : parseInt(c.index, 10);
    const lvl = Number.isInteger(c.level) ? c.level : parseInt(c.level, 10);
    if (!Number.isInteger(idx) || idx < 0 || idx >= matches.length) return;
    const h = matches[idx];
    if (h.kind !== 'wordOrdinal' && h.kind !== 'ocrMarked') return; // حماية إضافية: ممنوع تعديل عنوان مؤكد حتى لو النموذج حاول
    if (!Number.isInteger(lvl) || lvl < 0) return;
    if (lvl !== h.level) changed++;
    h.level = lvl;
  });

  return { matches, changed };
}

function parseBook(rawText) {
  const { lines, excluded, matches, n } = detectHeadingMatches(rawText);
  const { root, leaves, stats } = buildTreeFromMatches(matches);
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
  const italic = opts.italic ? '<w:i/>' : '';
  const styleTag = styleId ? `<w:pStyle w:val="${styleId}"/>` : '';
  const jc = opts.center ? '<w:jc w:val="center"/>' : '<w:jc w:val="right"/>';
  const lines = String(text || '').split(/\n+/);
  const runs = lines.map((ln, i) => {
    const br = i > 0 ? '<w:br/>' : '';
    return `<w:r><w:rPr><w:rtl/>${bold}${italic}</w:rPr>${br}<w:t xml:space="preserve">${escXml(ln)}</w:t></w:r>`;
  }).join('');
  return `<w:p><w:pPr><w:bidi/>${styleTag}${jc}</w:pPr>${runs}</w:p>`;
}
function pageBreakXml() { return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`; }

/**
 * يقسّم نص الشرح لفقرات مستقلة، ويميّز أي سطر توضيح عامي ("💡 خد بالك: ...")
 * بخط مائل عن باقي الشرح الأكاديمي، عشان الفرق بين الاتنين يبقى واضح بصريًا
 * في ملف Word (مش مجرد فرق في الأسلوب مدفون جوه نفس الفقرة).
 */
const KHOD_BALAK_REGEX = /^(?:💡\s*)?خد\s*بالك\s*[:：]/;
function renderExplanation(explanation) {
  const paragraphs = String(explanation || '').split(/\n+/).filter(p => p.trim());
  return paragraphs.map(p => {
    const isNote = KHOD_BALAK_REGEX.test(p.trim());
    return paraXml(p.trim(), 'Normal', isNote ? { italic: true } : {});
  }).join('');
}

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
          if (r.explanation) body += renderExplanation(r.explanation);
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
          if (r.tf && r.tf.length) {
            body += paraXml('أسئلة صح وخطأ', 'Normal', { bold: true });
            r.tf.forEach((q, qi) => {
              body += paraXml(`${qi + 1}. ${q.statement || ''}`, 'Normal');
              body += paraXml(`   ✔ الإجابة الصحيحة: ${q.answer ? 'صح' : 'خطأ'}`, 'Normal', { bold: true });
            });
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
  const verifyStats = { checked: 0, corrected: 0, failed: 0 };
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

    // مرحلة تحقق 1: مراجعة النص المستخرج بالرجوع لنفس الصورة، بالتركيز على الأرقام.
    if (VERIFY_EXTRACTION && !text.startsWith('[تعذّر استخراج')) {
      await sleep(pacingMs());
      let verified = null, vAttempts = 0, vBackoffs = 0;
      while (verified === null) {
        if (keyIdx >= API_KEYS.length) { verified = { text, changed: false }; verifyStats.failed++; break; }
        try {
          verified = await callGeminiVerifyExtraction(base64, text, currentKey(), MODEL);
        } catch (err) {
          if (isKeyExhaustedError(err.status, err.message)) { keyIdx++; vAttempts = 0; vBackoffs = 0; continue; }
          if (err.status === 429) {
            vBackoffs++;
            if (vBackoffs > 6) { keyIdx++; vBackoffs = 0; continue; }
            await sleep(Math.min(4000 * vBackoffs, 30000));
          } else {
            vAttempts++;
            if (vAttempts >= 2) { verified = { text, changed: false }; verifyStats.failed++; }
            else await sleep(1200);
          }
        }
      }
      verifyStats.checked++;
      if (verified.changed) {
        verifyStats.corrected++;
        console.log(`     ↳ ✏️  تم تصحيح تفاصيل (أرقام/كلمات) في الصفحة ${i + 1} بعد المراجعة`);
      }
      text = verified.text;
    }

    pageTexts.push(`--- الصفحة ${i + 1} ---\n${text}`);
    console.log(`  📄 صفحة ${i + 1}/${imagePaths.length} تم استخراجها`);
    await sleep(pacingMs());
  }
  return { fullText: pageTexts.join('\n\n'), verifyStats };
}

/* ================= المعالجة الكاملة لملف واحد ================= */

function excerptOf(text, maxChars) {
  if (!text) return '';
  const t = text.trim().replace(/\s+/g, ' ');
  return t.length > maxChars ? t.slice(0, maxChars) + '…' : t;
}

async function processOnePdf(pdfPath) {
  const baseName = path.basename(pdfPath, path.extname(pdfPath));
  console.log(`\n========== معالجة: ${baseName} ==========`);

  console.log('🖼️  جارٍ تحويل الصفحات لصور...');
  const pageDir = path.join(WORK_DIR, baseName);
  const imagePaths = convertPdfToImages(pdfPath, pageDir);
  console.log(`   عدد الصفحات: ${imagePaths.length}`);

  console.log('🔎 جارٍ استخراج النص عبر Gemini...' + (VERIFY_EXTRACTION ? ' (مع مراجعة كل صفحة)' : ''));
  const { fullText, verifyStats: extractionVerifyStats } = await extractTextFromImages(imagePaths);
  fs.writeFileSync(path.join(OUTPUT_DIR, `${baseName}-text.txt`), fullText, 'utf-8');
  console.log('   ✅ تم حفظ النص المستخرج');
  if (VERIFY_EXTRACTION) {
    console.log(`   🔍 تحقق الاستخراج: ${extractionVerifyStats.checked} صفحة رُوجعت، ${extractionVerifyStats.corrected} صفحة صُححت، ${extractionVerifyStats.failed} فشلت المراجعة (تم الإبقاء على النص الأصلي)`);
  }

  console.log('🌳 جارٍ تحليل الهيكل الهرمي...');
  const { lines: bookLines, excluded: bookExcluded, matches: headingMatches, n: bookLineCount } = detectHeadingMatches(fullText);
  const ambiguousHeadingsCount = headingMatches.filter(h => h.kind === 'wordOrdinal' || h.kind === 'ocrMarked').length;

  if (VERIFY_STRUCTURE && ambiguousHeadingsCount > 0) {
    console.log(`   🔎 ${ambiguousHeadingsCount} عنوان غامض (أولاً/ثانياً أو مكتشف بصريًا) — جارٍ تصنيف مستواها الصحيح بناءً على سياق الكتاب الكامل...`);
    let done = false, attempts = 0, rateLimitBackoffs = 0;
    while (!done) {
      if (keyIdx >= API_KEYS.length) { console.warn('   ⚠️ نفدت المفاتيح أثناء تصنيف الهيكل — سيُعتمد التخمين الأولي كما هو.'); done = true; break; }
      try {
        const { changed } = await callGeminiReclassifyHeadings(headingMatches, currentKey(), MODEL);
        if (changed > 0) console.log(`   ✏️  تم تصحيح مستوى ${changed} عنوان غامض بعد مراجعة السياق الكامل`);
        done = true;
      } catch (err) {
        if (isKeyExhaustedError(err.status, err.message)) { keyIdx++; attempts = 0; rateLimitBackoffs = 0; continue; }
        if (err.status === 429) {
          rateLimitBackoffs++;
          if (rateLimitBackoffs > 6) { keyIdx++; rateLimitBackoffs = 0; continue; }
          await sleep(Math.min(4000 * rateLimitBackoffs, 30000));
        } else {
          attempts++;
          if (attempts >= 2) { console.warn(`   ⚠️ فشلت مراجعة تصنيف الهيكل (${err.message}) — سيُعتمد التخمين الأولي كما هو.`); done = true; }
          else await sleep(1200);
        }
      }
    }
    await sleep(pacingMs());
  }

  const parsedBook = buildTreeFromMatches(headingMatches);
  console.log(`   ${parsedBook.stats.section} قسم، ${parsedBook.stats.part} باب، ${parsedBook.stats.chapter} فصل، ${parsedBook.stats.topic} مبحث، ${parsedBook.stats.subtopic} مطلب، ${parsedBook.stats.other} عنوان نسبي/بصري (أولاً.. أو مكتشف بصريًا)، ${parsedBook.leaves.length} جزء سيُشرح`);
  if (parsedBook.leaves.length === 0) {
    console.warn('   ⚠️ لم يتم العثور على أي عناوين بصيغة "الباب الأول" ونحوها — تأكد من صيغة عناوين الكتاب.');
  }
  assignPathLabels(parsedBook.root);

  console.log('✍️  جارٍ توليد الشرح والأسئلة لكل جزء...' + (VERIFY_CONTENT ? ' (مع تحقق مقابل النص الأصلي)' : ''));
  const aiResults = {};
  const contentVerifyStats = { checked: 0, corrected: 0, failed: 0 };
  const verificationLog = [];

  for (let i = 0; i < parsedBook.leaves.length; i++) {
    const unit = parsedBook.leaves[i];
    unit.pathLabel = unit.node.pathLabel || unit.title;

    const prevUnit = parsedBook.leaves[i - 1];
    const nextUnit = parsedBook.leaves[i + 1];
    const siblingContext = {
      prevTitle: prevUnit ? prevUnit.pathLabel : null,
      prevExcerpt: prevUnit && aiResults[prevUnit.id] ? excerptOf(aiResults[prevUnit.id].explanation, 220) : null,
      nextTitle: nextUnit ? nextUnit.pathLabel : null
    };

    let done = false, attempts = 0, rateLimitBackoffs = 0;
    let generated = null;
    while (!done) {
      if (keyIdx >= API_KEYS.length) throw new Error(`نفدت حدود كل المفاتيح عند الجزء ${i + 1}`);
      try {
        const parsed = await callGeminiForUnit(unit, currentKey(), MODEL, WORD_COUNT_OPTION, MCQ_COUNT, ESSAY_COUNT, TF_COUNT, siblingContext);
        generated = {
          explanation: parsed.explanation || '',
          mcq: Array.isArray(parsed.mcq) ? parsed.mcq : [],
          essay: Array.isArray(parsed.essay) ? parsed.essay : [],
          tf: Array.isArray(parsed.tf) ? parsed.tf : []
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
          if (attempts >= 2) { aiResults[unit.id] = { fail: true, note: err.message, explanation: '', mcq: [], essay: [], tf: [] }; done = true; }
          else await sleep(1200);
        }
      }
    }

    if (generated) {
      // مرحلة تحقق 2: مراجعة الشرح والأسئلة مقابل النص الأصلي فقط (تأليف/أرقام/إجابات/قابلية الإجابة).
      if (VERIFY_CONTENT) {
        await sleep(pacingMs());
        let vDone = false, vAttempts = 0, vBackoffs = 0;
        while (!vDone) {
          if (keyIdx >= API_KEYS.length) { vDone = true; contentVerifyStats.failed++; break; }
          try {
            const { result, changed, notes } = await callGeminiVerifyContent(unit, generated, currentKey(), MODEL, MCQ_COUNT, ESSAY_COUNT, TF_COUNT);
            generated = result;
            contentVerifyStats.checked++;
            if (changed) {
              contentVerifyStats.corrected++;
              verificationLog.push(`• ${unit.pathLabel}:\n  - ${notes.join('\n  - ')}`);
              console.log(`     ↳ ✏️  تم تصحيح ${notes.length} ملحوظة في "${unit.pathLabel}" بعد التحقق`);
            }
            vDone = true;
          } catch (err) {
            if (isKeyExhaustedError(err.status, err.message)) { keyIdx++; vAttempts = 0; vBackoffs = 0; continue; }
            if (err.status === 429) {
              vBackoffs++;
              if (vBackoffs > 6) { keyIdx++; vBackoffs = 0; continue; }
              await sleep(Math.min(4000 * vBackoffs, 30000));
            } else {
              vAttempts++;
              if (vAttempts >= 2) { vDone = true; contentVerifyStats.failed++; }
              else await sleep(1200);
            }
          }
        }
      }
      aiResults[unit.id] = { fail: false, ...generated };
    }

    const status = aiResults[unit.id].fail ? '✗ فشل' : '✓';
    console.log(`  ${i + 1}/${parsedBook.leaves.length} ${status} ${unit.pathLabel}`);
    await sleep(pacingMs());
  }

  if (VERIFY_CONTENT) {
    console.log(`   🔍 تحقق المحتوى: ${contentVerifyStats.checked} جزء رُوجع، ${contentVerifyStats.corrected} جزء صُحح، ${contentVerifyStats.failed} فشلت مراجعته (تم الإبقاء على المحتوى كما وُلّد)`);
  }

  console.log('📦 جارٍ بناء ملف Word...');
  const docxBuffer = await buildDocx(parsedBook, aiResults, baseName);
  const outPath = path.join(OUTPUT_DIR, `${baseName}-شرح-مبسط.docx`);
  fs.writeFileSync(outPath, docxBuffer);
  console.log(`✅ تم إنشاء: ${outPath}`);

  // تقرير تحقق مختصر يوضح كل ما تم تصحيحه فعليًا، عشان تراجعه بنفسك لو حابب.
  const reportLines = [];
  reportLines.push(`تقرير التحقق — ${baseName}`);
  reportLines.push('='.repeat(40));
  reportLines.push(`تحقق الاستخراج مفعّل: ${VERIFY_EXTRACTION ? 'نعم' : 'لا'}`);
  if (VERIFY_EXTRACTION) {
    reportLines.push(`  - صفحات رُوجعت: ${extractionVerifyStats.checked}`);
    reportLines.push(`  - صفحات صُححت فيها تفاصيل: ${extractionVerifyStats.corrected}`);
    reportLines.push(`  - صفحات فشلت مراجعتها (أُبقي نصها الأصلي): ${extractionVerifyStats.failed}`);
  }
  reportLines.push('');
  reportLines.push(`تحقق المحتوى مفعّل: ${VERIFY_CONTENT ? 'نعم' : 'لا'}`);
  if (VERIFY_CONTENT) {
    reportLines.push(`  - أجزاء رُوجعت: ${contentVerifyStats.checked}`);
    reportLines.push(`  - أجزاء صُححت: ${contentVerifyStats.corrected}`);
    reportLines.push(`  - أجزاء فشلت مراجعتها (أُبقي المحتوى كما وُلّد أول مرة): ${contentVerifyStats.failed}`);
    if (verificationLog.length) {
      reportLines.push('');
      reportLines.push('تفاصيل التصحيحات:');
      reportLines.push(verificationLog.join('\n'));
    }
  }
  fs.writeFileSync(path.join(OUTPUT_DIR, `${baseName}-تقرير-التحقق.txt`), reportLines.join('\n'), 'utf-8');
}

async function main() {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.mkdirSync(WORK_DIR, { recursive: true });

  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`❌ مجلد ${INPUT_DIR} غير موجود.`);
    process.exit(1);
  }
  let pdfs = fs.readdirSync(INPUT_DIR).filter(f => f.toLowerCase().endsWith('.pdf'));
  if (pdfs.length === 0) {
    console.error(`❌ لا يوجد أي ملف PDF داخل مجلد input/. ارفع ملف PDF هناك أولاً.`);
    process.exit(1);
  }

  // اختيار كتاب واحد بس: لو اتحدد BOOK_FILENAME (اسم كامل أو جزء منه)، نعالج
  // بس الملف/الملفات اللي بتطابقه. لو سيبته فاضي، نعالج كل الكتب زي المعتاد.
  const bookFilter = (process.env.BOOK_FILENAME || '').trim();
  if (bookFilter) {
    const filterLower = bookFilter.toLowerCase();
    const matched = pdfs.filter(f => f.toLowerCase().includes(filterLower));
    if (matched.length === 0) {
      console.error(`❌ لا يوجد أي ملف PDF في input/ يطابق "${bookFilter}".`);
      console.error(`الملفات الموجودة فعليًا:\n${pdfs.map(f => '  - ' + f).join('\n')}`);
      process.exit(1);
    }
    if (matched.length > 1) {
      console.error(`❌ فيه أكتر من ملف بيطابق "${bookFilter}" — حدد اسم أدق:\n${matched.map(f => '  - ' + f).join('\n')}`);
      process.exit(1);
    }
    pdfs = matched;
    console.log(`📖 تم اختيار كتاب واحد فقط: "${pdfs[0]}"`);
  }

  console.log(`النموذج: ${MODEL} | طول الشرح: ${WORD_COUNT_OPTION} | أسئلة اختيار: ${MCQ_COUNT} | أسئلة مقالية: ${ESSAY_COUNT} | أسئلة صح وخطأ: ${TF_COUNT}`);
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
