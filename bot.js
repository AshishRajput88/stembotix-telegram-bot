/**
 * STEMbotix Telegram AI Review Auto-Assign & Database Management Bot
 * =================================================================
 * Features:
 * 1. Bulk Document & Text Upload (PDF, Word, Excel, Text) -> splits 44+ employees & products
 * 2. Real-Time Employee Submission Alerts with Screenshot Proofs (SS) directly to Telegram!
 * 3. Review Management (Add, Edit/Update, Delete, Report & Stats)
 * 4. Interactive Inline Buttons & Live Firebase Firestore Sync
 * 5. Safe message chunking (Zero Telegram 400 Bad Request errors)
 * =================================================================
 */

require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const http = require('http');
const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const XLSX = require('xlsx');

// Configuration
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || '';
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const FIREBASE_PROJECT_ID = process.env.FIREBASE_PROJECT_ID || 'stembotix-riview-system';
const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY || 'AIzaSyDByqhHIqHzxXNVEp8lmLKcUuEW4Jv18R0';
const LIVE_DIRECTORY_URL = process.env.LIVE_DIRECTORY_URL || 'https://stembotix-riview-system.firebaseapp.com';

// Keepalive Web Server for 24/7 Free Cloud Hosting (Render, Koyeb, Railway)
const PORT = process.env.PORT || 3000;
const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end(`🚀 STEMbotix Telegram AI Bot is Running 24/7!\nStatus: ${botState.isSleeping ? '💤 SLEEP MODE' : '⚡ AWAKE & LISTENING'}\n`);
});
server.listen(PORT, () => {
  console.log(`🌐 24/7 Health Server active on port ${PORT}`);
});

// Self-Ping Keepalive to prevent Render from going to sleep
const SELF_PING_URL = process.env.RENDER_EXTERNAL_URL || 'https://stembotix-telegram-bot.onrender.com';
setInterval(() => {
  try {
    http.get(SELF_PING_URL, (res) => {}).on('error', () => {});
  } catch (e) {}
}, 5 * 60 * 1000); // Ping every 5 minutes

// Bot State Management (Sleep / Wake toggle)
const STATE_FILE = path.join(__dirname, 'bot-state.json');
let botState = { isSleeping: false };
try {
  if (fs.existsSync(STATE_FILE)) {
    botState = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
  }
} catch (e) {
  botState = { isSleeping: false };
}

function setBotSleepState(sleeping) {
  botState.isSleeping = !!sleeping;
  try {
    fs.writeFileSync(STATE_FILE, JSON.stringify(botState, null, 2), 'utf8');
  } catch (e) {}
}

if (!TELEGRAM_BOT_TOKEN) {
  console.error('\n❌ ERROR: TELEGRAM_BOT_TOKEN is missing in .env file!');
  console.log('Please open telegram-ai-bot/.env and paste your bot token from @BotFather.\n');
  process.exit(1);
}

// Track Admin Chat IDs for Instant Submission Notifications
const ADMIN_CHATS_FILE = path.join(__dirname, 'admin-chats.json');
let adminChatIds = new Set();
try {
  if (fs.existsSync(ADMIN_CHATS_FILE)) {
    const data = JSON.parse(fs.readFileSync(ADMIN_CHATS_FILE, 'utf8'));
    if (Array.isArray(data)) data.forEach(id => adminChatIds.add(String(id)));
  }
} catch (e) {
  console.warn('Could not read admin-chats.json:', e.message);
}

if (process.env.TELEGRAM_ADMIN_CHAT_ID) {
  adminChatIds.add(String(process.env.TELEGRAM_ADMIN_CHAT_ID));
}

function registerAdminChat(chatId) {
  if (!chatId) return;
  const strId = String(chatId);
  if (!adminChatIds.has(strId)) {
    adminChatIds.add(strId);
    try {
      fs.writeFileSync(ADMIN_CHATS_FILE, JSON.stringify(Array.from(adminChatIds), null, 2), 'utf8');
    } catch (e) { }
  }
}

// Load Team Directory
let DEPARTMENTS = [];
try {
  DEPARTMENTS = JSON.parse(fs.readFileSync(path.join(__dirname, 'team-data.json'), 'utf8'));
} catch (e) {
  console.error('Could not load team-data.json:', e);
}

const PEOPLE = {};
DEPARTMENTS.forEach(dept => {
  dept.members.forEach(member => {
    PEOPLE[member.slug] = {
      slug: member.slug,
      name: member.name,
      role: member.role,
      department: dept.name
    };
  });
});

function slugify(name) {
  return String(name).toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

function findBestEmployeeMatch(nameOrText) {
  if (!nameOrText) return '';
  let query = String(nameOrText).trim()
    .replace(/^\s*\d+[\.\)\:\-\s]+/g, '')
    .replace(/[\.\)\:\-\s]+\d+\s*$/g, '')
    .trim()
    .toLowerCase();

  if (!query) return '';
  const qSlug = slugify(query);
  if (PEOPLE[qSlug]) return qSlug;

  const keys = Object.keys(PEOPLE);
  for (let k of keys) {
    if (PEOPLE[k].name.toLowerCase() === query || PEOPLE[k].slug === qSlug) return PEOPLE[k].slug;
  }
  for (let k of keys) {
    const parts = PEOPLE[k].name.toLowerCase().split(/\s+/);
    for (let part of parts) {
      if (part.length > 2 && (query === part || query.startsWith(part) || query.endsWith(part) || query.includes(part))) {
        return PEOPLE[k].slug;
      }
    }
  }
  return '';
}

function cleanProductTitle(raw) {
  if (!raw) return 'Blockzie';
  let s = String(raw).trim();
  s = s.replace(/[★\u2605\u2B50]+/g, '');
  s = s.replace(/\s*\(\s*\d+\s*(?:star|stars|\/5)?\s*\)/gi, '');
  s = s.replace(/\s*5\s*star\s*/gi, '');
  s = s.replace(/^[-:–—\s]+|[-:–—\s]+$/g, '');
  return s.trim() || 'Blockzie';
}

/* ---------------- Firestore REST API Engine ---------------- */

async function fetchAllReviewsFromFirestore() {
  const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/reviews?key=${FIREBASE_API_KEY}&pageSize=1000`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Firestore fetch error: ${res.statusText}`);
  const data = await res.json();
  if (!data.documents) return [];
  return data.documents.map(doc => {
    const f = doc.fields || {};
    return {
      docName: doc.name,
      id: doc.name.split('/').pop(),
      person: f.person?.stringValue || '',
      product: f.product?.stringValue || '',
      text: f.text?.stringValue || '',
      createdAt: Number(f.createdAt?.integerValue || 0)
    };
  });
}

async function saveReviewToFirestore(item) {
  const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/reviews?key=${FIREBASE_API_KEY}`;
  const payload = {
    fields: {
      person: { stringValue: item.personSlug || item.person },
      product: { stringValue: item.product || 'STEMbotix Kit' },
      text: { stringValue: item.text },
      createdAt: { integerValue: String(Date.now()) }
    }
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Firestore save error: ${errText}`);
  }
  return await response.json();
}

async function deleteAllReviewsFromFirestore() {
  const allReviews = await fetchAllReviewsFromFirestore();
  if (allReviews.length === 0) return { totalDeleted: 0, previousCount: 0 };

  let totalDeleted = 0;
  for (let i = 0; i < allReviews.length; i += 10) {
    const batch = allReviews.slice(i, i + 10);
    await Promise.all(batch.map(async doc => {
      try {
        const url = `https://firestore.googleapis.com/v1/${doc.docName}?key=${FIREBASE_API_KEY}`;
        const res = await fetch(url, { method: 'DELETE' });
        if (res.ok) totalDeleted++;
      } catch (e) {
        console.error('Delete error for doc:', doc.id, e);
      }
    }));
  }

  return { totalDeleted, previousCount: allReviews.length };
}

async function deleteReviewsForEmployee(personSlug) {
  const allReviews = await fetchAllReviewsFromFirestore();
  const targetReviews = allReviews.filter(r => r.person === personSlug);
  if (targetReviews.length === 0) return { totalDeleted: 0 };

  let totalDeleted = 0;
  for (let r of targetReviews) {
    try {
      const url = `https://firestore.googleapis.com/v1/${r.docName}?key=${FIREBASE_API_KEY}`;
      const res = await fetch(url, { method: 'DELETE' });
      if (res.ok) totalDeleted++;
    } catch (e) {
      console.error('Delete error for doc:', r.id, e);
    }
  }
  return { totalDeleted };
}

async function updateReviewInFirestore(personSlug, newProduct, newReviewText) {
  const allReviews = await fetchAllReviewsFromFirestore();
  const targetReviews = allReviews.filter(r => r.person === personSlug);
  const cleanProduct = cleanProductTitle(newProduct) || 'Blockzie';

  if (targetReviews.length > 0) {
    let matchedReview = targetReviews.find(r => r.product.toLowerCase().includes(cleanProduct.toLowerCase()) || cleanProduct.toLowerCase().includes(r.product.toLowerCase()));
    if (!matchedReview) matchedReview = targetReviews[0];

    const patchUrl = `https://firestore.googleapis.com/v1/${matchedReview.docName}?updateMask.fieldPaths=text&updateMask.fieldPaths=product&updateMask.fieldPaths=createdAt&key=${FIREBASE_API_KEY}`;
    const payload = {
      fields: {
        person: { stringValue: personSlug },
        product: { stringValue: cleanProduct || matchedReview.product || 'Blockzie' },
        text: { stringValue: newReviewText },
        createdAt: { integerValue: String(Date.now()) }
      }
    };

    const res = await fetch(patchUrl, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      await fetch(`https://firestore.googleapis.com/v1/${matchedReview.docName}?key=${FIREBASE_API_KEY}`, { method: 'DELETE' });
      await saveReviewToFirestore({ personSlug, product: cleanProduct, text: newReviewText });
    }

    return {
      status: 'updated',
      docId: matchedReview.id,
      personSlug,
      personName: (PEOPLE[personSlug] && PEOPLE[personSlug].name) || personSlug,
      product: cleanProduct || matchedReview.product,
      text: newReviewText
    };
  } else {
    await saveReviewToFirestore({ personSlug, product: cleanProduct, text: newReviewText });
    return {
      status: 'created',
      personSlug,
      personName: (PEOPLE[personSlug] && PEOPLE[personSlug].name) || personSlug,
      product: cleanProduct,
      text: newReviewText
    };
  }
}

/* ---------------- Submissions Fetch & Real-Time Monitor ---------------- */

async function fetchAllSubmissionsFromFirestore() {
  try {
    const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/submissions?key=${FIREBASE_API_KEY}&pageSize=100`;
    const res = await fetch(url);
    if (!res.ok) return [];
    const data = await res.json();
    if (!data.documents) return [];

    return data.documents.map(doc => {
      const f = doc.fields || {};
      let reviewsList = [];
      if (f.reviews && f.reviews.arrayValue && f.reviews.arrayValue.values) {
        reviewsList = f.reviews.arrayValue.values.map(v => {
          const mv = v.mapValue?.fields || {};
          return {
            product: mv.product?.stringValue || '',
            text: mv.text?.stringValue || ''
          };
        });
      }

      let imagesList = [];
      if (f.images && f.images.arrayValue && f.images.arrayValue.values) {
        imagesList = f.images.arrayValue.values.map(v => v.stringValue || '').filter(Boolean);
      } else if (f.image && f.image.stringValue) {
        imagesList = [f.image.stringValue];
      }

      return {
        docName: doc.name,
        id: doc.name.split('/').pop(),
        person: f.person?.stringValue || '',
        personName: f.personName?.stringValue || '',
        submittedBy: f.submittedBy?.stringValue || '',
        department: f.department?.stringValue || '',
        where: f.where?.stringValue || '',
        note: f.note?.stringValue || '',
        submittedAt: Number(f.submittedAt?.integerValue || 0),
        reviews: reviewsList,
        images: imagesList
      };
    });
  } catch (err) {
    console.error('Submissions fetch error:', err.message);
    return [];
  }
}

let knownSubmissionIds = new Set();
let isInitialSubmissionsLoad = true;

async function checkAndNotifyNewSubmissions() {
  if (botState.isSleeping) return; // Mute notifications during Sleep Mode
  try {
    const subs = await fetchAllSubmissionsFromFirestore();
    if (isInitialSubmissionsLoad) {
      subs.forEach(s => knownSubmissionIds.add(s.id));
      isInitialSubmissionsLoad = false;
      return;
    }

    for (let s of subs) {
      if (!knownSubmissionIds.has(s.id)) {
        knownSubmissionIds.add(s.id);
        await broadcastSubmissionAlert(s);
      }
    }
  } catch (err) {
    console.warn('Submissions check error:', err.message);
  }
}

async function broadcastSubmissionAlert(s) {
  const timeStr = new Date(s.submittedAt || Date.now()).toLocaleString('en-IN', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

  const empName = s.personName || (PEOPLE[s.person] && PEOPLE[s.person].name) || s.person || 'Team Member';
  const deptName = s.department || (PEOPLE[s.person] && PEOPLE[s.person].department) || '';
  const byName = s.submittedBy || empName;
  const whereStatus = s.where || 'Yes';
  const noteText = s.note ? `\n📝 *Note:* ${s.note}` : '';

  let revListText = '';
  if (s.reviews && s.reviews.length > 0) {
    s.reviews.forEach(r => {
      const shortText = r.text.length > 80 ? r.text.slice(0, 80) + '…' : r.text;
      revListText += `  • *${r.product || 'Product'}*: "${shortText}"\n`;
    });
  }

  const ssCount = s.images ? s.images.length : 0;

  let msg = `🔔 *NEW EMPLOYEE REVIEW SUBMISSION RECEIVED!*\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `👤 *Employee:* ${empName}\n`;
  if (deptName) msg += `🏢 *Department:* ${deptName}\n`;
  msg += `✍️ *Submitted By:* ${byName}\n`;
  msg += `✅ *Review Posted?* ${whereStatus}\n`;
  msg += `📅 *Date & Time:* ${timeStr}\n`;
  if (noteText) msg += `${noteText}\n`;
  if (revListText) msg += `\n📦 *Reviews Completed (${s.reviews.length}):*\n${revListText}`;
  msg += `📸 *Screenshots Proof:* ${ssCount > 0 ? `${ssCount} Image(s) Attached` : 'No screenshot'}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━`;

  const chatList = Array.from(adminChatIds);
  for (let chatId of chatList) {
    try {
      await bot.sendMessage(chatId, msg, {
        parse_mode: 'Markdown',
        reply_markup: {
          inline_keyboard: [
            [{ text: '🌐 Open Admin Submissions', url: LIVE_DIRECTORY_URL + '/admin.html' }]
          ]
        }
      });

      // Send Screenshot Proofs
      if (s.images && s.images.length > 0) {
        for (let i = 0; i < s.images.length; i++) {
          const imgStr = s.images[i];
          if (imgStr.startsWith('data:image')) {
            const base64Data = imgStr.replace(/^data:image\/\w+;base64,/, '');
            const imgBuffer = Buffer.from(base64Data, 'base64');
            await bot.sendPhoto(chatId, imgBuffer, {
              caption: `📸 *Screenshot Proof ${i + 1} of ${s.images.length}* for *${empName}*`,
              parse_mode: 'Markdown'
            });
          } else if (imgStr.startsWith('http')) {
            await bot.sendPhoto(chatId, imgStr, {
              caption: `📸 *Screenshot Proof ${i + 1} of ${s.images.length}* for *${empName}*`,
              parse_mode: 'Markdown'
            });
          }
        }
      }
    } catch (sendErr) {
      console.error('Error sending submission alert to chat', chatId, sendErr.message);
    }
  }
}

// Start Real-time Submissions Polling (every 3 seconds)
setInterval(checkAndNotifyNewSubmissions, 3000);
checkAndNotifyNewSubmissions();

/* ---------------- Text / Document Heuristic Parser ---------------- */

function parseTextLocalHeuristic(rawText) {
  if (!rawText) return [];
  let text = String(rawText);

  text = text.replace(/(?=(?:\r?\n|\A)\s*\d+[\.\)\-]\s+[A-Za-z])/g, '\n__EMP_SPLIT__\n');
  text = text.replace(/(?=(?:App Name|Product Name|Product|Application|Kit Name|Item Name|Item|Kit|App)\s*[:=-])/gi, '\n__PROD_BREAK__ ');
  text = text.replace(/(?=(?:Employee Name|Employee|Staff Name|Team Member|Staff|Assigned To)\s*[:=-])/gi, '\n__EMP_BREAK__ ');

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const results = [];
  let curPersonSlug = '', curPersonName = '', curProduct = '', curTextParts = [];
  const teamKeys = Object.keys(PEOPLE);

  function commitCurrentReview() {
    if ((curPersonSlug || curPersonName) && curTextParts.length > 0) {
      const slugMatch = curPersonSlug || findBestEmployeeMatch(curPersonName) || (teamKeys[0] || '');
      const fullPersonName = (PEOPLE[slugMatch] && PEOPLE[slugMatch].name) || curPersonName || 'Team Member';
      const finalProduct = cleanProductTitle(curProduct);
      let finalReviewText = curTextParts.join(' ').trim();

      finalReviewText = finalReviewText.replace(/\s*\d+[\.\)\-]\s+[A-Za-z\s]+$/g, '').trim();

      if (finalReviewText.length > 0) {
        results.push({
          personSlug: slugMatch,
          personName: fullPersonName,
          product: finalProduct,
          text: finalReviewText
        });
      }
    }
    curTextParts = [];
  }

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];
    if (line === '__EMP_SPLIT__') continue;

    const empMatch = line.match(/^(?:__EMP_BREAK__\s*)?(?:Employee Name|Employee|Name|Person|Team Member|Assigned To|Staff)\s*[:=-]\s*(.+)$/i);
    const prodMatch = line.match(/^(?:__PROD_BREAK__\s*)?(?:App Name|Product Name|Product|Application|Kit Name|Item Name|App|Kit)\s*[:=-]\s*(.+)$/i);
    const directSlug = findBestEmployeeMatch(line);

    if (empMatch) {
      commitCurrentReview();
      const empVal = empMatch[1].trim();
      curPersonSlug = findBestEmployeeMatch(empVal);
      curPersonName = (PEOPLE[curPersonSlug] && PEOPLE[curPersonSlug].name) || empVal;
      curProduct = '';
    } else if (directSlug && line.length < 40 && (line.match(/^\d+[\.\)\-]\s+[A-Za-z\s]+$/) || line.match(/^[A-Za-z\s]+[:\-]$/) || !line.includes('★'))) {
      commitCurrentReview();
      curPersonSlug = directSlug;
      curPersonName = PEOPLE[directSlug].name;
      curProduct = '';
    } else if (prodMatch) {
      commitCurrentReview();
      const prodRest = prodMatch[1].trim();
      const starSplit = prodRest.split(/[★\u2605\u2B50]+/);
      if (starSplit.length > 1 && starSplit[1].trim().length > 10) {
        curProduct = starSplit[0].trim();
        curTextParts.push(starSplit.slice(1).join(' ').trim());
      } else {
        curProduct = prodRest;
      }
    } else {
      if (line.includes('★') && !curProduct) {
        const starIdx = line.indexOf('★');
        const beforeStars = line.slice(0, starIdx).trim();
        if (beforeStars.length > 2) curProduct = beforeStars;
        curTextParts.push(line.slice(starIdx).replace(/[★\u2605\u2B50]+/g, '').trim());
      } else {
        curTextParts.push(line);
      }
    }
  }
  commitCurrentReview();
  return results;
}

async function callGeminiAPI(rawDocText) {
  if (!GEMINI_API_KEY) {
    return parseTextLocalHeuristic(rawDocText);
  }

  const employeesList = Object.keys(PEOPLE).map(k => ({
    slug: PEOPLE[k].slug,
    name: PEOPLE[k].name,
    department: PEOPLE[k].department
  }));

  const prompt = `You are the STEMbotix AI Review Assignment Agent.
Your job is to analyze raw document text and extract every individual product and app review, assigning each review to the correct employee from the known STEMbotix team directory.

--- STEMBOTIX TEAM DIRECTORY (Valid Employees) ---
${JSON.stringify(employeesList, null, 2)}

--- DOCUMENT CONTENT ---
${rawDocText.slice(0, 30000)}

--- CRITICAL MULTI-PRODUCT & APP EXTRACTION RULES ---
1. MULTI-REVIEW SPLITTING: If one employee has multiple reviews (e.g. 'App Name: Blockzie Playstore' AND 'App Name: Botzie Playstore', or multiple products listed under one person), you MUST create a SEPARATE object in the JSON array for EACH product/app review. NEVER merge multiple products into a single review text.
2. PRODUCT / APP NAME: Extract the clean product or app name (e.g. 'Blockzie Playstore', 'Botzie Playstore') into the 'product' field. Remove decorative star symbols (★★★★★).
3. REVIEW TEXT: Clean and format the review text so it contains ONLY the specific review content for that product/app.
4. EMPLOYEE MATCHING: Match the person to their exact 'personSlug' from the directory.

--- OUTPUT FORMAT ---
Respond ONLY with a valid JSON array of objects:
[
  {
    "personSlug": "exact-employee-slug",
    "personName": "Employee Full Name",
    "product": "Clean Product or App Name",
    "text": "Clean review paragraph for this specific product..."
  }
]`;

  const models = ['gemini-1.5-flash', 'gemini-2.0-flash', 'gemini-2.5-flash-lite'];
  for (let model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`;
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1, topP: 0.95, maxOutputTokens: 8192 }
        })
      });
      if (!res.ok) continue;
      const data = await res.json();
      const reply = data?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!reply) continue;

      const cleaned = reply.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();
      const parsed = JSON.parse(cleaned);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    } catch (err) {
      console.warn(`Model ${model} error:`, err.message);
    }
  }

  return parseTextLocalHeuristic(rawDocText);
}

/* ---------------- Reports & Telegram Safe Chunking ---------------- */

async function sendDatabaseStatusReport(chatId) {
  try {
    const reviews = await fetchAllReviewsFromFirestore();
    const subs = await fetchAllSubmissionsFromFirestore();
    const now = new Date();
    const timeStr = now.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

    if (reviews.length === 0) {
      await bot.sendMessage(
        chatId,
        `📊 *STEMbotix Reviews Status Report*\n━━━━━━━━━━━━━━━━━━━━━━\n` +
        `📅 *Time:* ${timeStr}\n` +
        `ℹ️ *Total Reviews in Database:* 0\n` +
        `👥 *Employees with Reviews:* 0\n` +
        `📬 *Submissions Received:* ${subs.length}\n\n` +
        `_The database is currently clean. Send any document to auto-assign._`,
        {
          parse_mode: 'Markdown',
          reply_markup: {
            inline_keyboard: [
              [
                { text: '📋 View Submissions', callback_data: 'cmd_subs' },
                { text: '🌐 Live Directory', url: LIVE_DIRECTORY_URL }
              ]
            ]
          }
        }
      );
      return;
    }

    const byPerson = {};
    reviews.forEach(r => {
      const pSlug = r.person;
      const pName = (PEOPLE[pSlug] && PEOPLE[pSlug].name) || pSlug || 'Team Member';
      if (!byPerson[pName]) byPerson[pName] = [];
      byPerson[pName].push(r);
    });

    const empNames = Object.keys(byPerson);

    let headerMsg = `📊 *STEMbotix Reviews Status Report*\n`;
    headerMsg += `━━━━━━━━━━━━━━━━━━━━━━\n`;
    headerMsg += `📅 *Report Time:* ${timeStr}\n`;
    headerMsg += `✅ *Total Reviews in Database:* ${reviews.length}\n`;
    headerMsg += `👥 *Employees with Reviews:* ${empNames.length}\n`;
    headerMsg += `📬 *Employee Submissions:* ${subs.length}\n`;
    headerMsg += `━━━━━━━━━━━━━━━━━━━━━━\n`;

    await bot.sendMessage(chatId, headerMsg, { parse_mode: 'Markdown' });

    let currentChunk = '';
    for (let i = 0; i < empNames.length; i++) {
      const pName = empNames[i];
      const list = byPerson[pName];
      let empBlock = `👤 *${pName}* (${list.length} ${list.length === 1 ? 'review' : 'reviews'}):\n`;
      list.forEach(item => {
        const preview = item.text.length > 70 ? item.text.slice(0, 70) + '…' : item.text;
        empBlock += `  • *${item.product}*: "${preview}"\n`;
      });
      empBlock += `\n`;

      if ((currentChunk + empBlock).length > 3000) {
        await bot.sendMessage(chatId, currentChunk, { parse_mode: 'Markdown' });
        currentChunk = empBlock;
      } else {
        currentChunk += empBlock;
      }
    }

    if (currentChunk.trim()) {
      await bot.sendMessage(chatId, currentChunk, {
        parse_mode: 'Markdown',
        reply_markup: {
          inline_keyboard: [
            [
              { text: '📋 View Submissions', callback_data: 'cmd_subs' },
              { text: '🗑️ Delete All Reviews', callback_data: 'cmd_delete_all_prompt' }
            ],
            [
              { text: '🌐 Open Live Directory', url: LIVE_DIRECTORY_URL }
            ]
          ]
        }
      });
    }
  } catch (err) {
    console.error('Error generating report:', err);
    await bot.sendMessage(chatId, `❌ *Error generating report:* ${err.message}`, { parse_mode: 'Markdown' });
  }
}

async function sendSubmissionsSummaryReport(chatId) {
  try {
    const subs = await fetchAllSubmissionsFromFirestore();
    if (subs.length === 0) {
      await bot.sendMessage(chatId, 'ℹ️ No employee submissions received yet in Firebase database.');
      return;
    }

    let msg = `📬 *Recent Employee Submissions (${subs.length})*\n━━━━━━━━━━━━━━━━━━━━━━\n`;
    subs.slice(0, 8).forEach((s, idx) => {
      const timeStr = new Date(s.submittedAt || Date.now()).toLocaleDateString('en-IN', {
        day: 'numeric',
        month: 'short',
        hour: '2-digit',
        minute: '2-digit'
      });
      const empName = s.personName || (PEOPLE[s.person] && PEOPLE[s.person].name) || s.person || 'Team Member';
      msg += `${idx + 1}. 👤 *${empName}* (Submitted by ${s.submittedBy || empName})\n`;
      msg += `   • Posted: *${s.where || 'Yes'}* | 📸 Screenshots: *${s.images ? s.images.length : 0}*\n`;
      msg += `   • Time: ${timeStr}\n`;
      if (s.note) msg += `   • Note: "${s.note}"\n`;
      msg += `\n`;
    });

    await bot.sendMessage(chatId, msg, {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [{ text: '🌐 View in Admin Panel', url: LIVE_DIRECTORY_URL + '/admin.html' }]
        ]
      }
    });
  } catch (err) {
    await bot.sendMessage(chatId, `❌ *Error fetching submissions:* ${err.message}`, { parse_mode: 'Markdown' });
  }
}

async function sendChunkedReport(chatId, reviews, successCount) {
  const now = new Date();
  const timeStr = now.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

  const byPerson = {};
  reviews.forEach(r => {
    const pName = r.personName || (PEOPLE[r.personSlug] && PEOPLE[r.personSlug].name) || r.personSlug || 'Team Member';
    if (!byPerson[pName]) byPerson[pName] = [];
    byPerson[pName].push(r);
  });

  const totalEmployees = Object.keys(byPerson).length;

  let headerMsg = `🎉 *STEMbotix AI Auto-Assign Complete!*\n`;
  headerMsg += `━━━━━━━━━━━━━━━━━━━━━━\n`;
  headerMsg += `📅 *Assigned At:* ${timeStr}\n`;
  headerMsg += `✅ *Total Reviews Added:* ${successCount} of ${reviews.length}\n`;
  headerMsg += `👥 *Total Employees Processed:* ${totalEmployees}\n`;
  headerMsg += `━━━━━━━━━━━━━━━━━━━━━━`;

  await bot.sendMessage(chatId, headerMsg, { parse_mode: 'Markdown' });

  let currentChunk = '';
  const empNames = Object.keys(byPerson);

  for (let i = 0; i < empNames.length; i++) {
    const pName = empNames[i];
    let empBlock = `👤 *${pName}* (${byPerson[pName].length} reviews):\n`;
    byPerson[pName].forEach(item => {
      const shortText = item.text.length > 80 ? item.text.slice(0, 80) + '…' : item.text;
      empBlock += `  • *${item.product}*: "${shortText}"\n`;
    });
    empBlock += `\n`;

    if ((currentChunk + empBlock).length > 3000) {
      await bot.sendMessage(chatId, currentChunk, { parse_mode: 'Markdown' });
      currentChunk = empBlock;
    } else {
      currentChunk += empBlock;
    }
  }

  if (currentChunk.trim()) {
    await bot.sendMessage(chatId, currentChunk + `━━━━━━━━━━━━━━━━━━━━━━\n⚡ *100% Automated by STEMbotix AI Agent*`, {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [
            { text: '📊 View Current Status', callback_data: 'cmd_report' },
            { text: '🌐 Open Directory', url: LIVE_DIRECTORY_URL }
          ]
        ]
      }
    });
  }
}

/* ---------------- Initialize Telegram Bot ---------------- */

const bot = new TelegramBot(TELEGRAM_BOT_TOKEN, { polling: true });

bot.on('polling_error', (error) => {
  console.warn('Telegram Polling Warning:', error.code || error.message);
});

bot.on('error', (error) => {
  console.error('Telegram Bot General Error:', error.message);
});

console.log('\n======================================================');
console.log('🚀 STEMbotix Telegram AI Bot is RUNNING & LISTENING!');
console.log('⚡ Real-time Submission alerts with Screenshots active!');
console.log('======================================================\n');

function getMainMenuKeyboard() {
  if (botState.isSleeping) {
    return {
      inline_keyboard: [
        [{ text: '⚡ Wake Up Bot', callback_data: 'cmd_wake' }],
        [{ text: '🌐 Live Directory', url: LIVE_DIRECTORY_URL }]
      ]
    };
  }
  return {
    inline_keyboard: [
      [
        { text: '📊 Reviews Report', callback_data: 'cmd_report' },
        { text: '📬 View Submissions', callback_data: 'cmd_subs' }
      ],
      [
        { text: '💤 Sleep Bot', callback_data: 'cmd_sleep' },
        { text: '🗑️ Delete All', callback_data: 'cmd_delete_all_prompt' }
      ],
      [
        { text: '🌐 Live Directory', url: LIVE_DIRECTORY_URL }
      ]
    ]
  };
}

// Start / Help / Wake Command
bot.onText(/\/start|\/help|\/wake|\/resume|\/unmute/, (msg) => {
  const chatId = msg.chat.id;
  registerAdminChat(chatId);
  setBotSleepState(false);

  bot.sendMessage(
    chatId,
    `🚀 *STEMbotix AI Bot is AWAKE & ACTIVE!*\n\n` +
    `⚡ *Real-Time Features Live:*\n` +
    `1. 📬 *Instant Submission Alerts*: Employee proofs + Screenshots live direct to Telegram!\n` +
    `2. 📄 *Bulk Review Upload*: Drop any PDF, Word, Excel, or Text file.\n` +
    `3. ✏️ *Update Review*: Say *"Divy ka review update karde: Blockzie ★★★★★ new text"*\n` +
    `4. 📊 *Stats & Reports*: Ask *"kitne review add huye"* or *"submissions dikhao"*\n` +
    `5. 💤 *Sleep Bot*: Say *"sleep me jaa"* or type **/sleep** anytime to pause!`,
    {
      parse_mode: 'Markdown',
      reply_markup: getMainMenuKeyboard()
    }
  );
});

// Sleep / Pause Command
bot.onText(/\/sleep|\/pause|\/mute/, (msg) => {
  const chatId = msg.chat.id;
  registerAdminChat(chatId);
  setBotSleepState(true);

  bot.sendMessage(
    chatId,
    `💤 *Bot has entered Sleep Mode!*\n\n` +
    `• Review processing is paused.\n` +
    `• Real-time submission notifications are muted.\n\n` +
    `_Wapas activate karne ke liye **/start** ya **/wake** send karein ya button dabayein._`,
    {
      parse_mode: 'Markdown',
      reply_markup: getMainMenuKeyboard()
    }
  );
});

// Explicit Commands
bot.onText(/\/report|\/stats|\/status|\/count/, (msg) => {
  registerAdminChat(msg.chat.id);
  sendDatabaseStatusReport(msg.chat.id);
});

bot.onText(/\/submissions|\/subs/, (msg) => {
  registerAdminChat(msg.chat.id);
  sendSubmissionsSummaryReport(msg.chat.id);
});

bot.onText(/\/deleteall|\/clearall|\/clear/, async (msg) => {
  registerAdminChat(msg.chat.id);
  const chatId = msg.chat.id;
  await bot.sendMessage(
    chatId,
    `⚠️ *Are you sure you want to delete ALL reviews from Firebase?*\nThis will remove all employee assigned reviews permanently.`,
    {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [
            { text: '✅ Yes, Delete All Reviews', callback_data: 'cmd_delete_all_confirm' },
            { text: '❌ Cancel', callback_data: 'cmd_cancel' }
          ]
        ]
      }
    }
  );
});

// Callback Query Handler for Inline Buttons
bot.on('callback_query', async (query) => {
  const chatId = query.message.chat.id;
  registerAdminChat(chatId);
  const data = query.data;

  try {
    await bot.answerCallbackQuery(query.id);

    if (data === 'cmd_wake') {
      setBotSleepState(false);
      await bot.sendMessage(
        chatId,
        `⚡ *Bot WAKE UP ho gaya hai & 24/7 ACTIVE hai!*\nAb reviews upload ya commands use kar sakte hain.`,
        {
          parse_mode: 'Markdown',
          reply_markup: getMainMenuKeyboard()
        }
      );
    } else if (data === 'cmd_sleep') {
      setBotSleepState(true);
      await bot.sendMessage(
        chatId,
        `💤 *Bot Sleep Mode me chala gaya hai.*\nUthne ke liye **/start** ya **/wake** bhejein.`,
        {
          parse_mode: 'Markdown',
          reply_markup: getMainMenuKeyboard()
        }
      );
    } else if (data === 'cmd_report') {
      await sendDatabaseStatusReport(chatId);
    } else if (data === 'cmd_subs') {
      await sendSubmissionsSummaryReport(chatId);
    } else if (data === 'cmd_delete_all_prompt') {
      await bot.sendMessage(
        chatId,
        `⚠️ *Are you sure you want to delete ALL reviews from Firebase?*`,
        {
          parse_mode: 'Markdown',
          reply_markup: {
            inline_keyboard: [
              [
                { text: '✅ Yes, Delete All Reviews', callback_data: 'cmd_delete_all_confirm' },
                { text: '❌ Cancel', callback_data: 'cmd_cancel' }
              ]
            ]
          }
        }
      );
    } else if (data === 'cmd_delete_all_confirm') {
      await bot.sendMessage(chatId, '⏳ *Deleting all reviews from Firebase…*', { parse_mode: 'Markdown' });
      const res = await deleteAllReviewsFromFirestore();
      await bot.sendMessage(
        chatId,
        `🗑️ *All Reviews Deleted Successfully!*\n\n` +
        `✅ *Deleted:* ${res.totalDeleted} reviews\n` +
        `🌐 *Status:* Directory is now clean.\n\n` +
        `_You can now upload fresh review files._`,
        {
          parse_mode: 'Markdown',
          reply_markup: {
            inline_keyboard: [
              [{ text: '🌐 Open Live Directory', url: LIVE_DIRECTORY_URL }]
            ]
          }
        }
      );
    } else if (data === 'cmd_cancel') {
      await bot.sendMessage(chatId, '👍 Action cancelled. Reviews were not modified.');
    }
  } catch (err) {
    console.error('Callback error:', err);
    await bot.sendMessage(chatId, `❌ *Error:* ${err.message}`, { parse_mode: 'Markdown' });
  }
});

// Document Handler (PDF, Word, Excel, Text Files)
bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  registerAdminChat(chatId);

  if (botState.isSleeping) {
    await bot.sendMessage(
      chatId,
      `💤 *Bot Sleep Mode me hai!*\n\nDocument process karne ke liye bot ko wapas wake karein.\n\nType **/start** ya **/wake** ya niche button dabayein:`,
      {
        parse_mode: 'Markdown',
        reply_markup: getMainMenuKeyboard()
      }
    );
    return;
  }

  const doc = msg.document;

  try {
    await bot.sendMessage(chatId, '⏳ *Document received!* AI Agent is extracting and analyzing all reviews…', { parse_mode: 'Markdown' });

    const fileStream = bot.getFileStream(doc.file_id);
    const chunks = [];
    for await (const chunk of fileStream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    const fileName = (doc.file_name || '').toLowerCase();

    let extractedText = '';

    if (fileName.endsWith('.pdf') || doc.mime_type === 'application/pdf') {
      const pdfData = await pdfParse(buffer);
      extractedText = pdfData.text || '';
    } else if (fileName.endsWith('.docx') || fileName.endsWith('.doc')) {
      const docxData = await mammoth.extractRawText({ buffer });
      extractedText = docxData.value || '';
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv')) {
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const sheetsText = [];
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        sheetsText.push(XLSX.utils.sheet_to_csv(sheet));
      });
      extractedText = sheetsText.join('\n\n');
    } else {
      extractedText = buffer.toString('utf8');
    }

    if (!extractedText || !extractedText.trim()) {
      await bot.sendMessage(chatId, '⚠️ No readable text could be extracted from this document.');
      return;
    }

    const reviews = await callGeminiAPI(extractedText);

    if (!reviews || !reviews.length) {
      await bot.sendMessage(chatId, '⚠️ Could not detect any employee reviews in this document.');
      return;
    }

    let successCount = 0;
    for (let r of reviews) {
      try {
        await saveReviewToFirestore(r);
        successCount++;
      } catch (saveErr) {
        console.error('Failed to save review:', saveErr);
      }
    }

    await sendChunkedReport(chatId, reviews, successCount);

  } catch (err) {
    console.error('Error processing document:', err);
    await bot.sendMessage(chatId, `❌ *Error processing document:* ${err.message || 'Unknown error'}`, { parse_mode: 'Markdown' });
  }
});

// Natural Language Text Message Handler
bot.on('message', async (msg) => {
  if (msg.document || (msg.text && msg.text.startsWith('/'))) return;
  const chatId = msg.chat.id;
  registerAdminChat(chatId);
  const rawText = (msg.text || '').trim();
  const lowerText = rawText.toLowerCase();

  if (!rawText) return;

  // 0. Sleep / Wake NLP Commands
  const isWakeIntent = /^(?:wake\s*up|uth\s*ja|uth|start|chalu\s*ho\s*ja|chalu\s*kar|on\s*kar|resume|active\s*ho\s*ja|wake)$/i.test(lowerText);
  if (isWakeIntent) {
    setBotSleepState(false);
    await bot.sendMessage(
      chatId,
      `⚡ *Bot WAKE UP ho gaya hai & 24/7 ACTIVE hai!*\nAb reviews upload ya commands use kar sakte hain.`,
      {
        parse_mode: 'Markdown',
        reply_markup: getMainMenuKeyboard()
      }
    );
    return;
  }

  const isSleepIntent = /(?:sleep\s*me\s*ja|sleep\s*ho\s*ja|so\s*ja|sleep\s*kar|bot\s*sleep|pause|mute|bandh\s*ho\s*ja|band\s*ho\s*ja|sleep)/i.test(lowerText);
  if (isSleepIntent) {
    setBotSleepState(true);
    await bot.sendMessage(
      chatId,
      `💤 *Bot Sleep Mode me chala gaya hai.*\n\n• Reviews processing & alerts pause ho gaye hain.\n• Uthne ke liye **/start** ya **/wake** bhejein ya niche button dabayein.`,
      {
        parse_mode: 'Markdown',
        reply_markup: getMainMenuKeyboard()
      }
    );
    return;
  }

  if (botState.isSleeping) {
    await bot.sendMessage(
      chatId,
      `💤 *Bot abhi Sleep Mode me hai.*\n\nBot ko activate karne ke liye **/start** ya **/wake** send karein, ya niche button dabayein:`,
      {
        parse_mode: 'Markdown',
        reply_markup: getMainMenuKeyboard()
      }
    );
    return;
  }

  // 1. Check for "Delete All Reviews" intent
  const isDeleteAll =
    /(?:delete\s+all|all.*delete|clear\s+all|remove\s+all|sab.*delete|saare.*delete|sare.*delete|delete.*sare|delete.*sab|delete.*reviews|delete.*review|clear.*review|review.*delete)/i.test(lowerText) &&
    !/(?:of|for|ke|ki)\s+[a-z]+/i.test(lowerText);

  if (isDeleteAll) {
    await bot.sendMessage(
      chatId,
      `⚠️ *Delete All Confirmation*\nDo you want to delete ALL reviews from Firebase database?`,
      {
        parse_mode: 'Markdown',
        reply_markup: {
          inline_keyboard: [
            [
              { text: '🗑️ Yes, Delete All Reviews', callback_data: 'cmd_delete_all_confirm' },
              { text: '❌ Cancel', callback_data: 'cmd_cancel' }
            ]
          ]
        }
      }
    );
    return;
  }

  // 2. Check for Specific Employee Deletion intent
  if (/(?:delete|remove|hata|hatao)/i.test(lowerText) && /(?:review|reviews)/i.test(lowerText)) {
    const matchedSlug = findBestEmployeeMatch(lowerText);
    if (matchedSlug && PEOPLE[matchedSlug]) {
      const empName = PEOPLE[matchedSlug].name;
      await bot.sendMessage(chatId, `⏳ Deleting reviews for *${empName}*…`, { parse_mode: 'Markdown' });
      const delRes = await deleteReviewsForEmployee(matchedSlug);
      await bot.sendMessage(
        chatId,
        `🗑️ *Deleted ${delRes.totalDeleted} reviews for ${empName}* from Firebase.`,
        {
          parse_mode: 'Markdown',
          reply_markup: {
            inline_keyboard: [
              [{ text: '📊 View Status', callback_data: 'cmd_report' }]
            ]
          }
        }
      );
      return;
    }
  }

  // 3. Check for Submissions Intent
  const isSubmissionsIntent = /(?:submission|submissions|submition|proof|proofs|ss|form\s*submit)/i.test(lowerText);
  if (isSubmissionsIntent) {
    await sendSubmissionsSummaryReport(chatId);
    return;
  }

  // 4. Check for Employee Review Update intent
  const isUpdateIntent = /(?:update|change|badal|badlo|modify|edit)\s*(?:review|reviews)?/i.test(lowerText) ||
                         /(?:review|reviews)\s*(?:update|change|badal|badlo|modify|edit)/i.test(lowerText) ||
                         lowerText.startsWith('/update');

  if (isUpdateIntent) {
    const matchedSlug = findBestEmployeeMatch(lowerText);
    if (matchedSlug && PEOPLE[matchedSlug]) {
      const empName = PEOPLE[matchedSlug].name;
      
      let rawContent = rawText;
      rawContent = rawContent.replace(/^\/update\s*/i, '');
      rawContent = rawContent.replace(/^.*?(?:update|change|badal|badlo|modify|edit)\s*(?:karde|kardo|karo|please)?\s*[:=\-]?\s*/i, '');
      rawContent = rawContent.replace(/^(?:Employee|Person|Name|Staff)\s*[:=\-]\s*[A-Za-z\s]+[\n\r,]*/i, '');
      
      let extractedProduct = 'Blockzie';
      const prodMatch = rawContent.match(/(?:App Name|Product Name|Product|App|Kit)\s*[:=-]\s*([^\n★\r,]+)/i);
      if (prodMatch) {
        extractedProduct = cleanProductTitle(prodMatch[1]);
        rawContent = rawContent.replace(prodMatch[0], '');
      } else if (rawContent.includes('★')) {
        const starIdx = rawContent.indexOf('★');
        const beforeStars = rawContent.slice(0, starIdx).trim();
        if (beforeStars.length > 2 && beforeStars.length < 35) {
          extractedProduct = cleanProductTitle(beforeStars);
          rawContent = rawContent.slice(starIdx);
        }
      }

      let cleanedText = rawContent
        .replace(/^[★\u2605\u2B50\s\-:]+/g, '')
        .replace(/^(?:Review|Text|New Review)\s*[:=\-]\s*/i, '')
        .trim();

      if (cleanedText.length > 5) {
        await bot.sendMessage(chatId, `⏳ Updating review for *${empName}*…`, { parse_mode: 'Markdown' });
        try {
          const updRes = await updateReviewInFirestore(matchedSlug, extractedProduct, cleanedText);
          await bot.sendMessage(
            chatId,
            `✏️ *Review Successfully ${updRes.status === 'updated' ? 'Updated' : 'Added'}!*\n` +
            `━━━━━━━━━━━━━━━━━━━━━━\n` +
            `👤 *Employee:* ${updRes.personName}\n` +
            `📦 *Product:* ${updRes.product}\n` +
            `📝 *Review Text:*\n"${updRes.text}"\n` +
            `━━━━━━━━━━━━━━━━━━━━━━\n` +
            `⚡ *Updated in Firebase Database!*`,
            {
              parse_mode: 'Markdown',
              reply_markup: {
                inline_keyboard: [
                  [{ text: '📊 View Status Report', callback_data: 'cmd_report' }],
                  [{ text: '🌐 Open Live Directory', url: LIVE_DIRECTORY_URL }]
                ]
              }
            }
          );
          return;
        } catch (uErr) {
          console.error('Update error:', uErr);
          await bot.sendMessage(chatId, `❌ *Failed to update review:* ${uErr.message}`, { parse_mode: 'Markdown' });
          return;
        }
      }
    }
  }

  // 5. Check for Report / Stats / Status intent
  const isReportIntent = /(?:report|stat|stats|status|count|kitne|kitna|summary|total|kya\s+hai|batao)/i.test(lowerText);
  if (isReportIntent) {
    await sendDatabaseStatusReport(chatId);
    return;
  }

  // 6. Check for Greeting / Help
  const isGreeting = /^(?:hi|hello|hey|hola|namaste|kem cho|help|kya kar sakte ho)\b/i.test(lowerText);
  if (isGreeting) {
    await bot.sendMessage(
      chatId,
      `👋 *Hello! I am your STEMbotix Review & Submission Manager.*\n\n` +
      `⚡ *How I Help You:*\n` +
      `• 📬 *Live Submissions*: I alert you with full screenshots when an employee submits reviews.\n` +
      `• 📄 *Auto-Assign Reviews*: Drop PDF / Word / Excel / Text files.\n` +
      `• ✏️ *Update Review*: Say *"Divy ka review update karde: Blockzie ★★★★★ new text"*\n` +
      `• 📊 *Check Status*: Say *"kitne review add huye report de"*\n` +
      `• 🗑️ *Delete Reviews*: Say *"all review delete karde"*`,
      {
        parse_mode: 'Markdown',
        reply_markup: {
          inline_keyboard: [
            [
              { text: '📊 Reviews Report', callback_data: 'cmd_report' },
              { text: '📬 View Submissions', callback_data: 'cmd_subs' }
            ],
            [
              { text: '🗑️ Delete All Reviews', callback_data: 'cmd_delete_all_prompt' },
              { text: '🌐 Open Live Directory', url: LIVE_DIRECTORY_URL }
            ]
          ]
        }
      }
    );
    return;
  }

  // 7. If it looks like pasted review text (has stars ★ or numbered format or employee name)
  const isReviewContent = rawText.includes('★') || /^\s*\d+[\.\)\-]\s+[A-Za-z]/m.test(rawText) || /(?:App Name|Employee Name|Product)\s*[:=-]/i.test(rawText);

  if (isReviewContent && rawText.length > 25) {
    try {
      await bot.sendMessage(chatId, '⏳ *Review text detected!* Analyzing and auto-assigning…', { parse_mode: 'Markdown' });

      const reviews = await callGeminiAPI(rawText);
      if (!reviews || !reviews.length) {
        await bot.sendMessage(chatId, '⚠️ Could not detect any valid employee reviews in the provided text.');
        return;
      }

      let successCount = 0;
      for (let r of reviews) {
        try {
          await saveReviewToFirestore(r);
          successCount++;
        } catch (saveErr) {
          console.error('Failed to save review:', saveErr);
        }
      }

      await sendChunkedReport(chatId, reviews, successCount);
      return;
    } catch (err) {
      console.error('Error processing text review:', err);
      await bot.sendMessage(chatId, `❌ *Error:* ${err.message}`, { parse_mode: 'Markdown' });
      return;
    }
  }

  // 8. Default Fallback - Friendly Helpful Response with Interactive Buttons
  await bot.sendMessage(
    chatId,
    `💡 *I didn't quite catch that format.*\n\n` +
    `Here is what you can do:\n` +
    `• 📬 Type *"submissions"* to view recent employee submissions.\n` +
    `• 📄 *Drop a document* (PDF / Word / Excel / Text) to auto-assign.\n` +
    `• ✏️ Type *"update review of [Name]: [new text]"* to modify a review.\n` +
    `• 📊 Type *"report"* to see all assigned reviews.\n` +
    `• 🗑️ Type *"all review delete karde"* to clean the database.`,
    {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [
            { text: '📊 Reviews Report', callback_data: 'cmd_report' },
            { text: '📬 View Submissions', callback_data: 'cmd_subs' }
          ],
          [
            { text: '🗑️ Delete All Reviews', callback_data: 'cmd_delete_all_prompt' },
            { text: '🌐 Open Live Directory', url: LIVE_DIRECTORY_URL }
          ]
        ]
      }
    }
  );
});
