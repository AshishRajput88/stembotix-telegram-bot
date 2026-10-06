const fs = require('fs');

const DEPARTMENTS = JSON.parse(fs.readFileSync('./team-data.json', 'utf8'));
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
      if (part.length > 2 && (query === part || query.startsWith(part) || query.endsWith(part))) {
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

function parseTextLocalHeuristic(rawText) {
  if (!rawText) return [];
  let text = String(rawText);

  // Pre-split on lines that start with numbers + names like "43. Divy" or "44. Azam"
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

      // Clean trailing employee numbers like "44. Azam" at the end of review
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
      // Check if line contains stars indicating start of review
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

const sampleDoc = `42. Yukta
Blockzie ★★★★★ The drag-and-drop system is convenient and beginner-friendly.

43. Divy
App Name: Blockzie Playstore ★★★★★ Good educational app with a creative focus.
App Name: Botzie Playstore ★★★★★ Great robotics app.

44. Azam
Blockzie ★★★★★ Blockzie encourages curiosity and experimentation. It gives learners a space to try ideas, solve problems and understand coding through practice.`;

const parsed = parseTextLocalHeuristic(sampleDoc);
console.log('Total parsed reviews:', parsed.length);
console.log(JSON.stringify(parsed, null, 2));
