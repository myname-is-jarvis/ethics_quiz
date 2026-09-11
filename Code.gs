/**
 * 행사 퀴즈 백엔드 (Google Apps Script)
 *
 * 시트 구성 (setupSheets() 실행 시 자동 생성)
 *  Questions : id | question | opt1 | opt2 | opt3 | opt4 | answer(1~4) | active(Y/N, 비우면 Y)
 *  Results   : timestamp | name | phone | score | total | question_ids | choices | correct_flags | user_agent
 *  Stats     : buildStats() 실행 시 생성/갱신
 *
 * 배포: 배포 > 새 배포 > 유형 "웹 앱" > 실행 사용자 "나" > 액세스 "모든 사용자"
 * Code.gs를 수정한 뒤에는 "배포 관리 > 새 버전"으로 다시 배포해야 반영됩니다.
 * (Questions 시트의 문제 수정은 재배포 없이 즉시 반영됩니다.)
 */

const CONFIG = {
  QUESTIONS_PER_QUIZ: 10,
  SHEET_QUESTIONS: 'Questions',
  SHEET_RESULTS: 'Results',
  SHEET_STATS: 'Stats',
  ALLOW_DUPLICATE_PHONE: false, // true로 바꾸면 같은 번호로 여러 번 참여 허용
};

/* ---------- 최초 1회: 시트 뼈대 생성 ---------- */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let q = ss.getSheetByName(CONFIG.SHEET_QUESTIONS);
  if (!q) {
    q = ss.insertSheet(CONFIG.SHEET_QUESTIONS);
    q.appendRow(['id', 'question', 'opt1', 'opt2', 'opt3', 'opt4', 'answer', 'active']);
    q.appendRow(['Q001', '예시 문제: 이 행사의 개최 연도는?', '2024', '2025', '2026', '2027', 3, 'Y']);
    q.setFrozenRows(1);
  }
  if (!ss.getSheetByName(CONFIG.SHEET_RESULTS)) createResultsSheet_(ss);
}

function createResultsSheet_(ss) {
  const s = ss.insertSheet(CONFIG.SHEET_RESULTS);
  s.appendRow(['timestamp', 'name', 'phone', 'score', 'total', 'question_ids', 'choices', 'correct_flags', 'user_agent']);
  s.getRange('C:C').setNumberFormat('@'); // 전화번호 앞자리 0 보존
  s.setFrozenRows(1);
  return s;
}

/* ---------- HTTP 진입점 ---------- */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'questions';
    if (action === 'questions') return json_(getRandomQuestions_());
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad json' });
  }

  const name = String(body.name || '').trim().slice(0, 50);
  const phone = normalizePhone_(body.phone);
  const answers = Array.isArray(body.answers) ? body.answers : [];
  if (!name || !phone) return json_({ ok: false, error: 'invalid name or phone' });
  if (answers.length !== CONFIG.QUESTIONS_PER_QUIZ) return json_({ ok: false, error: 'answer count mismatch' });

  // 동시 제출 시 중복 검사와 행 추가가 섞이지 않도록 직렬화
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = ss.getSheetByName(CONFIG.SHEET_RESULTS) || createResultsSheet_(ss);

    if (!CONFIG.ALLOW_DUPLICATE_PHONE && isDuplicate_(results, phone)) {
      return json_({ ok: false, error: 'duplicate' });
    }

    // 비활성(active=N) 문제도 포함해서 채점: 행사 중 문제를 내려도 이미 받은 사람은 채점 가능
    const key = {};
    readQuestions_(true).forEach(q => { key[q.id] = q; });

    const graded = answers.map(a => {
      const q = key[String(a.id)];
      const choice = String(a.choice == null ? '' : a.choice).trim();
      return {
        id: String(a.id),
        choice,
        correct: !!q && choice === q.answer,
        answer: q ? q.answer : '',
      };
    });
    const score = graded.filter(g => g.correct).length;

    results.appendRow([
      new Date(),
      name,
      phone,
      score,
      graded.length,
      graded.map(g => g.id).join(','),
      JSON.stringify(graded.map(g => g.choice)),
      JSON.stringify(graded.map(g => (g.correct ? 1 : 0))),
      String(body.ua || '').slice(0, 200),
    ]);

    return json_({
      ok: true,
      score,
      total: graded.length,
      results: graded.map(g => ({ id: g.id, correct: g.correct, answer: g.answer })),
    });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 문제 ---------- */
function readQuestions_(includeInactive) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_QUESTIONS);
  if (!sh) throw new Error('Questions 시트가 없습니다. setupSheets()를 먼저 실행하세요.');
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const id = String(r[0] == null ? '' : r[0]).trim();
    const question = String(r[1] == null ? '' : r[1]).trim();
    if (!id || !question) continue;

    const active = String(r[7] == null ? '' : r[7]).trim().toUpperCase() !== 'N';
    if (!active && !includeInactive) continue;

    const options = [r[2], r[3], r[4], r[5]]
      .map(v => String(v == null ? '' : v).trim())
      .filter(v => v !== '');
    const ansIdx = Number(r[6]);
    if (options.length < 2 || !(ansIdx >= 1 && ansIdx <= 4)) continue;
    const answer = String(r[ansIdx + 1] == null ? '' : r[ansIdx + 1]).trim(); // opt1 = r[2]
    if (!answer) continue;

    out.push({ id, question, options, answer });
  }
  return out;
}

function getRandomQuestions_() {
  const pool = readQuestions_(false);
  if (pool.length < CONFIG.QUESTIONS_PER_QUIZ) {
    throw new Error('활성 문제가 ' + CONFIG.QUESTIONS_PER_QUIZ + '개보다 적습니다 (현재 ' + pool.length + '개)');
  }
  const picked = shuffle_(pool).slice(0, CONFIG.QUESTIONS_PER_QUIZ);
  // 정답은 내려보내지 않음 (채점은 서버에서)
  return {
    ok: true,
    questions: picked.map(q => ({
      id: q.id,
      question: q.question,
      options: shuffle_(q.options.slice()),
    })),
  };
}

function shuffle_(arr) { // Fisher–Yates
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}

/* ---------- 참가자 ---------- */
function isDuplicate_(results, phone) {
  const last = results.getLastRow();
  if (last < 2) return false;
  const phones = results.getRange(2, 3, last - 1, 1).getValues();
  return phones.some(row => normalizePhone_(row[0]) === phone);
}

function normalizePhone_(v) {
  const d = String(v == null ? '' : v).replace(/\D/g, '');
  return /^01[016789]\d{7,8}$/.test(d) ? d : '';
}

/* ---------- 통계: 스크립트 편집기에서 buildStats() 실행 → Stats 시트 ---------- */
function buildStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = ss.getSheetByName(CONFIG.SHEET_RESULTS);
  if (!results || results.getLastRow() < 2) throw new Error('아직 결과가 없습니다.');

  const rows = results.getRange(2, 1, results.getLastRow() - 1, 9).getValues();
  const qmap = {};
  readQuestions_(true).forEach(q => { qmap[q.id] = q.question; });

  const total = CONFIG.QUESTIONS_PER_QUIZ;
  const hist = new Array(total + 1).fill(0);
  const perQ = {};
  let sum = 0;

  rows.forEach(r => {
    const score = Number(r[3]) || 0;
    sum += score;
    if (score >= 0 && score <= total) hist[score]++;

    const ids = String(r[5]).split(',').map(s => s.trim()).filter(Boolean);
    let flags = [];
    try { flags = JSON.parse(r[7]); } catch (e) { /* 손상된 행은 문제별 집계에서 제외 */ }
    ids.forEach((id, i) => {
      if (!perQ[id]) perQ[id] = { asked: 0, correct: 0 };
      perQ[id].asked++;
      if (flags[i] === 1) perQ[id].correct++;
    });
  });

  const n = rows.length;
  const out = [];
  out.push(['집계 시각', new Date()]);
  out.push(['참여자 수', n]);
  out.push(['평균 점수', n ? Math.round((sum / n) * 100) / 100 : 0]);
  out.push([]);
  out.push(['점수', '인원']);
  hist.forEach((c, s) => out.push([s, c]));
  out.push([]);
  out.push(['문제 id', '문제', '출제 횟수', '정답 횟수', '정답률(%)']);
  Object.keys(perQ).sort().forEach(id => {
    const p = perQ[id];
    out.push([id, qmap[id] || '(시트에서 삭제된 문제)', p.asked, p.correct,
      p.asked ? Math.round((p.correct / p.asked) * 1000) / 10 : '']);
  });
  out.push([]);
  out.push(['순위', '이름', '전화번호', '점수', '제출 시각']); // 동점이면 먼저 제출한 사람이 위
  rows.slice()
    .sort((a, b) => (Number(b[3]) - Number(a[3])) || (new Date(a[0]) - new Date(b[0])))
    .forEach((r, i) => out.push([i + 1, r[1], r[2], r[3], r[0]]));

  let st = ss.getSheetByName(CONFIG.SHEET_STATS);
  if (!st) st = ss.insertSheet(CONFIG.SHEET_STATS);
  st.clearContents();
  const width = Math.max.apply(null, out.map(r => r.length));
  const padded = out.map(r => { const c = r.slice(); while (c.length < width) c.push(''); return c; });
  st.getRange(1, 1, padded.length, width).setValues(padded);
}

/* ---------- 편집기에서 동작 확인용 ---------- */
function testQuestions() {
  Logger.log(JSON.stringify(getRandomQuestions_(), null, 2));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
