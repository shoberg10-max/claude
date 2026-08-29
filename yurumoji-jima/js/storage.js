// ゆるもじ島 - セーブデータ管理 (localStorage)
const SAVE_KEY = "yurumoji-jima-save-v1";

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function defaultKanjiState() {
  return { read: 0, write: 0, word: 0, wrong: 0 };
}

function defaultState() {
  const kanjiState = {};
  KANJI_DATA.forEach(e => { kanjiState[e.k] = defaultKanjiState(); });
  return {
    xp: 0,
    coins: 0,
    streak: 0,
    lastPlayDate: null,
    restTicket: 1,
    totalCorrect: 0,
    chestProgress: 0,
    inventory: [],
    kanjiState,
    history: [], // [{date, correct, total}]
    missionsToday: 0,
    missionsTodayDate: null,
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(SAVE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    const def = defaultState();
    // マージ（新しい漢字が追加された場合に備える）
    const merged = Object.assign({}, def, parsed);
    merged.kanjiState = Object.assign({}, def.kanjiState, parsed.kanjiState || {});
    return merged;
  } catch (e) {
    return defaultState();
  }
}

function saveState(state) {
  try {
    localStorage.setItem(SAVE_KEY, JSON.stringify(state));
  } catch (e) {
    // localStorageが使えない環境では何もしない
  }
}

function resetState() {
  localStorage.removeItem(SAVE_KEY);
  return defaultState();
}
