// ゆるもじ島 - ゲームロジック
const XP_PER_LEVEL = 50;
const MASTERY_THRESHOLD = 2; // 各カテゴリでこの回数正解したら「おぼえた」
const CHEST_EVERY = 5; // 連続正解チェストの間隔

function levelFromXp(xp) {
  return Math.floor(xp / XP_PER_LEVEL) + 1;
}

function xpProgress(xp) {
  const level = levelFromXp(xp);
  const intoLevel = xp - (level - 1) * XP_PER_LEVEL;
  return { level, intoLevel, needed: XP_PER_LEVEL, percent: Math.round((intoLevel / XP_PER_LEVEL) * 100) };
}

function characterStage(level) {
  const stages = CHARACTERS[0].stages;
  let current = stages[0];
  for (const s of stages) {
    if (level >= s.lv) current = s;
  }
  return current;
}

function isMastered(state, kanjiChar) {
  const ks = state.kanjiState[kanjiChar];
  if (!ks) return false;
  return ks.read >= MASTERY_THRESHOLD && ks.write >= MASTERY_THRESHOLD && ks.word >= MASTERY_THRESHOLD;
}

function masteredCount(state, level) {
  const pool = level ? KANJI_BY_LEVEL[level] : KANJI_DATA;
  return pool.filter(e => isMastered(state, e.k)).length;
}

function recordAnswer(state, kanjiChar, category, correct) {
  const ks = state.kanjiState[kanjiChar] || defaultKanjiState();
  state.kanjiState[kanjiChar] = ks;
  if (correct) {
    ks[category] = (ks[category] || 0) + 1;
    state.totalCorrect += 1;
    state.chestProgress += 1;
    state.xp += 10;
    state.coins += 2;
  } else {
    ks.wrong = (ks.wrong || 0) + 1;
  }
}

function checkChestEarned(state) {
  if (state.chestProgress >= CHEST_EVERY) {
    state.chestProgress -= CHEST_EVERY;
    return true;
  }
  return false;
}

function islandStage(state) {
  const count = masteredCount(state);
  let stage = ISLAND_STAGES[0];
  for (const s of ISLAND_STAGES) {
    if (count >= s.min) stage = s;
  }
  return stage;
}

function tamagoStage(state) {
  let stage = TAMAGO_STAGES[0];
  for (const s of TAMAGO_STAGES) {
    if (state.totalCorrect >= s.min) stage = s;
  }
  return stage;
}

function drawGacha(state) {
  const notOwned = ITEMS.filter(i => !state.inventory.includes(i.id));
  let item, isNew;
  if (notOwned.length > 0) {
    item = notOwned[Math.floor(Math.random() * notOwned.length)];
    state.inventory.push(item.id);
    isNew = true;
  } else {
    item = ITEMS[Math.floor(Math.random() * ITEMS.length)];
    state.coins += 10;
    isNew = false;
  }
  return { item, isNew };
}

function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function pickMissionKanji(state, level, count) {
  const pool = level === "all" ? KANJI_DATA : KANJI_BY_LEVEL[level];
  const unmastered = shuffle(pool.filter(e => !isMastered(state, e.k)));
  const mastered = shuffle(pool.filter(e => isMastered(state, e.k)));
  return unmastered.concat(mastered).slice(0, count);
}

function addDays(dateStr, n) {
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

function updateStreak(state) {
  const t = todayStr();
  if (state.lastPlayDate === t) {
    return { changed: false, streak: state.streak, milestone: null };
  }
  let milestone = null;
  if (state.lastPlayDate === null) {
    state.streak = 1;
  } else if (addDays(state.lastPlayDate, 1) === t) {
    state.streak += 1;
  } else {
    if (state.restTicket > 0) {
      state.restTicket -= 1;
      state.streak += 1;
    } else {
      state.streak = 1;
    }
  }
  state.lastPlayDate = t;
  if (state.streak === 7 || state.streak === 30) {
    milestone = state.streak;
    if (state.streak === 7) state.restTicket += 1;
  }
  return { changed: true, streak: state.streak, milestone };
}

function recordMissionResult(state, correct, total) {
  const t = todayStr();
  if (state.missionsTodayDate !== t) {
    state.missionsTodayDate = t;
    state.missionsToday = 0;
  }
  state.missionsToday += 1;
  state.history.push({ date: t, correct, total });
  if (state.history.length > 60) state.history = state.history.slice(-60);
}

function kenteiUnlocked(state) {
  return masteredCount(state) >= 10;
}

function kenteiBand(percent) {
  if (percent >= 90) return { emoji: "🟢", label: "合格圏！", tone: "green" };
  if (percent >= 70) return { emoji: "🟡", label: "あと少し！", tone: "yellow" };
  return { emoji: "🔴", label: "もう一度特訓！", tone: "red" };
}

function weakKanji(state, limit) {
  return KANJI_DATA
    .map(e => ({ e, wrong: (state.kanjiState[e.k] || {}).wrong || 0 }))
    .filter(x => x.wrong > 0)
    .sort((a, b) => b.wrong - a.wrong)
    .slice(0, limit)
    .map(x => x.e);
}

function weeklyStudyDays(state) {
  const t = todayStr();
  const days = new Set();
  for (let i = 0; i < 7; i++) {
    const d = addDays(t, -i);
    if (state.history.some(h => h.date === d)) days.add(d);
  }
  return days.size;
}
