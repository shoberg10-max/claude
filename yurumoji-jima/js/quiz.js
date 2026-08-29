// ゆるもじ島 - 出題エンジン
function buildChoices(correctValue, candidateValues) {
  const distractors = shuffle(candidateValues.filter(v => v !== correctValue));
  const unique = [];
  for (const d of distractors) {
    if (!unique.includes(d)) unique.push(d);
    if (unique.length === 3) break;
  }
  return shuffle([correctValue, ...unique]);
}

function poolForLevel(level) {
  const pool = KANJI_BY_LEVEL[level];
  return pool && pool.length >= 4 ? pool : KANJI_DATA;
}

// 1つの漢字につき: なぞり(採点なし) → よみ → かきとり(かんじ再認) → じゅくご よみ
function generateMissionQuestions(entries) {
  const questions = [];
  entries.forEach(entry => {
    const pool = poolForLevel(entry.level);
    questions.push({ type: "trace", entry });
    questions.push({
      type: "read",
      entry,
      choices: buildChoices(entry.reading, pool.map(e => e.reading)),
      answer: entry.reading,
    });
    questions.push({
      type: "writeRecall",
      entry,
      choices: buildChoices(entry.k, pool.map(e => e.k)),
      answer: entry.k,
    });
    questions.push({
      type: "word",
      entry,
      choices: buildChoices(entry.wordReading, pool.map(e => e.wordReading)),
      answer: entry.wordReading,
    });
  });
  return questions;
}

function generateKenteiQuestions(state, count) {
  const attempted = KANJI_DATA.filter(e => {
    const ks = state.kanjiState[e.k];
    return ks && (ks.read + ks.write + ks.word) > 0;
  });
  const base = attempted.length >= count ? attempted : KANJI_DATA;
  const picked = shuffle(base).slice(0, Math.min(count, base.length));
  return picked.map(entry => {
    const pool = poolForLevel(entry.level);
    return {
      type: "read",
      entry,
      choices: buildChoices(entry.reading, pool.map(e => e.reading)),
      answer: entry.reading,
    };
  });
}
