// ゆるもじ島のクイズ生成（KANJI1 / KANJI2 を使う）
const KanjiGame = (() => {

  function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function datasetFor(level) {
    return level === 'k8' ? KANJI2 : KANJI1;
  }

  // まだ覚えていない漢字を優先しつつランダムに選ぶ
  function pickKanjiSet(level, count) {
    const dataset = datasetFor(level);
    const weighted = dataset.map(entry => {
      const lv = Storage.getIslandKanjiLevel(level, entry.k); // 0〜3
      return { entry, weight: 4 - lv };
    });
    const pool = [];
    weighted.forEach(w => {
      for (let i = 0; i < w.weight; i++) pool.push(w.entry);
    });
    const chosen = [];
    const used = new Set();
    let guard = 0;
    while (chosen.length < count && guard < 4000) {
      guard++;
      const e = pool[randInt(0, pool.length - 1)];
      if (!used.has(e.k)) {
        used.add(e.k);
        chosen.push(e);
      }
    }
    return chosen;
  }

  function allReadingsOf(entry) {
    return [...entry.on, ...entry.kun];
  }

  function randomReadingFor(entry) {
    const readings = entry.kun.length && Math.random() < 0.6 ? entry.kun : (entry.on.length ? entry.on : entry.kun);
    return readings[randInt(0, readings.length - 1)];
  }

  function makeReadingQuestion(level, entry) {
    const dataset = datasetFor(level);
    const correct = randomReadingFor(entry);
    const ownReadings = new Set(allReadingsOf(entry));
    const distractorPool = dataset
      .filter(e => e.k !== entry.k)
      .flatMap(allReadingsOf)
      .filter(r => r && !ownReadings.has(r));
    const distractors = new Set();
    let guard = 0;
    while (distractors.size < 3 && guard < 500) {
      guard++;
      distractors.add(distractorPool[randInt(0, distractorPool.length - 1)]);
    }
    return {
      kanji: entry.k,
      mode: 'reading',
      instruction: 'この かんじの よみかたは？',
      prompt: entry.k,
      choices: shuffle([correct, ...distractors]),
      answer: correct
    };
  }

  function makeFindQuestion(level, entry) {
    const dataset = datasetFor(level);
    const correct = entry.k;
    const reading = randomReadingFor(entry);
    const distractorPool = dataset.filter(e => e.k !== entry.k).map(e => e.k);
    const distractors = new Set();
    let guard = 0;
    while (distractors.size < 3 && guard < 500) {
      guard++;
      distractors.add(distractorPool[randInt(0, distractorPool.length - 1)]);
    }
    return {
      kanji: entry.k,
      mode: 'find',
      instruction: 'この よみかたの かんじは？',
      prompt: reading,
      choices: shuffle([correct, ...distractors]),
      answer: correct
    };
  }

  // ミッション用：読み・かんじさがしを混ぜたセットを作る
  function generateMission(level, count = 8) {
    const chosen = pickKanjiSet(level, count);
    return chosen.map(entry => {
      return Math.random() < 0.5 ? makeReadingQuestion(level, entry) : makeFindQuestion(level, entry);
    });
  }

  // 漢検チャレンジ用：全字からまんべんなく出題
  function generateChallenge(level, count = 20) {
    const dataset = shuffle(datasetFor(level)).slice(0, count);
    return dataset.map(entry => {
      return Math.random() < 0.5 ? makeReadingQuestion(level, entry) : makeFindQuestion(level, entry);
    });
  }

  function pickWritingKanji(level, fromEntries, n = 2) {
    return shuffle(fromEntries).slice(0, n);
  }

  return {
    datasetFor,
    generateMission,
    generateChallenge,
    pickWritingKanji
  };
})();
