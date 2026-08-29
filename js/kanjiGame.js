// ゆるもじ島のクイズ生成（KANJI_K9 / KANJI_K8 を使う）
// 出題データは claude/kanji-kentei-8-app-d3s15h の data.js（GRADE2_DATA / KANJI_DATA）と
// 同じ内容。1字ずつの音訓読みではなく、その字を使った熟語・ことば（word）と
// その読み方（reading）の組で出題する。
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
    return level === 'k8' ? KANJI_K8 : KANJI_K9;
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

  // 読み方クイズ：ことば（熟語など）を見て、正しい読みを選ぶ
  function makeReadingQuestion(level, entry) {
    const dataset = datasetFor(level);
    const distractorPool = dataset
      .filter(e => e.k !== entry.k && e.reading !== entry.reading)
      .map(e => e.reading);
    const distractors = new Set();
    let guard = 0;
    while (distractors.size < 3 && guard < 500) {
      guard++;
      distractors.add(distractorPool[randInt(0, distractorPool.length - 1)]);
    }
    return {
      kanji: entry.k,
      mode: 'reading',
      instruction: 'つぎの ことばの よみかたは？',
      prompt: entry.word,
      choices: shuffle([entry.reading, ...distractors]),
      answer: entry.reading
    };
  }

  // かんじさがしクイズ：読みを見て、正しいことば（漢字）を選ぶ
  function makeFindQuestion(level, entry) {
    const dataset = datasetFor(level);
    const distractorPool = dataset.filter(e => e.k !== entry.k).map(e => e.word);
    const distractors = new Set();
    let guard = 0;
    while (distractors.size < 3 && guard < 500) {
      guard++;
      distractors.add(distractorPool[randInt(0, distractorPool.length - 1)]);
    }
    return {
      kanji: entry.k,
      mode: 'find',
      instruction: 'つぎの よみかたの ことばは？',
      prompt: entry.reading,
      choices: shuffle([entry.word, ...distractors]),
      answer: entry.word
    };
  }

  // ミッション用：読み・かんじさがしを混ぜたセットを作る
  function generateMission(level, count = 8) {
    const chosen = pickKanjiSet(level, count);
    return chosen.map(entry => {
      return Math.random() < 0.5 ? makeReadingQuestion(level, entry) : makeFindQuestion(level, entry);
    });
  }

  // ふくしゅうミッション用：忘れかけの漢字だけを出す
  function generateReview(level, entries, count = 8) {
    return shuffle(entries).slice(0, count).map(entry => {
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
    generateReview,
    generateChallenge,
    pickWritingKanji
  };
})();
