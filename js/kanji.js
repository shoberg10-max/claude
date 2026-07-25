// 漢字クイズを作る（KANJI3 は js/data/kanji3.js で定義）
const KanjiQuiz = (() => {

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

  // まだ覚えていない漢字を優先しつつランダムに選ぶ
  function pickKanjiSet(count) {
    const weighted = KANJI3.map(entry => {
      const level = Storage.getKanjiLevel(entry.k); // 0〜3
      return { entry, weight: 4 - level };
    });
    const pool = [];
    weighted.forEach(w => {
      for (let i = 0; i < w.weight; i++) pool.push(w.entry);
    });
    const chosen = [];
    const used = new Set();
    let guard = 0;
    while (chosen.length < count && guard < 2000) {
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

  // 漢字を見て正しい読みを選ぶ
  function generateReadingQuiz(count = 10) {
    const chosen = pickKanjiSet(count);
    return chosen.map(entry => {
      const correct = randomReadingFor(entry);
      const ownReadings = new Set(allReadingsOf(entry));
      const distractorPool = KANJI3
        .filter(e => e.k !== entry.k)
        .flatMap(allReadingsOf)
        .filter(r => r && !ownReadings.has(r));
      const distractors = new Set();
      let guard = 0;
      while (distractors.size < 3 && guard < 500) {
        guard++;
        distractors.add(distractorPool[randInt(0, distractorPool.length - 1)]);
      }
      const choices = shuffle([correct, ...distractors]);
      return {
        kanji: entry.k,
        prompt: entry.k,
        mode: 'reading',
        choices,
        answer: correct
      };
    });
  }

  // 読みを見て正しい漢字を選ぶ
  function generateKanjiQuiz(count = 10) {
    const chosen = pickKanjiSet(count);
    return chosen.map(entry => {
      const correct = entry.k;
      const reading = randomReadingFor(entry);
      const distractorPool = KANJI3.filter(e => e.k !== entry.k).map(e => e.k);
      const distractors = new Set();
      let guard = 0;
      while (distractors.size < 3 && guard < 500) {
        guard++;
        distractors.add(distractorPool[randInt(0, distractorPool.length - 1)]);
      }
      const choices = shuffle([correct, ...distractors]);
      return {
        kanji: entry.k,
        prompt: reading,
        mode: 'kanji',
        choices,
        answer: correct
      };
    });
  }

  return {
    generateReadingQuiz,
    generateKanjiQuiz,
    all: KANJI3
  };
})();
