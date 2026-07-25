// localStorage を使った進み具合の保存
const Storage = (() => {
  const KEY = 'kids-drill-v1';

  function load() {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return defaultState();
      const parsed = JSON.parse(raw);
      return Object.assign(defaultState(), parsed);
    } catch (e) {
      return defaultState();
    }
  }

  function defaultState() {
    return {
      stars: 0,
      bestScores: {},   // topicId -> best correct count (out of 10)
      masteredKanji: {} // kanji -> correct streak count
    };
  }

  function save(state) {
    try {
      localStorage.setItem(KEY, JSON.stringify(state));
    } catch (e) {
      // ストレージが使えない場合は無視する
    }
  }

  let state = load();

  return {
    getStars() { return state.stars; },
    addStars(n) {
      state.stars += n;
      save(state);
      return state.stars;
    },
    getBestScore(topicId) {
      return state.bestScores[topicId] || 0;
    },
    setBestScore(topicId, score) {
      if (score > (state.bestScores[topicId] || 0)) {
        state.bestScores[topicId] = score;
        save(state);
      }
    },
    markKanjiResult(kanji, correct) {
      const cur = state.masteredKanji[kanji] || 0;
      state.masteredKanji[kanji] = correct ? Math.min(cur + 1, 3) : Math.max(cur - 1, 0);
      save(state);
    },
    getKanjiLevel(kanji) {
      return state.masteredKanji[kanji] || 0;
    },
    getMasteredCount() {
      return Object.values(state.masteredKanji).filter(v => v >= 3).length;
    }
  };
})();
