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
      masteredKanji: {}, // kanji -> correct streak count
      island: defaultIsland()
    };
  }

  function defaultIsland() {
    return {
      level: 'k9',            // 'k9'=9級(1年) / 'k8'=8級(2年)
      coins: 0,
      tickets: 0,
      charXp: 0,
      streak: 0,
      restTickets: 1,
      lastPlayDate: null,
      missionCount: { k9: 0, k8: 0 },
      kanjiState: {},          // `${level}:${kanji}` -> 0..3
      companions: ['mojimaru'],
      items: [],
      equippedItem: null,
      kankenBest: {},          // level -> { pct, readPct, findPct, at }
      dailyLog: {}             // 'YYYY-MM-DD' -> { total, correct }
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
  // 古いデータに island が無い場合の保険
  if (!state.island) state.island = defaultIsland();
  state.island = Object.assign(defaultIsland(), state.island);

  function todayStr() {
    return new Date().toISOString().slice(0, 10);
  }

  function charLevelFromXp(xp) {
    // レベル n に必要な累計経験値 = 50 * n * (n+1) / 2 くらいの緩やかな曲線
    let level = 1;
    while (xp >= 60 * level) {
      xp -= 60 * level;
      level++;
    }
    return { level, xpIntoLevel: xp, xpForNext: 60 * level };
  }

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
    },

    // ---------- ゆるもじ島 ----------
    getIsland() { return state.island; },

    setIslandLevel(level) {
      state.island.level = level;
      save(state);
    },

    getIslandKanjiLevel(level, kanji) {
      return state.island.kanjiState[level + ':' + kanji] || 0;
    },

    recordIslandKanjiResult(level, kanji, correct) {
      const key = level + ':' + kanji;
      const cur = state.island.kanjiState[key] || 0;
      state.island.kanjiState[key] = correct ? Math.min(cur + 1, 3) : Math.max(cur - 1, 0);
      save(state);
    },

    getIslandMasteredCount(level, kanjiList) {
      return kanjiList.filter(e => (state.island.kanjiState[level + ':' + e.k] || 0) >= 3).length;
    },

    addCoins(n) {
      state.island.coins += n;
      save(state);
      return state.island.coins;
    },

    addTickets(n) {
      state.island.tickets = Math.max(0, state.island.tickets + n);
      save(state);
      return state.island.tickets;
    },

    useTicket() {
      if (state.island.tickets <= 0) return false;
      state.island.tickets -= 1;
      save(state);
      return true;
    },

    addCharXp(n) {
      const before = charLevelFromXp(state.island.charXp).level;
      state.island.charXp += n;
      const after = charLevelFromXp(state.island.charXp).level;
      save(state);
      return { leveledUp: after > before, level: after };
    },

    getCharLevelInfo() {
      return charLevelFromXp(state.island.charXp);
    },

    incMissionCount(level) {
      state.island.missionCount[level] = (state.island.missionCount[level] || 0) + 1;
      save(state);
      return state.island.missionCount[level];
    },

    getMissionCount(level) {
      return state.island.missionCount[level] || 0;
    },

    addItem(itemId) {
      const isNew = !state.island.items.includes(itemId);
      if (isNew) state.island.items.push(itemId);
      save(state);
      return isNew;
    },

    getItems() { return state.island.items.slice(); },

    equipItem(itemId) {
      state.island.equippedItem = itemId;
      save(state);
    },

    getEquippedItem() { return state.island.equippedItem; },

    getCompanions() { return state.island.companions.slice(); },

    unlockCompanionIfEligible(companionId, need, totalMastered) {
      if (state.island.companions.includes(companionId)) return false;
      if (totalMastered >= need) {
        state.island.companions.push(companionId);
        save(state);
        return true;
      }
      return false;
    },

    // 毎日1回、アプリを開いたときに呼ぶ。連続日数を更新する
    touchStreak() {
      const today = todayStr();
      const last = state.island.lastPlayDate;
      if (last === today) return { streak: state.island.streak, changed: false };
      if (!last) {
        state.island.streak = 1;
      } else {
        const diffDays = Math.round((new Date(today) - new Date(last)) / 86400000);
        if (diffDays === 1) {
          state.island.streak += 1;
        } else if (diffDays === 2 && state.island.restTickets > 0) {
          state.island.restTickets -= 1;
          state.island.streak += 1;
        } else {
          state.island.streak = 1;
        }
      }
      // 7日ごとに「おやすみ券」を1枚補充
      if (state.island.streak > 0 && state.island.streak % 7 === 0) {
        state.island.restTickets = Math.min(state.island.restTickets + 1, 3);
      }
      state.island.lastPlayDate = today;
      save(state);
      return { streak: state.island.streak, changed: true };
    },

    getStreak() { return state.island.streak; },
    getRestTickets() { return state.island.restTickets; },

    addDailyLog(total, correct) {
      const today = todayStr();
      const cur = state.island.dailyLog[today] || { total: 0, correct: 0 };
      cur.total += total;
      cur.correct += correct;
      state.island.dailyLog[today] = cur;
      save(state);
    },

    getDailyLog(days = 7) {
      const out = [];
      for (let i = days - 1; i >= 0; i--) {
        const d = new Date();
        d.setDate(d.getDate() - i);
        const key = d.toISOString().slice(0, 10);
        out.push({ date: key, ...(state.island.dailyLog[key] || { total: 0, correct: 0 }) });
      }
      return out;
    },

    setKankenBest(level, result) {
      const cur = state.island.kankenBest[level];
      if (!cur || result.pct > cur.pct) {
        state.island.kankenBest[level] = result;
        save(state);
      }
    },

    getKankenBest(level) {
      return state.island.kankenBest[level] || null;
    }
  };
})();
