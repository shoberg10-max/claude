// 漢字検定8級 練習アプリ ロジック
(function () {
  "use strict";

  const STORAGE_KEY = "kanken8_stats_v1";

  const MODE_LABEL = {
    reading: "読み方",
    writing: "書き取り",
    strokes: "画数",
    choice: "漢字えらび",
    pair: "ことばのペア",
  };

  /* ---------- 状態 ---------- */
  let session = null; // { queue, index, correct, wrong: [], mode, count, isReview, metas? }
  let selectedMode = "reading";
  let selectedCount = 20;

  /* ---------- ユーティリティ ---------- */
  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function normalize(str) {
    return String(str)
      .trim()
      .replace(/[\s　]+/g, "")
      .replace(/[。.！!？?]+$/g, "");
  }

  function loadStats() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) throw new Error("empty");
      const parsed = JSON.parse(raw);
      if (!parsed.byCategory) throw new Error("bad shape");
      return parsed;
    } catch (e) {
      return {
        totalAnswered: 0,
        totalCorrect: 0,
        byCategory: {},
        weak: {}, // weakKey -> { wrongCount }
      };
    }
  }

  function saveStats(stats) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(stats));
  }

  // KANJI_DATA由来(読み方/書き取り/画数)は漢字1字をキーにしてまとめる。
  // 漢字えらび/ことばのペアは種類ごとに別キーにする。
  function weakKeyFor(meta) {
    if (meta.qtype === "choice") return `choice:${meta.entry.correct}`;
    if (meta.qtype === "pair") return `pair:${meta.entry.a.kanji}|${meta.entry.b.kanji}`;
    return meta.entry.kanji;
  }

  function weakLabel(key) {
    if (key.startsWith("choice:")) return key.slice(7);
    if (key.startsWith("pair:")) return key.slice(5).replace("|", "⇔");
    return key;
  }

  function metaFromWeakKey(key) {
    if (key.startsWith("choice:")) {
      const word = key.slice(7);
      const entry = CHOICE_DATA.find((e) => e.correct === word);
      return entry ? { qtype: "choice", entry } : null;
    }
    if (key.startsWith("pair:")) {
      const [ak, bk] = key.slice(5).split("|");
      const entry = PAIR_DATA.find((e) => e.a.kanji === ak && e.b.kanji === bk);
      return entry ? { qtype: "pair", entry, direction: Math.random() < 0.5 ? "a2b" : "b2a" } : null;
    }
    const entry = KANJI_DATA.find((e) => e.kanji === key);
    if (!entry) return null;
    const types = ["reading", "writing", "strokes"];
    return { qtype: types[Math.floor(Math.random() * types.length)], entry };
  }

  function recordAnswer(meta, isCorrect) {
    const stats = loadStats();
    stats.totalAnswered++;
    if (isCorrect) stats.totalCorrect++;
    const cat = meta.qtype;
    if (!stats.byCategory[cat]) stats.byCategory[cat] = { correct: 0, total: 0 };
    stats.byCategory[cat].total++;
    if (isCorrect) stats.byCategory[cat].correct++;

    const key = weakKeyFor(meta);
    if (!stats.weak[key]) stats.weak[key] = { wrongCount: 0 };
    if (isCorrect) {
      stats.weak[key].wrongCount = Math.max(0, stats.weak[key].wrongCount - 1);
      if (stats.weak[key].wrongCount === 0) delete stats.weak[key];
    } else {
      stats.weak[key].wrongCount += 2;
    }
    saveStats(stats);
  }

  /* ---------- 問題生成 ---------- */
  function buildQuestion(entry, type) {
    if (type === "reading") {
      return {
        type,
        kanji: entry.kanji,
        prompt: "次のことばの読み方を、ひらがなで書きましょう。",
        display: entry.word,
        answer: entry.reading,
        _meta: { qtype: type, entry },
      };
    }
    if (type === "writing") {
      return {
        type,
        kanji: entry.kanji,
        prompt: "次のひらがなを、漢字を使って書きましょう。",
        display: entry.reading,
        answer: entry.word,
        _meta: { qtype: type, entry },
      };
    }
    // strokes
    const correct = entry.strokes;
    const offsets = shuffle([-3, -2, -1, 1, 2, 3, 4]).slice(0, 3);
    const choiceSet = new Set([correct]);
    for (const off of offsets) {
      let v = correct + off;
      if (v < 1) v = correct + Math.abs(off) + 1;
      choiceSet.add(v);
      if (choiceSet.size >= 4) break;
    }
    while (choiceSet.size < 4) choiceSet.add(correct + choiceSet.size + 1);
    const choices = shuffle(Array.from(choiceSet)).slice(0, 4);
    return {
      type,
      kanji: entry.kanji,
      prompt: "次の漢字の総画数(すべての画数)を答えましょう。",
      display: entry.kanji,
      answer: correct,
      choices,
      _meta: { qtype: type, entry },
    };
  }

  function buildChoiceQuestion(entry) {
    const choices = shuffle([entry.correct, entry.wrong]);
    return {
      type: "choice",
      kanji: entry.correct,
      prompt: "次のひらがなを漢字で書くと、どちらが正しいですか。正しいほうを選びましょう。",
      display: entry.reading,
      answer: entry.correct,
      choices,
      _meta: { qtype: "choice", entry },
    };
  }

  function buildPairQuestion(entry, direction) {
    const dir = direction || (Math.random() < 0.5 ? "a2b" : "b2a");
    const src = dir === "a2b" ? entry.a : entry.b;
    const tgt = dir === "a2b" ? entry.b : entry.a;
    return {
      type: "pair",
      kanji: tgt.kanji,
      prompt: "次のことばと対になることばを、漢字で書きましょう。",
      display: src.kanji,
      hint: tgt.reading,
      answer: tgt.kanji,
      _meta: { qtype: "pair", entry, direction: dir },
    };
  }

  function rebuildQuestion(meta) {
    if (!meta) return null;
    if (meta.qtype === "choice") return buildChoiceQuestion(meta.entry);
    if (meta.qtype === "pair") return buildPairQuestion(meta.entry, meta.direction);
    return buildQuestion(meta.entry, meta.qtype);
  }

  function buildQueue(mode, count) {
    let candidates;
    if (mode === "mix") {
      candidates = [];
      KANJI_DATA.forEach((entry) => {
        candidates.push(buildQuestion(entry, "reading"));
        candidates.push(buildQuestion(entry, "writing"));
        candidates.push(buildQuestion(entry, "strokes"));
      });
      CHOICE_DATA.forEach((entry) => candidates.push(buildChoiceQuestion(entry)));
      PAIR_DATA.forEach((entry) => {
        candidates.push(buildPairQuestion(entry, "a2b"));
        candidates.push(buildPairQuestion(entry, "b2a"));
      });
    } else if (mode === "choice") {
      candidates = CHOICE_DATA.map((e) => buildChoiceQuestion(e));
    } else if (mode === "pair") {
      candidates = PAIR_DATA.flatMap((e) => [buildPairQuestion(e, "a2b"), buildPairQuestion(e, "b2a")]);
    } else {
      candidates = KANJI_DATA.map((e) => buildQuestion(e, mode));
    }
    candidates = shuffle(candidates);
    const n = count === "all" ? candidates.length : Math.min(count, candidates.length);
    return candidates.slice(0, n);
  }

  /* ---------- 画面切り替え ---------- */
  const screens = {
    home: document.getElementById("screen-home"),
    quiz: document.getElementById("screen-quiz"),
    result: document.getElementById("screen-result"),
  };

  function showScreen(name) {
    Object.entries(screens).forEach(([key, el]) => {
      el.hidden = key !== name;
    });
    window.scrollTo(0, 0);
  }

  /* ---------- ホーム画面 ---------- */
  const modeGrid = document.getElementById("mode-grid");
  const countGrid = document.getElementById("count-grid");

  modeGrid.addEventListener("click", (e) => {
    const btn = e.target.closest(".mode-btn");
    if (!btn) return;
    selectedMode = btn.dataset.mode;
    [...modeGrid.children].forEach((b) => b.classList.toggle("selected", b === btn));
  });

  countGrid.addEventListener("click", (e) => {
    const btn = e.target.closest(".count-btn");
    if (!btn) return;
    selectedCount = btn.dataset.count === "all" ? "all" : Number(btn.dataset.count);
    [...countGrid.children].forEach((b) => b.classList.toggle("selected", b === btn));
  });
  modeGrid.querySelector('[data-mode="reading"]').classList.add("selected");

  document.getElementById("start-btn").addEventListener("click", () => {
    startSession(selectedMode, selectedCount);
  });

  document.getElementById("weak-review-btn").addEventListener("click", () => {
    const stats = loadStats();
    const metas = Object.keys(stats.weak).map(metaFromWeakKey).filter(Boolean);
    if (metas.length === 0) return;
    startReviewFromMetas(metas);
  });

  document.getElementById("reset-stats-btn").addEventListener("click", () => {
    if (confirm("これまでの記録をすべてリセットしますか？")) {
      localStorage.removeItem(STORAGE_KEY);
      renderHome();
    }
  });

  function renderHome() {
    const stats = loadStats();

    // 成績
    const statsBody = document.getElementById("stats-body");
    if (stats.totalAnswered === 0) {
      statsBody.innerHTML = '<p class="sub small">まだ記録がありません。問題を解いてみよう！</p>';
    } else {
      const pct = Math.round((stats.totalCorrect / stats.totalAnswered) * 100);
      let html = `<div class="stat-row"><span>全体の正答率</span><b>${pct}% (${stats.totalCorrect}/${stats.totalAnswered})</b></div>`;
      for (const cat of Object.keys(MODE_LABEL)) {
        const c = stats.byCategory[cat] || { correct: 0, total: 0 };
        if (c.total === 0) continue;
        const p = Math.round((c.correct / c.total) * 100);
        html += `<div class="stat-row"><span>${MODE_LABEL[cat]}</span><b>${p}% (${c.correct}/${c.total})</b></div>`;
      }
      statsBody.innerHTML = html;
    }

    // にがて
    const weakCard = document.getElementById("weak-card");
    const weakList = document.getElementById("weak-list");
    const weakEntries = Object.entries(stats.weak).sort((a, b) => b[1].wrongCount - a[1].wrongCount);
    if (weakEntries.length === 0) {
      weakCard.hidden = true;
    } else {
      weakCard.hidden = false;
      weakList.innerHTML = weakEntries
        .slice(0, 16)
        .map(([k]) => `<span class="weak-chip">${weakLabel(k)}</span>`)
        .join("");
    }
  }

  /* ---------- クイズ画面 ---------- */
  const questionType = document.getElementById("question-type");
  const questionPrompt = document.getElementById("question-prompt");
  const questionDisplay = document.getElementById("question-display");
  const answerForm = document.getElementById("answer-form");
  const answerInput = document.getElementById("answer-input");
  const choiceGrid = document.getElementById("choice-grid");
  const submitBtn = document.getElementById("submit-btn");
  const feedback = document.getElementById("feedback");
  const feedbackMark = document.getElementById("feedback-mark");
  const feedbackText = document.getElementById("feedback-text");
  const nextBtn = document.getElementById("next-btn");
  const progressFill = document.getElementById("progress-fill");
  const quizCounter = document.getElementById("quiz-counter");

  let chosenChoice = null;

  function startSession(mode, count) {
    const queue = buildQueue(mode, count);
    if (queue.length === 0) return;
    session = { queue, index: 0, correct: 0, wrong: [], mode, count, isReview: false };
    showScreen("quiz");
    renderQuestion();
  }

  function startReviewFromMetas(metas) {
    const queue = shuffle(metas.map(rebuildQuestion).filter(Boolean));
    if (queue.length === 0) return;
    session = { queue, index: 0, correct: 0, wrong: [], mode: "review", count: "all", isReview: true, metas };
    showScreen("quiz");
    renderQuestion();
  }

  function isChoiceLike(type) {
    return type === "strokes" || type === "choice";
  }

  function renderQuestion() {
    const q = session.queue[session.index];
    chosenChoice = null;
    questionType.textContent = MODE_LABEL[q.type];
    questionPrompt.textContent = q.prompt;
    if (q.type === "pair") {
      questionDisplay.innerHTML = `${q.display}<span class="furigana-hint">→（　${q.hint}　）</span>`;
    } else {
      questionDisplay.textContent = q.display;
    }

    feedback.hidden = true;
    answerForm.hidden = false;
    answerInput.value = "";

    if (isChoiceLike(q.type)) {
      answerInput.hidden = true;
      choiceGrid.hidden = false;
      const label = q.type === "strokes" ? (c) => `${c}画` : (c) => c;
      choiceGrid.innerHTML = q.choices
        .map((c) => `<button type="button" class="choice-btn" data-value="${c}">${label(c)}</button>`)
        .join("");
    } else {
      answerInput.hidden = false;
      choiceGrid.hidden = true;
      choiceGrid.innerHTML = "";
      setTimeout(() => answerInput.focus(), 50);
    }

    const total = session.queue.length;
    progressFill.style.width = `${(session.index / total) * 100}%`;
    quizCounter.textContent = `${session.index + 1} / ${total}`;
  }

  choiceGrid.addEventListener("click", (e) => {
    const btn = e.target.closest(".choice-btn");
    if (!btn) return;
    chosenChoice = btn.dataset.value;
    [...choiceGrid.children].forEach((b) => b.classList.toggle("chosen", b === btn));
  });

  answerForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const q = session.queue[session.index];
    let userAnswer, isCorrect;

    if (isChoiceLike(q.type)) {
      if (chosenChoice === null) return;
      if (q.type === "strokes") {
        userAnswer = Number(chosenChoice);
        isCorrect = userAnswer === q.answer;
      } else {
        userAnswer = chosenChoice;
        isCorrect = userAnswer === q.answer;
      }
    } else {
      userAnswer = answerInput.value;
      if (normalize(userAnswer) === "") return;
      isCorrect = normalize(userAnswer) === normalize(q.answer);
    }

    recordAnswer(q._meta, isCorrect);

    if (isCorrect) {
      session.correct++;
    } else {
      session.wrong.push({ q, userAnswer });
    }

    feedback.hidden = false;
    submitBtn.disabled = true;

    if (isCorrect) {
      feedbackMark.textContent = "○ せいかい！";
      feedbackMark.className = "feedback-mark correct";
      feedbackText.innerHTML = "";
    } else {
      feedbackMark.textContent = "✕ ざんねん";
      feedbackMark.className = "feedback-mark wrong";
      const ansText = q.type === "strokes" ? `${q.answer}画` : q.answer;
      feedbackText.innerHTML = `正しいこたえ: <b>${ansText}</b>`;
    }
  });

  nextBtn.addEventListener("click", () => {
    submitBtn.disabled = false;
    session.index++;
    if (session.index >= session.queue.length) {
      finishSession();
    } else {
      renderQuestion();
    }
  });

  document.getElementById("quit-btn").addEventListener("click", () => {
    if (confirm("練習をやめてホームにもどりますか？")) {
      renderHome();
      showScreen("home");
    }
  });

  /* ---------- 結果画面 ---------- */
  const resultScore = document.getElementById("result-score");
  const resultBarFill = document.getElementById("result-bar-fill");
  const resultComment = document.getElementById("result-comment");
  const wrongReview = document.getElementById("wrong-review");
  const retryWrongBtn = document.getElementById("retry-wrong-btn");

  function finishSession() {
    const total = session.queue.length;
    const pct = Math.round((session.correct / total) * 100);
    resultScore.textContent = `${session.correct} / ${total} 問正解`;
    resultBarFill.style.width = `${pct}%`;

    let comment;
    if (pct === 100) comment = "パーフェクト！すばらしい！";
    else if (pct >= 80) comment = "よくできました！";
    else if (pct >= 50) comment = "もう少し！がんばろう。";
    else comment = "復習して、またチャレンジしよう。";
    resultComment.textContent = comment;

    if (session.wrong.length > 0) {
      wrongReview.innerHTML =
        `<h2>まちがえた問題 (${session.wrong.length})</h2>` +
        session.wrong
          .map(({ q, userAnswer }) => {
            const ans = q.type === "strokes" ? `${q.answer}画` : q.answer;
            const mine = q.type === "strokes" ? `${userAnswer}画` : userAnswer || "(未回答)";
            return `<div class="wrong-item"><span class="w-q">${q.display}<br><span style="font-size:.8em">あなた: ${mine}</span></span><span class="w-a">${ans}</span></div>`;
          })
          .join("");
      retryWrongBtn.hidden = false;
    } else {
      wrongReview.innerHTML = "";
      retryWrongBtn.hidden = true;
    }

    renderHome();
    showScreen("result");
  }

  retryWrongBtn.addEventListener("click", () => {
    const seen = new Set();
    const metas = [];
    session.wrong.forEach(({ q }) => {
      const id = `${q._meta.qtype}:${weakKeyFor(q._meta)}`;
      if (!seen.has(id)) {
        seen.add(id);
        metas.push(q._meta);
      }
    });
    startReviewFromMetas(metas);
  });

  document.getElementById("retry-same-btn").addEventListener("click", () => {
    if (session.isReview) {
      startReviewFromMetas(session.metas);
    } else {
      startSession(session.mode, session.count);
    }
  });

  document.getElementById("home-btn").addEventListener("click", () => {
    renderHome();
    showScreen("home");
  });

  /* ---------- 初期化 ---------- */
  renderHome();
  showScreen("home");
})();
