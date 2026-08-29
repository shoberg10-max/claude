// 漢字検定 8級・9級 練習アプリ ロジック
(function () {
  "use strict";

  const STORAGE_KEY = "kanken89_stats_v1";
  const INPUT_MODE_KEY = "kanken89_write_input_mode";
  const HW_GUIDE_KEY = "kanken89_write_guide_on";

  const MODE_LABEL = {
    reading: "読み方",
    writing: "書き取り",
    strokes: "画数",
    choice: "漢字えらび",
    pair: "ことばのペア",
    radical: "部首なかま",
    strokeorder: "筆順",
  };

  // その漢字に筆順データがあり、かつ4択の distractor を確実に作れる(総画数4以上)か
  function hasStrokeOrderData(kanji) {
    const paths = STROKE_ORDER_DATA[kanji];
    return !!paths && paths.length >= 4;
  }

  // レベルごとのデータセット。「radical」は学年をまたぐ共通データなのでここには含めない。
  const LEVELS = {
    9: { label: "9級", sub: "小学2年生修了程度", kanjiData: GRADE2_DATA, choiceData: GRADE2_CHOICE_DATA, pairData: GRADE2_PAIR_DATA },
    8: { label: "8級", sub: "小学3年生修了程度", kanjiData: KANJI_DATA, choiceData: CHOICE_DATA, pairData: PAIR_DATA },
  };

  /* ---------- 状態 ---------- */
  let session = null; // { queue, index, correct, wrong: [], mode, count, level, isReview, metas? }
  let selectedMode = "reading";
  let selectedCount = 20;
  let selectedLevel = "8";

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

  // 画数クイズの4択(正解1つ+誤り3つ、すべて1以上の重複なし整数)を作る。
  function generateStrokeChoices(correct) {
    const pool = [];
    for (let off = -4; off <= 4; off++) {
      if (off === 0) continue;
      const v = correct + off;
      if (v >= 1 && v !== correct) pool.push(v);
    }
    const wrong = shuffle(pool).slice(0, 3);
    let extra = 1;
    while (wrong.length < 3) {
      const v = correct + extra;
      if (v !== correct && !wrong.includes(v)) wrong.push(v);
      extra++;
    }
    return [correct, ...wrong];
  }

  // 筆順クイズの4択(正解の画数位置1つ+誤り3つ、1〜総画数の範囲で重複なし)を作る。
  function generateStrokePositionChoices(correctPos, totalStrokes) {
    const pool = [];
    for (let i = 1; i <= totalStrokes; i++) if (i !== correctPos) pool.push(i);
    const wrong = shuffle(pool).slice(0, 3);
    return [correctPos, ...wrong];
  }

  // 指定した画(1始まり)だけ強調したSVGのHTML文字列を作る。
  function renderStrokeSvg(paths, highlightIndex) {
    const strokeEls = paths
      .map((d, i) => {
        const isHi = i + 1 === highlightIndex;
        const color = isHi ? "#e5484d" : "#c3cbd6";
        const width = isHi ? 6 : 3;
        return `<path d="${d}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linecap="round" stroke-linejoin="round"/>`;
      })
      .join("");
    return `<svg viewBox="0 0 109 109" class="stroke-svg">${strokeEls}</svg>`;
  }

  // 選択肢・正解表示の文字列フォーマット(問題タイプごとに単位が違う)
  function formatAnswerValue(type, value) {
    if (type === "strokes") return `${value}画`;
    if (type === "strokeorder") return `${value}画目`;
    return value;
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

  // KANJI_DATA/GRADE2_DATA由来(読み方/書き取り/画数)は漢字1字をキーにしてまとめる。
  // 学年で漢字がかぶることはないため、9級・8級を分けなくても衝突しない。
  function weakKeyFor(meta) {
    if (meta.qtype === "choice") return `choice:${meta.entry.correct}`;
    if (meta.qtype === "pair") return `pair:${meta.entry.a.kanji}|${meta.entry.b.kanji}`;
    if (meta.qtype === "radical") return `radical:${meta.shown}`;
    return meta.entry.kanji;
  }

  function weakLabel(key) {
    if (key.startsWith("choice:")) return key.slice(7);
    if (key.startsWith("pair:")) return key.slice(5).replace("|", "⇔");
    if (key.startsWith("radical:")) return key.slice(8);
    return key;
  }

  function findKanjiEntry(kanji) {
    return GRADE2_DATA.find((e) => e.kanji === kanji) || KANJI_DATA.find((e) => e.kanji === kanji);
  }
  function findChoiceEntry(correctWord) {
    return CHOICE_DATA.find((e) => e.correct === correctWord) || GRADE2_CHOICE_DATA.find((e) => e.correct === correctWord);
  }
  function findPairEntry(ak, bk) {
    return (
      PAIR_DATA.find((e) => e.a.kanji === ak && e.b.kanji === bk) ||
      GRADE2_PAIR_DATA.find((e) => e.a.kanji === ak && e.b.kanji === bk)
    );
  }
  function findRadicalGroup(kanji) {
    return RADICAL_GROUPS.find((g) => g.kanji.includes(kanji));
  }

  function metaFromWeakKey(key) {
    if (key.startsWith("choice:")) {
      const entry = findChoiceEntry(key.slice(7));
      return entry ? { qtype: "choice", entry } : null;
    }
    if (key.startsWith("pair:")) {
      const [ak, bk] = key.slice(5).split("|");
      const entry = findPairEntry(ak, bk);
      return entry ? { qtype: "pair", entry, direction: Math.random() < 0.5 ? "a2b" : "b2a" } : null;
    }
    if (key.startsWith("radical:")) {
      const shown = key.slice(8);
      const group = findRadicalGroup(shown);
      return group ? { qtype: "radical", group, shown } : null;
    }
    const entry = findKanjiEntry(key);
    if (!entry) return null;
    const types = ["reading", "writing", "strokes"];
    if (hasStrokeOrderData(entry.kanji)) types.push("strokeorder");
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
    const choices = shuffle(generateStrokeChoices(correct));
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

  function buildRadicalQuestion(group, shownKanji) {
    const shown = shownKanji || group.kanji[Math.floor(Math.random() * group.kanji.length)];
    const sameGroupOthers = group.kanji.filter((k) => k !== shown);
    const correct = sameGroupOthers[Math.floor(Math.random() * sameGroupOthers.length)];
    const otherGroups = RADICAL_GROUPS.filter((g) => g !== group);
    const wrongPool = shuffle(otherGroups.flatMap((g) => g.kanji));
    const wrongChoices = [];
    for (const k of wrongPool) {
      if (wrongChoices.length >= 3) break;
      if (!wrongChoices.includes(k)) wrongChoices.push(k);
    }
    const choices = shuffle([correct, ...wrongChoices]);
    return {
      type: "radical",
      kanji: correct,
      prompt: `次の「${shown}」と同じ部首（${group.radical}）のなかまの漢字はどれでしょう。`,
      display: shown,
      answer: correct,
      choices,
      _meta: { qtype: "radical", group, shown },
    };
  }

  function buildStrokeOrderQuestion(entry) {
    const paths = STROKE_ORDER_DATA[entry.kanji];
    const total = paths.length;
    const correctPos = 1 + Math.floor(Math.random() * total);
    const choices = shuffle(generateStrokePositionChoices(correctPos, total));
    return {
      type: "strokeorder",
      kanji: entry.kanji,
      prompt: `次の「${entry.kanji}」の太い画は、何画目に書きますか。`,
      display: renderStrokeSvg(paths, correctPos),
      reviewLabel: entry.kanji,
      answer: correctPos,
      choices,
      _meta: { qtype: "strokeorder", entry },
    };
  }

  function rebuildQuestion(meta) {
    if (!meta) return null;
    if (meta.qtype === "choice") return buildChoiceQuestion(meta.entry);
    if (meta.qtype === "pair") return buildPairQuestion(meta.entry, meta.direction);
    if (meta.qtype === "radical") return buildRadicalQuestion(meta.group, meta.shown);
    if (meta.qtype === "strokeorder") return buildStrokeOrderQuestion(meta.entry);
    return buildQuestion(meta.entry, meta.qtype);
  }

  function allRadicalCandidates() {
    const list = [];
    RADICAL_GROUPS.forEach((group) => {
      group.kanji.forEach((shown) => list.push(buildRadicalQuestion(group, shown)));
    });
    return list;
  }

  function buildQueue(mode, count, level) {
    const L = LEVELS[level];
    let candidates;
    if (mode === "mix") {
      candidates = [];
      L.kanjiData.forEach((entry) => {
        candidates.push(buildQuestion(entry, "reading"));
        candidates.push(buildQuestion(entry, "writing"));
        candidates.push(buildQuestion(entry, "strokes"));
        if (hasStrokeOrderData(entry.kanji)) candidates.push(buildStrokeOrderQuestion(entry));
      });
      L.choiceData.forEach((entry) => candidates.push(buildChoiceQuestion(entry)));
      L.pairData.forEach((entry) => {
        candidates.push(buildPairQuestion(entry, "a2b"));
        candidates.push(buildPairQuestion(entry, "b2a"));
      });
      candidates.push(...allRadicalCandidates());
    } else if (mode === "choice") {
      candidates = L.choiceData.map((e) => buildChoiceQuestion(e));
    } else if (mode === "pair") {
      candidates = L.pairData.flatMap((e) => [buildPairQuestion(e, "a2b"), buildPairQuestion(e, "b2a")]);
    } else if (mode === "radical") {
      candidates = allRadicalCandidates();
    } else if (mode === "strokeorder") {
      candidates = L.kanjiData.filter((e) => hasStrokeOrderData(e.kanji)).map((e) => buildStrokeOrderQuestion(e));
    } else {
      candidates = L.kanjiData.map((e) => buildQuestion(e, mode));
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
  const levelGrid = document.getElementById("level-grid");
  const modeGrid = document.getElementById("mode-grid");
  const countGrid = document.getElementById("count-grid");

  levelGrid.addEventListener("click", (e) => {
    const btn = e.target.closest(".level-btn");
    if (!btn) return;
    selectedLevel = btn.dataset.level;
    [...levelGrid.children].forEach((b) => b.classList.toggle("selected", b === btn));
  });

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
  levelGrid.querySelector('[data-level="8"]').classList.add("selected");

  document.getElementById("start-btn").addEventListener("click", () => {
    startSession(selectedMode, selectedCount, selectedLevel);
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

  const inputModeToggle = document.getElementById("input-mode-toggle");
  const handwriteArea = document.getElementById("handwrite-area");
  const hwCanvas = document.getElementById("handwrite-canvas");
  const hwCtx = hwCanvas.getContext("2d");
  const hwGuideToggle = document.getElementById("hw-guide-toggle");
  const hwClearBtn = document.getElementById("hw-clear-btn");
  const hwRevealBtn = document.getElementById("hw-reveal-btn");
  const hwRevealArea = document.getElementById("hw-reveal-area");
  const hwAnswerWord = document.getElementById("hw-answer-word");
  const hwSelfcheck = document.getElementById("hw-selfcheck");
  const hwCorrectBtn = document.getElementById("hw-correct-btn");
  const hwWrongBtn = document.getElementById("hw-wrong-btn");

  const HW_CANVAS_SIZE = 280; // 論理座標(CSS px)での一辺のサイズ
  let writeInputMode = localStorage.getItem(INPUT_MODE_KEY) || "write";
  let hwGuideOn = localStorage.getItem(HW_GUIDE_KEY) === "1";
  let hwDrawing = false;

  function setupHandwriteCanvas() {
    const dpr = window.devicePixelRatio || 1;
    hwCanvas.width = HW_CANVAS_SIZE * dpr;
    hwCanvas.height = HW_CANVAS_SIZE * dpr;
    hwCtx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }
  setupHandwriteCanvas();

  function hwCanvasPos(e) {
    const rect = hwCanvas.getBoundingClientRect();
    return {
      x: ((e.clientX - rect.left) / rect.width) * HW_CANVAS_SIZE,
      y: ((e.clientY - rect.top) / rect.height) * HW_CANVAS_SIZE,
    };
  }

  function drawHwGuide() {
    if (!hwGuideOn || !session) return;
    const q = session.queue[session.index];
    if (!q || q.type !== "writing") return;
    const paths = STROKE_ORDER_DATA[q._meta.entry.kanji];
    if (!paths) return;
    const scale = HW_CANVAS_SIZE / 109;
    hwCtx.save();
    hwCtx.globalAlpha = 0.16;
    hwCtx.strokeStyle = "#4a5568";
    hwCtx.lineWidth = 4 / scale;
    hwCtx.lineCap = "round";
    hwCtx.lineJoin = "round";
    hwCtx.scale(scale, scale);
    paths.forEach((d) => hwCtx.stroke(new Path2D(d)));
    hwCtx.restore();
  }

  function hwClear() {
    hwCtx.clearRect(0, 0, HW_CANVAS_SIZE, HW_CANVAS_SIZE);
    drawHwGuide();
  }

  function hwStrokeStart(e) {
    hwDrawing = true;
    hwCanvas.setPointerCapture(e.pointerId);
    const p = hwCanvasPos(e);
    hwCtx.beginPath();
    hwCtx.moveTo(p.x, p.y);
    e.preventDefault();
  }
  function hwStrokeMove(e) {
    if (!hwDrawing) return;
    const p = hwCanvasPos(e);
    hwCtx.globalAlpha = 1;
    hwCtx.strokeStyle = "#26313d";
    hwCtx.lineWidth = 9;
    hwCtx.lineCap = "round";
    hwCtx.lineJoin = "round";
    hwCtx.lineTo(p.x, p.y);
    hwCtx.stroke();
    e.preventDefault();
  }
  function hwStrokeEnd() {
    hwDrawing = false;
  }
  hwCanvas.addEventListener("pointerdown", hwStrokeStart);
  hwCanvas.addEventListener("pointermove", hwStrokeMove);
  hwCanvas.addEventListener("pointerup", hwStrokeEnd);
  hwCanvas.addEventListener("pointercancel", hwStrokeEnd);
  hwCanvas.addEventListener("pointerleave", hwStrokeEnd);

  function updateHwGuideBtn() {
    hwGuideToggle.classList.toggle("active", hwGuideOn);
  }
  updateHwGuideBtn();

  hwGuideToggle.addEventListener("click", () => {
    hwGuideOn = !hwGuideOn;
    localStorage.setItem(HW_GUIDE_KEY, hwGuideOn ? "1" : "0");
    updateHwGuideBtn();
    hwClear();
  });

  hwClearBtn.addEventListener("click", () => hwClear());

  hwRevealBtn.addEventListener("click", () => {
    const q = session.queue[session.index];
    hwAnswerWord.textContent = q.answer;
    hwRevealArea.hidden = false;
    hwRevealBtn.hidden = true;
  });

  hwCorrectBtn.addEventListener("click", () => {
    const q = session.queue[session.index];
    submitAnswer(q, "（手書き）できた", true);
  });
  hwWrongBtn.addEventListener("click", () => {
    const q = session.queue[session.index];
    submitAnswer(q, "（手書き）まちがえた", false);
  });

  function updateInputModeBtns() {
    [...inputModeToggle.children].forEach((b) => {
      b.classList.toggle("selected", b.dataset.inputMode === writeInputMode);
    });
  }

  inputModeToggle.addEventListener("click", (e) => {
    const btn = e.target.closest(".input-mode-btn");
    if (!btn || !session) return;
    writeInputMode = btn.dataset.inputMode;
    localStorage.setItem(INPUT_MODE_KEY, writeInputMode);
    updateInputModeBtns();
    showWritingInput(session.queue[session.index]);
  });

  function showWritingInput(q) {
    hwRevealArea.hidden = true;
    hwRevealBtn.hidden = false;
    if (writeInputMode === "write") {
      answerForm.hidden = true;
      handwriteArea.hidden = false;
      hwClear();
    } else {
      answerForm.hidden = false;
      handwriteArea.hidden = true;
      answerInput.hidden = false;
      choiceGrid.hidden = true;
      choiceGrid.innerHTML = "";
      setTimeout(() => answerInput.focus(), 50);
    }
  }

  let chosenChoice = null;

  function startSession(mode, count, level) {
    const queue = buildQueue(mode, count, level);
    if (queue.length === 0) return;
    session = { queue, index: 0, correct: 0, wrong: [], mode, count, level, isReview: false };
    showScreen("quiz");
    renderQuestion();
  }

  function startReviewFromMetas(metas) {
    const queue = shuffle(metas.map(rebuildQuestion).filter(Boolean));
    if (queue.length === 0) return;
    session = { queue, index: 0, correct: 0, wrong: [], mode: "review", count: "all", level: null, isReview: true, metas };
    showScreen("quiz");
    renderQuestion();
  }

  function isChoiceLike(type) {
    return type === "strokes" || type === "choice" || type === "radical" || type === "strokeorder";
  }

  function renderQuestion() {
    const q = session.queue[session.index];
    chosenChoice = null;
    questionType.textContent = MODE_LABEL[q.type];
    questionPrompt.textContent = q.prompt;
    if (q.type === "pair") {
      questionDisplay.innerHTML = `${q.display}<span class="furigana-hint">→（　${q.hint}　）</span>`;
    } else if (q.type === "strokeorder") {
      questionDisplay.innerHTML = q.display;
    } else {
      questionDisplay.textContent = q.display;
    }

    feedback.hidden = true;
    answerForm.hidden = false;
    answerInput.value = "";
    inputModeToggle.hidden = true;
    handwriteArea.hidden = true;

    if (isChoiceLike(q.type)) {
      answerInput.hidden = true;
      choiceGrid.hidden = false;
      choiceGrid.innerHTML = q.choices
        .map((c) => `<button type="button" class="choice-btn" data-value="${c}">${formatAnswerValue(q.type, c)}</button>`)
        .join("");
    } else if (q.type === "writing") {
      inputModeToggle.hidden = false;
      updateInputModeBtns();
      showWritingInput(q);
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

  function submitAnswer(q, userAnswer, isCorrect) {
    recordAnswer(q._meta, isCorrect);

    if (isCorrect) {
      session.correct++;
    } else {
      session.wrong.push({ q, userAnswer });
    }

    inputModeToggle.hidden = true;
    handwriteArea.hidden = true;
    feedback.hidden = false;
    submitBtn.disabled = true;

    if (isCorrect) {
      feedbackMark.textContent = "○ せいかい！";
      feedbackMark.className = "feedback-mark correct";
      feedbackText.innerHTML = "";
    } else {
      feedbackMark.textContent = "✕ ざんねん";
      feedbackMark.className = "feedback-mark wrong";
      const ansText = formatAnswerValue(q.type, q.answer);
      feedbackText.innerHTML = `正しいこたえ: <b>${ansText}</b>`;
    }
  }

  answerForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const q = session.queue[session.index];
    let userAnswer, isCorrect;

    if (isChoiceLike(q.type)) {
      if (chosenChoice === null) return;
      if (q.type === "strokes" || q.type === "strokeorder") {
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

    submitAnswer(q, userAnswer, isCorrect);
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
            const ans = formatAnswerValue(q.type, q.answer);
            const mine = userAnswer === "" || userAnswer == null ? "(未回答)" : formatAnswerValue(q.type, userAnswer);
            const label = q.reviewLabel || q.display;
            return `<div class="wrong-item"><span class="w-q">${label}<br><span style="font-size:.8em">あなた: ${mine}</span></span><span class="w-a">${ans}</span></div>`;
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
      startSession(session.mode, session.count, session.level);
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
