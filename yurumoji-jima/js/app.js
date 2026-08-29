// ゆるもじ島 - 画面制御・アプリ本体
let state = null;
let session = null; // {mode:'mission'|'kentei', questions, index, correctCount, xpBefore, coinsBefore, chestEarned}
let currentZukanLevel = 9;

function $(sel) { return document.querySelector(sel); }
function $all(sel) { return Array.from(document.querySelectorAll(sel)); }

function showScreen(id) {
  $all(".screen").forEach(s => s.classList.remove("active"));
  document.getElementById(id).classList.add("active");
}

function renderStatusBar() {
  const bar = document.getElementById("statusbar");
  const info = xpProgress(state.xp);
  const stage = characterStage(info.level);
  document.getElementById("stat-level").textContent = `Lv.${info.level} ${stage.emoji}`;
  document.getElementById("stat-coins").textContent = state.coins;
  document.getElementById("stat-streak").textContent = `${state.streak}日`;
  bar.hidden = false;
}

/* ===== 島MAP ===== */
function renderMap() {
  const stage = islandStage(state);
  document.getElementById("island-label").textContent = stage.label;
  const grid = document.getElementById("island-grid");
  grid.innerHTML = "";
  stage.tiles.forEach(t => {
    const d = document.createElement("div");
    d.textContent = t;
    grid.appendChild(d);
  });
  const tg = tamagoStage(state);
  document.getElementById("tamago-emoji").textContent = tg.emoji;
  document.getElementById("tamago-label").textContent = tg.label;
  renderStatusBar();
}

/* ===== ミッション選択 ===== */
function renderMissionSelect() {
  document.getElementById("mission-progress-9").textContent =
    `${masteredCount(state, 9)} / ${KANJI_BY_LEVEL[9].length} おぼえた`;
  document.getElementById("mission-progress-8").textContent =
    `${masteredCount(state, 8)} / ${KANJI_BY_LEVEL[8].length} おぼえた`;
}

function startMission(level) {
  const entries = pickMissionKanji(state, level, 3);
  session = {
    mode: "mission",
    entries,
    questions: generateMissionQuestions(entries),
    index: 0,
    correctCount: 0,
    xpBefore: state.xp,
    coinsBefore: state.coins,
  };
  showScreen("screen-quiz");
  renderQuizProgress();
  renderCurrentQuestion();
}

/* ===== 漢検チャレンジ ===== */
function openKenteiIntro() {
  const unlocked = kenteiUnlocked(state);
  const text = document.getElementById("kentei-intro-text");
  const btn = document.getElementById("btn-start-kentei");
  if (unlocked) {
    text.textContent = "漢検チャレンジに 挑戦できます！";
    btn.hidden = false;
  } else {
    const need = 10 - masteredCount(state);
    text.textContent = `あと ${need}こ かんじを おぼえると ちょうせんできるよ！`;
    btn.hidden = true;
  }
  showScreen("screen-kentei-intro");
}

function startKentei() {
  session = {
    mode: "kentei",
    questions: generateKenteiQuestions(state, 10),
    index: 0,
    correctCount: 0,
    xpBefore: state.xp,
    coinsBefore: state.coins,
  };
  showScreen("screen-quiz");
  renderQuizProgress();
  renderCurrentQuestion();
}

/* ===== クイズ共通 ===== */
function renderQuizProgress() {
  const wrap = document.getElementById("quiz-progress");
  wrap.innerHTML = "";
  session.questions.forEach((q, i) => {
    const dot = document.createElement("span");
    if (i < session.index) dot.classList.add("done");
    else if (i === session.index) dot.classList.add("current");
    wrap.appendChild(dot);
  });
}

function renderCurrentQuestion() {
  const q = session.questions[session.index];
  const card = document.getElementById("quiz-card");
  card.innerHTML = "";

  if (q.type === "trace") {
    card.innerHTML = `
      <div class="quiz-question-label">なぞって かいてみよう！</div>
      <div class="trace-canvas-wrap">
        <div class="trace-frame" style="position:relative;width:200px;height:200px;">
          <div style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:140px;color:rgba(0,0,0,0.18);">${q.entry.k}</div>
          <canvas id="trace-canvas" width="200" height="200" style="position:absolute;inset:0;"></canvas>
        </div>
      </div>
      <p class="feedback-line">ゆびや マウスで なぞってみよう</p>
      <button class="btn btn-primary btn-large" id="btn-trace-done">できた！つぎへ</button>
    `;
    initTraceCanvas();
    document.getElementById("btn-trace-done").addEventListener("click", nextQuestion);
    return;
  }

  let label = "";
  let bigDisplay = "";
  let choiceClass = "";
  if (q.type === "read") {
    label = "なんて よむ？";
    bigDisplay = `<div class="quiz-kanji-big">${q.entry.k}</div>`;
  } else if (q.type === "writeRecall") {
    label = `「${q.entry.meaning}」の かんじは どれ？`;
    bigDisplay = `<div class="quiz-question-label">よみ：${q.entry.reading}</div>`;
    choiceClass = "quiz-kanji-big";
  } else if (q.type === "word") {
    label = "なんて よむ？";
    bigDisplay = `<div class="quiz-kanji-big">${q.entry.word}</div>`;
  }

  const choicesHtml = q.choices
    .map(c => `<button class="choice-btn" data-value="${c}">${c}</button>`)
    .join("");

  card.innerHTML = `
    <div class="quiz-question-label">${label}</div>
    ${bigDisplay}
    <div class="quiz-choices">${choicesHtml}</div>
  `;
  if (choiceClass) {
    $all("#quiz-card .choice-btn").forEach(b => (b.style.fontSize = "28px"));
  }

  $all("#quiz-card .choice-btn").forEach(btn => {
    btn.addEventListener("click", () => handleChoice(btn, q));
  });
}

function categoryForType(type) {
  if (type === "read") return "read";
  if (type === "writeRecall") return "write";
  if (type === "word") return "word";
  return null;
}

function handleChoice(btn, q) {
  $all("#quiz-card .choice-btn").forEach(b => (b.disabled = true));
  const chosen = btn.dataset.value;
  const correct = chosen === q.answer;
  btn.classList.add(correct ? "correct" : "wrong");
  if (!correct) {
    $all("#quiz-card .choice-btn").forEach(b => {
      if (b.dataset.value === q.answer) b.classList.add("correct");
    });
  }
  const category = categoryForType(q.type);
  recordAnswer(state, q.entry.k, category, correct);
  if (correct) session.correctCount += 1;
  saveState(state);
  renderStatusBar();
  setTimeout(nextQuestion, 700);
}

function initTraceCanvas() {
  const canvas = document.getElementById("trace-canvas");
  const ctx = canvas.getContext("2d");
  ctx.strokeStyle = "#ff8a3d";
  ctx.lineWidth = 6;
  ctx.lineCap = "round";
  let drawing = false;

  function pos(e) {
    const rect = canvas.getBoundingClientRect();
    const p = e.touches ? e.touches[0] : e;
    return { x: p.clientX - rect.left, y: p.clientY - rect.top };
  }
  function start(e) {
    drawing = true;
    const p = pos(e);
    ctx.beginPath();
    ctx.moveTo(p.x, p.y);
    e.preventDefault();
  }
  function move(e) {
    if (!drawing) return;
    const p = pos(e);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    e.preventDefault();
  }
  function end() { drawing = false; }

  canvas.addEventListener("mousedown", start);
  canvas.addEventListener("mousemove", move);
  window.addEventListener("mouseup", end);
  canvas.addEventListener("touchstart", start, { passive: false });
  canvas.addEventListener("touchmove", move, { passive: false });
  canvas.addEventListener("touchend", end);
}

function nextQuestion() {
  session.index += 1;
  if (session.index >= session.questions.length) {
    if (session.mode === "mission") finishMission();
    else finishKentei();
    return;
  }
  renderQuizProgress();
  renderCurrentQuestion();
}

/* ===== ミッション結果 ===== */
let pendingChest = false;

function finishMission() {
  const total = session.questions.filter(q => q.type !== "trace").length;
  recordMissionResult(state, session.correctCount, total);
  const streakResult = updateStreak(state);
  pendingChest = checkChestEarned(state);
  const xpGained = state.xp - session.xpBefore;
  const coinsGained = state.coins - session.coinsBefore;
  const leveledUp = levelFromXp(session.xpBefore) < levelFromXp(state.xp);
  saveState(state);

  document.getElementById("result-score").textContent = `${session.correctCount} / ${total} せいかい`;
  document.getElementById("result-rewards").innerHTML =
    `<span>✨ +${xpGained}XP</span><span>🪙 +${coinsGained}</span>`;

  const levelupBox = document.getElementById("result-levelup");
  if (leveledUp) {
    const info = xpProgress(state.xp);
    const stage = characterStage(info.level);
    levelupBox.hidden = false;
    levelupBox.textContent = `✨ しんか！ Lv.${info.level} ${stage.name}になった！`;
  } else {
    levelupBox.hidden = true;
  }

  const streakBox = document.getElementById("result-streak");
  if (streakResult.changed) {
    let txt = `🔥 ${streakResult.streak}日れんぞく学習中！`;
    if (streakResult.milestone) txt += ` 🎉 ${streakResult.milestone}日たっせいボーナス！`;
    streakBox.textContent = txt;
  } else {
    streakBox.textContent = "";
  }

  showScreen("screen-result");
}

function finishKentei() {
  const total = session.questions.length;
  const percent = Math.round((session.correctCount / total) * 100);
  const band = kenteiBand(percent);
  saveState(state);

  document.getElementById("kentei-result-emoji").textContent = band.emoji;
  document.getElementById("kentei-result-label").textContent = band.label;
  document.getElementById("kentei-result-score").textContent =
    `${session.correctCount} / ${total} 問正解（${percent}%）`;

  const hint = document.getElementById("kentei-result-hint");
  if (band.tone === "yellow") {
    const needCorrect = Math.ceil(total * 0.9) - session.correctCount;
    hint.textContent = `あと ${Math.max(needCorrect, 1)}問 正解できると 合格圏だよ！`;
  } else if (band.tone === "green") {
    hint.textContent = "すごい！このちょうしで がんばろう！";
  } else {
    hint.textContent = "もういちど 冒険して、にがてな かんじを れんしゅうしよう！";
  }
  showScreen("screen-kentei-result");
}

/* ===== 宝箱 ===== */
function openTreasureScreen() {
  document.getElementById("treasure-reveal").hidden = true;
  document.getElementById("btn-open-treasure").hidden = false;
  document.getElementById("btn-treasure-done").hidden = true;
  document.getElementById("treasure-title").textContent = "たからばこ発見！";
  showScreen("screen-treasure");
}

function revealTreasure() {
  const { item, isNew } = drawGacha(state);
  saveState(state);
  document.getElementById("treasure-item-emoji").textContent = item.emoji;
  document.getElementById("treasure-item-name").textContent = isNew
    ? `NEW！「${item.name}」を てにいれた！`
    : `「${item.name}」（もっていたので +10コイン）`;
  document.getElementById("treasure-reveal").hidden = false;
  document.getElementById("btn-open-treasure").hidden = true;
  document.getElementById("btn-treasure-done").hidden = false;
  renderStatusBar();
}

/* ===== かんじ図鑑 ===== */
function renderZukan(level) {
  currentZukanLevel = level;
  $all(".zukan-tab").forEach(t => t.classList.toggle("active", Number(t.dataset.level) === level));
  const grid = document.getElementById("zukan-grid");
  grid.innerHTML = "";
  document.getElementById("zukan-detail").hidden = true;
  KANJI_BY_LEVEL[level].forEach(entry => {
    const ks = state.kanjiState[entry.k] || { read: 0, write: 0, word: 0 };
    const attempted = ks.read + ks.write + ks.word > 0;
    const mastered = isMastered(state, entry.k);
    const tile = document.createElement("div");
    tile.className = "zukan-item " + (mastered ? "mastered" : attempted ? "" : "locked");
    tile.textContent = attempted ? entry.k : "？";
    tile.addEventListener("click", () => showZukanDetail(entry, attempted));
    grid.appendChild(tile);
  });
}

function showZukanDetail(entry, attempted) {
  const box = document.getElementById("zukan-detail");
  box.hidden = false;
  if (!attempted) {
    box.innerHTML = `<p>まだ であっていない かんじだよ。冒険に行って みつけよう！</p>`;
    return;
  }
  const ks = state.kanjiState[entry.k];
  box.innerHTML = `
    <div class="quiz-kanji-big">${entry.k}</div>
    <p>音：${entry.on.join("、") || "－"} ／ 訓：${entry.kun.join("、") || "－"}</p>
    <p>いみ：${entry.meaning}</p>
    <p>じゅくご：${entry.word}（${entry.wordReading}）</p>
    <p>よみ${ks.read}回・かき${ks.write}回・いみ${ks.word}回 せいかい</p>
  `;
}

/* ===== 保護者画面 ===== */
function renderParent() {
  const days = weeklyStudyDays(state);
  const mastered9 = masteredCount(state, 9);
  const mastered8 = masteredCount(state, 8);
  const total9 = KANJI_BY_LEVEL[9].length;
  const total8 = KANJI_BY_LEVEL[8].length;
  const weak = weakKanji(state, 5);
  const overallPercent = Math.round(
    ((mastered9 + mastered8) / (total9 + total8)) * 100
  );
  const stars = Math.max(1, Math.min(5, Math.round(overallPercent / 20)));

  const card = document.getElementById("parent-card");
  card.innerHTML = `
    <div class="parent-row"><b>今週の学習日数</b><span>${days} / 7日</span></div>
    <div class="parent-row"><b>連続学習</b><span>🔥 ${state.streak}日</span></div>
    <div class="parent-row"><b>学習回数（累計）</b><span>${state.history.length}回</span></div>
    <div class="parent-row"><b>9級 おぼえた漢字</b>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round((mastered9/total9)*100)}%"></div></div>
    </div>
    <div class="parent-row"><b>8級 おぼえた漢字</b>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round((mastered8/total8)*100)}%"></div></div>
    </div>
    <div>
      <b>苦手な漢字</b>
      <div class="parent-weak-list">${
        weak.length ? weak.map(e => `<span>${e.k}</span>`).join("") : "<span>（まだありません）</span>"
      }</div>
    </div>
    <div class="parent-row"><b>合格の目安</b><span>${"★".repeat(stars)}${"☆".repeat(5 - stars)}</span></div>
  `;
}

/* ===== 初期化・イベント配線 ===== */
function init() {
  state = loadState();

  document.getElementById("btn-start").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });
  document.getElementById("btn-reset").addEventListener("click", () => {
    if (confirm("さいしょから やりなおしますか？（きろくが きえます）")) {
      state = resetState();
      renderMap();
      showScreen("screen-map");
    }
  });

  document.getElementById("btn-parent").addEventListener("click", () => {
    renderParent();
    showScreen("screen-parent");
  });
  document.getElementById("btn-back-from-parent").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });

  document.getElementById("btn-go-mission").addEventListener("click", () => {
    renderMissionSelect();
    showScreen("screen-mission-select");
  });
  document.getElementById("btn-back-from-mission-select").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });
  $all(".mission-choice").forEach(btn => {
    btn.addEventListener("click", () => startMission(btn.dataset.level));
  });

  document.getElementById("btn-result-next").addEventListener("click", () => {
    if (pendingChest) {
      pendingChest = false;
      openTreasureScreen();
    } else {
      renderMap();
      showScreen("screen-map");
    }
  });
  document.getElementById("btn-open-treasure").addEventListener("click", revealTreasure);
  document.getElementById("btn-treasure-done").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });

  document.getElementById("btn-go-kentei").addEventListener("click", openKenteiIntro);
  document.getElementById("btn-start-kentei").addEventListener("click", startKentei);
  document.getElementById("btn-back-from-kentei").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });
  document.getElementById("btn-kentei-done").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });

  document.getElementById("btn-go-zukan").addEventListener("click", () => {
    renderZukan(9);
    showScreen("screen-zukan");
  });
  document.getElementById("btn-back-from-zukan").addEventListener("click", () => {
    renderMap();
    showScreen("screen-map");
  });
  $all(".zukan-tab").forEach(tab => {
    tab.addEventListener("click", () => renderZukan(Number(tab.dataset.level)));
  });

  showScreen("screen-opening");
}

document.addEventListener("DOMContentLoaded", init);
