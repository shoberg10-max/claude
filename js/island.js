// ゆるもじ島の漢字たんけん隊：メインゲームループ
const Island = (() => {
  const el = UI.el;
  const screenEl = UI.screenEl;

  const LEVELS = {
    k9: { key: 'k9', label: '9級', grade: '1年生の漢字', total: 80 },
    k8: { key: 'k8', label: '8級', grade: '2年生の漢字', total: 160 }
  };
  const TERRAIN = ['🌳', '🌸', '🌴', '⛰️', '🌊', '🏠', '🌻', '🏯', '🌵', '🐚', '🌼', '⭐'];

  function currentLevel() { return Storage.getIsland().level; }
  function dataset(level) { return KanjiGame.datasetFor(level); }
  function masteredCount(level) { return Storage.getIslandMasteredCount(level, dataset(level)); }
  function totalMastered() { return masteredCount('k9') + masteredCount('k8'); }

  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  // ---------- 共通：ステータスバー ----------
  function renderStatBar() {
    const info = Storage.getCharLevelInfo();
    const island = Storage.getIsland();
    UI.setStatBar(el('div', { className: 'statBarInner' }, [
      el('span', { className: 'statChip', text: '🪙' + island.coins }),
      el('span', { className: 'statChip', text: '🎫' + island.tickets }),
      el('span', { className: 'statChip', text: '🔥' + island.streak }),
      el('span', { className: 'statChip', text: 'Lv.' + info.level })
    ]));
  }

  // ---------- 共通：ボトムナビ ----------
  function bottomNav(active) {
    const items = [
      { id: 'home', emoji: '🏠', label: 'ホーム', onClick: renderHome },
      { id: 'mission', emoji: '📖', label: 'ミッション', onClick: renderMissionIntro },
      { id: 'treasure', emoji: '🎁', label: 'たからばこ', onClick: renderTreasure },
      { id: 'char', emoji: '👤', label: 'キャラ', onClick: renderCharacter }
    ];
    return el('nav', { className: 'bottomNav' }, items.map(it =>
      el('button', {
        className: 'navBtn' + (it.id === active ? ' navBtnActive' : ''),
        onClick: it.onClick
      }, [
        el('div', { className: 'navBtnEmoji', text: it.emoji }),
        el('div', { className: 'navBtnLabel', text: it.label })
      ])
    ));
  }

  function levelTabs(onSwitch) {
    const level = currentLevel();
    const row = el('div', { className: 'levelTabs' });
    Object.values(LEVELS).forEach(lv => {
      row.appendChild(el('button', {
        className: 'levelTab' + (lv.key === level ? ' levelTabActive' : ''),
        onClick: () => { Storage.setIslandLevel(lv.key); onSwitch(); }
      }, [
        el('div', { className: 'levelTabTitle', text: lv.label + 'の島' }),
        el('div', { className: 'levelTabSub', text: lv.grade })
      ]));
    });
    return row;
  }

  // ---------- ホーム（島マップ） ----------
  function renderHome() {
    UI.setHomeHandler(renderHome);
    UI.setTitle('ゆるもじ島マップ');
    renderStatBar();
    UI.clearScreen();

    const level = currentLevel();
    const meta = LEVELS[level];
    const mastered = masteredCount(level);
    const ratio = mastered / meta.total;
    const revealedCount = Math.max(1, Math.round(ratio * TERRAIN.length));

    const mapGrid = el('div', { className: 'islandMapGrid' });
    TERRAIN.forEach((emoji, i) => {
      mapGrid.appendChild(el('div', {
        className: 'mapTile' + (i < revealedCount ? '' : ' mapTileFog'),
        text: i < revealedCount ? emoji : '☁️'
      }));
    });

    screenEl.appendChild(levelTabs(renderHome));
    screenEl.appendChild(mapGrid);
    screenEl.appendChild(el('div', { className: 'mapProgress', text: `おぼえた かんじ：${mastered} / ${meta.total}` }));

    const tiles = el('div', { className: 'menuGrid homeTilesGrid' }, [
      el('button', { className: 'menuCard', onClick: renderMissionIntro }, [
        el('div', { className: 'menuCardEmoji', text: '📖' }),
        el('div', { className: 'menuCardTitle', text: '今日の冒険' }),
        el('div', { className: 'menuCardScore', text: `ミッション ${Storage.getMissionCount(level) + 1}` })
      ]),
      el('button', { className: 'menuCard', onClick: renderZukan }, [
        el('div', { className: 'menuCardEmoji', text: '📚' }),
        el('div', { className: 'menuCardTitle', text: 'かんじ図鑑' }),
        el('div', { className: 'menuCardScore', text: `${mastered}/${meta.total} 字` })
      ]),
      el('button', { className: 'menuCard', onClick: renderChallengeIntro }, [
        el('div', { className: 'menuCardEmoji', text: '🏆' }),
        el('div', { className: 'menuCardTitle', text: '漢検チャレンジ' }),
        el('div', { className: 'menuCardScore', text: (() => { const b = Storage.getKankenBest(level); return b ? `ベスト ${Math.round(b.pct)}%` : 'ちょうせん！'; })() })
      ]),
      el('button', { className: 'menuCard', onClick: renderParentGate }, [
        el('div', { className: 'menuCardEmoji', text: '👪' }),
        el('div', { className: 'menuCardTitle', text: 'おうちの方へ' }),
        el('div', { className: 'menuCardScore', text: 'がくしゅう記録' })
      ])
    ]);
    screenEl.appendChild(tiles);

    screenEl.appendChild(el('button', {
      className: 'backBtn',
      text: '🧮 さんすう・3年生のドリルはこちら',
      onClick: () => App.renderHome()
    }));

    screenEl.appendChild(bottomNav('home'));
  }

  // ---------- ミッション導入 ----------
  function renderMissionIntro() {
    UI.setHomeHandler(renderHome);
    UI.setTitle('今日のミッション');
    renderStatBar();
    UI.clearScreen();

    const level = currentLevel();
    const missionNo = Storage.getMissionCount(level) + 1;

    const card = el('div', { className: 'quizCard missionIntroCard' }, [
      el('div', { className: 'bigCardEmoji', text: '🌱' }),
      el('div', { className: 'questionText', text: `ミッション ${missionNo}` }),
      el('div', { className: 'heartsRow', text: '❤️❤️❤️' }),
      el('div', { className: 'quizInstruction', text: 'かんじを 8もん とこう！さいごに 書く れんしゅうもあるよ。' }),
      el('button', { className: 'primaryBtn', text: 'スタート！', onClick: startMission })
    ]);
    screenEl.appendChild(card);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
    screenEl.appendChild(bottomNav('mission'));
  }

  function startMission() {
    const level = currentLevel();
    const questions = KanjiGame.generateMission(level, 8);
    runMission(level, questions, 0, 0);
  }

  function runMission(level, questions, index, correctCount) {
    UI.setTitle('ミッション');
    renderStatBar();
    UI.clearScreen();

    if (index >= questions.length) {
      startWritingBonus(level, questions, correctCount);
      return;
    }

    const q = questions[index];
    const progress = el('div', { className: 'progress', text: `もんだい ${index + 1} / ${questions.length}　　せいかい ${correctCount}こ` });
    const card = el('div', { className: 'quizCard' });
    card.appendChild(el('div', { className: 'quizInstruction', text: q.instruction }));
    card.appendChild(el('div', { className: q.mode === 'reading' ? 'kanjiPrompt' : 'readingPrompt', text: q.prompt }));

    const feedback = el('div', { className: 'feedback' });
    let answered = false;

    function goNext(wasCorrect) {
      Storage.recordIslandKanjiResult(level, q.kanji, wasCorrect);
      runMission(level, questions, index + 1, correctCount + (wasCorrect ? 1 : 0));
    }

    function showFeedback(isCorrect, correctText) {
      feedback.innerHTML = '';
      feedback.appendChild(el('div', {
        className: isCorrect ? 'feedbackOk' : 'feedbackNg',
        text: isCorrect ? '⭕ せいかい！ もじまる「やったね！」' : `✕ ざんねん… こたえは ${correctText}`
      }));
      const nextBtn = el('button', { className: 'nextBtn', text: index + 1 >= questions.length ? 'つぎへ →' : 'つぎへ →' });
      nextBtn.addEventListener('click', () => goNext(isCorrect));
      feedback.appendChild(nextBtn);
    }

    const choiceRow = el('div', { className: 'choiceRow' });
    q.choices.forEach(choice => {
      const btn = el('button', {
        className: q.mode === 'find' ? 'choiceBtn choiceKanji' : 'choiceBtn',
        text: choice,
        onClick: () => {
          if (answered) return;
          answered = true;
          [...choiceRow.children].forEach(b => b.disabled = true);
          const isCorrect = choice === q.answer;
          btn.classList.add(isCorrect ? 'choiceCorrect' : 'choiceWrong');
          showFeedback(isCorrect, q.answer);
        }
      });
      choiceRow.appendChild(btn);
    });
    card.appendChild(choiceRow);
    card.appendChild(feedback);

    screenEl.appendChild(progress);
    screenEl.appendChild(card);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ（きろくはのこりません）', onClick: renderHome }));
  }

  // ---------- 書きれんしゅう（ボーナス） ----------
  function startWritingBonus(level, questions, correctCount) {
    const uniqueKanji = [];
    const seen = new Set();
    questions.forEach(q => {
      if (!seen.has(q.kanji)) {
        seen.add(q.kanji);
        const entry = dataset(level).find(e => e.k === q.kanji);
        if (entry) uniqueKanji.push(entry);
      }
    });
    const picks = KanjiGame.pickWritingKanji(level, uniqueKanji, 2);
    renderWritingPractice(level, picks, 0, () => finishMission(level, questions.length, correctCount));
  }

  function renderWritingPractice(level, list, idx, onDone) {
    UI.setTitle('書いてみよう');
    renderStatBar();
    UI.clearScreen();

    if (idx >= list.length) { onDone(); return; }
    const entry = list[idx];
    const readingHint = [...entry.kun, ...entry.on][0] || '';

    const card = el('div', { className: 'quizCard' });
    card.appendChild(el('div', { className: 'quizInstruction', text: '書いてみよう！　なぞって れんしゅうしよう' }));

    const stage = el('div', { className: 'writingStage' });
    const guide = el('div', { className: 'writingGuide', text: entry.k });
    const canvas = document.createElement('canvas');
    canvas.className = 'writingCanvas';
    canvas.width = 260;
    canvas.height = 260;
    stage.appendChild(guide);
    stage.appendChild(canvas);
    card.appendChild(stage);

    const ctx = canvas.getContext('2d');
    ctx.lineWidth = 10;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#3aa0ff';
    let drawing = false;
    let last = null;

    function posFromEvent(e) {
      const rect = canvas.getBoundingClientRect();
      const scaleX = canvas.width / rect.width;
      const scaleY = canvas.height / rect.height;
      return { x: (e.clientX - rect.left) * scaleX, y: (e.clientY - rect.top) * scaleY };
    }
    canvas.addEventListener('pointerdown', e => {
      drawing = true;
      last = posFromEvent(e);
      canvas.setPointerCapture(e.pointerId);
    });
    canvas.addEventListener('pointermove', e => {
      if (!drawing) return;
      const p = posFromEvent(e);
      ctx.beginPath();
      ctx.moveTo(last.x, last.y);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      last = p;
    });
    ['pointerup', 'pointerleave', 'pointercancel'].forEach(ev =>
      canvas.addEventListener(ev, () => { drawing = false; })
    );

    const hintRow = el('div', { className: 'writingHint', text: `よみかた：${readingHint}` });
    card.appendChild(hintRow);

    const btnRow = el('div', { className: 'resultBtnRow' }, [
      el('button', {
        className: 'secondaryBtn',
        text: 'けす',
        onClick: () => ctx.clearRect(0, 0, canvas.width, canvas.height)
      }),
      el('button', {
        className: 'primaryBtn',
        text: idx + 1 >= list.length ? 'できた！つぎへ →' : 'できた！つぎへ →',
        onClick: () => renderWritingPractice(level, list, idx + 1, onDone)
      })
    ]);
    card.appendChild(btnRow);

    screenEl.appendChild(el('div', { className: 'progress', text: `かきとり ${idx + 1} / ${list.length}` }));
    screenEl.appendChild(card);
  }

  // ---------- ミッション結果・ごほうび ----------
  function finishMission(level, total, correctCount) {
    const coins = correctCount * 10;
    const xp = correctCount * 15;
    const earnedTicket = correctCount >= Math.ceil(total * 0.75) ? 1 : 0;

    Storage.addCoins(coins);
    const xpResult = Storage.addCharXp(xp);
    if (earnedTicket) Storage.addTickets(earnedTicket);
    Storage.addDailyLog(total, correctCount);
    Storage.incMissionCount(level);

    const newlyUnlocked = [];
    const totalNow = totalMastered();
    COMPANIONS.forEach(c => {
      if (c.always) return;
      if (Storage.unlockCompanionIfEligible(c.id, c.need, totalNow)) newlyUnlocked.push(c);
    });

    const steps = [];
    steps.push(next => renderMissionResultStep(correctCount, total, coins, xp, next));
    if (xpResult.leveledUp) steps.push(next => renderLevelUpStep(xpResult.level, next));
    newlyUnlocked.forEach(c => steps.push(next => renderCompanionUnlockStep(c, next)));
    if (earnedTicket) steps.push(next => renderTicketOfferStep(next));

    playSequence(steps);
  }

  function playSequence(steps) {
    let i = 0;
    function next() {
      if (i >= steps.length) { renderHome(); return; }
      const step = steps[i++];
      step(next);
    }
    next();
  }

  function renderMissionResultStep(correctCount, total, coins, xp, next) {
    UI.setTitle('ミッションクリア！');
    renderStatBar();
    UI.clearScreen();
    const starCount = Math.max(1, Math.round((correctCount / total) * 3));
    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: '🌟'.repeat(starCount) }),
      el('div', { className: 'resultScore', text: 'すごい！よくできたね！' }),
      el('div', { className: 'resultScore', text: `${correctCount} / ${total} もん せいかい` }),
      el('div', { className: 'rewardRow' }, [
        el('span', { className: 'rewardChip', text: `🪙 コイン +${coins}` }),
        el('span', { className: 'rewardChip', text: `✨ けいけんち +${xp}` })
      ]),
      el('button', { className: 'primaryBtn', text: 'つぎへ →', onClick: next })
    ]));
  }

  function renderLevelUpStep(level, next) {
    const stage = charStageForLevel(level);
    UI.setTitle('レベルアップ！');
    renderStatBar();
    UI.clearScreen();
    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: stage.emoji }),
      el('div', { className: 'resultScore', text: `もじまるが Lv.${level} に なった！` }),
      el('div', { className: 'resultScore', text: stage.title }),
      el('button', { className: 'primaryBtn', text: 'つぎへ →', onClick: next })
    ]));
  }

  function renderCompanionUnlockStep(companion, next) {
    UI.setTitle('あたらしいなかま！');
    renderStatBar();
    UI.clearScreen();
    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: companion.emoji }),
      el('div', { className: 'resultScore', text: `${companion.name} が なかまになった！` }),
      el('div', { className: 'quizInstruction', text: companion.desc }),
      el('button', { className: 'primaryBtn', text: 'つぎへ →', onClick: next })
    ]));
  }

  function renderTicketOfferStep(next) {
    UI.setTitle('たからばこチケット GET！');
    renderStatBar();
    UI.clearScreen();
    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: '🎫' }),
      el('div', { className: 'resultScore', text: 'たからばこチケットを もらった！' }),
      el('div', { className: 'resultBtnRow' }, [
        el('button', { className: 'primaryBtn', text: '宝箱をあけてみる', onClick: () => renderTreasureRevealCard(next) }),
        el('button', { className: 'secondaryBtn', text: 'あとで', onClick: next })
      ])
    ]));
  }

  // ---------- たからばこ ----------
  function renderTreasure() {
    UI.setHomeHandler(renderHome);
    UI.setTitle('たからばこ');
    renderStatBar();
    UI.clearScreen();

    const island = Storage.getIsland();
    const ownedCount = Storage.getItems().length;
    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: '🎁' }),
      el('div', { className: 'resultScore', text: `チケット：🎫 ${island.tickets}まい` }),
      el('div', { className: 'quizInstruction', text: `あつめた アイテム：${ownedCount} / ${TREASURE_ITEMS.length}` }),
      el('button', {
        className: 'primaryBtn',
        text: '宝箱をあける',
        disabled: island.tickets <= 0,
        onClick: () => renderTreasureRevealCard(renderTreasure)
      })
    ]));
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
    screenEl.appendChild(bottomNav('treasure'));
  }

  function renderTreasureRevealCard(next) {
    UI.setTitle('たからばこ');
    renderStatBar();
    UI.clearScreen();

    if (!Storage.useTicket()) {
      screenEl.appendChild(el('div', { className: 'resultCard' }, [
        el('div', { className: 'resultScore', text: 'チケットが ないよ…' }),
        el('button', { className: 'primaryBtn', text: 'もどる', onClick: next })
      ]));
      return;
    }
    const item = rollTreasureItem();
    const isNew = Storage.addItem(item.id);
    if (!isNew) Storage.addCoins(20);

    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: item.emoji }),
      el('div', { className: 'resultScore', text: isNew ? `NEW！「${item.name}」を手に入れた！` : `「${item.name}」はもう持っていたので コイン+20！` }),
      el('div', { className: 'quizInstruction', text: 'レア度：' + (item.rarity === 'epic' ? '✨ エピック' : item.rarity === 'rare' ? '🌟 レア' : '⭐ ノーマル') }),
      el('button', { className: 'primaryBtn', text: 'つぎへ →', onClick: next })
    ]));
  }

  // ---------- キャラクター ----------
  function renderCharacter() {
    UI.setHomeHandler(renderHome);
    UI.setTitle('キャラクター');
    renderStatBar();
    UI.clearScreen();

    const info = Storage.getCharLevelInfo();
    const stage = charStageForLevel(info.level);
    const equipped = Storage.getEquippedItem();
    const equippedItem = TREASURE_ITEMS.find(i => i.id === equipped);
    const xpPct = Math.round((info.xpIntoLevel / info.xpForNext) * 100);

    screenEl.appendChild(el('div', { className: 'quizCard charScreenTop' }, [
      el('div', { className: 'bigCardEmoji', text: (equippedItem ? equippedItem.emoji + ' ' : '') + stage.emoji }),
      el('div', { className: 'questionText', text: `もじまる Lv.${info.level}` }),
      el('div', { className: 'quizInstruction', text: stage.title }),
      el('div', { className: 'xpBar' }, [el('div', { className: 'xpBarFill', attrs: { style: `width:${xpPct}%` } })]),
      el('div', { className: 'quizInstruction', text: `つぎのレベルまで ${info.xpForNext - info.xpIntoLevel} けいけんち` })
    ]));

    screenEl.appendChild(el('div', { className: 'sectionTitle', text: 'なかまたち' }));
    const unlockedIds = Storage.getCompanions();
    const companionGrid = el('div', { className: 'companionGrid' });
    COMPANIONS.forEach(c => {
      const unlocked = c.always || unlockedIds.includes(c.id);
      companionGrid.appendChild(el('div', { className: 'companionCard' + (unlocked ? '' : ' companionLocked') }, [
        el('div', { className: 'bigCardEmoji', text: unlocked ? c.emoji : '🔒' }),
        el('div', { className: 'menuCardTitle', text: unlocked ? c.name : '？？？' }),
        el('div', { className: 'menuCardScore', text: unlocked ? c.desc : c.hint })
      ]));
    });
    screenEl.appendChild(companionGrid);

    screenEl.appendChild(el('div', { className: 'sectionTitle', text: 'かぶりものコレクション' }));
    const ownedItems = Storage.getItems();
    const itemGrid = el('div', { className: 'itemGrid' });
    TREASURE_ITEMS.forEach(item => {
      const owned = ownedItems.includes(item.id);
      itemGrid.appendChild(el('button', {
        className: 'itemCell' + (owned ? ' itemOwned' : '') + (equipped === item.id ? ' itemEquipped' : ''),
        text: owned ? item.emoji : '❔',
        onClick: () => {
          if (!owned) return;
          Storage.equipItem(equipped === item.id ? null : item.id);
          renderCharacter();
        }
      }));
    });
    screenEl.appendChild(itemGrid);

    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
    screenEl.appendChild(bottomNav('char'));
  }

  // ---------- かんじ図鑑 ----------
  function renderZukan() {
    UI.setHomeHandler(renderHome);
    const level = currentLevel();
    UI.setTitle(`かんじ図鑑（${LEVELS[level].label}）`);
    renderStatBar();
    UI.clearScreen();

    screenEl.appendChild(levelTabs(renderZukan));

    const grid = el('div', { className: 'kanjiListGrid' });
    dataset(level).forEach(entry => {
      const lv = Storage.getIslandKanjiLevel(level, entry.k);
      const readings = [...entry.on.map(r => 'オ' + r), ...entry.kun.map(r => 'ク' + r)].join('　');
      grid.appendChild(el('div', { className: 'kanjiListItem' }, [
        el('div', { className: 'kanjiListChar' + (lv >= 3 ? ' mastered' : ''), text: entry.k }),
        el('div', { className: 'kanjiListReading', text: readings })
      ]));
    });
    screenEl.appendChild(el('div', { className: 'kanjiListLegend', text: 'オ＝音読み　ク＝訓読み　金色＝おぼえたかんじ' }));
    screenEl.appendChild(grid);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
    screenEl.appendChild(bottomNav('home'));
  }

  // ---------- 漢検チャレンジ ----------
  function renderChallengeIntro() {
    UI.setHomeHandler(renderHome);
    const level = currentLevel();
    UI.setTitle('漢検チャレンジ');
    renderStatBar();
    UI.clearScreen();

    const best = Storage.getKankenBest(level);
    screenEl.appendChild(el('div', { className: 'quizCard missionIntroCard' }, [
      el('div', { className: 'bigCardEmoji', text: '🏯' }),
      el('div', { className: 'questionText', text: `漢検${LEVELS[level].label} もぎテスト` }),
      el('div', { className: 'quizInstruction', text: '全20問！読み・かんじさがしミックスで実力チェック！' }),
      el('div', { className: 'menuCardScore', text: best ? `ベスト正答率：${Math.round(best.pct)}%` : 'はじめてのちょうせん！' }),
      el('button', { className: 'primaryBtn', text: 'チャレンジスタート！', onClick: startChallenge })
    ]));
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
  }

  function startChallenge() {
    const level = currentLevel();
    const questions = KanjiGame.generateChallenge(level, 20);
    runChallenge(level, questions, 0, { reading: { c: 0, t: 0 }, find: { c: 0, t: 0 } });
  }

  function runChallenge(level, questions, index, stats) {
    UI.setTitle('漢検チャレンジ');
    renderStatBar();
    UI.clearScreen();

    if (index >= questions.length) {
      finishChallenge(level, questions.length, stats);
      return;
    }

    const q = questions[index];
    screenEl.appendChild(el('div', { className: 'progress', text: `第 ${index + 1} 問 / ${questions.length}` }));
    const card = el('div', { className: 'quizCard' });
    card.appendChild(el('div', { className: 'quizInstruction', text: q.instruction }));
    card.appendChild(el('div', { className: q.mode === 'reading' ? 'kanjiPrompt' : 'readingPrompt', text: q.prompt }));

    const feedback = el('div', { className: 'feedback' });
    let answered = false;

    function goNext(wasCorrect) {
      Storage.recordIslandKanjiResult(level, q.kanji, wasCorrect);
      stats[q.mode].t += 1;
      if (wasCorrect) stats[q.mode].c += 1;
      runChallenge(level, questions, index + 1, stats);
    }

    function showFeedback(isCorrect, correctText) {
      feedback.innerHTML = '';
      feedback.appendChild(el('div', {
        className: isCorrect ? 'feedbackOk' : 'feedbackNg',
        text: isCorrect ? '⭕ せいかい！' : `✕ こたえは ${correctText}`
      }));
      const nextBtn = el('button', { className: 'nextBtn', text: index + 1 >= questions.length ? 'けっかを見る →' : 'つぎへ →' });
      nextBtn.addEventListener('click', () => goNext(isCorrect));
      feedback.appendChild(nextBtn);
    }

    const choiceRow = el('div', { className: 'choiceRow' });
    q.choices.forEach(choice => {
      const btn = el('button', {
        className: q.mode === 'find' ? 'choiceBtn choiceKanji' : 'choiceBtn',
        text: choice,
        onClick: () => {
          if (answered) return;
          answered = true;
          [...choiceRow.children].forEach(b => b.disabled = true);
          const isCorrect = choice === q.answer;
          btn.classList.add(isCorrect ? 'choiceCorrect' : 'choiceWrong');
          showFeedback(isCorrect, q.answer);
        }
      });
      choiceRow.appendChild(btn);
    });
    card.appendChild(choiceRow);
    card.appendChild(feedback);
    screenEl.appendChild(card);
  }

  function finishChallenge(level, total, stats) {
    const totalCorrect = stats.reading.c + stats.find.c;
    const pct = (totalCorrect / total) * 100;
    const readPct = stats.reading.t ? (stats.reading.c / stats.reading.t) * 100 : 0;
    const findPct = stats.find.t ? (stats.find.c / stats.find.t) * 100 : 0;

    Storage.setKankenBest(level, { pct, readPct, findPct, at: new Date().toISOString().slice(0, 10) });
    Storage.addCoins(totalCorrect * 5);

    let tier, message;
    if (pct >= 80) { tier = 'green'; message = '🎉 合格圏！ このちょうしで がんばろう！'; }
    else if (pct >= 60) {
      tier = 'yellow';
      const needed = Math.ceil(total * 0.8) - totalCorrect;
      message = `👍 あと少し！ あと ${Math.max(needed, 1)} 問正解できれば合格圏だよ！`;
    } else {
      tier = 'red';
      message = '💪 もう一度、ミッションで特訓しよう！';
    }

    UI.setTitle('漢検チャレンジ：けっか');
    renderStatBar();
    UI.clearScreen();
    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'bigCardEmoji', text: tier === 'green' ? '🟢' : tier === 'yellow' ? '🟡' : '🔴' }),
      el('div', { className: 'resultScore', text: `総合正答率：${Math.round(pct)}%` }),
      el('div', { className: 'rewardRow' }, [
        el('span', { className: 'rewardChip', text: `よみ ${Math.round(readPct)}%` }),
        el('span', { className: 'rewardChip', text: `かんじさがし ${Math.round(findPct)}%` })
      ]),
      el('div', { className: 'quizInstruction', text: message }),
      el('div', { className: 'resultBtnRow' }, [
        el('button', { className: 'primaryBtn', text: 'もう一度 →', onClick: startChallenge }),
        el('button', { className: 'secondaryBtn', text: 'しまマップへ', onClick: renderHome })
      ])
    ]));
  }

  // ---------- おうちの方 ----------
  function renderParentGate() {
    UI.setHomeHandler(renderHome);
    UI.setTitle('おうちの方へ');
    renderStatBar();
    UI.clearScreen();

    const a = 3 + Math.floor(Math.random() * 6);
    const b = 3 + Math.floor(Math.random() * 6);
    const card = el('div', { className: 'quizCard' });
    card.appendChild(el('div', { className: 'quizInstruction', text: 'おうちの方はこちらの計算にお答えください' }));
    card.appendChild(el('div', { className: 'questionText', text: `${a} + ${b} = ?` }));
    const input = el('input', { className: 'numInput', attrs: { type: 'number', inputmode: 'numeric' } });
    const msg = el('div', { className: 'feedback' });
    const submit = el('button', {
      className: 'submitBtn',
      text: 'すすむ',
      onClick: () => {
        if (Number(input.value) === a + b) renderParentDashboard();
        else {
          msg.innerHTML = '';
          msg.appendChild(el('div', { className: 'feedbackNg', text: 'ちがいます。もう一度お試しください。' }));
        }
      }
    });
    card.appendChild(el('div', { className: 'inputRow' }, [input, submit]));
    card.appendChild(msg);
    screenEl.appendChild(card);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
  }

  function renderParentDashboard() {
    UI.setTitle('がくしゅう記録');
    renderStatBar();
    UI.clearScreen();

    const log = Storage.getDailyLog(7);
    const weekTotal = log.reduce((s, d) => s + d.total, 0);
    const weekCorrect = log.reduce((s, d) => s + d.correct, 0);
    const learningDays = log.filter(d => d.total > 0).length;

    screenEl.appendChild(el('div', { className: 'resultCard' }, [
      el('div', { className: 'resultScore', text: '今週のがくしゅう' }),
      el('div', { className: 'rewardRow' }, [
        el('span', { className: 'rewardChip', text: `学習日 ${learningDays}/7日` }),
        el('span', { className: 'rewardChip', text: `といた問題 ${weekTotal}問` }),
        el('span', { className: 'rewardChip', text: `せいかい ${weekCorrect}問` })
      ])
    ]));

    Object.values(LEVELS).forEach(meta => {
      const mastered = masteredCount(meta.key);
      const pct = Math.round((mastered / meta.total) * 100);
      const best = Storage.getKankenBest(meta.key);
      const stars = best ? '★'.repeat(Math.max(1, Math.round(best.pct / 20))) + '☆'.repeat(5 - Math.max(1, Math.round(best.pct / 20))) : '☆☆☆☆☆';

      const weak = dataset(meta.key)
        .map(e => ({ e, lv: Storage.getIslandKanjiLevel(meta.key, e.k) }))
        .filter(x => x.lv > 0 && x.lv < 3)
        .sort((x, y) => x.lv - y.lv)
        .slice(0, 5)
        .map(x => x.e.k);

      screenEl.appendChild(el('div', { className: 'resultCard' }, [
        el('div', { className: 'resultScore', text: `漢検${meta.label}（${meta.grade}）` }),
        el('div', { className: 'quizInstruction', text: `おぼえたかんじ：${mastered} / ${meta.total}（${pct}%）` }),
        el('div', { className: 'quizInstruction', text: `もぎテスト推定合格度：${stars}${best ? '（ベスト ' + Math.round(best.pct) + '%）' : '（未挑戦）'}` }),
        el('div', { className: 'quizInstruction', text: weak.length ? `苦手なかんじ：${weak.join('　')}` : '苦手なかんじ：とくになし！' })
      ]));
    });

    screenEl.appendChild(el('button', { className: 'backBtn', text: '← しまマップへ', onClick: renderHome }));
  }

  return {
    start() {
      Storage.touchStreak();
      renderHome();
    },
    renderHome
  };
})();

Island.start();
