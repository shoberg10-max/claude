// 画面の切りかえと全体の動き（３年生 さんすう・かんじドリル：レガシー機能）
const App = (() => {
  const screenEl = UI.screenEl;
  const el = UI.el;
  const setTitle = UI.setTitle;
  const clearScreen = UI.clearScreen;

  function refreshStars() {
    UI.setStatBar(el('div', { className: 'starCount' }, [
      document.createTextNode('⭐ ' + Storage.getStars())
    ]));
  }

  // ---------- ホーム画面 ----------
  function renderHome() {
    setTitle('３年生 さんすう・かんじ ドリル');
    UI.setHomeHandler(() => Island.renderHome());
    refreshStars();
    clearScreen();
    const wrap = el('div', { className: 'homeGrid' }, [
      el('button', {
        className: 'bigCard mathCard',
        onClick: renderMathMenu
      }, [
        el('div', { className: 'bigCardEmoji', text: '🧮' }),
        el('div', { className: 'bigCardTitle', text: 'さんすう' })
      ]),
      el('button', {
        className: 'bigCard kanjiCard',
        onClick: renderKanjiMenu
      }, [
        el('div', { className: 'bigCardEmoji', text: '📖' }),
        el('div', { className: 'bigCardTitle', text: 'かんじ（3年生）' })
      ])
    ]);
    screenEl.appendChild(wrap);
    screenEl.appendChild(el('button', {
      className: 'backBtn',
      text: '← ゆるもじ島にもどる',
      onClick: () => Island.renderHome()
    }));
  }

  // ---------- さんすう：メニュー ----------
  function renderMathMenu() {
    setTitle('さんすう：もんだいをえらぼう');
    clearScreen();
    const grid = el('div', { className: 'menuGrid' });
    MathTopics.list.forEach(topic => {
      const best = Storage.getBestScore('math_' + topic.id);
      grid.appendChild(el('button', {
        className: 'menuCard',
        onClick: () => startMathQuiz(topic.id)
      }, [
        el('div', { className: 'menuCardEmoji', text: topic.emoji }),
        el('div', { className: 'menuCardTitle', text: topic.title }),
        el('div', { className: 'menuCardScore', text: best ? `ベスト ${best}/10 ⭐` : 'ちょうせん！' })
      ]));
    });
    screenEl.appendChild(grid);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← もどる', onClick: renderHome }));
  }

  // ---------- さんすう：クイズ ----------
  function startMathQuiz(topicId) {
    const topic = MathTopics.getTopic(topicId);
    const questions = topic.gen();
    runMathQuiz(topic, questions, 0, 0);
  }

  function runMathQuiz(topic, questions, index, correctCount) {
    setTitle(topic.title);
    clearScreen();

    if (index >= questions.length) {
      finishMathQuiz(topic, correctCount, questions.length);
      return;
    }

    const q = questions[index];
    const progress = el('div', { className: 'progress', text: `もんだい ${index + 1} / ${questions.length}　　せいかい ${correctCount}こ` });
    const card = el('div', { className: 'quizCard' });
    const qText = el('div', { className: 'questionText', text: q.text });
    card.appendChild(qText);

    const feedback = el('div', { className: 'feedback' });
    let answered = false;

    function goNext(wasCorrect) {
      if (wasCorrect) Storage.addStars(1);
      refreshStars();
      runMathQuiz(topic, questions, index + 1, correctCount + (wasCorrect ? 1 : 0));
    }

    function showFeedback(isCorrect, correctText) {
      feedback.innerHTML = '';
      feedback.appendChild(el('div', {
        className: isCorrect ? 'feedbackOk' : 'feedbackNg',
        text: isCorrect ? '⭕ せいかい！' : `✕ ざんねん… こたえは ${correctText}`
      }));
      const nextBtn = el('button', { className: 'nextBtn', text: index + 1 >= questions.length ? 'けっかを見る →' : 'つぎへ →' });
      nextBtn.addEventListener('click', () => goNext(isCorrect));
      feedback.appendChild(nextBtn);
    }

    if (q.kind === 'number') {
      const input = el('input', { className: 'numInput', attrs: { type: 'number', inputmode: 'numeric', placeholder: '?' } });
      const submitBtn = el('button', {
        className: 'submitBtn',
        text: 'こたえる',
        onClick: () => {
          if (answered) return;
          answered = true;
          submitBtn.disabled = true;
          input.disabled = true;
          const val = Number(input.value);
          const isCorrect = val === q.answer;
          showFeedback(isCorrect, String(q.answer));
        }
      });
      input.addEventListener('keydown', e => { if (e.key === 'Enter') submitBtn.click(); });
      card.appendChild(el('div', { className: 'inputRow' }, [input, submitBtn]));
    } else if (q.kind === 'div_rem') {
      const qInput = el('input', { className: 'numInput small', attrs: { type: 'number', inputmode: 'numeric', placeholder: '商' } });
      const rInput = el('input', { className: 'numInput small', attrs: { type: 'number', inputmode: 'numeric', placeholder: 'あまり' } });
      const submitBtn = el('button', {
        className: 'submitBtn',
        text: 'こたえる',
        onClick: () => {
          if (answered) return;
          answered = true;
          submitBtn.disabled = true;
          qInput.disabled = true;
          rInput.disabled = true;
          const isCorrect = Number(qInput.value) === q.answer.q && Number(rInput.value) === q.answer.r;
          showFeedback(isCorrect, `${q.answer.q} あまり ${q.answer.r}`);
        }
      });
      card.appendChild(el('div', { className: 'inputRow' }, [qInput, el('span', { className: 'divRemLabel', text: 'あまり' }), rInput, submitBtn]));
    } else if (q.kind === 'mc') {
      const choiceRow = el('div', { className: 'choiceRow' });
      q.choices.forEach(choice => {
        const btn = el('button', {
          className: 'choiceBtn',
          text: choice,
          onClick: () => {
            if (answered) return;
            answered = true;
            [...choiceRow.children].forEach(b => b.disabled = true);
            const isCorrect = choice === q.answer;
            btn.classList.add(isCorrect ? 'choiceCorrect' : 'choiceWrong');
            showFeedback(isCorrect, String(q.answer));
          }
        });
        choiceRow.appendChild(btn);
      });
      card.appendChild(choiceRow);
    }

    card.appendChild(feedback);
    screenEl.appendChild(progress);
    screenEl.appendChild(card);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← もどる（きろくはのこりません）', onClick: renderMathMenu }));
  }

  function finishMathQuiz(topic, correctCount, total) {
    setTitle(topic.title + '：けっか');
    clearScreen();
    Storage.setBestScore('math_' + topic.id, correctCount);
    screenEl.appendChild(resultCard(correctCount, total, () => startMathQuiz(topic.id), renderMathMenu));
  }

  function resultCard(correctCount, total, onRetry, onMenu) {
    const stars = '⭐'.repeat(Math.max(1, Math.round((correctCount / total) * 5)));
    return el('div', { className: 'resultCard' }, [
      el('div', { className: 'resultScore', text: `${correctCount} / ${total} もんせいかい！` }),
      el('div', { className: 'resultStars', text: stars }),
      el('div', { className: 'resultBtnRow' }, [
        el('button', { className: 'primaryBtn', text: 'もう一回 →', onClick: onRetry }),
        el('button', { className: 'secondaryBtn', text: 'メニューへ', onClick: onMenu })
      ])
    ]);
  }

  // ---------- かんじ：メニュー ----------
  function renderKanjiMenu() {
    setTitle('かんじ：れんしゅうをえらぼう');
    clearScreen();
    const mastered = Storage.getMasteredCount();
    const info = el('div', { className: 'kanjiInfo', text: `おぼえた かんじ：${mastered} / ${KanjiQuiz.all.length}` });

    const grid = el('div', { className: 'menuGrid' }, [
      el('button', {
        className: 'menuCard',
        onClick: () => startKanjiQuiz('reading')
      }, [
        el('div', { className: 'menuCardEmoji', text: '🈶' }),
        el('div', { className: 'menuCardTitle', text: 'よみかたクイズ' }),
        el('div', { className: 'menuCardScore', text: (() => { const b = Storage.getBestScore('kanji_reading'); return b ? `ベスト ${b}/10 ⭐` : 'ちょうせん！'; })() })
      ]),
      el('button', {
        className: 'menuCard',
        onClick: () => startKanjiQuiz('kanji')
      }, [
        el('div', { className: 'menuCardEmoji', text: '🔎' }),
        el('div', { className: 'menuCardTitle', text: 'かんじさがしクイズ' }),
        el('div', { className: 'menuCardScore', text: (() => { const b = Storage.getBestScore('kanji_kanji'); return b ? `ベスト ${b}/10 ⭐` : 'ちょうせん！'; })() })
      ]),
      el('button', {
        className: 'menuCard',
        onClick: renderKanjiBrowse
      }, [
        el('div', { className: 'menuCardEmoji', text: '📚' }),
        el('div', { className: 'menuCardTitle', text: 'かんじ一覧' }),
        el('div', { className: 'menuCardScore', text: '200字ぜんぶ見る' })
      ])
    ]);
    screenEl.appendChild(info);
    screenEl.appendChild(grid);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← もどる', onClick: renderHome }));
  }

  // ---------- かんじ：クイズ ----------
  function startKanjiQuiz(mode) {
    const questions = mode === 'reading' ? KanjiQuiz.generateReadingQuiz(10) : KanjiQuiz.generateKanjiQuiz(10);
    runKanjiQuiz(mode, questions, 0, 0);
  }

  function runKanjiQuiz(mode, questions, index, correctCount) {
    const title = mode === 'reading' ? 'よみかたクイズ' : 'かんじさがしクイズ';
    setTitle(title);
    clearScreen();

    if (index >= questions.length) {
      finishKanjiQuiz(mode, correctCount, questions.length);
      return;
    }

    const q = questions[index];
    const progress = el('div', { className: 'progress', text: `もんだい ${index + 1} / ${questions.length}　　せいかい ${correctCount}こ` });
    const card = el('div', { className: 'quizCard' });

    const promptClass = mode === 'reading' ? 'kanjiPrompt' : 'readingPrompt';
    const questionLabel = mode === 'reading' ? 'この かんじの 読みかたは？' : 'この 読みかたの かんじは？';
    card.appendChild(el('div', { className: 'quizInstruction', text: questionLabel }));
    card.appendChild(el('div', { className: promptClass, text: q.prompt }));

    const feedback = el('div', { className: 'feedback' });
    let answered = false;

    function goNext(wasCorrect) {
      Storage.markKanjiResult(q.kanji, wasCorrect);
      if (wasCorrect) Storage.addStars(1);
      refreshStars();
      runKanjiQuiz(mode, questions, index + 1, correctCount + (wasCorrect ? 1 : 0));
    }

    function showFeedback(isCorrect, correctText) {
      feedback.innerHTML = '';
      feedback.appendChild(el('div', {
        className: isCorrect ? 'feedbackOk' : 'feedbackNg',
        text: isCorrect ? '⭕ せいかい！' : `✕ ざんねん… こたえは ${correctText}`
      }));
      const nextBtn = el('button', { className: 'nextBtn', text: index + 1 >= questions.length ? 'けっかを見る →' : 'つぎへ →' });
      nextBtn.addEventListener('click', () => goNext(isCorrect));
      feedback.appendChild(nextBtn);
    }

    const choiceRow = el('div', { className: 'choiceRow' });
    q.choices.forEach(choice => {
      const btn = el('button', {
        className: mode === 'kanji' ? 'choiceBtn choiceKanji' : 'choiceBtn',
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
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← もどる（きろくはのこりません）', onClick: renderKanjiMenu }));
  }

  function finishKanjiQuiz(mode, correctCount, total) {
    setTitle('けっか');
    clearScreen();
    Storage.setBestScore('kanji_' + mode, correctCount);
    screenEl.appendChild(resultCard(correctCount, total, () => startKanjiQuiz(mode), renderKanjiMenu));
  }

  // ---------- かんじ一覧 ----------
  function renderKanjiBrowse() {
    setTitle('かんじ一覧（200字）');
    clearScreen();
    const grid = el('div', { className: 'kanjiListGrid' });
    KanjiQuiz.all.forEach(entry => {
      const level = Storage.getKanjiLevel(entry.k);
      const readings = [...entry.on.map(r => 'オ' + r), ...entry.kun.map(r => 'ク' + r)].join('　');
      grid.appendChild(el('div', { className: 'kanjiListItem' }, [
        el('div', { className: 'kanjiListChar' + (level >= 3 ? ' mastered' : ''), text: entry.k }),
        el('div', { className: 'kanjiListReading', text: readings })
      ]));
    });
    screenEl.appendChild(el('div', { className: 'kanjiListLegend', text: 'オ＝音読み　ク＝訓読み　金色＝おぼえたかんじ' }));
    screenEl.appendChild(grid);
    screenEl.appendChild(el('button', { className: 'backBtn', text: '← もどる', onClick: renderKanjiMenu }));
  }

  return {
    renderHome
  };
})();
