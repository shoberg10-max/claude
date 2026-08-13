// 画面遷移とUI制御（3ステップのウィザード）
(function () {
  const SAMPLE_NOTES = `# 新規顧客管理システム導入に関する検討報告

## [pattern: pyramid] （１）背景・目的｜①課題認識と対応方針
> 顧客情報の分散管理が対応品質の低下と解約率上昇を招いており、**クラウド型CRM**による一元化が急務である
>> 現行体制のまま推移した場合、来期の解約率はさらに悪化する見込みである
>> 対応品質のばらつきは、担当者間の情報共有不足が主因と分析している
- 課題：顧客対応の質と速度が低下し、顧客満足度調査のスコアが**前年比8ポイント低下**している
- 原因：顧客情報が担当者個人のPC・Excelに分散し、組織的な検索・共有ができていない
- 対策：クラウド型CRMを導入し、顧客情報と対応履歴を全社で一元管理する体制を構築する
出所：社内顧客満足度調査（2026年1月実施）

## [pattern: compare-vertical] （２）論点の整理｜①比較軸別に見た候補ツールの違い
> 初期費用・運用コスト・拡張性など5つの軸で3社を整理すると、**B社CRM**が総合的にバランスが良い
>> 特にコストと拡張性のトレードオフの観点で、A社・C社に対する優位性が明確である
- 初期費用：A社30万円、B社50万円、C社120万円と幅がある
- 月額運用コスト：A社1万円、B社3万円、C社8万円
- 機能の拡張性：A社は標準機能のみ、B社はオプションで拡張可能、C社は高機能だが過剰投資気味
- サポート体制：A社はメールのみ、B社は電話・チャット対応、C社は専任担当者付き
- 導入期間：A社1ヶ月、B社2ヶ月、C社4ヶ月

## [pattern: box-compare] （３）提案｜①候補ツールの比較と推奨案
> 3社を比較した結果、機能とコストのバランスから**B社CRM**が最も現実的な選択肢である
- A社CRM：低コストだが分析機能が限定的で、将来的な拡張性に課題がある
- B社CRM：★推奨 コストは中程度で、必要な機能を標準搭載しており拡張性も高い
- C社CRM：高コストだが分析機能が充実しており、大規模組織向けの構成である

## [pattern: kpi-summary] （３）提案｜②導入効果の見込み
> 対応品質と業務効率の両面で定量的な改善効果が見込まれる
- 顧客対応リードタイム：50%短縮（平均対応時間 30分→15分）
- 解約率：20%改善（対応履歴の一元化により解約予兆を早期検知）
- 営業一人あたり商談数：15%増加（情報検索時間の削減により商談準備時間を確保）
出所：NRI「CRM導入効果調査」2024年版

## [pattern: before-after] （３）提案｜③情報検索にかかる時間の変化
> 情報検索の一元化により、担当者1人あたり年間換算で約50時間の業務時間削減が見込まれる
- 現状：情報を探すために複数の担当者に確認する必要があり、平均15分を要している
- 導入後：CRM上で即座に検索でき、平均2分に短縮される見込みである

## [pattern: title-message] （４）意思決定｜①本日ご確認いただきたい事項
> 本提案について、以下3点のご承認をいただいたうえで導入プロジェクトを開始したい
- B社CRMを導入ベンダーとして正式に決定する
- 初期費用50万円・月額運用コスト3万円の予算枠を確保する
- プロジェクトオーナーを情報システム部長とし、9月から準備に着手する

## [pattern: action-plan-table] （５）実行計画｜①導入までのアクションプラン
> ご承認後、以下のスケジュールで各タスクを実行する
- 要件定義・ベンダー最終契約を完了する｜情報システム部　鈴木｜2026年9月末
- 既存データ移行・初期設定を行う｜情報システム部　佐藤｜2026年10月中旬
- 管理者・現場担当者向け研修を実施する｜情報システム部　佐藤｜2026年10月末
- 一部部署での試験運用を行い課題を洗い出す｜営業部　田中｜2026年11月末
- 全社展開および利用ルールを周知徹底する｜情報システム部　鈴木｜2026年12月末
`;

  const state = {
    step: 1,
    notesText: '',
    outline: { title: '', slides: [] },
    storyType: 'decision',
    deepDiveKind: 'issues',
    deepDiveTheme: '',
  };

  const screen = document.getElementById('screen');

  function uid() {
    return 'slide-' + Math.random().toString(36).slice(2, 9);
  }

  function ensureIds(slides) {
    slides.forEach((s) => {
      if (!s.id) s.id = uid();
      if (!s.bullets) s.bullets = [];
    });
  }

  // 構成案の見出しに "[pattern: xxx]" タグが付いていれば（js/promptBuilder.js が
  // 生成するプロンプトの出力フォーマット）、それを社内LLMによるパターン提案として採用する。
  function applyTaggedPatterns(outline) {
    outline.slides.forEach((s) => {
      if (s.taggedPatternId && DocAssist.patternById[s.taggedPatternId]) {
        s.patternId = s.taggedPatternId;
        s.patternSource = 'ai';
        s.patternReason = '社内LLMが構成案の中で提案したパターンです。';
      }
      delete s.taggedPatternId;
    });
  }

  // patternSource: 'rule'（自動選択） / 'ai'（社内LLM提案） / 'manual'（手動選択）
  // 'ai' と 'manual' はユーザーが確定させた選択なので、構成案編集画面に戻って
  // 内容を直しても上書きしない。
  function ensurePatterns() {
    state.outline.slides.forEach((s) => {
      const confirmed = s.patternSource === 'ai' || s.patternSource === 'manual';
      if (!confirmed || !DocAssist.patternById[s.patternId]) {
        const result = DocAssist.selectPattern(s);
        s.patternId = result.pattern.id;
        s.patternSource = 'rule';
        s.patternReason = '';
      }
    });
  }

  // 単純な画面切り替え（データの再生成はしない）。Step間を行き来しても
  // 構成案の編集内容やパターン選択が失われないようにするため、
  // 「メモから構成案を生成する」処理と「パターンを確定する」処理は
  // それぞれ Step 1 / Step 2 の「次へ」ボタンからのみ呼び出す。
  function goToStep(n) {
    state.step = n;
    render();
  }

  function generateOutlineAndProceed() {
    if (!state.notesText.trim()) {
      renderStep1('会議メモを入力してください。');
      return;
    }
    const outline = DocAssist.outlineProvider(state.notesText);
    ensureIds(outline.slides);
    applyTaggedPatterns(outline);
    outline.arrangeEnabled = true;
    outline.arrangeSeed = 0;
    state.outline = outline;
    if (!state.outline.slides.length) {
      renderStep1('構成案を読み取れませんでした。見出しや箇条書きを増やして再度お試しください。');
      return;
    }
    goToStep(2);
  }

  function proceedToStep3() {
    if (!state.outline.slides.length) return;
    ensurePatterns();
    goToStep(3);
  }

  function render() {
    renderStepPills();
    if (state.step === 1) renderStep1();
    else if (state.step === 2) renderStep2();
    else renderStep3();
  }

  function renderStepPills() {
    const labels = ['会議メモを入力', '構成案を確認・編集', 'レイアウト生成 → PPTX書き出し'];
    const pills = document.getElementById('stepPills');
    pills.innerHTML = labels
      .map((label, i) => {
        const n = i + 1;
        const cls = n === state.step ? 'active' : n < state.step ? 'done' : '';
        return `<div class="step-pill ${cls}"><span class="num">${n}</span>${label}</div>`;
      })
      .join('');
  }

  // ---------------- Step 1 ----------------
  function renderStep1(errorMsg) {
    screen.innerHTML = `
      <div class="panel">
        <h2>Step 1. 会議メモ・議事録を入力</h2>
        <p class="hint">次回会議に向けた資料の元になるメモや議事録を貼り付けてください。見出し（# や ##）や箇条書き（- ・）があると構成案の精度が上がりますが、フリーテキストでも解析します。</p>
        <p class="hint">資料化の元ネタがまだ薄い場合は、まず「🔍 テーマ深掘り用プロンプトを作成」でテーマそのものを社内LLMに掘り下げてもらってください。その出力をこの欄に貼り付けたら、「🪄 社内LLM用プロンプトを作成」でストーリー型（SCQA法・課題原因解決策など）を選んで構成案に整形し、出力結果をこの下の欄に貼り付け直してから「次へ」を押してください。</p>
        <textarea class="notes-input" id="notesInput" placeholder="例：\n# 会議タイトル\n## 現状の課題\n- 課題A\n- 課題B\n## 比較：A案 vs B案\nA案：...\nB案：...">${escapeHtml(state.notesText)}</textarea>
        ${errorMsg ? `<p class="status-msg error">${escapeHtml(errorMsg)}</p>` : ''}
        <div class="btn-row">
          <button class="btn ghost" id="sampleBtn">サンプルを試す</button>
          <div style="display:flex;gap:10px;flex-wrap:wrap;">
            <button class="btn secondary" id="deepDiveBtn">🔍 テーマ深掘り用プロンプトを作成</button>
            <button class="btn secondary" id="promptBtn">🪄 社内LLM用プロンプトを作成</button>
            <button class="btn" id="nextBtn">次へ：構成案を生成</button>
          </div>
        </div>
      </div>
    `;
    const input = document.getElementById('notesInput');
    input.addEventListener('input', () => {
      state.notesText = input.value;
    });
    document.getElementById('sampleBtn').addEventListener('click', () => {
      state.notesText = SAMPLE_NOTES;
      input.value = SAMPLE_NOTES;
    });
    document.getElementById('deepDiveBtn').addEventListener('click', () => openDeepDiveModal());
    document.getElementById('promptBtn').addEventListener('click', () => openPromptModal(input.value));
    document.getElementById('nextBtn').addEventListener('click', generateOutlineAndProceed);
  }

  // テーマ深掘り用プロンプト（Step 0相当）：資料化の元ネタがまだ薄いときに、
  // テーマそのものを社内LLMで掘り下げるための3つの型（論点整理／提案書向け／MECE）から選ぶ。
  function openDeepDiveModal() {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    const typeOptions = DocAssist.DEEP_DIVE_TYPE_ORDER.map(
      (id) => `<option value="${id}"${id === state.deepDiveKind ? ' selected' : ''}>${escapeHtml(DocAssist.DEEP_DIVE_TYPES[id].label)}</option>`
    ).join('');
    overlay.innerHTML = `
      <div class="modal-box modal-box-wide">
        <h3>🔍 テーマ深掘り用プロンプト</h3>
        <p class="hint">会議メモがまだ無い、または議論が浅い段階で、テーマそのものを社内LLMに掘り下げてもらうためのプロンプトです。出力結果は次のステップ（構成案整形プロンプト）の元ネタとして使います。</p>
        <label for="deepDiveTheme">テーマ・議題</label>
        <input type="text" id="deepDiveTheme" placeholder="例：新規顧客管理システムの導入方針" value="${escapeAttr(state.deepDiveTheme)}">
        <label for="deepDiveKind">型を選ぶ</label>
        <select id="deepDiveKind">${typeOptions}</select>
        <p class="hint" id="deepDiveHint"></p>
        <textarea class="prompt-output" id="deepDiveOutput" readonly></textarea>
        <div class="btn-row">
          <span class="status-msg" id="deepDiveCopyStatus"></span>
          <div style="display:flex;gap:10px;">
            <button class="btn ghost" id="deepDiveClose">閉じる</button>
            <button class="btn" id="deepDiveCopy">コピーする</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const close = () => document.body.removeChild(overlay);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    overlay.querySelector('#deepDiveClose').addEventListener('click', close);

    const themeInput = overlay.querySelector('#deepDiveTheme');
    const kindSelect = overlay.querySelector('#deepDiveKind');
    const hintEl = overlay.querySelector('#deepDiveHint');
    const textarea = overlay.querySelector('#deepDiveOutput');

    function refresh() {
      state.deepDiveTheme = themeInput.value;
      state.deepDiveKind = kindSelect.value;
      hintEl.textContent = DocAssist.DEEP_DIVE_TYPES[state.deepDiveKind].hint;
      textarea.value = DocAssist.buildDeepDivePrompt(state.deepDiveTheme, state.deepDiveKind);
    }
    themeInput.addEventListener('input', refresh);
    kindSelect.addEventListener('change', refresh);
    refresh();

    overlay.querySelector('#deepDiveCopy').addEventListener('click', async () => {
      const statusEl = overlay.querySelector('#deepDiveCopyStatus');
      try {
        await navigator.clipboard.writeText(textarea.value);
        statusEl.textContent = 'コピーしました。';
        statusEl.classList.remove('error');
      } catch (e) {
        textarea.select();
        statusEl.textContent = '自動コピーに失敗しました。テキストを選択したので Ctrl+C（Mac は Cmd+C）でコピーしてください。';
        statusEl.classList.add('error');
      }
    });
  }

  function openPromptModal(rawNotes) {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    const typeOptions = DocAssist.STORY_TYPE_ORDER.map(
      (id) => `<option value="${id}"${id === state.storyType ? ' selected' : ''}>${escapeHtml(DocAssist.STORY_TYPES[id].label)}</option>`
    ).join('');
    overlay.innerHTML = `
      <div class="modal-box modal-box-wide">
        <h3>🪄 社内LLM用 構成案整形プロンプト</h3>
        <p class="hint">下のテキストをコピーして社内LLMのチャットに貼り付けて実行してください。出力された構成案をコピーし、Step 1の入力欄に貼り付け直してから「次へ」を押すと、このアプリの形式にきれいに読み込まれます（見出しに含まれる [pattern: ...] タグはデザインパターンの提案として自動的に反映されます）。</p>
        <label for="storyType">資料全体のストーリー型</label>
        <select id="storyType">${typeOptions}</select>
        <p class="hint" id="storyTypeHint"></p>
        <textarea class="prompt-output" id="promptOutput" readonly></textarea>
        <div class="btn-row">
          <span class="status-msg" id="copyStatus"></span>
          <div style="display:flex;gap:10px;">
            <button class="btn ghost" id="promptClose">閉じる</button>
            <button class="btn" id="promptCopy">コピーする</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const close = () => document.body.removeChild(overlay);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    overlay.querySelector('#promptClose').addEventListener('click', close);

    const storySelect = overlay.querySelector('#storyType');
    const hintEl = overlay.querySelector('#storyTypeHint');
    const textarea = overlay.querySelector('#promptOutput');

    function refresh() {
      state.storyType = storySelect.value;
      hintEl.textContent = DocAssist.STORY_TYPES[state.storyType].hint;
      textarea.value = DocAssist.buildOutlineFormatPrompt(rawNotes, state.storyType);
    }
    storySelect.addEventListener('change', refresh);
    refresh();

    overlay.querySelector('#promptCopy').addEventListener('click', async () => {
      const statusEl = overlay.querySelector('#copyStatus');
      try {
        await navigator.clipboard.writeText(textarea.value);
        statusEl.textContent = 'コピーしました。';
        statusEl.classList.remove('error');
      } catch (e) {
        textarea.select();
        statusEl.textContent = '自動コピーに失敗しました。テキストを選択したので Ctrl+C（Mac は Cmd+C）でコピーしてください。';
        statusEl.classList.add('error');
      }
    });
  }

  // ---------------- Step 2 ----------------
  function renderStep2() {
    const o = state.outline;
    screen.innerHTML = `
      <div class="panel">
        <h2>Step 2. 構成案を確認・編集</h2>
        <p class="hint">メモから自動生成した構成案です。見出しや箇条書きを自由に編集・追加・削除・並べ替えできます。</p>
        <div class="title-field">
          <label for="docTitle">資料タイトル</label>
          <input type="text" id="docTitle" value="${escapeAttr(o.title)}">
        </div>
        <div class="outline-list" id="outlineList"></div>
        <div class="add-slide-row">
          <button class="btn secondary" id="addSlideBtn">＋ スライドを追加</button>
        </div>
        <div class="btn-row">
          <button class="btn ghost" id="backBtn">← 戻る</button>
          <button class="btn" id="nextBtn">次へ：レイアウトを自動生成</button>
        </div>
      </div>
    `;

    document.getElementById('docTitle').addEventListener('input', (e) => {
      o.title = e.target.value;
    });

    const list = document.getElementById('outlineList');
    o.slides.forEach((slide, idx) => {
      list.appendChild(renderOutlineCard(slide, idx, o.slides.length));
    });

    document.getElementById('addSlideBtn').addEventListener('click', () => {
      o.slides.push({ id: uid(), heading: '新しいスライド', message: '', subMessage: '', bullets: [] });
      renderStep2();
    });
    document.getElementById('backBtn').addEventListener('click', () => goToStep(1));
    document.getElementById('nextBtn').addEventListener('click', proceedToStep3);
  }

  function renderOutlineCard(slide, idx, total) {
    const card = document.createElement('div');
    card.className = 'outline-card';
    card.innerHTML = `
      <div class="outline-card-head">
        <span class="idx">${idx + 1}</span>
        <input type="text" class="heading" placeholder="見出し（トピックラベル）" value="${escapeAttr(slide.heading)}">
        <button class="icon-btn" data-act="up" title="上へ" ${idx === 0 ? 'disabled' : ''}>↑</button>
        <button class="icon-btn" data-act="down" title="下へ" ${idx === total - 1 ? 'disabled' : ''}>↓</button>
        <button class="icon-btn danger" data-act="del" title="削除">✕</button>
      </div>
      <input type="text" class="message" placeholder="リード文（このスライドの結論を一文で。未入力の場合は先頭の箇条書きを使用）" value="${escapeAttr(slide.message || '')}">
      <textarea class="sub-message" rows="2" placeholder="サブリード文（リード文を補足する箇条書き。1行に1項目、任意）">${escapeHtml(slide.subMessage || '')}</textarea>
      <textarea class="bullets" placeholder="箇条書きを1行ずつ入力（メッセージの根拠となる情報）">${escapeHtml((slide.bullets || []).join('\n'))}</textarea>
    `;
    card.querySelector('.heading').addEventListener('input', (e) => {
      slide.heading = e.target.value;
    });
    card.querySelector('.message').addEventListener('input', (e) => {
      slide.message = e.target.value;
    });
    card.querySelector('.sub-message').addEventListener('input', (e) => {
      slide.subMessage = e.target.value;
    });
    card.querySelector('.bullets').addEventListener('input', (e) => {
      slide.bullets = e.target.value.split('\n').map((s) => s.trim()).filter(Boolean);
    });
    card.querySelector('[data-act="del"]').addEventListener('click', () => {
      const i = state.outline.slides.indexOf(slide);
      state.outline.slides.splice(i, 1);
      renderStep2();
    });
    card.querySelector('[data-act="up"]').addEventListener('click', () => {
      const i = state.outline.slides.indexOf(slide);
      if (i > 0) {
        [state.outline.slides[i - 1], state.outline.slides[i]] = [state.outline.slides[i], state.outline.slides[i - 1]];
        renderStep2();
      }
    });
    card.querySelector('[data-act="down"]').addEventListener('click', () => {
      const i = state.outline.slides.indexOf(slide);
      if (i < state.outline.slides.length - 1) {
        [state.outline.slides[i + 1], state.outline.slides[i]] = [state.outline.slides[i], state.outline.slides[i + 1]];
        renderStep2();
      }
    });
    return card;
  }

  // ---------------- Step 3 ----------------
  function renderStep3() {
    const o = state.outline;
    screen.innerHTML = `
      <div class="panel">
        <h2>Step 3. デザインパターンを自動選択 → PowerPoint書き出し</h2>
        <p class="hint">各スライドの内容から、あらかじめ用意したコンサルスライドのデザインパターンを自動選択しました。「🤖 AI選択」は社内LLMがStep 1のプロンプトで提案したパターンです。プレビュー右上のプルダウンから手動で変更することもできます。</p>
        <details>
          <summary style="cursor:pointer;color:var(--accent);font-size:13px;">利用可能なデザインパターン一覧（${DocAssist.patterns.length}種類）</summary>
          ${renderPatternLegend()}
        </details>
        <div class="slide-grid" id="slideGrid"></div>
        <div class="arrange-row">
          <label class="arrange-toggle">
            <input type="checkbox" id="arrangeToggle"${o.arrangeEnabled !== false ? ' checked' : ''}>
            🎲 出力時に章ごとの色味・見出し装飾を少しアレンジする
          </label>
          <button type="button" class="btn ghost" id="arrangeRerollBtn"${o.arrangeEnabled === false ? ' disabled' : ''}>🔄 アレンジ案を変える</button>
          <span class="hint arrange-hint">デザインパターン自体は変わりません。書き出したPowerPointにのみ反映され、この画面のプレビューは変わりません。</span>
        </div>
        <div class="export-row">
          <input type="text" class="filename" id="fileName" value="${escapeAttr((o.title || '資料構成案') + '.pptx')}">
          <button class="btn" id="exportBtn">PowerPointを書き出す</button>
          <span class="status-msg" id="exportStatus"></span>
        </div>
        <div class="btn-row">
          <button class="btn ghost" id="backBtn">← 戻る</button>
        </div>
      </div>
    `;

    const grid = document.getElementById('slideGrid');
    o.slides.forEach((slide, idx) => {
      grid.appendChild(renderSlidePreviewCard(slide, idx));
    });

    document.getElementById('backBtn').addEventListener('click', () => goToStep(2));
    document.getElementById('arrangeToggle').addEventListener('change', (e) => {
      state.outline.arrangeEnabled = e.target.checked;
      document.getElementById('arrangeRerollBtn').disabled = !e.target.checked;
    });
    document.getElementById('arrangeRerollBtn').addEventListener('click', () => {
      state.outline.arrangeSeed = (state.outline.arrangeSeed || 0) + 1;
      const statusEl = document.getElementById('exportStatus');
      statusEl.textContent = `アレンジ案を切り替えました（案 ${state.outline.arrangeSeed + 1}）。書き出すと反映されます。`;
      statusEl.classList.remove('error');
    });
    document.getElementById('exportBtn').addEventListener('click', async () => {
      const statusEl = document.getElementById('exportStatus');
      const fileNameInput = document.getElementById('fileName');
      statusEl.textContent = '生成中…';
      statusEl.classList.remove('error');
      try {
        await DocAssist.exportPptx(state.outline, fileNameInput.value.trim() || undefined);
        statusEl.textContent = 'PowerPointファイルを書き出しました。';
      } catch (e) {
        console.error(e);
        statusEl.textContent = '書き出しに失敗しました: ' + e.message;
        statusEl.classList.add('error');
      }
    });
  }

  function sourceLabel(source) {
    if (source === 'ai') return '🤖 AI選択';
    if (source === 'manual') return '✋ 手動選択';
    return '📐 自動選択（ルール）';
  }

  function replaceCard(slide) {
    const idx = state.outline.slides.indexOf(slide);
    const grid = document.getElementById('slideGrid');
    if (grid && grid.children[idx]) {
      grid.replaceChild(renderSlidePreviewCard(slide, idx), grid.children[idx]);
    }
  }

  // 比較表・ボックス比較など「1行=1項目」の箇条書きを使うパターンで、Step 3の
  // プレビューを見ながら項目数を素早く増減できるようにする。減らす場合は末尾から
  // 削る。増やす場合は「項目N：説明を入力してください」の仮の行を追加する
  // （KV形式のため、box-compare/comparison-table系のパターンでそのまま表示できる。
  // 他のパターンでもラベル付き箇条書きとして表示され、空にはならない）。
  function setBulletCount(slide, n) {
    const target = Math.max(1, n);
    const bullets = (slide.bullets || []).slice();
    if (target <= bullets.length) {
      slide.bullets = bullets.slice(0, target);
    } else {
      let i = bullets.length;
      while (bullets.length < target) {
        i++;
        bullets.push(`項目${i}：説明を入力してください`);
      }
      slide.bullets = bullets;
    }
    replaceCard(slide);
  }

  function renderPatternLegend() {
    const groups = [];
    const byCategory = {};
    DocAssist.patterns.forEach((p) => {
      const cat = p.category || 'その他';
      if (!byCategory[cat]) {
        byCategory[cat] = [];
        groups.push(cat);
      }
      byCategory[cat].push(p);
    });
    return groups
      .map(
        (cat) => `
        <div class="pattern-legend-group">
          <div class="pattern-legend-cat">${escapeHtml(cat)}</div>
          <div class="pattern-legend">
            ${byCategory[cat].map((p) => `<div class="item"><b>${escapeHtml(p.name)}</b>${escapeHtml(p.description)}</div>`).join('')}
          </div>
        </div>`
      )
      .join('');
  }

  // ---------------- テンプレート選択ギャラリー ----------------
  // 用途（シーン）と役割で絞り込みながら、テンプレートを一覧から選べるようにする。
  // サムネイルは pattern.renderBody() の出力をそのまま縮小表示しているだけなので、
  // js/patterns.js にパターンを1つ追加すればギャラリーにも自動的に並ぶ
  // （サムネイル画像を別途用意する必要がない）。
  // そのスライドの実データで描けるパターンは実データで、描けないパターン
  // （ガント・移行図など固有の記法が要るもの）はサンプルでプレビューする。
  const THUMB_W = 460;

  function chipHtml(axis, value, label, active) {
    return `<button type="button" class="tpl-chip${active ? ' is-active' : ''}" data-axis="${axis}" data-value="${
      value == null ? '' : escapeAttr(value)
    }">${escapeHtml(label)}</button>`;
  }

  function openTemplatePicker(slide) {
    const filters = { scene: null, role: null, q: '' };

    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="tpl-modal">
        <div class="tpl-head">
          <h3>テンプレート</h3>
          <input type="search" class="tpl-search" id="tplSearch" placeholder="テンプレート名・説明で検索">
          <button class="btn ghost" id="tplClose">閉じる</button>
        </div>
        <div class="tpl-filters">
          <div class="tpl-chip-row">
            <span class="tpl-axis-label">用途</span>
            ${chipHtml('scene', null, 'すべて', true)}
            ${DocAssist.patternScenes.map((s) => chipHtml('scene', s, s, false)).join('')}
          </div>
          <div class="tpl-chip-row">
            <span class="tpl-axis-label">役割</span>
            ${chipHtml('role', null, 'すべて', true)}
            ${DocAssist.patternRoles.map((r) => chipHtml('role', r, r, false)).join('')}
          </div>
        </div>
        <div class="tpl-actions">
          <span class="tpl-count" id="tplCount"></span>
          <div class="tpl-action-btns">
            <button type="button" class="btn secondary" id="tplNew">＋ テンプレートを追加</button>
            <button type="button" class="btn ghost" id="tplImport">読み込み</button>
            <button type="button" class="btn ghost" id="tplExport">書き出し</button>
          </div>
        </div>
        <div class="tpl-grid" id="tplGrid"></div>
        <input type="file" id="tplFile" accept="application/json,.json" hidden>
      </div>
    `;
    document.body.appendChild(overlay);

    function close() {
      document.removeEventListener('keydown', onKey);
      window.removeEventListener('resize', scaleThumbs);
      if (overlay.parentNode) document.body.removeChild(overlay);
    }
    function onKey(e) {
      // テンプレート作成画面が上に開いているときは、そちらのEscを優先する
      if (e.key === 'Escape' && !document.querySelector('.tb-modal')) close();
    }

    // サムネイルは実寸（460px幅）で組んでから、カード幅に合わせてCSSで縮小する。
    // こうすると本番のプレビューと同じマークアップ・同じCSSをそのまま使い回せる。
    function scaleThumbs() {
      overlay.querySelectorAll('.tpl-thumb').forEach((thumb) => {
        const inner = thumb.querySelector('.tpl-thumb-scale');
        if (inner && thumb.clientWidth) inner.style.transform = `scale(${thumb.clientWidth / THUMB_W})`;
      });
    }

    function matches(p) {
      if (filters.scene && !(p.scenes || []).includes(filters.scene)) return false;
      if (filters.role && p.role !== filters.role) return false;
      if (filters.q) {
        const hay = `${p.name} ${p.description} ${p.category} ${p.role} ${(p.scenes || []).join(' ')}`.toLowerCase();
        if (!hay.includes(filters.q.toLowerCase())) return false;
      }
      return true;
    }

    function safeRender(pattern, bullets) {
      try {
        return pattern.renderBody(bullets) || '';
      } catch (e) {
        return '';
      }
    }

    function renderGrid() {
      const grid = overlay.querySelector('#tplGrid');
      const list = DocAssist.patterns.filter(matches);
      overlay.querySelector('#tplCount').textContent = `${list.length}件のテンプレート`;
      if (!list.length) {
        grid.innerHTML = '<p class="hint">条件に合うテンプレートがありません。絞り込みを解除してください。</p>';
        return;
      }
      const own = slide.bullets || [];
      const lead = DocAssist.analyze.stripEmphasis(DocAssist.analyze.effectiveMessage(slide));
      grid.innerHTML = list
        .map((p) => {
          let body = own.length ? safeRender(p, own) : '';
          const usedOwn = !!body && !/pv-empty/.test(body);
          if (!usedOwn) body = safeRender(p, p.sample);
          return `
          <div role="button" tabindex="0" class="tpl-card${p.id === slide.patternId ? ' is-selected' : ''}" data-id="${p.id}" title="${escapeAttr(
            p.description
          )}">
            <div class="tpl-thumb">
              <div class="tpl-thumb-scale">
                <div class="tpl-thumb-title"><span class="bar"></span>${escapeHtml(slide.heading || '見出しが入ります')}</div>
                <div class="tpl-thumb-lead">${escapeHtml(lead || 'このスライドの結論が1文で入ります')}</div>
                <div class="tpl-thumb-body">${body}</div>
              </div>
            </div>
            <div class="tpl-card-foot">
              <span class="tpl-card-name">${escapeHtml(p.name)}</span>
              <span class="tpl-card-tag">${escapeHtml(p.role)}</span>
            </div>
            <span class="tpl-card-src">${usedOwn ? 'このスライドの内容で表示' : 'サンプル内容で表示'}</span>
            ${
              p.isCustom
                ? `<span class="tpl-card-custom">自作</span>
                   <span class="tpl-card-actions">
                     <button type="button" class="tpl-mini" data-act="edit" title="編集">編集</button>
                     <button type="button" class="tpl-mini danger" data-act="del" title="削除">削除</button>
                   </span>`
                : ''
            }
          </div>`;
        })
        .join('');
      scaleThumbs();
      grid.querySelectorAll('.tpl-card').forEach((card) => {
        const apply = () => {
          slide.patternId = card.dataset.id;
          slide.patternSource = 'manual';
          slide.patternReason = '';
          close();
          replaceCard(slide);
        };
        card.addEventListener('click', (e) => {
          const actionBtn = e.target.closest('[data-act]');
          if (!actionBtn) {
            apply();
            return;
          }
          e.stopPropagation();
          const id = card.dataset.id;
          if (actionBtn.dataset.act === 'edit') {
            openTemplateBuilder(DocAssist.customTemplates.getSpec(id), () => renderGrid());
          } else if (actionBtn.dataset.act === 'del') {
            const p = DocAssist.patternById[id];
            if (window.confirm(`テンプレート「${p ? p.name : id}」を削除します。よろしいですか？`)) {
              DocAssist.customTemplates.deleteSpec(id);
              // 削除したテンプレートを使っていたスライドは既定のパターンに戻す
              state.outline.slides.forEach((s) => {
                if (s.patternId === id) {
                  s.patternId = 'title-message';
                  s.patternSource = 'auto';
                }
              });
              renderGrid();
            }
          }
        });
        card.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            apply();
          }
        });
      });
      // 適用中のテンプレートは一覧の途中にあることが多いので、開いた時点で見える位置まで送る
      const selected = grid.querySelector('.tpl-card.is-selected');
      if (selected) selected.scrollIntoView({ block: 'nearest' });
    }

    overlay.querySelectorAll('.tpl-chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const axis = chip.dataset.axis;
        const value = chip.dataset.value || null;
        filters[axis] = value;
        overlay.querySelectorAll('.tpl-chip').forEach((c) => {
          if (c.dataset.axis === axis) c.classList.toggle('is-active', (c.dataset.value || null) === value);
        });
        renderGrid();
      });
    });
    const search = overlay.querySelector('#tplSearch');
    search.addEventListener('input', () => {
      filters.q = search.value.trim();
      renderGrid();
    });
    overlay.querySelector('#tplClose').addEventListener('click', close);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    document.addEventListener('keydown', onKey);
    window.addEventListener('resize', scaleThumbs);

    // 新規作成
    overlay.querySelector('#tplNew').addEventListener('click', () => {
      openTemplateBuilder(null, (savedSpec) => {
        renderGrid();
        // 作った直後はそのテンプレートを見つけやすいよう、名前で絞り込んでおく
        if (savedSpec) {
          search.value = savedSpec.name;
          filters.q = savedSpec.name;
          renderGrid();
        }
      });
    });

    // JSONファイルとして書き出し（同僚への共有・バックアップ用）
    overlay.querySelector('#tplExport').addEventListener('click', () => {
      const specs = DocAssist.customTemplates.listSpecs();
      if (!specs.length) {
        window.alert('書き出せる自作テンプレートがありません。先に「＋ テンプレートを追加」で作成してください。');
        return;
      }
      const blob = new Blob([DocAssist.customTemplates.exportJson()], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'docassist-templates.json';
      a.click();
      URL.revokeObjectURL(url);
    });

    // JSONファイルの読み込み
    const fileInput = overlay.querySelector('#tplFile');
    overlay.querySelector('#tplImport').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const r = DocAssist.customTemplates.importJson(String(reader.result));
          renderGrid();
          window.alert(`テンプレートを読み込みました（新規${r.added}件 / 更新${r.updated}件${r.skipped ? ` / 対象外${r.skipped}件` : ''}）。`);
        } catch (e) {
          window.alert('読み込みに失敗しました：' + e.message);
        }
        fileInput.value = '';
      };
      reader.readAsText(file);
    });

    renderGrid();
  }

  // ---------------- テンプレート作成・編集 ----------------
  // レイアウトの仕様（js/templateSpec.js）をフォームで組み立て、右側にその場で
  // プレビューしながら自作テンプレートを作る画面。コードを書かずに型を増やせる
  // ようにするための入口。
  const FROM_OPTIONS = [
    ['key', '見出し（「A：B」のA）'],
    ['value', '内容（「A：B」のB）'],
    ['field:0', '内容の1番目（｜区切り）'],
    ['field:1', '内容の2番目（｜区切り）'],
    ['field:2', '内容の3番目（｜区切り）'],
    ['field:3', '内容の4番目（｜区切り）'],
    ['index', '連番'],
  ];
  const FILL_OPTIONS = [
    ['primary', 'ネイビー'],
    ['primaryLight', 'ミディアムブルー'],
    ['accent', 'ライトブルー'],
    ['gray', 'グレー'],
    ['light', '淡いブルー（塗り）'],
    ['lighter', 'さらに淡いブルー'],
    ['white', '白'],
    ['highlight', 'オレンジ（強調）'],
  ];

  function optionsHtml(pairs, selected) {
    return pairs
      .map(([v, label]) => `<option value="${escapeAttr(v)}"${v === selected ? ' selected' : ''}>${escapeHtml(label)}</option>`)
      .join('');
  }

  function openTemplateBuilder(existingSpec, onSaved) {
    const editing = !!existingSpec;
    const spec = DocAssist.templateSpec.normalizeSpec(
      existingSpec || {
        name: '',
        category: '比較',
        role: '比較',
        scenes: ['提案書'],
        layout: { kind: 'cards', direction: 'row', head: { show: true, from: 'key', fill: 'primary' }, body: { from: 'value', fill: 'light' } },
      }
    );
    const L = spec.layout;

    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="tb-modal">
        <div class="tb-head">
          <h3>${editing ? 'テンプレートを編集' : 'テンプレートを追加'}</h3>
          <span class="hint">左で設定すると右のプレビューがその場で変わります。</span>
        </div>
        <div class="tb-cols">
          <div class="tb-form" id="tbForm">

            <div class="tb-section">基本情報</div>
            <label class="tb-field"><span>テンプレート名</span>
              <input type="text" name="name" value="${escapeAttr(spec.name)}" placeholder="例：3案の横並び比較"></label>
            <label class="tb-field"><span>説明</span>
              <input type="text" name="description" value="${escapeAttr(spec.description)}" placeholder="どんなときに使う型か"></label>
            <div class="tb-row">
              <label class="tb-field"><span>カテゴリ</span>
                <select name="category">${optionsHtml(DocAssist.patternCategories.map((c) => [c, c]), spec.category)}</select></label>
              <label class="tb-field"><span>役割</span>
                <select name="role">${optionsHtml(DocAssist.patternRoles.map((r) => [r, r]), spec.role)}</select></label>
            </div>
            <div class="tb-field"><span>用途（ギャラリーの絞り込み）</span>
              <div class="tb-chips">${DocAssist.patternScenes
                .map(
                  (s) =>
                    `<label class="tb-chip"><input type="checkbox" name="scene" value="${escapeAttr(s)}"${
                      spec.scenes.indexOf(s) !== -1 ? ' checked' : ''
                    }><span>${escapeHtml(s)}</span></label>`
                )
                .join('')}</div></div>

            <div class="tb-section">レイアウト</div>
            <div class="tb-field"><span>種類</span>
              <div class="tb-radios">
                <label><input type="radio" name="kind" value="cards"${L.kind === 'cards' ? ' checked' : ''}> カード（1項目＝1枚）</label>
                <label><input type="radio" name="kind" value="table"${L.kind === 'table' ? ' checked' : ''}> 表（1項目＝1行）</label>
              </div></div>

            <div id="tbCards" class="tb-subform">
              <div class="tb-row">
                <label class="tb-field"><span>並べる向き</span>
                  <select name="direction">${optionsHtml([['row', '横に並べる'], ['column', '縦に並べる']], L.direction)}</select></label>
                <label class="tb-field"><span>図形</span>
                  <select name="shape">${optionsHtml([['rect', '四角'], ['roundRect', '角丸'], ['chevron', '矢羽根（横のみ）']], L.shape)}</select></label>
              </div>
              <div class="tb-row">
                <label class="tb-check"><input type="checkbox" name="headShow"${L.head && L.head.show ? ' checked' : ''}> 見出し帯をつける</label>
                <label class="tb-check"><input type="checkbox" name="bullet"${L.body && L.body.bullet ? ' checked' : ''}> 本文に■をつける</label>
              </div>
              <div class="tb-row">
                <label class="tb-field"><span>見出しに使うデータ</span>
                  <select name="headFrom">${optionsHtml(FROM_OPTIONS, L.head && L.head.from)}</select></label>
                <label class="tb-field"><span>見出しの色</span>
                  <select name="headFill">${optionsHtml(FILL_OPTIONS, L.head && L.head.fill)}</select></label>
              </div>
              <div class="tb-row">
                <label class="tb-field"><span>本文に使うデータ</span>
                  <select name="bodyFrom">${optionsHtml(FROM_OPTIONS, L.body && L.body.from)}</select></label>
                <label class="tb-field"><span>本文の塗り</span>
                  <select name="bodyFill">${optionsHtml(FILL_OPTIONS, L.body && L.body.fill)}</select></label>
              </div>
              <div class="tb-row">
                <label class="tb-check"><input type="checkbox" name="numbered"${L.numbered ? ' checked' : ''}> 連番をつける</label>
                <label class="tb-check"><input type="checkbox" name="icon"${L.icon ? ' checked' : ''}> アイコンを自動でつける</label>
                <label class="tb-check"><input type="checkbox" name="connector"${L.connector === 'arrow' ? ' checked' : ''}> 矢印でつなぐ</label>
              </div>
            </div>

            <div id="tbTable" class="tb-subform">
              <label class="tb-check"><input type="checkbox" name="zebra"${L.zebra !== false ? ' checked' : ''}> 1行おきに色を敷く</label>
              <div class="tb-field"><span>列の設定</span><div id="tbCols" class="tb-cols-editor"></div>
                <button type="button" class="btn ghost tb-addcol" id="tbAddCol">＋ 列を追加</button></div>
            </div>

            <div class="tb-section">自動選択とサンプル</div>
            <label class="tb-field"><span>自動選択のキーワード（読点区切り・空欄なら手動選択のみ）</span>
              <input type="text" name="keywords" value="${escapeAttr(spec.keywords.join('、'))}" placeholder="例：比較、検討、案"></label>
            <div class="tb-row">
              <label class="tb-field"><span>最小件数</span><input type="number" name="minItems" min="1" max="12" value="${spec.minItems}"></label>
              <label class="tb-field"><span>最大件数</span><input type="number" name="maxItems" min="1" max="12" value="${spec.maxItems}"></label>
            </div>
            <label class="tb-field"><span>プレビュー用サンプル（1行に1項目）</span>
              <textarea name="sample" rows="4">${escapeHtml(spec.sample.join('\n'))}</textarea></label>
          </div>

          <div class="tb-preview">
            <div class="tb-preview-label">プレビュー</div>
            <div class="slide-canvas">
              <div class="slide-title-bar"><span class="bar"></span><span class="heading-text">見出しが入ります</span></div>
              <div class="slide-message-block">
                <div class="slide-message-lead">このスライドの結論が1文で入ります</div>
              </div>
              <div class="slide-body" id="tbBody"></div>
              <div class="slide-footer"><span class="source-note"></span><span class="page-num">1</span></div>
            </div>
            <p class="hint" id="tbWarn"></p>
          </div>
        </div>
        <div class="btn-row">
          <span class="status-msg" id="tbStatus"></span>
          <div style="display:flex;gap:10px;">
            ${editing ? '<button type="button" class="btn ghost" id="tbDelete">削除</button>' : ''}
            <button type="button" class="btn ghost" id="tbCancel">キャンセル</button>
            <button type="button" class="btn" id="tbSave">${editing ? '上書き保存' : '保存して追加'}</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const form = overlay.querySelector('#tbForm');
    const colsEditor = overlay.querySelector('#tbCols');
    let columns = (L.columns || [{ label: '項目', from: 'key', width: 0.3 }, { label: '内容', from: 'value', width: 0.7 }]).map((c) =>
      Object.assign({}, c)
    );

    function renderColumns() {
      colsEditor.innerHTML = columns
        .map(
          (c, i) => `
        <div class="tb-col-row" data-i="${i}">
          <input type="text" class="tb-col-label" value="${escapeAttr(c.label)}" placeholder="列見出し">
          <select class="tb-col-from">${optionsHtml(FROM_OPTIONS, c.from)}</select>
          <select class="tb-col-badge">${optionsHtml(
            [['', 'バッジなし'], ['status', '状況バッジ'], ['severity', '影響度バッジ']],
            c.badge || ''
          )}</select>
          <input type="number" class="tb-col-width" min="1" max="100" value="${Math.round((c.width || 0.2) * 100)}" title="幅の比率">
          <button type="button" class="tpl-mini danger tb-col-del"${columns.length <= 1 ? ' disabled' : ''}>✕</button>
        </div>`
        )
        .join('');
    }

    function readColumns() {
      return Array.from(colsEditor.querySelectorAll('.tb-col-row')).map((row) => ({
        label: row.querySelector('.tb-col-label').value,
        from: row.querySelector('.tb-col-from').value,
        badge: row.querySelector('.tb-col-badge').value,
        width: Math.max(1, parseInt(row.querySelector('.tb-col-width').value, 10) || 20) / 100,
      }));
    }

    function val(name) {
      const el = form.querySelector(`[name="${name}"]`);
      if (!el) return '';
      return el.type === 'checkbox' ? el.checked : el.value;
    }

    function readForm() {
      const kind = form.querySelector('[name="kind"]:checked').value;
      const sample = String(val('sample'))
        .split('\n')
        .map((s) => s.trim())
        .filter(Boolean);
      const layout =
        kind === 'table'
          ? { kind: 'table', zebra: !!val('zebra'), columns: readColumns() }
          : {
              kind: 'cards',
              direction: val('direction'),
              shape: val('shape'),
              numbered: !!val('numbered'),
              icon: !!val('icon'),
              connector: val('connector') ? 'arrow' : 'none',
              head: { show: !!val('headShow'), from: val('headFrom'), fill: val('headFill') },
              body: { from: val('bodyFrom'), fill: val('bodyFill'), bullet: !!val('bullet') },
            };
      return {
        id: spec.id && editing ? spec.id : undefined,
        name: String(val('name')).trim(),
        description: String(val('description')).trim(),
        category: val('category'),
        role: val('role'),
        scenes: Array.from(form.querySelectorAll('[name="scene"]:checked')).map((c) => c.value),
        keywords: String(val('keywords'))
          .split(/[、,]/)
          .map((s) => s.trim())
          .filter(Boolean),
        minItems: parseInt(val('minItems'), 10) || 1,
        maxItems: parseInt(val('maxItems'), 10) || 6,
        sample: sample.length ? sample : undefined,
        layout,
      };
    }

    function syncSubforms() {
      const kind = form.querySelector('[name="kind"]:checked').value;
      overlay.querySelector('#tbCards').style.display = kind === 'cards' ? '' : 'none';
      overlay.querySelector('#tbTable').style.display = kind === 'table' ? '' : 'none';
      // 矢羽根は横並び専用なので、選ばれたら向きの選択を固定する
      const dirSel = form.querySelector('[name="direction"]');
      const isChevron = val('shape') === 'chevron';
      if (dirSel) {
        dirSel.disabled = isChevron;
        if (isChevron) dirSel.value = 'row';
      }
    }

    function refreshPreview() {
      syncSubforms();
      const draft = readForm();
      const pattern = DocAssist.templateSpec.createPatternFromSpec(draft);
      const body = overlay.querySelector('#tbBody');
      try {
        body.innerHTML = pattern.renderBody(pattern.sample);
      } catch (e) {
        body.innerHTML = '<div class="pv-empty">プレビューを描画できませんでした</div>';
      }
      const warn = overlay.querySelector('#tbWarn');
      const n = pattern.sample.length;
      const notes = [];
      if (n > draft.maxItems) notes.push(`サンプルが${n}件ありますが、最大件数${draft.maxItems}件までしか表示されません。`);
      if (!draft.keywords.length) notes.push('キーワードが未設定のため、このテンプレートは自動選択の候補になりません（ギャラリーから手動で選べます）。');
      warn.textContent = notes.join(' ');
    }

    renderColumns();
    refreshPreview();

    form.addEventListener('input', refreshPreview);
    form.addEventListener('change', refreshPreview);
    colsEditor.addEventListener('input', refreshPreview);
    colsEditor.addEventListener('change', refreshPreview);
    colsEditor.addEventListener('click', (e) => {
      const del = e.target.closest('.tb-col-del');
      if (!del) return;
      const i = parseInt(del.closest('.tb-col-row').dataset.i, 10);
      columns = readColumns();
      columns.splice(i, 1);
      renderColumns();
      refreshPreview();
    });
    overlay.querySelector('#tbAddCol').addEventListener('click', () => {
      columns = readColumns();
      columns.push({ label: '列' + (columns.length + 1), from: 'value', width: 0.2 });
      renderColumns();
      refreshPreview();
    });

    function closeBuilder() {
      document.removeEventListener('keydown', onBuilderKey);
      if (overlay.parentNode) document.body.removeChild(overlay);
    }
    function onBuilderKey(e) {
      if (e.key === 'Escape') closeBuilder();
    }
    document.addEventListener('keydown', onBuilderKey);
    overlay.querySelector('#tbCancel').addEventListener('click', closeBuilder);

    overlay.querySelector('#tbSave').addEventListener('click', () => {
      const draft = readForm();
      const statusEl = overlay.querySelector('#tbStatus');
      if (!draft.name) {
        statusEl.textContent = 'テンプレート名を入力してください。';
        statusEl.classList.add('error');
        return;
      }
      const saved = DocAssist.customTemplates.saveSpec(draft);
      closeBuilder();
      if (onSaved) onSaved(saved);
    });

    const delBtn = overlay.querySelector('#tbDelete');
    if (delBtn) {
      delBtn.addEventListener('click', () => {
        if (!window.confirm(`テンプレート「${spec.name}」を削除します。よろしいですか？`)) return;
        DocAssist.customTemplates.deleteSpec(spec.id);
        state.outline.slides.forEach((s) => {
          if (s.patternId === spec.id) {
            s.patternId = 'title-message';
            s.patternSource = 'auto';
          }
        });
        closeBuilder();
        if (onSaved) onSaved(null);
      });
    }
  }

  function renderSlidePreviewCard(slide, idx) {
    const wrap = document.createElement('div');
    wrap.className = 'slide-preview-card';

    const pattern = DocAssist.patternById[slide.patternId] || DocAssist.patternById['title-message'];
    const reasonAttr = slide.patternReason ? ` title="${escapeAttr(slide.patternReason)}"` : '';
    const message = DocAssist.analyze.effectiveMessage(slide);
    const subLines = (slide.subMessage || '').split('\n').map((s) => s.trim()).filter(Boolean);

    const bulletCountForControl = (slide.bullets || []).length;
    wrap.innerHTML = `
      <div class="slide-preview-toolbar">
        <span class="label"${reasonAttr}>${sourceLabel(slide.patternSource)}</span>
        <button type="button" class="tpl-open-btn">
          <span class="tpl-open-name">${escapeHtml(pattern.name)}</span>
          <span class="tpl-open-hint">テンプレートを選ぶ</span>
        </button>
      </div>
      <div class="item-count-control">
        <span class="item-count-label">項目数</span>
        <div class="item-count-quick">
          ${[2, 3, 4, 5].map((n) => `<button type="button" class="item-count-quick-btn${bulletCountForControl === n ? ' is-active' : ''}" data-count="${n}">${n}</button>`).join('')}
        </div>
        <div class="item-count-stepper">
          <button type="button" class="item-count-step-btn" data-act="dec" title="項目を1つ減らす">－</button>
          <span class="item-count-value">${bulletCountForControl}</span>
          <button type="button" class="item-count-step-btn" data-act="inc" title="項目を1つ増やす">＋</button>
        </div>
      </div>
      <div class="slide-canvas">
        <div class="slide-title-bar">
          <span class="bar"></span>
          <span class="heading-text">${escapeHtml(slide.heading || '(見出し未設定)')}</span>
        </div>
        <div class="slide-message-block">
          <div class="slide-message-lead">${emphasisHtml(message)}</div>
          ${subLines.length ? `<div class="slide-message-sub"><ul class="pv-bullets pv-bullets-square">${subLines.map((l) => `<li>${emphasisHtml(l)}</li>`).join('')}</ul></div>` : ''}
        </div>
        <div class="slide-body"></div>
        <div class="slide-footer">
          <span class="source-note">${slide.sourceNote ? '出所：' + escapeHtml(slide.sourceNote) : ''}</span>
          <span class="page-num">${idx + 1}</span>
        </div>
      </div>
    `;

    const bullets = slide.bullets || [];
    let bodyHtml = pattern.renderBody(bullets);
    // テンプレートギャラリーから手動で切り替えた場合など、箇条書きの形がそのパターンの
    // 想定と合わず何も描画できなかった（pv-emptyのプレースホルダーになった）ケースの
    // 安全網。内容自体はあるので、汎用フォールバック（タイトル＋メッセージ）で
    // 最低限の表示を保証し、書き出し（js/pptxExport.js）と同じ理由づけのヒントを添える。
    let fellBack = false;
    if (bullets.length && /class="pv-empty"/.test(bodyHtml)) {
      bodyHtml = DocAssist.patternById['title-message'].renderBody(bullets);
      fellBack = true;
    }
    wrap.querySelector('.slide-body').innerHTML = bodyHtml;
    if (fellBack) {
      const hint = document.createElement('p');
      hint.className = 'pattern-fallback-hint';
      hint.textContent = '⚠ このパターンの形式に箇条書きが合わないため、簡易表示（タイトル＋メッセージ）にしています。内容を編集するか、別のパターンを選んでください。';
      wrap.querySelector('.slide-canvas').insertAdjacentElement('afterend', hint);
    }
    wrap.querySelector('.tpl-open-btn').addEventListener('click', () => openTemplatePicker(slide));
    wrap.querySelectorAll('.item-count-quick-btn').forEach((btn) => {
      btn.addEventListener('click', () => setBulletCount(slide, Number(btn.dataset.count)));
    });
    wrap.querySelector('[data-act="dec"]').addEventListener('click', () => setBulletCount(slide, bullets.length - 1));
    wrap.querySelector('[data-act="inc"]').addEventListener('click', () => setBulletCount(slide, bullets.length + 1));

    return wrap;
  }

  // ---------------- utils ----------------
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }
  function escapeAttr(s) {
    return escapeHtml(s);
  }
  function emphasisHtml(s) {
    return DocAssist.analyze
      .parseEmphasisTokens(s)
      .map((t) => (t.bold ? `<b class="em">${escapeHtml(t.text)}</b>` : escapeHtml(t.text)))
      .join('');
  }

  // ---------------- デザインテンプレートの取り込み ----------------
  // 既存のPowerPointファイルをテンプレートとして読み込み、js/pptxImport.js が抽出した
  // 配色・フォント・箇条書き記号を DocAssist.theme に反映する。HTMLプレビューの色は
  // すべてCSS変数（var(--primary)等）経由で描かれているため、適用後に個々のカードを
  // 再描画しなくても見た目は自動的に切り替わる（renderThemeBar()だけ更新すればよい）。
  function renderThemeBar() {
    const bar = document.getElementById('themeBar');
    if (!bar) return;
    const saved = DocAssist.pptxImport.getSaved();
    const sampleBtnHtml = `<button type="button" class="btn ghost theme-bar-btn" id="sampleExportBtn">📦 全レイアウトをサンプル出力</button>
         <span class="status-msg theme-bar-status" id="sampleExportStatus"></span>`;
    bar.innerHTML = saved
      ? `<span class="theme-bar-swatch" style="background:#${escapeAttr(saved.primary)};"></span>
         <span class="theme-bar-label">テンプレート適用中：${escapeHtml(saved.sourceFileName)}</span>
         <button type="button" class="btn ghost theme-bar-btn" id="themeChangeBtn">変更</button>
         <button type="button" class="btn ghost theme-bar-btn" id="themeResetBtn">既定に戻す</button>
         ${sampleBtnHtml}`
      : `<span class="theme-bar-label">配色：官公庁報告書スタイル（既定）</span>
         <button type="button" class="btn ghost theme-bar-btn" id="themeChangeBtn">🎨 テンプレートから取り込む</button>
         ${sampleBtnHtml}`;
    document.getElementById('themeChangeBtn').addEventListener('click', openTemplateThemeModal);
    const resetBtn = document.getElementById('themeResetBtn');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        DocAssist.pptxImport.resetTheme();
        renderThemeBar();
      });
    }
    document.getElementById('sampleExportBtn').addEventListener('click', async () => {
      const statusEl = document.getElementById('sampleExportStatus');
      statusEl.textContent = '生成中…';
      statusEl.classList.remove('error');
      try {
        await DocAssist.exportSamplePptx();
        statusEl.textContent = `全${DocAssist.patterns.length}パターンのサンプルを書き出しました。`;
      } catch (e) {
        console.error(e);
        statusEl.textContent = '書き出しに失敗しました: ' + e.message;
        statusEl.classList.add('error');
      }
    });
  }

  function openTemplateThemeModal() {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal-box tt-modal">
        <h3>🎨 テンプレートから配色を取り込む</h3>
        <p class="hint">お手元のPowerPointファイル（.pptx）を選ぶと、スライドマスターのテーマ色・本文フォント・箇条書き記号を読み取ります。抽出した内容を確認してから適用できます。</p>
        <p class="hint">※ このアプリの4層構成（ヘッダー／メッセージ／ボディ／フッター）自体は変わりません。取り込むのは色・フォント・箇条書き記号のみです。テーマの副次的なアクセント色（赤・オレンジ等を含みうる）はそのまま使わず、ブランドカラー1色から青系統と同じ考え方で階調を作り直します。</p>
        <input type="file" id="ttFile" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation">
        <div id="ttResult"></div>
        <div class="btn-row">
          <span class="status-msg" id="ttStatus"></span>
          <div style="display:flex;gap:10px;">
            <button type="button" class="btn ghost" id="ttClose">閉じる</button>
            <button type="button" class="btn" id="ttApply" disabled>適用する</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    function close() {
      if (overlay.parentNode) document.body.removeChild(overlay);
    }
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    overlay.querySelector('#ttClose').addEventListener('click', close);

    let parsed = null;
    const resultEl = overlay.querySelector('#ttResult');
    const statusEl = overlay.querySelector('#ttStatus');
    const applyBtn = overlay.querySelector('#ttApply');

    overlay.querySelector('#ttFile').addEventListener('change', async (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      parsed = null;
      applyBtn.disabled = true;
      resultEl.innerHTML = '';
      statusEl.textContent = '読み込み中…';
      statusEl.classList.remove('error');
      try {
        parsed = await DocAssist.pptxImport.parseFile(file);
        const scale = [0, 0.25, 0.42, 0.58, 0.7, 0.82].map((pct) => DocAssist.pptxImport.tint(parsed.primary, pct));
        resultEl.innerHTML = `
          <div class="tt-preview">
            <div class="tt-row">
              <span class="tt-row-label">抽出した配色</span>
              <div class="tt-swatches">${scale.map((hex) => `<span class="tt-swatch" style="background:#${hex};" title="#${hex}"></span>`).join('')}</div>
            </div>
            <div class="tt-row">
              <span class="tt-row-label">本文フォント</span>
              <span class="tt-value" style="font-family:'${escapeAttr(parsed.fontFace || 'Yu Gothic UI')}';">${escapeHtml(
          parsed.fontFace || 'Yu Gothic UI（フォント情報を検出できなかったため既定のまま）'
        )}</span>
            </div>
            <div class="tt-row">
              <span class="tt-row-label">箇条書き記号</span>
              <span class="tt-value tt-bullet">${escapeHtml(parsed.bulletChar || '■')}</span>
              ${!parsed.bulletChar ? '<span class="tt-note">（記号を検出できなかった、または装飾フォント依存のため既定の■を使用）</span>' : ''}
            </div>
            ${
              parsed.rawAccents.length
                ? `<div class="tt-row">
                     <span class="tt-row-label">参考：テーマの全アクセント色</span>
                     <div class="tt-swatches">${parsed.rawAccents.map((hex) => `<span class="tt-swatch" style="background:#${hex};" title="#${hex}"></span>`).join('')}</div>
                     <span class="tt-note">（本文には最初の1色のみ使用し、残りは適用しません）</span>
                   </div>`
                : ''
            }
          </div>`;
        statusEl.textContent = '内容を確認して「適用する」を押してください。';
        applyBtn.disabled = false;
      } catch (err) {
        statusEl.textContent = '読み込みに失敗しました：' + err.message;
        statusEl.classList.add('error');
      }
    });

    applyBtn.addEventListener('click', () => {
      if (!parsed) return;
      DocAssist.pptxImport.applyTheme(parsed);
      close();
      renderThemeBar();
    });
  }

  renderThemeBar();
  render();
})();
