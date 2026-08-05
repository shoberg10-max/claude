// 画面遷移とUI制御（3ステップのウィザード）
(function () {
  const SAMPLE_NOTES = `# 新規顧客管理システム導入に関する検討報告

## [pattern: pyramid] （１）現状の課題｜①課題認識と対応方針
> 顧客情報の分散管理が対応品質の低下と解約率上昇を招いており、**クラウド型CRM**による一元化が急務である
>> 現行体制のまま推移した場合、来期の解約率はさらに悪化する見込みである
- 課題：顧客対応の質と速度が低下し、顧客満足度調査のスコアが**前年比8ポイント低下**している
- 原因：顧客情報が担当者個人のPC・Excelに分散し、組織的な検索・共有ができていない
- 対策：クラウド型CRMを導入し、顧客情報と対応履歴を全社で一元管理する体制を構築する
出所：社内顧客満足度調査（2026年1月実施）

## [pattern: box-compare] （２）導入検討｜①候補ツールの比較検討
> 3社を比較した結果、機能とコストのバランスから**B社CRM**が最も現実的な選択肢である
- A社CRM：低コストだが分析機能が限定的で、将来的な拡張性に課題がある
- B社CRM：コストは中程度で、必要な機能を標準搭載しており拡張性も高い
- C社CRM：高コストだが分析機能が充実しており、大規模組織向けの構成である

## [pattern: timeline] （２）導入検討｜②導入スケジュール
> 4フェーズ・約4ヶ月の計画とし、全社展開前に試験運用で定着度を確認する
- フェーズ1（2026年9月）：要件定義・ベンダー選定・契約締結
- フェーズ2（2026年10月）：既存データ移行・初期設定・管理者研修
- フェーズ3（2026年11月）：一部部署での試験運用と課題の洗い出し
- フェーズ4（2026年12月）：全社展開および利用ルールの周知徹底

## [pattern: kpi-summary] （３）導入効果｜①定量効果の見込み
> 対応品質と業務効率の両面で定量的な改善効果が見込まれる
- 顧客対応リードタイム：50%短縮（平均対応時間 30分→15分）
- 解約率：20%改善（対応履歴の一元化により解約予兆を早期検知）
- 営業一人あたり商談数：15%増加（情報検索時間の削減により商談準備時間を確保）
出所：NRI「CRM導入効果調査」2024年版

## [pattern: before-after] （３）導入効果｜②情報検索にかかる時間の変化
> 情報検索の一元化により、担当者1人あたり年間換算で約50時間の業務時間削減が見込まれる
現状：情報を探すために複数の担当者に確認する必要があり、平均15分を要している
導入後：CRM上で即座に検索でき、平均2分に短縮される見込みである
`;

  const state = {
    step: 1,
    notesText: '',
    outline: { title: '', slides: [] },
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
        <p class="hint">構成案がうまく組み立てられない場合は、「🪄 社内LLM用プロンプトを作成」でこのアプリの形式に合わせて整形するためのプロンプトを作成できます。社内LLMに貼り付けて実行し、出力結果をこの下の欄に貼り付け直してから「次へ」を押してください。</p>
        <textarea class="notes-input" id="notesInput" placeholder="例：\n# 会議タイトル\n## 現状の課題\n- 課題A\n- 課題B\n## 比較：A案 vs B案\nA案：...\nB案：...">${escapeHtml(state.notesText)}</textarea>
        ${errorMsg ? `<p class="status-msg error">${escapeHtml(errorMsg)}</p>` : ''}
        <div class="btn-row">
          <button class="btn ghost" id="sampleBtn">サンプルを試す</button>
          <div style="display:flex;gap:10px;">
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
    document.getElementById('promptBtn').addEventListener('click', () => openPromptModal(input.value));
    document.getElementById('nextBtn').addEventListener('click', generateOutlineAndProceed);
  }

  function openPromptModal(rawNotes) {
    const prompt = DocAssist.buildOutlineFormatPrompt(rawNotes);
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal-box modal-box-wide">
        <h3>🪄 社内LLM用 構成案整形プロンプト</h3>
        <p class="hint">下のテキストをコピーして社内LLMのチャットに貼り付けて実行してください。出力された構成案をコピーし、Step 1の入力欄に貼り付け直してから「次へ」を押すと、このアプリの形式にきれいに読み込まれます（見出しに含まれる [pattern: ...] タグはデザインパターンの提案として自動的に反映されます）。</p>
        <textarea class="prompt-output" id="promptOutput" readonly>${escapeHtml(prompt)}</textarea>
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
    const textarea = overlay.querySelector('#promptOutput');
    overlay.querySelector('#promptCopy').addEventListener('click', async () => {
      const statusEl = overlay.querySelector('#copyStatus');
      try {
        await navigator.clipboard.writeText(prompt);
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
      <input type="text" class="sub-message" placeholder="サブメッセージ（リード文の補足。任意）" value="${escapeAttr(slide.subMessage || '')}">
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

  // 同じカテゴリ（用途）のパターンを <optgroup> でまとめ、兄弟デザインを
  // 見つけやすくする。同一カテゴリが複数登録されているほど選択肢が増える。
  function groupedPatternOptions(selectedId) {
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
      .map((cat) => {
        const opts = byCategory[cat]
          .map((p) => `<option value="${p.id}" ${p.id === selectedId ? 'selected' : ''}>${escapeHtml(p.name)}</option>`)
          .join('');
        return `<optgroup label="${escapeAttr(cat)}">${opts}</optgroup>`;
      })
      .join('');
  }

  function renderSlidePreviewCard(slide, idx) {
    const wrap = document.createElement('div');
    wrap.className = 'slide-preview-card';

    const options = groupedPatternOptions(slide.patternId);
    const reasonAttr = slide.patternReason ? ` title="${escapeAttr(slide.patternReason)}"` : '';
    const message = DocAssist.analyze.effectiveMessage(slide);
    const subMessage = (slide.subMessage || '').trim();

    wrap.innerHTML = `
      <div class="slide-preview-toolbar">
        <span class="label"${reasonAttr}>${sourceLabel(slide.patternSource)}</span>
        <select class="pattern-select">${options}</select>
      </div>
      <div class="slide-canvas">
        <div class="slide-title-bar">
          <span class="bar"></span>
          <span class="heading-text">${escapeHtml(slide.heading || '(見出し未設定)')}</span>
        </div>
        <div class="slide-message-block">
          <div class="slide-message-lead">${emphasisHtml(message)}</div>
          ${subMessage ? `<div class="slide-message-sub">${emphasisHtml(subMessage)}</div>` : ''}
        </div>
        <div class="slide-body"></div>
        <div class="slide-footer">
          <span class="source-note">${slide.sourceNote ? '出所：' + escapeHtml(slide.sourceNote) : ''}</span>
          <span class="page-num">${idx + 1}</span>
        </div>
      </div>
    `;

    const body = wrap.querySelector('.slide-body');
    const pattern = DocAssist.patternById[slide.patternId] || DocAssist.patternById['title-message'];
    body.innerHTML = pattern.renderBody(slide.bullets || []);

    wrap.querySelector('.pattern-select').addEventListener('change', (e) => {
      slide.patternId = e.target.value;
      slide.patternSource = 'manual';
      slide.patternReason = '';
      replaceCard(slide);
    });

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

  render();
})();
