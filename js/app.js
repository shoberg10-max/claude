// 画面遷移とUI制御（3ステップのウィザード）
(function () {
  const SAMPLE_NOTES = `# 新規顧客管理システム導入 次回会議資料

## 現状の課題
- 顧客情報がExcelで属人的に管理されている
- 対応履歴の共有が遅く、二重対応が発生している
- 解約予兆の把握が遅れ、解約率が上昇している

## 課題の原因と対策
課題：顧客対応の質と速度が低下している
原因：情報が個人PCに分散し検索性が低い
対策：クラウド型CRMを導入し情報を一元化する

## 導入候補ツールの比較
A社CRM：低コストだが機能が限定的
B社CRM：コストは中程度で機能も標準的
C社CRM：高コストだが分析機能が充実

## 導入スケジュール
- フェーズ1（2026年9月）：要件定義・ベンダー選定
- フェーズ2（2026年10月）：データ移行・設定
- フェーズ3（2026年11月）：試験運用
- フェーズ4（2026年12月）：全社展開

## 導入後の主要KPI
- 顧客対応リードタイム：50%短縮
- 解約率：20%改善
- 営業一人あたり商談数：15%増加

## 現状と導入後の比較
現状：情報検索に平均15分かかっている
導入後：情報検索が平均2分に短縮される
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

  // patternSource: 'rule'（自動選択） / 'ai'（AI判定） / 'manual'（手動選択）
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
        <p class="hint">次回会議に向けた資料の元になるメモや議事録を貼り付けてください。見出し（# や ##）や箇条書き（- ・）があると構成案の精度が上がりますが、フリーテキストでも解析します。他のLLMで先に作成した構成案テキストをそのまま貼り付けても構いません。</p>
        <textarea class="notes-input" id="notesInput" placeholder="例：\n# 会議タイトル\n## 現状の課題\n- 課題A\n- 課題B\n## 比較：A案 vs B案\nA案：...\nB案：...">${escapeHtml(state.notesText)}</textarea>
        ${errorMsg ? `<p class="status-msg error">${escapeHtml(errorMsg)}</p>` : ''}
        <div class="btn-row">
          <button class="btn ghost" id="sampleBtn">サンプルを試す</button>
          <button class="btn" id="nextBtn">次へ：構成案を生成</button>
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
    document.getElementById('nextBtn').addEventListener('click', generateOutlineAndProceed);
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
      o.slides.push({ id: uid(), heading: '新しいスライド', bullets: [] });
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
        <input type="text" class="heading" value="${escapeAttr(slide.heading)}">
        <button class="icon-btn" data-act="up" title="上へ" ${idx === 0 ? 'disabled' : ''}>↑</button>
        <button class="icon-btn" data-act="down" title="下へ" ${idx === total - 1 ? 'disabled' : ''}>↓</button>
        <button class="icon-btn danger" data-act="del" title="削除">✕</button>
      </div>
      <textarea class="bullets" placeholder="箇条書きを1行ずつ入力">${escapeHtml((slide.bullets || []).join('\n'))}</textarea>
    `;
    card.querySelector('.heading').addEventListener('input', (e) => {
      slide.heading = e.target.value;
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
    const cfg = DocAssist.loadLlmConfig();
    const aiReady = !!cfg.endpoint;
    screen.innerHTML = `
      <div class="panel">
        <div class="step3-head">
          <div>
            <h2>Step 3. デザインパターンを自動選択 → PowerPoint書き出し</h2>
            <p class="hint">各スライドの内容から、あらかじめ用意したコンサルスライドのデザインパターンを自動選択しました（ルールベース）。プレビュー右上のプルダウンで手動変更するか、社内LLM等に接続してAIに判定させることもできます。</p>
          </div>
          <div class="step3-actions">
            <button class="btn ghost" id="settingsBtn">⚙ AI設定</button>
            <button class="btn secondary" id="aiAllBtn" ${aiReady ? '' : 'disabled'} title="${aiReady ? 'すべてのスライドをAIで再判定します' : '先に「⚙ AI設定」でエンドポイントを設定してください'}">🤖 AIで一括判定</button>
          </div>
        </div>
        <details>
          <summary style="cursor:pointer;color:var(--accent);font-size:13px;">利用可能なデザインパターン一覧（${DocAssist.patterns.length}種類）</summary>
          <div class="pattern-legend">
            ${DocAssist.patterns.map((p) => `<div class="item"><b>${escapeHtml(p.name)}</b>${escapeHtml(p.description)}</div>`).join('')}
          </div>
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
    o.slides.forEach((slide) => {
      grid.appendChild(renderSlidePreviewCard(slide));
    });

    document.getElementById('backBtn').addEventListener('click', () => goToStep(2));
    document.getElementById('settingsBtn').addEventListener('click', openSettingsModal);
    document.getElementById('aiAllBtn').addEventListener('click', runAiSelectionForAll);
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
      grid.replaceChild(renderSlidePreviewCard(slide), grid.children[idx]);
    }
  }

  async function runAiSelectionForAll() {
    const btn = document.getElementById('aiAllBtn');
    const o = state.outline;
    const total = o.slides.length;
    const originalLabel = btn.textContent;
    btn.disabled = true;
    let errors = 0;
    for (let i = 0; i < total; i++) {
      btn.textContent = `判定中…(${i + 1}/${total})`;
      const slide = o.slides[i];
      try {
        const result = await DocAssist.selectPatternWithLLM(slide, DocAssist.patterns);
        slide.patternId = result.pattern.id;
        slide.patternSource = 'ai';
        slide.patternReason = result.reason;
        slide.lastAiError = '';
      } catch (e) {
        console.error('AIパターン判定エラー:', slide.heading, e);
        slide.lastAiError = e.message;
        errors++;
      }
      replaceCard(slide);
    }
    btn.disabled = false;
    btn.textContent = originalLabel;
    const statusEl = document.getElementById('exportStatus');
    if (statusEl) {
      if (errors) {
        statusEl.textContent = `AI判定: ${total - errors}/${total}件成功（${errors}件は失敗のため既存の選択のままです。詳細は各カードを確認してください）`;
        statusEl.classList.add('error');
      } else {
        statusEl.textContent = `AI判定が完了しました（${total}件）`;
        statusEl.classList.remove('error');
      }
    }
  }

  function openSettingsModal() {
    const cfg = DocAssist.loadLlmConfig();
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal-box">
        <h3>⚙ AI設定（パターン自動判定用のLLM接続先）</h3>
        <p class="hint">社内の隔離LLM環境など、OpenAI互換またはAnthropic互換のAPIエンドポイントを設定すると、Step 3で「AIによるパターン判定」が使えるようになります。入力内容はこのブラウザのlocalStorageにのみ保存され、外部には送信されません。</p>
        <label>リクエスト形式
          <select id="cfgFormat">
            <option value="openai" ${cfg.format === 'openai' ? 'selected' : ''}>OpenAI互換（Chat Completions）</option>
            <option value="anthropic" ${cfg.format === 'anthropic' ? 'selected' : ''}>Anthropic互換（Messages API）</option>
          </select>
        </label>
        <label>エンドポイントURL
          <input type="text" id="cfgEndpoint" value="${escapeAttr(cfg.endpoint)}" placeholder="https://internal-llm.example.com/v1/chat/completions">
        </label>
        <label>モデル名
          <input type="text" id="cfgModel" value="${escapeAttr(cfg.model)}" placeholder="例：internal-gpt-4o">
        </label>
        <label>APIキー（任意・リクエストヘッダーに設定されます）
          <input type="password" id="cfgApiKey" value="${escapeAttr(cfg.apiKey)}">
        </label>
        <div class="btn-row">
          <button class="btn ghost" id="cfgCancel">キャンセル</button>
          <button class="btn" id="cfgSave">保存</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const close = () => document.body.removeChild(overlay);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    overlay.querySelector('#cfgCancel').addEventListener('click', close);
    overlay.querySelector('#cfgSave').addEventListener('click', () => {
      DocAssist.saveLlmConfig({
        format: overlay.querySelector('#cfgFormat').value,
        endpoint: overlay.querySelector('#cfgEndpoint').value.trim(),
        model: overlay.querySelector('#cfgModel').value.trim(),
        apiKey: overlay.querySelector('#cfgApiKey').value,
      });
      close();
      renderStep3();
    });
  }

  function renderSlidePreviewCard(slide) {
    const wrap = document.createElement('div');
    wrap.className = 'slide-preview-card';

    const options = DocAssist.patterns
      .map((p) => `<option value="${p.id}" ${p.id === slide.patternId ? 'selected' : ''}>${escapeHtml(p.name)}</option>`)
      .join('');
    const reasonAttr = slide.patternReason ? ` title="${escapeAttr(slide.patternReason)}"` : '';

    wrap.innerHTML = `
      <div class="slide-preview-toolbar">
        <span class="label"${reasonAttr}>${sourceLabel(slide.patternSource)}</span>
        <div class="toolbar-actions">
          <button class="icon-btn ai-btn" title="この1枚をAIで判定">🤖</button>
          <select class="pattern-select">${options}</select>
        </div>
      </div>
      ${slide.lastAiError ? `<div class="card-error">${escapeHtml(slide.lastAiError)}</div>` : ''}
      <div class="slide-canvas">
        <div class="slide-title-bar">${escapeHtml(slide.heading || '(見出し未設定)')}</div>
        <div class="slide-body"></div>
      </div>
    `;

    const body = wrap.querySelector('.slide-body');
    const pattern = DocAssist.patternById[slide.patternId] || DocAssist.patternById['title-message'];
    body.innerHTML = pattern.renderBody(slide.bullets || []);

    wrap.querySelector('.pattern-select').addEventListener('change', (e) => {
      slide.patternId = e.target.value;
      slide.patternSource = 'manual';
      slide.patternReason = '';
      slide.lastAiError = '';
      replaceCard(slide);
    });

    wrap.querySelector('.ai-btn').addEventListener('click', async (e) => {
      const btn = e.currentTarget;
      btn.disabled = true;
      btn.textContent = '…';
      try {
        const result = await DocAssist.selectPatternWithLLM(slide, DocAssist.patterns);
        slide.patternId = result.pattern.id;
        slide.patternSource = 'ai';
        slide.patternReason = result.reason;
        slide.lastAiError = '';
      } catch (err) {
        console.error('AIパターン判定エラー:', slide.heading, err);
        slide.lastAiError = err.message;
      }
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

  render();
})();
