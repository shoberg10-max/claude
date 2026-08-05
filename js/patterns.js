// デザインパターン（コンサルスライドのレイアウト）ライブラリ。
//
// 新しいパターンを追加するときは、このファイルの末尾で
// DocAssist.patterns 配列に以下の形のオブジェクトを追加するだけでよい。
// UI（プレビュー・手動切り替えドロップダウン）と PPTX 書き出しの両方に
// 自動的に反映される。
//
//   {
//     id: 'kebab-case-id',           // 一意なID
//     name: '表示名',
//     category: '比較',              // 手動切り替えプルダウンのグループ名（同カテゴリ＝デザインの兄弟候補）
//     description: '説明文（プレビュー画面に表示）',
//     score(section) -> number,      // {heading, bullets} を見て 0〜10 程度で採点
//     renderBody(bullets) -> HTML文字列,   // プレビュー用（本文のみ。見出しは共通枠が描く）
//     buildBody(pptxSlide, bullets, theme, box) -> void,  // PPTX書き出し用
//   }
//
// 既存カテゴリ: 汎用／比較／マトリクス／ポジショニング／フロー・プロセス／
//              構造・ロジック／変化・対比／時系列／数値・グラフ
// 同じカテゴリに複数のパターンを登録しておくと、手動切り替えプルダウンで
// 「同じ用途の別デザイン」としてまとめて選べるようになる。
window.DocAssist = window.DocAssist || {};

(function () {
  const A = DocAssist.analyze;

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }

  function emptyBody(message) {
    return `<div class="pv-empty">${esc(message || '内容がありません')}</div>`;
  }

  // 「**強調**」構文を含む文字列を、プレビュー用HTML（<b class="em">）に変換する。
  function emphasisHtml(text) {
    return A.parseEmphasisTokens(text)
      .map((t) => (t.bold ? `<b class="em">${esc(t.text)}</b>` : esc(t.text)))
      .join('');
  }

  // ■ を箇条書きマーカーに使う（NRI報告書の本文で多用される記号）。
  const SQUARE_BULLET = { code: '25A0', indent: 14 };

  // 箇条書き配列（**強調**構文を含みうる）を、PptxGenJSのtext run配列に変換する。
  // 強調部分は太字＋ネイビーになる。複数のパターンから共通で使う。
  function bulletTextRuns(bullets, theme, opts) {
    const fontSize = (opts && opts.fontSize) || 14;
    const spaceAfter = (opts && opts.spaceAfter) || 12;
    const bulletChar = (opts && opts.bulletChar) || SQUARE_BULLET;
    const runs = [];
    (bullets || []).forEach((b) => {
      const tokens = A.parseEmphasisTokens(b);
      tokens.forEach((t, ti) => {
        const isLast = ti === tokens.length - 1;
        const runOpts = { fontSize, bullet: bulletChar, color: t.bold ? theme.primary : theme.text };
        if (t.bold) runOpts.bold = true;
        if (isLast) {
          runOpts.breakLine = true;
          runOpts.paraSpaceAfter = spaceAfter;
        }
        runs.push({ text: t.text, options: runOpts });
      });
    });
    return runs;
  }

  // ---------- タイトル＋メッセージ（フォールバック） ----------
  // リード文はスライド上部の共通メッセージ帯（pptxExport.js）が既に表示しているので、
  // ここでは箇条書きをシンプルに並べるだけにする（メッセージを重複表示しない）。
  const titleMessage = {
    id: 'title-message',
    name: 'タイトル＋メッセージ',
    category: '汎用',
    description: '補足の箇条書きのみのシンプルな本文。どんな内容にも使える汎用パターン。',
    score(section) {
      const n = (section.bullets || []).length;
      return n <= 4 ? 3 : 2;
    },
    renderBody(bullets) {
      if (!bullets.length) return emptyBody();
      return `<ul class="pv-bullets pv-bullets-square">${bullets.map((b) => `<li>${emphasisHtml(b)}</li>`).join('')}</ul>`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      slide.addText(bulletTextRuns(bullets, theme), { x: box.x, y: box.y, w: box.w, h: box.h, valign: 'top' });
    },
  };

  // ---------- ボックス比較 ----------
  const boxCompare = {
    id: 'box-compare',
    name: 'ボックス比較',
    category: '比較',
    description: '2〜4個の項目を横並びのボックスで比較する。',
    score(section) {
      const bullets = section.bullets || [];
      const ratio = A.kvRatio(bullets);
      const n = bullets.length;
      if (n >= 2 && n <= 4 && ratio >= 0.5) return 8;
      if (n >= 2 && n <= 4) return 4;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return emptyBody();
      return `<div class="pv-box-row">${items
        .map(
          (it) => `
        <div class="pv-box">
          <div class="pv-box-head">${esc(it.key)}</div>
          <div class="pv-box-body">${esc(it.value)}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return;
      const n = items.length;
      const gap = 0.25;
      const boxW = (box.w - gap * (n - 1)) / n;
      items.forEach((it, i) => {
        const x = box.x + i * (boxW + gap);
        slide.addShape('rect', { x, y: box.y, w: boxW, h: 0.5, fill: { color: theme.primary } });
        slide.addText(it.key, { x, y: box.y, w: boxW, h: 0.5, fontSize: 13, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        slide.addShape('rect', { x, y: box.y + 0.5, w: boxW, h: box.h - 0.5, fill: { color: theme.light }, line: { type: 'none' } });
        slide.addText(
          [{ text: it.value, options: { bullet: SQUARE_BULLET } }],
          { x: x + 0.15, y: box.y + 0.62, w: boxW - 0.28, h: box.h - 0.74, fontSize: 11.5, color: theme.text, valign: 'top' }
        );
      });
    },
  };

  // ---------- 2軸マトリクス ----------
  const matrix2x2 = {
    id: 'matrix-2x2',
    name: '2軸マトリクス',
    category: 'マトリクス／ポジショニング',
    description: '4象限のマトリクスで整理する（4項目向け）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['軸', 'マトリクス', '象限']);
      if (bullets.length === 4 && kw) return 9;
      if (bullets.length === 4) return 5;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return emptyBody();
      return `<div class="pv-matrix">${items
        .map(
          (it, i) => `
        <div class="pv-matrix-cell pv-matrix-${i}">
          <div class="pv-matrix-key">${esc(it.key)}</div>
          <div class="pv-matrix-value">${esc(it.value)}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return;
      const gap = 0.2;
      const cw = (box.w - gap) / 2;
      const ch = (box.h - gap) / 2;
      const positions = [
        [box.x, box.y],
        [box.x + cw + gap, box.y],
        [box.x, box.y + ch + gap],
        [box.x + cw + gap, box.y + ch + gap],
      ];
      const colors = [theme.accent, theme.primary, theme.primary, theme.accent];
      items.forEach((it, i) => {
        const [x, y] = positions[i];
        slide.addShape('rect', { x, y, w: cw, h: ch, fill: { color: colors[i] }, line: { color: theme.white, width: 2 } });
        slide.addText(
          [
            { text: it.key + '\n', options: { bold: true, fontSize: 13 } },
            { text: it.value, options: { fontSize: 11 } },
          ],
          { x: x + 0.15, y: y + 0.15, w: cw - 0.3, h: ch - 0.3, color: theme.white, valign: 'top' }
        );
      });
    },
  };

  // ---------- プロセス／フロー ----------
  const processFlow = {
    id: 'process-flow',
    name: 'プロセス／フロー',
    category: 'フロー・プロセス',
    description: '手順やステップを左から右へ順番に示す。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ステップ', 'フロー', '手順', 'プロセス', 'フェーズ']);
      const seq = A.sequenceWordCount(text) + A.numberedBulletCount(bullets);
      const n = bullets.length;
      if (n >= 3 && n <= 6 && kw && seq > 0) return 9;
      if (n >= 3 && n <= 6 && (kw || seq > 0)) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return emptyBody();
      return `<div class="pv-flow">${items
        .map((it, i) => {
          const label = it.value === it.key ? it.key : `${it.key}：${it.value}`;
          return `
        <div class="pv-flow-step">
          <div class="pv-flow-num">${i + 1}</div>
          <div class="pv-flow-label">${esc(label)}</div>
        </div>
        ${i < items.length - 1 ? '<div class="pv-flow-arrow">→</div>' : ''}`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return;
      const n = items.length;
      const arrowW = 0.35;
      const stepW = (box.w - arrowW * (n - 1)) / n;
      items.forEach((it, i) => {
        const x = box.x + i * (stepW + arrowW);
        const label = it.value === it.key ? it.key : `${it.key}\n${it.value}`;
        slide.addShape('rect', { x, y: box.y + 0.6, w: stepW, h: box.h - 0.6, fill: { color: theme.accent }, line: { type: 'none' } });
        slide.addShape('oval', { x: x + stepW / 2 - 0.25, y: box.y, w: 0.5, h: 0.5, fill: { color: theme.primary } });
        slide.addText(String(i + 1), { x: x + stepW / 2 - 0.25, y: box.y, w: 0.5, h: 0.5, fontSize: 14, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        slide.addText(label, { x: x + 0.1, y: box.y + 0.7, w: stepW - 0.2, h: box.h - 0.8, fontSize: 11, color: theme.white, align: 'center', valign: 'top' });
        if (i < n - 1) {
          slide.addText('→', { x: x + stepW, y: box.y + 0.6, w: arrowW, h: box.h - 0.6, fontSize: 18, bold: true, color: theme.primary, align: 'center', valign: 'middle' });
        }
      });
    },
  };

  // ---------- ピラミッド構造 ----------
  const pyramid = {
    id: 'pyramid',
    name: 'ピラミッド構造',
    category: '構造・ロジック',
    description: '結論（メッセージ）を頂点に、根拠・要点を土台として示す。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.countAny(text, ['課題', '原因', '対策', '解決策', '結論', '要因']);
      const n = (section.bullets || []).length;
      if (kw >= 2 && n >= 2) return 9;
      if (kw >= 1 && n >= 2) return 6;
      return 0;
    },
    renderBody(bullets) {
      if (!bullets.length) return emptyBody();
      const [apex, ...base] = bullets;
      const baseItems = base.length ? base : [apex];
      const apexText = base.length ? apex : '結論';
      return `
        <div class="pv-pyramid">
          <div class="pv-pyramid-apex">${emphasisHtml(apexText)}</div>
          <div class="pv-pyramid-base">${baseItems
            .slice(0, 4)
            .map((b) => `<div class="pv-pyramid-cell">${emphasisHtml(b)}</div>`)
            .join('')}</div>
        </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      const [apex, ...base] = bullets;
      const baseItems = (base.length ? base : [apex]).slice(0, 4);
      const apexText = base.length ? apex : '結論';
      const apexW = box.w * 0.55;
      slide.addShape('rect', { x: box.x + (box.w - apexW) / 2, y: box.y, w: apexW, h: 0.9, fill: { color: theme.primary } });
      slide.addText(A.emphasisRuns(apexText, null, { bold: true }), { x: box.x + (box.w - apexW) / 2 + 0.15, y: box.y, w: apexW - 0.3, h: 0.9, fontSize: 15, color: theme.white, valign: 'middle' });
      const n = baseItems.length;
      const gap = 0.2;
      const cw = (box.w - gap * (n - 1)) / n;
      const y2 = box.y + 1.3;
      baseItems.forEach((b, i) => {
        const x = box.x + i * (cw + gap);
        slide.addShape('rect', { x, y: y2, w: cw, h: box.h - 1.3, fill: { color: theme.light }, line: { type: 'none' } });
        slide.addText(A.emphasisRuns(b, theme.primary, { fontSize: 12 }), { x: x + 0.1, y: y2 + 0.1, w: cw - 0.2, h: box.h - 1.5, color: theme.text, valign: 'top' });
      });
    },
  };

  // ---------- Before / After ----------
  const beforeAfter = {
    id: 'before-after',
    name: 'Before／After',
    category: '変化・対比',
    description: '現状（Before）と施策後（After）を対比させる。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['Before', 'After', 'before', 'after', 'As-Is', 'To-Be', '現状', '改善後', '従来']);
      if (bullets.length === 2 && kw) return 9;
      if (bullets.length === 2) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 2);
      if (items.length < 2) return emptyBody();
      const labels = ['Before', 'After'];
      return `<div class="pv-ba">
        ${items
          .map((it, i) => {
            const label = it.key !== `項目${i + 1}` ? it.key : labels[i];
            return `
          <div class="pv-ba-box pv-ba-${i}">
            <div class="pv-ba-label">${esc(label)}</div>
            <div class="pv-ba-value">${esc(it.value)}</div>
          </div>
          ${i === 0 ? '<div class="pv-ba-arrow">→</div>' : ''}`;
          })
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 2);
      if (items.length < 2) return;
      const labels = ['Before', 'After'];
      const arrowW = 0.6;
      const boxW = (box.w - arrowW) / 2;
      items.forEach((it, i) => {
        const x = box.x + i * (boxW + arrowW);
        const label = it.key !== `項目${i + 1}` ? it.key : labels[i];
        slide.addShape('rect', { x, y: box.y, w: boxW, h: 0.6, fill: { color: i === 0 ? theme.subtext : theme.primaryLight } });
        slide.addText(label, { x, y: box.y, w: boxW, h: 0.6, fontSize: 14, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        slide.addShape('rect', { x, y: box.y + 0.6, w: boxW, h: box.h - 0.6, fill: { color: theme.light }, line: { color: theme.border, width: 1 } });
        slide.addText(it.value, { x: x + 0.15, y: box.y + 0.75, w: boxW - 0.3, h: box.h - 0.9, fontSize: 13, color: theme.text, valign: 'top' });
      });
      slide.addText('→', { x: box.x + boxW, y: box.y + box.h / 2 - 0.4, w: arrowW, h: 0.8, fontSize: 22, bold: true, color: theme.primary, align: 'center', valign: 'middle' });
    },
  };

  // ---------- タイムライン／ロードマップ ----------
  const timeline = {
    id: 'timeline',
    name: 'タイムライン／ロードマップ',
    category: '時系列',
    description: '日付・フェーズを時系列に並べて示す。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['スケジュール', 'ロードマップ', 'タイムライン']);
      const dates = A.dateTokenCount(text);
      const n = bullets.length;
      if (n >= 2 && n <= 6 && kw && dates > 0) return 9;
      if (n >= 2 && n <= 6 && dates >= n * 0.5) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return emptyBody();
      return `<div class="pv-timeline">
        <div class="pv-timeline-line"></div>
        ${items
          .map(
            (it) => `
          <div class="pv-timeline-item">
            <div class="pv-timeline-dot"></div>
            <div class="pv-timeline-label"><b>${esc(it.key)}</b>${it.value !== it.key ? '<br>' + esc(it.value) : ''}</div>
          </div>`
          )
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return;
      const n = items.length;
      const lineY = box.y + 0.4;
      slide.addShape('line', { x: box.x, y: lineY, w: box.w, h: 0, line: { color: theme.accent, width: 3 } });
      const stepW = box.w / n;
      items.forEach((it, i) => {
        const cx = box.x + stepW * i + stepW / 2;
        slide.addShape('oval', { x: cx - 0.12, y: lineY - 0.12, w: 0.24, h: 0.24, fill: { color: theme.primary } });
        const label = it.value !== it.key ? `${it.key}\n${it.value}` : it.key;
        slide.addText(label, { x: cx - stepW / 2 + 0.1, y: lineY + 0.25, w: stepW - 0.2, h: box.h - 0.7, fontSize: 11, color: theme.text, align: 'center', valign: 'top' });
      });
    },
  };

  // ---------- 比較表 ----------
  const comparisonTable = {
    id: 'comparison-table',
    name: '比較表',
    category: '比較',
    description: '5項目以上を一覧で比較する表形式。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const ratio = A.kvRatio(bullets);
      const kw = A.hasAny(text, ['比較表', '一覧']);
      if (bullets.length >= 5 && ratio >= 0.5) return kw ? 10 : 8;
      if (kw) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets);
      if (!items.length) return emptyBody();
      return `<table class="pv-table"><thead><tr><th>項目</th><th>内容</th></tr></thead><tbody>
        ${items.map((it) => `<tr><td>${esc(it.key)}</td><td>${esc(it.value)}</td></tr>`).join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets);
      if (!items.length) return;
      const rows = [
        [
          { text: '項目', options: { bold: true, color: theme.white, fill: { color: theme.primary } } },
          { text: '内容', options: { bold: true, color: theme.white, fill: { color: theme.primary } } },
        ],
        ...items.map((it, i) => [
          { text: it.key, options: { fill: { color: i % 2 ? theme.light : theme.white } } },
          { text: it.value, options: { fill: { color: i % 2 ? theme.light : theme.white } } },
        ]),
      ];
      slide.addTable(rows, {
        x: box.x,
        y: box.y,
        w: box.w,
        colW: [box.w * 0.3, box.w * 0.7],
        fontSize: 12,
        color: theme.text,
        border: { type: 'solid', color: theme.white, pt: 1.5 },
        valign: 'middle',
        autoPage: false,
      });
    },
  };

  // ---------- KPIサマリー ----------
  const kpiSummary = {
    id: 'kpi-summary',
    name: 'KPIサマリー',
    category: '数値・グラフ',
    description: '数値・指標をカード形式で並べて見せる。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['KPI', '指標', '実績']);
      const numHits = bullets.filter((b) => A.numberTokenCount(b) > 0).length;
      if (numHits >= 3 && numHits >= bullets.length * 0.6) return kw ? 10 : 8;
      if (kw && numHits > 0) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return emptyBody();
      return `<div class="pv-kpi-grid">${items
        .map((it) => {
          const numMatch = it.value.match(/[\d.,]+\s*(?:%|％|億|万|千|件|人|円|pt)?/);
          const big = numMatch ? numMatch[0] : it.value;
          return `<div class="pv-kpi-card">
          <div class="pv-kpi-num">${esc(big)}</div>
          <div class="pv-kpi-label">${esc(it.key)}</div>
        </div>`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return;
      const n = items.length;
      const cols = n <= 3 ? n : Math.ceil(n / 2);
      const rowsN = Math.ceil(n / cols);
      const gap = 0.3;
      const rowGap = 0.25;
      const cw = (box.w - gap * (cols - 1)) / cols;
      const ch = (box.h - rowGap * (rowsN - 1)) / rowsN;
      items.forEach((it, i) => {
        const col = i % cols;
        const row = Math.floor(i / cols);
        const x = box.x + col * (cw + gap);
        const y = box.y + row * (ch + rowGap);
        const numMatch = it.value.match(/[\d.,]+\s*(?:%|％|億|万|千|件|人|円|pt)?/);
        const big = numMatch ? numMatch[0] : it.value;
        slide.addText(big, { x, y, w: cw, h: ch * 0.6, fontSize: 26, bold: true, color: theme.primary, align: 'center', valign: 'bottom' });
        slide.addShape('line', { x: x + cw * 0.3, y: y + ch * 0.64, w: cw * 0.4, h: 0, line: { color: theme.border, width: 0.75 } });
        slide.addText(it.key, { x, y: y + ch * 0.68, w: cw, h: ch * 0.3, fontSize: 10.5, color: theme.subtext, align: 'center', valign: 'top' });
      });
    },
  };

  // ---------- サイクル図（循環プロセス） ----------
  function circlePositions(n, radiusPct) {
    const pos = [];
    for (let i = 0; i < n; i++) {
      const angle = ((-90 + (360 / n) * i) * Math.PI) / 180;
      pos.push({ x: 50 + radiusPct * Math.cos(angle), y: 50 + radiusPct * Math.sin(angle) });
    }
    return pos;
  }

  const cycle = {
    id: 'cycle',
    name: 'サイクル図（循環プロセス）',
    category: 'フロー・プロセス',
    description: 'PDCAなど繰り返すプロセスを円環で示す（3〜6項目）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['サイクル', 'PDCA', '循環', '繰り返し']);
      const n = bullets.length;
      if (kw && n >= 3 && n <= 6) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (items.length < 2) return emptyBody();
      const pos = circlePositions(items.length, 36);
      const nodes = items
        .map(
          (it, i) => `
        <div class="pv-cycle-node" style="left:${pos[i].x}%;top:${pos[i].y}%;">
          <div class="pv-cycle-num">${i + 1}</div>
          <div class="pv-cycle-label">${esc(it.key)}</div>
        </div>`
        )
        .join('');
      return `<div class="pv-cycle"><div class="pv-cycle-ring"></div>${nodes}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (items.length < 2) return;
      const n = items.length;
      const cx = box.x + box.w / 2;
      const cy = box.y + box.h / 2;
      const R = (Math.min(box.w, box.h) / 2) * 0.66;
      slide.addShape('oval', { x: cx - R, y: cy - R, w: R * 2, h: R * 2, fill: { type: 'none' }, line: { color: theme.accent, width: 1.5, dashType: 'dash' } });
      const nodeW = Math.min(2.4, (box.w / n) * 1.3);
      const nodeH = 0.9;
      items.forEach((it, i) => {
        const angle = ((-90 + (360 / n) * i) * Math.PI) / 180;
        const x = cx + R * Math.cos(angle) - nodeW / 2;
        const y = cy + R * Math.sin(angle) - nodeH / 2;
        slide.addShape('roundRect', { x, y, w: nodeW, h: nodeH, fill: { color: theme.primary }, rectRadius: 0.5 });
        slide.addText(`${i + 1}. ${it.key}`, { x, y, w: nodeW, h: nodeH, fontSize: 11, bold: true, color: theme.white, align: 'center', valign: 'middle' });
      });
    },
  };

  // ---------- ロジックツリー ----------
  const logicTree = {
    id: 'logic-tree',
    name: 'ロジックツリー',
    category: '構造・ロジック',
    description: 'テーマを複数の要素に分解して示す（MECE整理）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ロジックツリー', '要素分解', 'MECE', '内訳', '分解']);
      if (kw && bullets.length >= 2) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 5);
      if (!items.length) return emptyBody();
      return `<div class="pv-tree">
        <div class="pv-tree-trunk"></div>
        <div class="pv-tree-branches">${items
          .map(
            (it) => `
          <div class="pv-tree-branch">
            <div class="pv-tree-connector"></div>
            <div class="pv-tree-node">${esc(it.key)}</div>
            ${it.value !== it.key ? `<div class="pv-tree-leaf">${esc(it.value)}</div>` : ''}
          </div>`
          )
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 5);
      if (!items.length) return;
      const n = items.length;
      const trunkY = box.y + 0.3;
      slide.addShape('line', { x: box.x + box.w / 2, y: box.y, w: 0, h: 0.3, line: { color: theme.accent, width: 2 } });
      slide.addShape('line', { x: box.x, y: trunkY, w: box.w, h: 0, line: { color: theme.accent, width: 2 } });
      const gap = 0.2;
      const colW = (box.w - gap * (n - 1)) / n;
      items.forEach((it, i) => {
        const x = box.x + i * (colW + gap);
        const cx = x + colW / 2;
        slide.addShape('line', { x: cx, y: trunkY, w: 0, h: 0.25, line: { color: theme.accent, width: 2 } });
        const nodeY = trunkY + 0.25;
        slide.addShape('rect', { x, y: nodeY, w: colW, h: 0.6, fill: { color: theme.primary } });
        slide.addText(it.key, { x, y: nodeY, w: colW, h: 0.6, fontSize: 11, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        if (it.value !== it.key) {
          const leafY = nodeY + 0.75;
          slide.addShape('rect', { x, y: leafY, w: colW, h: box.y + box.h - leafY, fill: { color: theme.light }, line: { color: theme.accent, width: 1 } });
          slide.addText(it.value, { x: x + 0.05, y: leafY + 0.1, w: colW - 0.1, h: box.y + box.h - leafY - 0.1, fontSize: 10, color: theme.text, align: 'center', valign: 'top' });
        }
      });
    },
  };

  // ---------- ファネル図 ----------
  function firstNumber(str) {
    const m = String(str).match(/[\d,]+(?:\.\d+)?/);
    return m ? parseFloat(m[0].replace(/,/g, '')) : null;
  }
  function isDecreasingSeq(items) {
    const nums = items.map((it) => firstNumber(it.value));
    if (nums.some((n) => n === null) || nums.length < 3) return false;
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] > nums[i - 1]) return false;
    }
    return true;
  }

  const funnel = {
    id: 'funnel',
    name: 'ファネル図',
    category: '変化・対比',
    description: '歩留まり・コンバージョンなど段階的に絞り込まれるデータを示す。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ファネル', '歩留まり', 'コンバージョン']);
      const items = A.extractItems(bullets);
      const n = bullets.length;
      const dec = n >= 3 && isDecreasingSeq(items);
      if (n >= 3 && n <= 6 && kw && dec) return 10;
      if (n >= 3 && n <= 6 && dec) return 8;
      if (kw && n >= 3) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (items.length < 2) return emptyBody();
      const n = items.length;
      return `<div class="pv-funnel">${items
        .map((it, i) => {
          const widthPct = 100 - i * (60 / (n - 1 || 1));
          return `<div class="pv-funnel-row" style="width:${widthPct}%;"><span>${esc(it.key)}</span><span>${esc(it.value)}</span></div>`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (items.length < 2) return;
      const n = items.length;
      const gap = 0.12;
      const rowH = (box.h - gap * (n - 1)) / n;
      const maxW = box.w * 0.92;
      items.forEach((it, i) => {
        const w = maxW * (1 - i * (0.6 / (n - 1 || 1)));
        const x = box.x + (box.w - w) / 2;
        const y = box.y + i * (rowH + gap);
        slide.addShape('trapezoid', { x, y, w, h: rowH, fill: { color: theme.accent }, line: { type: 'none' } });
        slide.addText(`${it.key}：${it.value}`, { x, y, w, h: rowH, fontSize: 11, color: theme.white, align: 'center', valign: 'middle' });
      });
    },
  };

  // ---------- ウォーターフォール（増減内訳） ----------
  function parseSignedNumber(str) {
    const s = String(str);
    const m = s.match(/([+\-−▲△])?\s*([\d,]+(?:\.\d+)?)/);
    if (!m) return null;
    const val = parseFloat(m[2].replace(/,/g, ''));
    const negMark = m[1] === '-' || m[1] === '−' || m[1] === '▲' || m[1] === '△';
    const neg = negMark || /減/.test(s);
    const pos = m[1] === '+' || /増/.test(s);
    if (neg && !pos) return -Math.abs(val);
    return Math.abs(val);
  }
  function isTotalLabel(key) {
    return /(開始|期首|期末|終了|合計|総計|計)/.test(key);
  }
  function computeWaterfallBars(items) {
    let cumulative = 0;
    return items.map((it, i) => {
      const isTotal = isTotalLabel(it.key) || i === 0 || i === items.length - 1;
      const parsed = parseSignedNumber(it.value);
      let from;
      let to;
      if (isTotal) {
        const val = parsed != null ? Math.abs(parsed) : cumulative;
        from = 0;
        to = val;
        cumulative = val;
      } else {
        const delta = parsed || 0;
        from = cumulative;
        to = cumulative + delta;
        cumulative = to;
      }
      return { key: it.key, value: it.value, from: Math.min(from, to), to: Math.max(from, to), isTotal, isIncrease: isTotal ? true : to >= from };
    });
  }
  function hasSignedNumber(str) {
    return /[+\-−▲△]\s*\d|\d\s*(増|減)/.test(String(str));
  }

  const waterfall = {
    id: 'waterfall',
    name: 'ウォーターフォール（増減内訳）',
    category: '変化・対比',
    description: '数値の増減要因を積み上げ棒で示す（開始値→各要因→終了値）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ウォーターフォール', '増減', 'ブリッジ']);
      const items = A.extractItems(bullets);
      const signedCount = items.filter((it) => hasSignedNumber(it.value)).length;
      const n = bullets.length;
      if (n >= 3 && n <= 7 && kw && signedCount >= 2) return 10;
      if (n >= 3 && n <= 7 && signedCount >= Math.ceil(n * 0.5)) return 8;
      if (kw && n >= 3) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (items.length < 2) return emptyBody();
      const bars = computeWaterfallBars(items);
      const maxVal = Math.max(...bars.map((b) => b.to), 1);
      return `<div class="pv-waterfall">${bars
        .map((b) => {
          const heightPct = Math.max(4, ((b.to - b.from) / maxVal) * 100);
          const bottomPct = (b.from / maxVal) * 100;
          const cls = b.isTotal ? 'total' : b.isIncrease ? 'up' : 'down';
          return `<div class="pv-wf-col">
            <div class="pv-wf-bar-track"><div class="pv-wf-bar ${cls}" style="height:${heightPct}%;bottom:${bottomPct}%;"></div></div>
            <div class="pv-wf-label">${esc(b.key)}<br>${esc(b.value)}</div>
          </div>`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (items.length < 2) return;
      const bars = computeWaterfallBars(items);
      const maxVal = Math.max(...bars.map((b) => b.to), 1);
      const n = bars.length;
      const gap = 0.15;
      const colW = (box.w - gap * (n - 1)) / n;
      const chartH = box.h - 0.9;
      const chartBottom = box.y + chartH;
      bars.forEach((b, i) => {
        const x = box.x + i * (colW + gap);
        const h = Math.max(0.15, ((b.to - b.from) / maxVal) * chartH);
        const y = chartBottom - (b.to / maxVal) * chartH;
        const color = b.isTotal ? theme.primary : b.isIncrease ? theme.primaryLight : theme.accent;
        slide.addShape('rect', { x, y, w: colW, h, fill: { color } });
        slide.addText(`${b.key}\n${b.value}`, { x, y: chartBottom + 0.05, w: colW, h: 0.8, fontSize: 8, color: theme.text, align: 'center', valign: 'top' });
      });
    },
  };

  // ---------- SWOT分析 ----------
  const SWOT_LABELS = ['強み', '弱み', '機会', '脅威'];
  // 強み=ネイビー／弱み=ミディアムブルー／機会=青紫／脅威=オレンジ（NRIパレットの強調色）
  const SWOT_COLORS = ['000F78', '3C64AA', '5A63A7', 'F59637'];
  function mapToSwot(items) {
    const slots = [null, null, null, null];
    const rest = [];
    items.forEach((it) => {
      const idx = SWOT_LABELS.findIndex((l) => it.key.includes(l));
      if (idx !== -1 && !slots[idx]) slots[idx] = it;
      else rest.push(it);
    });
    let ri = 0;
    for (let i = 0; i < 4; i++) {
      if (!slots[i] && ri < rest.length) slots[i] = rest[ri++];
    }
    return slots;
  }

  const swot = {
    id: 'swot',
    name: 'SWOT分析',
    category: 'マトリクス／ポジショニング',
    description: '強み・弱み・機会・脅威の4象限で整理する。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.countAny(text, ['SWOT', '強み', '弱み', '機会', '脅威']);
      if (kw >= 2 && bullets.length >= 2) return 10;
      if (kw >= 1 && bullets.length >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets);
      const slots = mapToSwot(items);
      if (slots.every((s) => !s)) return emptyBody();
      return `<div class="pv-matrix">${slots
        .map(
          (s, i) => `
        <div class="pv-matrix-cell pv-swot-${i}">
          <div class="pv-matrix-key">${esc(SWOT_LABELS[i])}</div>
          <div class="pv-matrix-value">${s ? esc(s.value) : ''}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets);
      const slots = mapToSwot(items);
      if (slots.every((s) => !s)) return;
      const gap = 0.2;
      const cw = (box.w - gap) / 2;
      const ch = (box.h - gap) / 2;
      const positions = [
        [box.x, box.y],
        [box.x + cw + gap, box.y],
        [box.x, box.y + ch + gap],
        [box.x + cw + gap, box.y + ch + gap],
      ];
      slots.forEach((s, i) => {
        const [x, y] = positions[i];
        slide.addShape('rect', { x, y, w: cw, h: ch, fill: { color: SWOT_COLORS[i] }, line: { color: theme.white, width: 2 } });
        slide.addText(
          [
            { text: SWOT_LABELS[i] + '\n', options: { bold: true, fontSize: 13 } },
            { text: s ? s.value : '', options: { fontSize: 11 } },
          ],
          { x: x + 0.15, y: y + 0.15, w: cw - 0.3, h: ch - 0.3, color: theme.white, valign: 'top' }
        );
      });
    },
  };

  // ---------- アジェンダ／目次 ----------
  const agenda = {
    id: 'agenda',
    name: 'アジェンダ／目次',
    category: '汎用',
    description: '本日お伝えする項目を番号付きリストで示す（表紙・目次スライド向け）。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['アジェンダ', '目次', 'もくじ', '本日お伝えすること', 'Agenda']);
      return kw ? 10 : 0;
    },
    renderBody(bullets) {
      const list = (bullets || []).slice(0, 8);
      if (!list.length) return emptyBody();
      return `<ol class="pv-agenda">${list.map((b) => `<li>${emphasisHtml(b)}</li>`).join('')}</ol>`;
    },
    buildBody(slide, bullets, theme, box) {
      const list = (bullets || []).slice(0, 8);
      if (!list.length) return;
      const rowH = box.h / list.length;
      list.forEach((b, i) => {
        const y = box.y + i * rowH;
        slide.addShape('oval', { x: box.x, y: y + rowH / 2 - 0.2, w: 0.4, h: 0.4, fill: { color: theme.primary } });
        slide.addText(String(i + 1), { x: box.x, y: y + rowH / 2 - 0.2, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        slide.addText(A.emphasisRuns(b, theme.primary), { x: box.x + 0.6, y, w: box.w - 0.6, h: rowH, fontSize: 14, color: theme.text, valign: 'middle' });
        if (i < list.length - 1) {
          slide.addShape('line', { x: box.x, y: y + rowH, w: box.w, h: 0, line: { color: theme.border, width: 0.75 } });
        }
      });
    },
  };

  // ---------- 縦型プロセスフロー（ステップ数が多い場合向け） ----------
  const flowVertical = {
    id: 'flow-vertical',
    name: '縦型プロセスフロー',
    category: 'フロー・プロセス',
    description: '手順・ステップを縦に並べて示す。ステップ数が多い・説明が長い場合に横型より読みやすい。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ステップ', 'フロー', '手順', 'プロセス', 'フェーズ']);
      const seq = A.sequenceWordCount(text) + A.numberedBulletCount(bullets);
      const n = bullets.length;
      if (n >= 5 && n <= 8 && kw && seq > 0) return 9;
      if (n >= 5 && n <= 8 && (kw || seq > 0)) return 7;
      if (n >= 3 && n <= 4 && kw && seq > 0) return 5;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return emptyBody();
      return `<div class="pv-flow-v">
        <div class="pv-flow-v-line"></div>
        ${items
          .map((it, i) => {
            const label = it.value !== it.key ? `${it.key}：${it.value}` : it.key;
            return `<div class="pv-flow-v-row"><div class="pv-flow-v-num">${i + 1}</div><div class="pv-flow-v-text">${esc(label)}</div></div>`;
          })
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return;
      const n = items.length;
      const rowH = box.h / n;
      const lineX = box.x + 0.22;
      slide.addShape('line', { x: lineX, y: box.y + 0.1, w: 0, h: box.h - 0.2, line: { color: theme.accent, width: 2 } });
      items.forEach((it, i) => {
        const y = box.y + i * rowH;
        const cy = y + rowH / 2;
        slide.addShape('oval', { x: lineX - 0.18, y: cy - 0.18, w: 0.36, h: 0.36, fill: { color: theme.primary } });
        slide.addText(String(i + 1), { x: lineX - 0.18, y: cy - 0.18, w: 0.36, h: 0.36, fontSize: 11, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        const label = it.value !== it.key ? `${it.key}：${it.value}` : it.key;
        slide.addText(label, { x: lineX + 0.35, y: y + 0.03, w: box.x + box.w - (lineX + 0.35), h: rowH - 0.06, fontSize: 11, color: theme.text, valign: 'middle' });
      });
    },
  };

  // ---------- 縦型比較（項目数が多い・説明が長い場合向け） ----------
  const compareVertical = {
    id: 'compare-vertical',
    name: '縦型比較',
    category: '比較',
    description: '複数の項目を縦に並べて比較する。項目数が多い・説明が長い場合にボックス比較より読みやすい。',
    score(section) {
      const bullets = section.bullets || [];
      const ratio = A.kvRatio(bullets);
      const n = bullets.length;
      const avgLen = bullets.length ? bullets.join('').length / bullets.length : 0;
      if (n >= 5 && n <= 6 && ratio >= 0.5) return 8;
      if (n >= 2 && n <= 6 && ratio >= 0.5 && avgLen > 28) return 7;
      if (n >= 5 && n <= 6) return 4;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return emptyBody();
      return `<div class="pv-compare-v">${items
        .map(
          (it) => `
        <div class="pv-compare-v-row">
          <div class="pv-compare-v-label">${esc(it.key)}</div>
          <div class="pv-compare-v-desc">${esc(it.value)}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return;
      const n = items.length;
      const gap = 0.12;
      const rowH = (box.h - gap * (n - 1)) / n;
      const labelW = box.w * 0.26;
      items.forEach((it, i) => {
        const y = box.y + i * (rowH + gap);
        slide.addShape('rect', { x: box.x, y, w: labelW, h: rowH, fill: { color: theme.primary } });
        slide.addText(it.key, { x: box.x + 0.08, y, w: labelW - 0.16, h: rowH, fontSize: 11, bold: true, color: theme.white, valign: 'middle' });
        slide.addShape('rect', { x: box.x + labelW, y, w: box.w - labelW, h: rowH, fill: { color: theme.light }, line: { color: theme.border, width: 1 } });
        slide.addText(it.value, { x: box.x + labelW + 0.1, y, w: box.w - labelW - 0.2, h: rowH, fontSize: 10.5, color: theme.text, valign: 'middle' });
      });
    },
  };

  // ---------- 縦型タイムライン（マイルストーンが多い場合向け） ----------
  const timelineVertical = {
    id: 'timeline-vertical',
    name: '縦型タイムライン',
    category: '時系列',
    description: 'マイルストーンが多い・説明が詳しい場合向けの縦型タイムライン。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['スケジュール', 'ロードマップ', 'タイムライン']);
      const dates = A.dateTokenCount(text);
      const n = bullets.length;
      if (n >= 5 && n <= 8 && kw && dates > 0) return 9;
      if (n >= 5 && n <= 8 && dates >= n * 0.5) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return emptyBody();
      return `<div class="pv-tl-v">
        <div class="pv-tl-v-line"></div>
        ${items
          .map(
            (it) => `
          <div class="pv-tl-v-row">
            <div class="pv-tl-v-dot"></div>
            <div class="pv-tl-v-text"><b>${esc(it.key)}</b>${it.value !== it.key ? '<br>' + esc(it.value) : ''}</div>
          </div>`
          )
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return;
      const n = items.length;
      const rowH = box.h / n;
      const lineX = box.x + 0.16;
      slide.addShape('line', { x: lineX, y: box.y + 0.1, w: 0, h: box.h - 0.2, line: { color: theme.accent, width: 2 } });
      items.forEach((it, i) => {
        const y = box.y + i * rowH;
        const cy = y + Math.min(0.3, rowH / 2);
        slide.addShape('oval', { x: lineX - 0.09, y: cy - 0.09, w: 0.18, h: 0.18, fill: { color: theme.primary } });
        const label = it.value !== it.key ? `${it.key}\n${it.value}` : it.key;
        slide.addText(label, { x: lineX + 0.25, y: y + 0.03, w: box.x + box.w - (lineX + 0.25), h: rowH - 0.06, fontSize: 10.5, color: theme.text, valign: 'top' });
      });
    },
  };

  // ---------- ポジショニングマップ ----------
  const SCATTER_POS = [
    [0.22, 0.24],
    [0.72, 0.2],
    [0.3, 0.7],
    [0.78, 0.68],
    [0.5, 0.15],
    [0.15, 0.5],
    [0.85, 0.5],
    [0.5, 0.85],
  ];

  const positioningMap = {
    id: 'positioning-map',
    name: 'ポジショニングマップ',
    category: 'マトリクス／ポジショニング',
    description: '2軸の座標平面上に複数項目をプロットし、位置づけを示す。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      // 「マップ」は「ロードマップ」等と誤マッチしやすいため、より具体的な語のみ使う
      const kw = A.hasAny(text, ['ポジショニングマップ', 'ポジショニング', '位置づけ', '位置付け']);
      const n = bullets.length;
      if (kw && n >= 3 && n <= 8) return 10;
      if (A.hasAny(text, ['軸', 'マトリクス']) && n >= 5 && n <= 8) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (items.length < 2) return emptyBody();
      return `<div class="pv-posmap">
        <div class="pv-posmap-axis-x"></div>
        <div class="pv-posmap-axis-y"></div>
        <div class="pv-posmap-label top">軸2（高）</div>
        <div class="pv-posmap-label bottom">軸2（低）</div>
        <div class="pv-posmap-label left">軸1（低）</div>
        <div class="pv-posmap-label right">軸1（高）</div>
        ${items
          .map((it, i) => {
            const [x, y] = SCATTER_POS[i % SCATTER_POS.length];
            return `<div class="pv-posmap-dot" style="left:${x * 100}%;top:${y * 100}%;"><span class="dot"></span><span class="lbl">${esc(it.key)}</span></div>`;
          })
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (items.length < 2) return;
      const cx = box.x + box.w / 2;
      const cy = box.y + box.h / 2;
      slide.addShape('line', { x: box.x + box.w * 0.04, y: cy, w: box.w * 0.92, h: 0, line: { color: theme.border, width: 1 } });
      slide.addShape('line', { x: cx, y: box.y + box.h * 0.04, w: 0, h: box.h * 0.92, line: { color: theme.border, width: 1 } });
      slide.addText('軸2（高）', { x: box.x, y: box.y, w: box.w, h: 0.22, fontSize: 9, color: theme.subtext, align: 'center' });
      slide.addText('軸2（低）', { x: box.x, y: box.y + box.h - 0.22, w: box.w, h: 0.22, fontSize: 9, color: theme.subtext, align: 'center' });
      slide.addText('軸1（低）', { x: box.x, y: cy - 0.15, w: 1.0, h: 0.3, fontSize: 9, color: theme.subtext, valign: 'middle' });
      slide.addText('軸1（高）', { x: box.x + box.w - 1.0, y: cy - 0.15, w: 1.0, h: 0.3, fontSize: 9, color: theme.subtext, align: 'right', valign: 'middle' });
      items.forEach((it, i) => {
        const [px, py] = SCATTER_POS[i % SCATTER_POS.length];
        const x = box.x + box.w * px;
        const y = box.y + box.h * py;
        slide.addShape('oval', { x: x - 0.07, y: y - 0.07, w: 0.14, h: 0.14, fill: { color: theme.accent } });
        slide.addText(it.key, { x: x - 0.7, y: y + 0.09, w: 1.4, h: 0.28, fontSize: 9, color: theme.text, align: 'center' });
      });
    },
  };

  // ---------- 階層ピラミッド（ビジョン→戦略→施策など） ----------
  const pyramidTiered = {
    id: 'pyramid-tiered',
    name: '階層ピラミッド',
    category: '構造・ロジック',
    description: 'ビジョン→戦略→施策など、2〜3段階の階層を三角形の積み上げで示す。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ビジョン', '戦略', '階層', 'ピラミッド', '重点', '基本方針']);
      const n = (section.bullets || []).length;
      if (kw && n >= 2 && n <= 3) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 3);
      if (!items.length) return emptyBody();
      const n = items.length;
      return `<div class="pv-pyr-t">${items
        .map((it, i) => {
          const widthPct = n > 1 ? 35 + i * (65 / (n - 1)) : 70;
          return `<div class="pv-pyr-t-row" style="width:${widthPct}%;"><div class="pv-pyr-t-key">${esc(it.key)}</div><div class="pv-pyr-t-val">${esc(it.value)}</div></div>`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 3);
      if (!items.length) return;
      const n = items.length;
      const gap = 0.15;
      const rowH = (box.h - gap * (n - 1)) / n;
      const colors = [theme.primary, theme.primaryLight, theme.accent];
      items.forEach((it, i) => {
        const widthFrac = n > 1 ? 0.35 + i * (0.65 / (n - 1)) : 0.7;
        const w = box.w * widthFrac;
        const x = box.x + (box.w - w) / 2;
        const y = box.y + i * (rowH + gap);
        slide.addShape('rect', { x, y, w, h: rowH, fill: { color: colors[i % colors.length] } });
        slide.addText(
          [
            { text: it.key + '\n', options: { bold: true, fontSize: 12 } },
            { text: it.value, options: { fontSize: 10 } },
          ],
          { x: x + 0.1, y, w: w - 0.2, h: rowH, color: theme.white, align: 'center', valign: 'middle' }
        );
      });
    },
  };

  // ---------- 棒グラフ（ネイティブグラフ） ----------
  const barChart = {
    id: 'bar-chart',
    name: '棒グラフ',
    category: '数値・グラフ',
    description: '複数項目の数値を棒グラフで比較する（PowerPointのネイティブグラフとして生成）。',
    score(section) {
      const bullets = section.bullets || [];
      const items = A.extractItems(bullets);
      const validNums = items.filter((it) => firstNumber(it.value) !== null).length;
      const kw = A.hasAny(A.fullText(section), ['グラフ', '棒グラフ']);
      const n = bullets.length;
      if (n >= 3 && n <= 8 && validNums === n && kw) return 10;
      if (n >= 3 && n <= 8 && validNums === n) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (items.length < 2) return emptyBody();
      const values = items.map((it) => firstNumber(it.value) || 0);
      const maxV = Math.max(...values, 1);
      return `<div class="pv-barchart">${items
        .map((it, i) => {
          const h = Math.max(4, (values[i] / maxV) * 100);
          return `<div class="pv-bar-col">
            <div class="pv-bar-val">${esc(String(values[i]))}</div>
            <div class="pv-bar-track"><div class="pv-bar" style="height:${h}%;"></div></div>
            <div class="pv-bar-label">${esc(it.key)}</div>
          </div>`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (items.length < 2) return;
      const labels = items.map((it) => it.key);
      const values = items.map((it) => firstNumber(it.value) || 0);
      slide.addChart('bar', [{ name: '値', labels, values }], {
        x: box.x, y: box.y, w: box.w, h: box.h,
        barDir: 'col',
        showLegend: false,
        showValue: true,
        dataLabelColor: theme.text,
        dataLabelFontSize: 10,
        catAxisLabelColor: theme.text,
        catAxisLabelFontSize: 9,
        valAxisLabelColor: theme.subtext,
        valAxisLabelFontSize: 8,
        chartColors: [theme.accent],
      });
    },
  };

  // ---------- 折れ線グラフ（ネイティブグラフ） ----------
  const lineChart = {
    id: 'line-chart',
    name: '折れ線グラフ',
    category: '数値・グラフ',
    description: '数値の時系列推移を折れ線グラフで示す（PowerPointのネイティブグラフとして生成）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const items = A.extractItems(bullets);
      const dateNum = items.filter((it) => A.dateTokenCount(it.key + it.value) > 0 && firstNumber(it.value) !== null).length;
      const kw = A.hasAny(text, ['推移', 'トレンド', '折れ線']);
      const n = bullets.length;
      if (n >= 3 && n <= 8 && dateNum >= Math.ceil(n * 0.6) && kw) return 10;
      if (n >= 3 && n <= 8 && dateNum >= Math.ceil(n * 0.6)) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (items.length < 2) return emptyBody();
      const values = items.map((it) => firstNumber(it.value) || 0);
      const maxV = Math.max(...values, 1);
      const minV = Math.min(...values, 0);
      const range = maxV - minV || 1;
      const w = 100;
      const h = 60;
      const stepX = w / (items.length - 1 || 1);
      const points = values.map((v, i) => `${(i * stepX).toFixed(1)},${(h - ((v - minV) / range) * h).toFixed(1)}`).join(' ');
      return `<div class="pv-linechart">
        <svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" class="pv-line-svg">
          <polyline points="${points}" fill="none" stroke="var(--accent)" stroke-width="2" vector-effect="non-scaling-stroke" />
          ${values
            .map((v, i) => `<circle cx="${(i * stepX).toFixed(1)}" cy="${(h - ((v - minV) / range) * h).toFixed(1)}" r="1.6" fill="var(--primary)" />`)
            .join('')}
        </svg>
        <div class="pv-line-labels">${items.map((it) => `<span>${esc(it.key)}</span>`).join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (items.length < 2) return;
      const labels = items.map((it) => it.key);
      const values = items.map((it) => firstNumber(it.value) || 0);
      slide.addChart('line', [{ name: '値', labels, values }], {
        x: box.x, y: box.y, w: box.w, h: box.h,
        showLegend: false,
        showValue: true,
        lineDataSymbol: 'circle',
        lineSize: 2,
        chartColors: [theme.accent],
        dataLabelColor: theme.text,
        dataLabelFontSize: 9,
        catAxisLabelColor: theme.text,
        catAxisLabelFontSize: 9,
        valAxisLabelColor: theme.subtext,
        valAxisLabelFontSize: 8,
      });
    },
  };

  // ---------- 円グラフ／構成比（ネイティブグラフ） ----------
  const pieChart = {
    id: 'pie-chart',
    name: '円グラフ／構成比',
    category: '数値・グラフ',
    description: '構成比・内訳をドーナツ／円グラフで示す（PowerPointのネイティブグラフとして生成）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['内訳', '構成比', 'シェア', '割合', '円グラフ']);
      const items = A.extractItems(bullets);
      const pctCount = items.filter((it) => /%|％/.test(it.value)).length;
      const n = bullets.length;
      if (n >= 2 && n <= 6 && kw && pctCount >= Math.ceil(n * 0.5)) return 10;
      if (n >= 2 && n <= 6 && pctCount === n) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (items.length < 2) return emptyBody();
      const values = items.map((it) => firstNumber(it.value) || 0);
      const total = values.reduce((a, b) => a + b, 0) || 1;
      const colors = ['var(--primary)', 'var(--accent)', 'var(--primary-light)', '#8fb2d8', 'var(--negative)', '#a9b3c1'];
      let acc = 0;
      const stops = values
        .map((v, i) => {
          const start = (acc / total) * 360;
          acc += v;
          const end = (acc / total) * 360;
          return `${colors[i % colors.length]} ${start}deg ${end}deg`;
        })
        .join(', ');
      return `<div class="pv-piechart">
        <div class="pv-pie-donut" style="background: conic-gradient(${stops});"></div>
        <div class="pv-pie-legend">${items
          .map((it, i) => `<div class="pv-pie-legend-item"><span class="sw" style="background:${colors[i % colors.length]};"></span>${esc(it.key)}：${esc(it.value)}</div>`)
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (items.length < 2) return;
      const labels = items.map((it) => it.key);
      const values = items.map((it) => firstNumber(it.value) || 0);
      const colors = [theme.primary, theme.primaryLight, theme.accent2, theme.light, theme.accent, 'A9B3C1'];
      slide.addChart('doughnut', [{ name: '構成比', labels, values }], {
        x: box.x, y: box.y, w: box.w, h: box.h,
        showLegend: true,
        legendPos: 'r',
        legendColor: theme.text,
        legendFontSize: 10,
        showPercent: true,
        dataLabelColor: theme.white,
        dataLabelFontSize: 9,
        chartColors: colors,
      });
    },
  };

  DocAssist.patterns = [
    titleMessage,
    agenda,
    boxCompare,
    compareVertical,
    swot,
    matrix2x2,
    positioningMap,
    processFlow,
    flowVertical,
    cycle,
    funnel,
    waterfall,
    pyramid,
    pyramidTiered,
    logicTree,
    beforeAfter,
    timeline,
    timelineVertical,
    comparisonTable,
    kpiSummary,
    barChart,
    lineChart,
    pieChart,
  ];
  DocAssist.patternById = {};
  DocAssist.patterns.forEach((p) => {
    DocAssist.patternById[p.id] = p;
  });
})();
