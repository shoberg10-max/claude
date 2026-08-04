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
//     description: '説明文（プレビュー画面に表示）',
//     score(section) -> number,      // {heading, bullets} を見て 0〜10 程度で採点
//     renderBody(bullets) -> HTML文字列,   // プレビュー用（本文のみ。見出しは共通枠が描く）
//     buildBody(pptxSlide, bullets, theme, box) -> void,  // PPTX書き出し用
//   }
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

  // ---------- タイトル＋メッセージ（フォールバック） ----------
  const titleMessage = {
    id: 'title-message',
    name: 'タイトル＋メッセージ',
    description: 'キーメッセージ1行と補足の箇条書き。どんな内容にも使える汎用パターン。',
    score(section) {
      const n = (section.bullets || []).length;
      return n <= 4 ? 3 : 2;
    },
    renderBody(bullets) {
      if (!bullets.length) return emptyBody();
      const [msg, ...rest] = bullets;
      const restHtml = rest.length
        ? `<ul class="pv-bullets">${rest.map((b) => `<li>${esc(b)}</li>`).join('')}</ul>`
        : '';
      return `<div class="pv-message-box">${esc(msg)}</div>${restHtml}`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      const [msg, ...rest] = bullets;
      slide.addShape('roundRect', { x: box.x, y: box.y, w: box.w, h: 1.0, fill: { color: theme.light }, line: { color: theme.accent, width: 1 }, rectRadius: 0.06 });
      slide.addText(msg, { x: box.x + 0.2, y: box.y, w: box.w - 0.4, h: 1.0, fontSize: 18, bold: true, color: theme.primary, valign: 'middle' });
      if (rest.length) {
        slide.addText(
          rest.map((b) => ({ text: b, options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } })),
          { x: box.x + 0.1, y: box.y + 1.3, w: box.w - 0.2, h: box.h - 1.3, fontSize: 14, color: theme.text, valign: 'top' }
        );
      }
    },
  };

  // ---------- ボックス比較 ----------
  const boxCompare = {
    id: 'box-compare',
    name: 'ボックス比較',
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
        slide.addShape('rect', { x, y: box.y, w: boxW, h: 0.6, fill: { color: theme.primary } });
        slide.addText(it.key, { x, y: box.y, w: boxW, h: 0.6, fontSize: 13, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        slide.addShape('rect', { x, y: box.y + 0.6, w: boxW, h: box.h - 0.6, fill: { color: theme.light }, line: { color: theme.border, width: 1 } });
        slide.addText(it.value, { x: x + 0.12, y: box.y + 0.75, w: boxW - 0.24, h: box.h - 0.9, fontSize: 12, color: theme.text, valign: 'top' });
      });
    },
  };

  // ---------- 2軸マトリクス ----------
  const matrix2x2 = {
    id: 'matrix-2x2',
    name: '2軸マトリクス',
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
        slide.addShape('roundRect', { x, y: box.y + 0.6, w: stepW, h: box.h - 0.6, fill: { color: theme.accent }, line: { type: 'none' }, rectRadius: 0.06 });
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
          <div class="pv-pyramid-apex">${esc(apexText)}</div>
          <div class="pv-pyramid-base">${baseItems
            .slice(0, 4)
            .map((b) => `<div class="pv-pyramid-cell">${esc(b)}</div>`)
            .join('')}</div>
        </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      const [apex, ...base] = bullets;
      const baseItems = (base.length ? base : [apex]).slice(0, 4);
      const apexText = base.length ? apex : '結論';
      const apexW = box.w * 0.55;
      slide.addShape('roundRect', { x: box.x + (box.w - apexW) / 2, y: box.y, w: apexW, h: 0.9, fill: { color: theme.primary }, rectRadius: 0.06 });
      slide.addText(apexText, { x: box.x + (box.w - apexW) / 2 + 0.15, y: box.y, w: apexW - 0.3, h: 0.9, fontSize: 15, bold: true, color: theme.white, valign: 'middle' });
      const n = baseItems.length;
      const gap = 0.2;
      const cw = (box.w - gap * (n - 1)) / n;
      const y2 = box.y + 1.3;
      baseItems.forEach((b, i) => {
        const x = box.x + i * (cw + gap);
        slide.addShape('rect', { x, y: y2, w: cw, h: box.h - 1.3, fill: { color: theme.light }, line: { color: theme.accent, width: 1 } });
        slide.addText(b, { x: x + 0.1, y: y2 + 0.1, w: cw - 0.2, h: box.h - 1.5, fontSize: 12, color: theme.text, valign: 'top' });
      });
    },
  };

  // ---------- Before / After ----------
  const beforeAfter = {
    id: 'before-after',
    name: 'Before／After',
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
        slide.addShape('rect', { x, y: box.y, w: boxW, h: 0.6, fill: { color: i === 0 ? theme.subtext : theme.accent2 } });
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
        color: '333333',
        border: { type: 'solid', color: theme.border, pt: 0.5 },
        valign: 'middle',
        autoPage: false,
      });
    },
  };

  // ---------- KPIサマリー ----------
  const kpiSummary = {
    id: 'kpi-summary',
    name: 'KPIサマリー',
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
      const gap = 0.2;
      const cw = (box.w - gap * (cols - 1)) / cols;
      const ch = (box.h - gap * (rowsN - 1)) / rowsN;
      items.forEach((it, i) => {
        const col = i % cols;
        const row = Math.floor(i / cols);
        const x = box.x + col * (cw + gap);
        const y = box.y + row * (ch + gap);
        const numMatch = it.value.match(/[\d.,]+\s*(?:%|％|億|万|千|件|人|円|pt)?/);
        const big = numMatch ? numMatch[0] : it.value;
        slide.addShape('roundRect', { x, y, w: cw, h: ch, fill: { color: theme.light }, line: { color: theme.accent, width: 1 }, rectRadius: 0.06 });
        slide.addText(big, { x: x + 0.1, y: y + 0.1, w: cw - 0.2, h: ch * 0.6, fontSize: 22, bold: true, color: theme.accent2, align: 'center', valign: 'bottom' });
        slide.addText(it.key, { x: x + 0.1, y: y + ch * 0.65, w: cw - 0.2, h: ch * 0.3, fontSize: 11, color: theme.text, align: 'center', valign: 'top' });
      });
    },
  };

  DocAssist.patterns = [
    titleMessage,
    boxCompare,
    matrix2x2,
    processFlow,
    pyramid,
    beforeAfter,
    timeline,
    comparisonTable,
    kpiSummary,
  ];
  DocAssist.patternById = {};
  DocAssist.patterns.forEach((p) => {
    DocAssist.patternById[p.id] = p;
  });
})();
