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
//              構造・ロジック／変化・対比／時系列／数値・グラフ／アクションプラン
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

  // 洗練された見た目のための共通スタイル定数。
  // 参考にしたパーツテンプレート集の特徴：角はごくわずかに丸め、罫線は髪の毛のように細く、
  // 番号は塗りつぶさず輪郭の丸で示す。数値を1箇所にまとめ、全パターンで統一する。
  const ROUND = 0.04;        // コンテンツボックスの角丸（帯・シェブロン・表には使わない）
  const HAIRLINE = 0.75;     // 標準の罫線の太さ（pt）
  const HAIRLINE_THIN = 0.5; // 補助的な罫線（目盛り・区切り）

  // 輪郭だけの番号丸。塗りつぶしバッジより軽く見え、淡い背景の上でも沈まない。
  function addNumberCircle(slide, n, box, theme, opts) {
    const o = opts || {};
    const color = o.color || theme.primary;
    slide.addShape('oval', {
      x: box.x, y: box.y, w: box.w, h: box.h,
      fill: { color: o.fill || theme.white },
      line: { color, width: o.width || 1 },
    });
    slide.addText(String(n), {
      x: box.x, y: box.y, w: box.w, h: box.h,
      fontSize: o.fontSize || 11, bold: true, color,
      align: 'center', valign: 'middle',
    });
  }

  // 同系色を段階的に変化させる（帯や円弧を「進むほど濃く」見せるための階調）。
  const BLUE_SCALE = ['light', 'softBlue', 'accent', 'mid', 'primaryLight', 'primary'];
  function scaleColor(theme, i, n, reverse) {
    if (n <= 1) return theme.primaryLight;
    const t = reverse ? 1 - i / (n - 1) : i / (n - 1);
    const idx = Math.round(t * (BLUE_SCALE.length - 1));
    return theme[BLUE_SCALE[idx]];
  }
  function scaleCssVar(i, n, reverse) {
    const MAP = {
      light: 'var(--light)', softBlue: 'var(--soft-blue)', accent: 'var(--accent)',
      mid: 'var(--mid)', primaryLight: 'var(--primary-light)', primary: 'var(--primary)',
    };
    if (n <= 1) return MAP.primaryLight;
    const t = reverse ? 1 - i / (n - 1) : i / (n - 1);
    return MAP[BLUE_SCALE[Math.round(t * (BLUE_SCALE.length - 1))]];
  }

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

  // ---------- アイコン ----------
  // PowerPoint標準のプリセット図形でアイコンを描く。画像を貼り込まないので、
  // 書き出したあともPowerPoint上で色・サイズをそのまま編集できる。
  // プレビュー側は同じ意味を持つ線画SVGで揃える。
  const ICONS = {
    process: { shape: 'flowChartProcess', svg: '<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M3 10h18"/>' },
    doc: { shape: 'flowChartDocument', svg: '<path d="M6 2h8l4 4v16H6z"/><path d="M14 2v4h4"/>' },
    db: { shape: 'flowChartMagneticDisk', svg: '<ellipse cx="12" cy="6" rx="8" ry="3"/><path d="M4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6"/>' },
    server: { shape: 'can', svg: '<ellipse cx="12" cy="5" rx="7" ry="2.5"/><path d="M5 5v14c0 1.4 3 2.5 7 2.5s7-1.1 7-2.5V5"/><path d="M5 12c0 1.4 3 2.5 7 2.5s7-1.1 7-2.5"/>' },
    cloud: { shape: 'cloud', svg: '<path d="M7 19h10a4 4 0 0 0 .3-8 6 6 0 0 0-11.4-1.4A4 4 0 0 0 7 19z"/>' },
    gear: { shape: 'gear6', svg: '<circle cx="12" cy="12" r="3.5"/><path d="M12 2.5v3M12 18.5v3M21.5 12h-3M5.5 12h-3M18.7 5.3l-2.1 2.1M7.4 16.6l-2.1 2.1M18.7 18.7l-2.1-2.1M7.4 7.4L5.3 5.3"/>' },
    folder: { shape: 'folderCorner', svg: '<path d="M3 6h6l2 2h10v12H3z"/>' },
    decision: { shape: 'flowChartDecision', svg: '<path d="M12 3l9 9-9 9-9-9z"/>' },
  };

  // ラベルの語句からふさわしいアイコンを推測する（構成図で明示指定させないための補助）。
  const ICON_HINTS = [
    [/クラウド|SaaS|Cloud|AWS|Azure|GCP/i, 'cloud'],
    [/DB|ＤＢ|データベース|データ|Oracle|SQL|マスタ/i, 'db'],
    [/サーバ|基盤|インフラ|オンプレ/, 'server'],
    [/帳票|書類|ファイル|ドキュメント|資料|Excel/i, 'doc'],
    [/業務|運用|プロセス|手作業|自動化/, 'gear'],
    [/フォルダ|共有|格納|ストレージ/, 'folder'],
    [/判断|審査|承認|チェック/, 'decision'],
  ];
  function guessIcon(label) {
    const hit = ICON_HINTS.find(([re]) => re.test(String(label || '')));
    return hit ? hit[1] : 'process';
  }
  function iconSvg(key) {
    const icon = ICONS[key] || ICONS.process;
    return `<svg class="pv-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${icon.svg}</svg>`;
  }
  function addIcon(slide, key, box, theme) {
    const icon = ICONS[key] || ICONS.process;
    slide.addShape(icon.shape, {
      x: box.x, y: box.y, w: box.w, h: box.h,
      fill: { color: theme.light }, line: { color: theme.primary, width: 1 },
    });
  }

  // 複合レイアウトで使う共通の区切り。「／」は1グループ内の要素、「｜」は列・段階を表す。
  function splitGroupItems(value) {
    return String(value == null ? '' : value).split(/[／/]/).map((s) => s.trim()).filter(Boolean);
  }
  function splitFields(value) {
    return String(value == null ? '' : value).split(/[｜|]/).map((s) => s.trim()).filter(Boolean);
  }
  // 「顧客DB（Oracle）」→ { main: '顧客DB', sub: 'Oracle' }。括弧書きを補足行として扱う。
  function splitItemSub(text) {
    const m = String(text == null ? '' : text).match(/^(.*?)[（(]([^）)]*)[）)]\s*$/);
    if (m && m[1].trim()) return { main: m[1].trim(), sub: m[2].trim() };
    return { main: String(text == null ? '' : text).trim(), sub: '' };
  }

  // 進捗系パターンの状態バッジ色。「遅延」は例外的にレッドを使う（本当に注意喚起が
  // 必要な箇所のみ予約色を使う、というテーマ方針に沿った唯一の用途）。
  function statusTone(text, theme) {
    const s = String(text == null ? '' : text);
    if (/遅延|遅れ|超過|危険|Delayed|At ?Risk/i.test(s)) return theme.critical;
    if (/完了|済|Done|クローズ|Closed/i.test(s)) return theme.primary;
    if (/進行|対応中|着手|WIP|In ?Progress/i.test(s)) return theme.primaryLight;
    return theme.gray;
  }
  // 影響度・重要度バッジ色（高＝レッド、中＝ブルー、低＝グレー）。
  function severityTone(text, theme) {
    const s = String(text == null ? '' : text);
    if (/^\s*(高|大|重大|High|Critical)/i.test(s)) return theme.critical;
    if (/^\s*(中|Medium|Mid)/i.test(s)) return theme.primaryLight;
    return theme.gray;
  }
  // プレビューHTML側は同じ判定をCSSクラスに落とす（PPTX側の色と1対1で対応させる）。
  function statusCls(text) {
    const s = String(text == null ? '' : text);
    if (/遅延|遅れ|超過|危険|Delayed|At ?Risk/i.test(s)) return 'is-critical';
    if (/完了|済|Done|クローズ|Closed/i.test(s)) return 'is-primary';
    if (/進行|対応中|着手|WIP|In ?Progress/i.test(s)) return 'is-primary-light';
    return 'is-gray';
  }
  function severityCls(text) {
    const s = String(text == null ? '' : text);
    if (/^\s*(高|大|重大|High|Critical)/i.test(s)) return 'is-critical';
    if (/^\s*(中|Medium|Mid)/i.test(s)) return 'is-primary-light';
    return 'is-gray';
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
    description: '2〜5個の項目を横並びのボックスで比較する。',
    score(section) {
      const bullets = section.bullets || [];
      const ratio = A.kvRatio(bullets);
      const n = bullets.length;
      if (n >= 2 && n <= 5 && ratio >= 0.5) return 8;
      if (n >= 2 && n <= 5) return 4;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 5);
      if (!items.length) return emptyBody();
      return `<div class="pv-box-row">${items
        .map(
          (it) => `
        <div class="pv-box${it.highlight ? ' pv-box-highlight' : ''}">
          <div class="pv-box-head">${it.highlight ? '★ ' : ''}${esc(it.key)}</div>
          <div class="pv-box-body">${esc(it.value)}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 5);
      if (!items.length) return;
      const n = items.length;
      const gap = n >= 5 ? 0.16 : 0.25;
      const boxW = (box.w - gap * (n - 1)) / n;
      const headFontSize = n >= 5 ? 12 : 13;
      const bodyFontSize = n >= 5 ? 10.5 : 11.5;
      items.forEach((it, i) => {
        const x = box.x + i * (boxW + gap);
        const headColor = it.highlight ? theme.highlight : theme.primary;
        slide.addShape('rect', { x, y: box.y, w: boxW, h: 0.5, fill: { color: headColor } });
        slide.addText(`${it.highlight ? '★ ' : ''}${it.key}`, { x, y: box.y, w: boxW, h: 0.5, fontSize: headFontSize, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        slide.addShape('rect', {
          x, y: box.y + 0.5, w: boxW, h: box.h - 0.5, fill: { color: theme.light },
          line: it.highlight ? { color: theme.highlight, width: 1.5 } : { type: 'none' },
        });
        slide.addText(
          [{ text: it.value, options: { bullet: SQUARE_BULLET } }],
          { x: x + 0.15, y: box.y + 0.62, w: boxW - 0.28, h: box.h - 0.74, fontSize: bodyFontSize, color: theme.text, valign: 'top' }
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
      // ベタ塗り＋白文字ではなく、淡い塗り＋濃い文字にする（参考パーツ集の4象限の作法）。
      // 象限ごとの区別は塗りの濃淡だけで付け、文字は常に読みやすい濃さを保つ。
      const fills = [theme.light, theme.lighter, theme.lighter, theme.pale];
      items.forEach((it, i) => {
        const [x, y] = positions[i];
        slide.addShape('rect', { x, y, w: cw, h: ch, fill: { color: fills[i] }, line: { color: theme.white, width: 2 } });
        slide.addText(
          [
            { text: it.key + '\n', options: { bold: true, fontSize: 13, color: theme.primary } },
            { text: it.value, options: { fontSize: 11, color: theme.text } },
          ],
          { x: x + 0.18, y: y + 0.15, w: cw - 0.36, h: ch - 0.3, valign: 'top' }
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
      if (n >= 3 && n <= 7 && kw && seq > 0) return 9;
      if (n >= 3 && n <= 7 && (kw || seq > 0)) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (!items.length) return emptyBody();
      return `<div class="pv-flow">${items
        .map((it, i) => {
          const label = it.value === it.key ? it.key : `${it.key}：${it.value}`;
          return `
        <div class="pv-flow-step">
          <div class="pv-flow-num">${i + 1}</div>
          <div class="pv-flow-label">${esc(label)}</div>
        </div>
        ${i < items.length - 1 ? '<div class="pv-flow-arrow"></div>' : ''}`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (!items.length) return;
      const n = items.length;
      const arrowW = 0.3;
      const stepW = (box.w - arrowW * (n - 1)) / n;
      const labelFontSize = n >= 6 ? 10 : 11;
      items.forEach((it, i) => {
        const x = box.x + i * (stepW + arrowW);
        const label = it.value === it.key ? it.key : `${it.key}\n${it.value}`;
        // 淡い塗り＋濃い文字の角丸ボックスに、輪郭だけの番号丸を重ねる
        slide.addShape('roundRect', {
          x, y: box.y + 0.6, w: stepW, h: box.h - 0.6, rectRadius: ROUND,
          fill: { color: theme.light }, line: { type: 'none' },
        });
        addNumberCircle(slide, i + 1, { x: x + stepW / 2 - 0.25, y: box.y, w: 0.5, h: 0.5 }, theme, { fontSize: 13 });
        slide.addText(label, { x: x + 0.12, y: box.y + 0.72, w: stepW - 0.24, h: box.h - 0.84, fontSize: labelFontSize, color: theme.text, align: 'center', valign: 'top' });
        if (i < n - 1) {
          // 「→」の文字ではなく小さな三角形にする（参考パーツ集の区切り記号）
          slide.addShape('triangle', {
            x: x + stepW + arrowW / 2 - 0.09, y: box.y + 0.6 + (box.h - 0.6) / 2 - 0.09, w: 0.18, h: 0.18,
            rotate: 90, fill: { color: theme.primary }, line: { type: 'none' },
          });
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
            .slice(0, 5)
            .map((b) => `<div class="pv-pyramid-cell">${emphasisHtml(b)}</div>`)
            .join('')}</div>
        </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      const [apex, ...base] = bullets;
      const baseItems = (base.length ? base : [apex]).slice(0, 5);
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
        ${items
          .map(
            (it) => `<tr${it.highlight ? ' class="pv-table-highlight"' : ''}><td>${it.highlight ? '★ ' : ''}${esc(it.key)}</td><td>${esc(it.value)}</td></tr>`
          )
          .join('')}
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
        ...items.map((it, i) => {
          const fillColor = it.highlight ? theme.light : i % 2 ? theme.light : theme.white;
          return [
            {
              text: `${it.highlight ? '★ ' : ''}${it.key}`,
              options: {
                fill: { color: fillColor },
                bold: !!it.highlight,
                color: it.highlight ? theme.highlight : theme.text,
              },
            },
            { text: it.value, options: { fill: { color: fillColor }, bold: !!it.highlight } },
          ];
        }),
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

  // ---------- アクションプラン（タスク・担当・期限） ----------
  // 論点への意思決定が済んだあと、「誰が・何を・いつまでに」やるかを一覧化するための表。
  // 各箇条書きは「タスク｜担当｜期限」の形式（全角パイプ区切り）を想定し、
  // 区切りが無い行は担当・期限を空欄のままタスク名だけ表示する（壊れないための保険）。
  function parseActionItems(bullets) {
    return (bullets || []).map((b) => {
      const stripped = A.stripEmphasis(b);
      const parts = stripped.split('｜').map((s) => s.trim()).filter((s) => s.length);
      if (parts.length >= 3) return { task: parts[0], owner: parts[1], due: parts[2] };
      if (parts.length === 2) return { task: parts[0], owner: '', due: parts[1] };
      const kv = A.splitKV(stripped);
      if (kv) return { task: kv.key, owner: '', due: kv.value };
      return { task: stripped, owner: '', due: '' };
    });
  }

  const actionPlanTable = {
    id: 'action-plan-table',
    name: 'アクションプラン（タスク・担当・期限）',
    category: 'アクションプラン',
    description: '意思決定後の実行計画を「タスク／担当／期限」の一覧表で示す。各行は「タスク｜担当｜期限」形式。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['アクションプラン', 'ネクストアクション', 'Next Action', '実行計画', 'ToDo', 'To Do', 'todo', '今後の進め方']);
      const items = parseActionItems(bullets);
      const structured = items.filter((it) => it.owner && it.due).length;
      const n = bullets.length;
      if (n >= 2 && n <= 8 && structured >= Math.ceil(n * 0.6) && kw) return 10;
      if (n >= 2 && n <= 8 && structured >= Math.ceil(n * 0.6)) return 8;
      if (kw && n >= 2) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = parseActionItems(bullets).slice(0, 8);
      if (!items.length) return emptyBody();
      return `<table class="pv-table"><thead><tr><th style="width:8%;">No</th><th>タスク</th><th style="width:20%;">担当</th><th style="width:22%;">期限</th></tr></thead><tbody>
        ${items
          .map(
            (it, i) => `<tr><td>${i + 1}</td><td>${esc(it.task)}</td><td>${esc(it.owner)}</td><td>${esc(it.due)}</td></tr>`
          )
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = parseActionItems(bullets).slice(0, 8);
      if (!items.length) return;
      const headOpts = { bold: true, color: theme.white, fill: { color: theme.primary } };
      const rows = [
        [
          { text: 'No', options: headOpts },
          { text: 'タスク', options: headOpts },
          { text: '担当', options: headOpts },
          { text: '期限', options: headOpts },
        ],
        ...items.map((it, i) => {
          const fill = { color: i % 2 ? theme.light : theme.white };
          return [
            { text: String(i + 1), options: { fill, align: 'center' } },
            { text: it.task, options: { fill } },
            { text: it.owner, options: { fill, align: 'center' } },
            { text: it.due, options: { fill, align: 'center' } },
          ];
        }),
      ];
      slide.addTable(rows, {
        x: box.x,
        y: box.y,
        w: box.w,
        colW: [box.w * 0.08, box.w * 0.5, box.w * 0.2, box.w * 0.22],
        fontSize: 11.5,
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
      const items = A.extractItems(bullets).slice(0, 9);
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
      const items = A.extractItems(bullets).slice(0, 9);
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
      // 破線ではなく淡い実線の輪でつなぐ。ノードは進むほど濃くして順序を色でも示す。
      const ringW = 0.14;
      slide.addShape('donut', {
        x: cx - R - ringW, y: cy - R - ringW, w: (R + ringW) * 2, h: (R + ringW) * 2,
        fill: { color: theme.lighter }, line: { type: 'none' },
      });
      const nodeW = Math.min(2.4, (box.w / n) * 1.3);
      const nodeH = 0.86;
      items.forEach((it, i) => {
        const angle = ((-90 + (360 / n) * i) * Math.PI) / 180;
        const x = cx + R * Math.cos(angle) - nodeW / 2;
        const y = cy + R * Math.sin(angle) - nodeH / 2;
        const fill = scaleColor(theme, i, n);
        slide.addShape('roundRect', { x, y, w: nodeW, h: nodeH, fill: { color: fill }, line: { color: theme.white, width: 1.5 }, rectRadius: 0.4 });
        slide.addText(`${i + 1}. ${it.key}`, {
          x: x + 0.08, y, w: nodeW - 0.16, h: nodeH,
          fontSize: 11, bold: true, color: i <= 1 ? theme.text : theme.white, align: 'center', valign: 'middle',
        });
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
      const items = A.extractItems(bullets).slice(0, 6);
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
      const items = A.extractItems(bullets).slice(0, 6);
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
        const color = b.isTotal ? theme.primary : b.isIncrease ? theme.primaryLight : theme.gray;
        slide.addShape('rect', { x, y, w: colW, h, fill: { color } });
        slide.addText(`${b.key}\n${b.value}`, { x, y: chartBottom + 0.05, w: colW, h: 0.8, fontSize: 8, color: theme.text, align: 'center', valign: 'top' });
      });
    },
  };

  // ---------- SWOT分析 ----------
  const SWOT_LABELS = ['強み', '弱み', '機会', '脅威'];
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
      // SWOT_LABELSの並び順（強み・弱み・機会・脅威）に対応。
      // 塗りは淡くし、ポジティブ（強み・機会）=青系、要注意（弱み・脅威）=グレー系という
      // 区別は見出し文字の色で付ける。ベタ塗りより落ち着き、本文も読みやすい。
      const swotFills = [theme.light, theme.paleGray, theme.lighter, 'F4F4F4'];
      const swotLabelColors = [theme.primary, '6B7480', theme.primaryLight, '6B7480'];
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
        slide.addShape('rect', { x, y, w: cw, h: ch, fill: { color: swotFills[i] }, line: { color: theme.white, width: 2 } });
        slide.addText(
          [
            { text: SWOT_LABELS[i] + '\n', options: { bold: true, fontSize: 13, color: swotLabelColors[i] } },
            { text: s ? s.value : '', options: { fontSize: 11, color: theme.text } },
          ],
          { x: x + 0.18, y: y + 0.15, w: cw - 0.36, h: ch - 0.3, valign: 'top' }
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
        addNumberCircle(slide, i + 1, { x: box.x, y: y + rowH / 2 - 0.2, w: 0.4, h: 0.4 }, theme, { fontSize: 11 });
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
        addNumberCircle(slide, i + 1, { x: lineX - 0.18, y: cy - 0.18, w: 0.36, h: 0.36 }, theme, { fontSize: 10 });
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
      if (n >= 5 && n <= 8 && ratio >= 0.5) return 8;
      if (n >= 2 && n <= 8 && ratio >= 0.5 && avgLen > 28) return 7;
      if (n >= 5 && n <= 8) return 4;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return emptyBody();
      return `<div class="pv-compare-v">${items
        .map(
          (it) => `
        <div class="pv-compare-v-row${it.highlight ? ' pv-compare-v-highlight' : ''}">
          <div class="pv-compare-v-label">${it.highlight ? '★ ' : ''}${esc(it.key)}</div>
          <div class="pv-compare-v-desc">${esc(it.value)}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return;
      const n = items.length;
      const gap = n >= 7 ? 0.08 : 0.12;
      const rowH = (box.h - gap * (n - 1)) / n;
      const labelW = box.w * 0.26;
      const fontSize = n >= 7 ? 10 : 10.5;
      items.forEach((it, i) => {
        const y = box.y + i * (rowH + gap);
        const labelColor = it.highlight ? theme.highlight : theme.primary;
        slide.addShape('rect', { x: box.x, y, w: labelW, h: rowH, fill: { color: labelColor } });
        slide.addText(`${it.highlight ? '★ ' : ''}${it.key}`, { x: box.x + 0.08, y, w: labelW - 0.16, h: rowH, fontSize, bold: true, color: theme.white, valign: 'middle' });
        slide.addShape('rect', {
          x: box.x + labelW, y, w: box.w - labelW, h: rowH, fill: { color: theme.light },
          line: it.highlight ? { color: theme.highlight, width: 1.5 } : { color: theme.border, width: 1 },
        });
        slide.addText(it.value, { x: box.x + labelW + 0.1, y, w: box.w - labelW - 0.2, h: rowH, fontSize, color: theme.text, valign: 'middle' });
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
    [0.35, 0.4],
    [0.65, 0.4],
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
      if (kw && n >= 3 && n <= 10) return 10;
      if (A.hasAny(text, ['軸', 'マトリクス']) && n >= 5 && n <= 10) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 10);
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
      if (n >= 3 && n <= 10 && validNums === n && kw) return 10;
      if (n >= 3 && n <= 10 && validNums === n) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 10);
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
      const items = A.extractItems(bullets).slice(0, 10);
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
      if (n >= 3 && n <= 10 && dateNum >= Math.ceil(n * 0.6) && kw) return 10;
      if (n >= 3 && n <= 10 && dateNum >= Math.ceil(n * 0.6)) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 10);
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
      const items = A.extractItems(bullets).slice(0, 10);
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
      const colors = ['var(--primary)', 'var(--primary-light)', 'var(--accent2)', 'var(--light)', 'var(--accent)', '#a9b3c1'];
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

  // ---------- As-Is / To-Be 移行図 ----------
  // 現行構成と移行後構成を左右のゾーンで対比し、中央に移行の向き、下部に移行ステップを置く複合図。
  // 入力形式：
  //   - As-Is：基幹システム（オンプレ）／顧客DB（Oracle）
  //   - To-Be：SaaS基幹システム／クラウドDB（PostgreSQL）
  //   - 移行ステップ：現行分析｜データ移行｜並行稼働
  // 「／」でゾーン内の要素を、「｜」で移行ステップを区切る。括弧書きは各要素の補足行になる。
  function parseAsIsToBe(bullets) {
    const items = A.extractItems(bullets);
    const pick = (re) => items.find((it) => re.test(it.key));
    const asIs = pick(/as-?is|現状|現行|旧/i);
    const toBe = pick(/to-?be|将来|あるべき|移行後|新/i);
    const steps = pick(/移行|ステップ|経路|フェーズ|段階/);
    const rest = items.filter((it) => it !== asIs && it !== toBe && it !== steps);
    const fallback = (i) => (rest[i] ? rest[i].value : '');
    return {
      asIsLabel: asIs ? asIs.key : 'As-Is',
      asIsItems: splitGroupItems(asIs ? asIs.value : fallback(0)).slice(0, 3),
      toBeLabel: toBe ? toBe.key : 'To-Be',
      toBeItems: splitGroupItems(toBe ? toBe.value : fallback(1)).slice(0, 3),
      steps: splitFields(steps ? steps.value : '').slice(0, 4),
    };
  }

  const asIsToBe = {
    id: 'as-is-to-be',
    name: 'As-Is／To-Be 移行図',
    category: '変化・対比',
    description: '現行構成と移行後構成をアイコン付きのゾーンで左右に対比し、下部に移行ステップを並べる構成図。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.countAny(text, ['As-Is', 'as-is', 'To-Be', 'to-be', '移行', '現行', 'あるべき姿', '構成図']);
      const hasZone = /[／/]/.test(text);
      if (kw >= 2 && hasZone) return 10;
      if (kw >= 2) return 8;
      return 0;
    },
    renderBody(bullets) {
      const d = parseAsIsToBe(bullets);
      if (!d.asIsItems.length && !d.toBeItems.length) return emptyBody();
      const zone = (label, list, toneCls) => `
        <div class="pv-zone">
          <div class="pv-zone-chip ${toneCls}">${esc(label)}</div>
          <div class="pv-zone-cards">${list
            .map((raw) => {
              const it = splitItemSub(raw);
              return `<div class="pv-zone-card">${iconSvg(guessIcon(it.main))}<div class="pv-zone-card-main">${esc(it.main)}</div>${
                it.sub ? `<div class="pv-zone-card-sub">${esc(it.sub)}</div>` : ''
              }</div>`;
            })
            .join('')}</div>
        </div>`;
      return `<div class="pv-a2b">
        <div class="pv-a2b-row">
          ${zone(d.asIsLabel, d.asIsItems, 'is-gray')}
          <div class="pv-a2b-mid"><div class="pv-a2b-mid-label">移行</div><div class="pv-a2b-arrow">▶</div></div>
          ${zone(d.toBeLabel, d.toBeItems, 'is-navy')}
        </div>
        ${
          d.steps.length
            ? `<div class="pv-a2b-steps">${d.steps
                .map((s, i) => `<div class="pv-a2b-step"><b>${i + 1}.</b> ${esc(s)}</div>`)
                .join('')}</div>`
            : ''
        }
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const d = parseAsIsToBe(bullets);
      if (!d.asIsItems.length && !d.toBeItems.length) return;
      const stepH = d.steps.length ? 0.62 : 0;
      const zoneH = box.h - stepH;
      const midW = Math.min(1.35, box.w * 0.12);
      const gap = 0.16;
      const zoneW = (box.w - midW - gap * 2) / 2;
      const zones = [
        { x: box.x, label: d.asIsLabel, items: d.asIsItems, tone: theme.gray },
        { x: box.x + zoneW + gap + midW + gap, label: d.toBeLabel, items: d.toBeItems, tone: theme.primary },
      ];
      zones.forEach((z) => {
        slide.addShape('rect', { x: z.x, y: box.y, w: zoneW, h: zoneH, fill: { color: theme.lighter }, line: { type: 'none' } });
        const chipW = Math.min(1.15, zoneW * 0.42);
        slide.addShape('rect', { x: z.x + 0.1, y: box.y + 0.1, w: chipW, h: 0.26, fill: { color: z.tone }, line: { type: 'none' } });
        slide.addText(z.label, { x: z.x + 0.1, y: box.y + 0.1, w: chipW, h: 0.26, fontSize: 9.5, bold: true, color: theme.white, align: 'center', valign: 'middle' });
        const n = z.items.length || 1;
        const cardGap = 0.12;
        const cardW = (zoneW - 0.24 - cardGap * (n - 1)) / n;
        const cardY = box.y + 0.46;
        const cardH = zoneH - 0.58;
        z.items.forEach((raw, i) => {
          const it = splitItemSub(raw);
          const x = z.x + 0.12 + i * (cardW + cardGap);
          slide.addShape('rect', { x, y: cardY, w: cardW, h: cardH, fill: { color: theme.white }, line: { color: theme.border, width: 0.75 } });
          const iconSize = Math.min(0.4, cardW * 0.38, cardH * 0.34);
          addIcon(slide, guessIcon(it.main), { x: x + cardW / 2 - iconSize / 2, y: cardY + 0.14, w: iconSize, h: iconSize }, theme);
          const textY = cardY + iconSize + 0.22;
          slide.addText(it.main, { x: x + 0.05, y: textY, w: cardW - 0.1, h: 0.4, fontSize: 9.5, bold: true, color: theme.text, align: 'center', valign: 'top' });
          if (it.sub) {
            slide.addText(it.sub, { x: x + 0.05, y: textY + 0.38, w: cardW - 0.1, h: 0.28, fontSize: 8, color: theme.subtext, align: 'center', valign: 'top' });
          }
        });
      });
      const midX = box.x + zoneW + gap;
      slide.addShape('rect', { x: midX, y: box.y, w: midW, h: zoneH, fill: { color: theme.white }, line: { color: theme.border, width: 0.75 } });
      slide.addText('移行', { x: midX, y: box.y + 0.1, w: midW, h: 0.26, fontSize: 9, bold: true, color: theme.primary, align: 'center', valign: 'middle' });
      slide.addShape('rightArrow', { x: midX + midW * 0.14, y: box.y + zoneH / 2 - 0.15, w: midW * 0.72, h: 0.3, fill: { color: theme.primary }, line: { type: 'none' } });
      if (d.steps.length) {
        const sy = box.y + zoneH + 0.16;
        const sn = d.steps.length;
        const sGap = 0.12;
        const sw = (box.w - sGap * (sn - 1)) / sn;
        d.steps.forEach((s, i) => {
          const x = box.x + i * (sw + sGap);
          slide.addShape('rect', { x, y: sy, w: sw, h: 0.34, fill: { color: theme.light }, line: { type: 'none' } });
          slide.addText(`${i + 1}. ${s}`, { x: x + 0.08, y: sy, w: sw - 0.16, h: 0.34, fontSize: 9, bold: true, color: theme.primary, valign: 'middle' });
        });
      }
    },
  };

  // ---------- 計画/実績ガント ----------
  // 入力形式：
  //   - 期間：4月｜5月｜6月｜7月｜8月     ← 列見出し（省略時は 1,2,3… になる）
  //   - 要件定義：1-2                     ← 計画のみ
  //   - 設計・開発：2-4｜2-5               ← 計画｜実績（2本目を書くと実績バーが下に並ぶ）
  function parseSpan(text) {
    const m = String(text == null ? '' : text).match(/(\d+)\s*[-–~〜]\s*(\d+)/);
    if (!m) return null;
    const from = parseInt(m[1], 10);
    const to = parseInt(m[2], 10);
    if (!(from >= 1) || !(to >= from)) return null;
    return { from, to };
  }
  function parseGantt(bullets) {
    const items = A.extractItems(bullets);
    let cols = [];
    const rows = [];
    items.forEach((it) => {
      if (/^(期間|列|軸|スケール|月)/.test(it.key) && !parseSpan(it.value)) {
        cols = splitFields(it.value);
        return;
      }
      const spans = splitFields(it.value).map(parseSpan).filter(Boolean);
      if (spans.length) rows.push({ label: it.key, plan: spans[0], actual: spans[1] || null, highlight: it.highlight });
    });
    const maxCol = rows.reduce((m, r) => Math.max(m, r.plan.to, r.actual ? r.actual.to : 0), 1);
    if (!cols.length) cols = Array.from({ length: maxCol }, (_, i) => String(i + 1));
    while (cols.length < maxCol) cols.push(String(cols.length + 1));
    return { cols, rows: rows.slice(0, 8) };
  }

  const ganttChart = {
    id: 'gantt-chart',
    name: '計画/実績ガント',
    category: '時系列',
    description: 'タスクごとの計画期間（と実績期間）を横棒で並べるガントチャート。「タスク：2-4｜2-5」の形式で書く。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ガント', '工程', 'スケジュール', '計画/実績', '計画・実績']);
      const g = parseGantt(bullets);
      if (g.rows.length >= 3 && kw) return 10;
      if (g.rows.length >= 3) return 8;
      return 0;
    },
    renderBody(bullets) {
      const g = parseGantt(bullets);
      if (!g.rows.length) return emptyBody();
      const n = g.cols.length;
      const pct = (v) => (v / n) * 100;
      return `<div class="pv-gantt">
        <div class="pv-gantt-head"><div class="pv-gantt-label"></div><div class="pv-gantt-cols">${g.cols
          .map((c) => `<div class="pv-gantt-col">${esc(c)}</div>`)
          .join('')}</div></div>
        ${g.rows
          .map(
            (r) => `<div class="pv-gantt-row">
          <div class="pv-gantt-label${r.highlight ? ' is-highlight' : ''}">${esc(r.label)}</div>
          <div class="pv-gantt-track">
            ${g.cols.map((_, i) => `<span class="pv-gantt-grid" style="left:${pct(i)}%;"></span>`).join('')}
            <div class="pv-gantt-bar${r.highlight ? ' is-highlight' : ''}" style="left:${pct(r.plan.from - 1)}%;width:${pct(r.plan.to - r.plan.from + 1)}%;"></div>
            ${
              r.actual
                ? `<div class="pv-gantt-bar is-actual" style="left:${pct(r.actual.from - 1)}%;width:${pct(r.actual.to - r.actual.from + 1)}%;"></div>`
                : ''
            }
          </div>
        </div>`
          )
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const g = parseGantt(bullets);
      if (!g.rows.length) return;
      const hasActual = g.rows.some((r) => r.actual);
      const legendH = hasActual ? 0.24 : 0;
      const labelW = box.w * 0.26;
      const gridX = box.x + labelW;
      const gridW = box.w - labelW;
      const colW = gridW / g.cols.length;
      const headH = 0.3;
      const bodyH = box.h - headH - legendH;
      const rowH = Math.min(0.5, bodyH / g.rows.length);
      const gridBottom = box.y + headH + rowH * g.rows.length;
      g.cols.forEach((c, i) => {
        slide.addText(c, { x: gridX + i * colW, y: box.y, w: colW, h: headH, fontSize: 8.5, color: theme.subtext, align: 'center', valign: 'middle' });
        slide.addShape('line', { x: gridX + i * colW, y: box.y + headH, w: 0, h: gridBottom - (box.y + headH), line: { color: theme.border, width: 0.5 } });
      });
      slide.addShape('line', { x: gridX + gridW, y: box.y + headH, w: 0, h: gridBottom - (box.y + headH), line: { color: theme.border, width: 0.5 } });
      slide.addShape('line', { x: box.x, y: box.y + headH, w: box.w, h: 0, line: { color: theme.border, width: 0.75 } });
      g.rows.forEach((r, i) => {
        const y = box.y + headH + i * rowH;
        slide.addText(r.label, {
          x: box.x, y, w: labelW - 0.12, h: rowH,
          fontSize: 9.5, bold: !!r.highlight, color: r.highlight ? theme.highlight : theme.text, valign: 'middle',
        });
        const barH = r.actual ? rowH * 0.28 : rowH * 0.4;
        const planY = r.actual ? y + rowH * 0.16 : y + rowH / 2 - barH / 2;
        slide.addShape('rect', {
          x: gridX + (r.plan.from - 1) * colW + 0.02, y: planY,
          w: (r.plan.to - r.plan.from + 1) * colW - 0.04, h: barH,
          fill: { color: r.highlight ? theme.highlight : theme.primary }, line: { type: 'none' },
        });
        if (r.actual) {
          slide.addShape('rect', {
            x: gridX + (r.actual.from - 1) * colW + 0.02, y: planY + barH + rowH * 0.08,
            w: (r.actual.to - r.actual.from + 1) * colW - 0.04, h: barH,
            fill: { color: theme.accent }, line: { type: 'none' },
          });
        }
        slide.addShape('line', { x: box.x, y: y + rowH, w: box.w, h: 0, line: { color: theme.border, width: 0.4 } });
      });
      if (hasActual) {
        const ly = gridBottom + 0.06;
        slide.addShape('rect', { x: gridX, y: ly + 0.05, w: 0.18, h: 0.1, fill: { color: theme.primary }, line: { type: 'none' } });
        slide.addText('計画', { x: gridX + 0.24, y: ly, w: 0.6, h: 0.2, fontSize: 8, color: theme.subtext, valign: 'middle' });
        slide.addShape('rect', { x: gridX + 0.86, y: ly + 0.05, w: 0.18, h: 0.1, fill: { color: theme.accent }, line: { type: 'none' } });
        slide.addText('実績', { x: gridX + 1.1, y: ly, w: 0.6, h: 0.2, fontSize: 8, color: theme.subtext, valign: 'middle' });
      }
    },
  };

  // ---------- マイルストーン（シェブロン） ----------
  // 入力形式： - 要件確定：2026年3月  ／  - M1：要件確定｜2026年3月
  function parseMilestones(bullets) {
    return A.extractItems(bullets)
      .slice(0, 6)
      .map((it) => {
        const f = splitFields(it.value);
        return f.length >= 2
          ? { name: it.key, title: f[0], date: f[1], highlight: it.highlight }
          : { name: it.key, title: '', date: f[0] || '', highlight: it.highlight };
      });
  }

  const milestoneChevron = {
    id: 'milestone-chevron',
    name: 'マイルストーン（シェブロン）',
    category: '時系列',
    description: '節目を矢羽根（シェブロン）で連ねて示す。到達点と時期を短く並べたいときに使う。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['マイルストーン', '節目', 'ゲート', 'Milestone']);
      const n = bullets.length;
      if (kw && n >= 3 && n <= 6) return 10;
      if (kw && n >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const ms = parseMilestones(bullets);
      if (!ms.length) return emptyBody();
      return `<div class="pv-ms">${ms
        .map(
          (m) => `<div class="pv-ms-step${m.highlight ? ' is-highlight' : ''}">
          <div class="pv-ms-name">${esc(m.name)}</div>
          ${m.title ? `<div class="pv-ms-title">${esc(m.title)}</div>` : ''}
          ${m.date ? `<div class="pv-ms-date">${esc(m.date)}</div>` : ''}
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const ms = parseMilestones(bullets);
      if (!ms.length) return;
      const n = ms.length;
      const overlap = 0.12;
      const stepW = (box.w + overlap * (n - 1)) / n;
      const h = Math.min(1.05, box.h);
      const y = box.y + (box.h - h) / 2;
      ms.forEach((m, i) => {
        const x = box.x + i * (stepW - overlap);
        const fill = m.highlight ? theme.highlight : i % 2 === 0 ? theme.primary : theme.primaryLight;
        slide.addShape(i === 0 ? 'homePlate' : 'chevron', { x, y, w: stepW, h, fill: { color: fill }, line: { type: 'none' } });
        const runs = [{ text: m.name, options: { bold: true, fontSize: 10.5, breakLine: true } }];
        if (m.title) runs.push({ text: m.title, options: { fontSize: 9, breakLine: true } });
        if (m.date) runs.push({ text: m.date, options: { fontSize: 8.5 } });
        slide.addText(runs, {
          x: x + overlap + 0.06, y, w: stepW - overlap * 2 - 0.12, h,
          color: theme.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.1,
        });
      });
    },
  };

  // ---------- タスク状況表 ----------
  // 入力形式： - 要件定義：山田｜完了｜3/10   （タスク名：担当｜状況｜期限）
  const taskStatusTable = {
    id: 'task-status-table',
    name: 'タスク状況表',
    category: 'アクションプラン',
    description: '進捗MTG向けに、タスクの担当・状況・期限を一覧化する。状況は完了／進行中／未着手／遅延で色分けされる。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['進捗', 'ステータス', '状況', 'タスク一覧']);
      const items = A.extractItems(bullets);
      const withStatus = items.filter((it) => /完了|進行|未着手|遅延|対応中/.test(it.value)).length;
      const n = bullets.length;
      if (n >= 3 && withStatus >= Math.ceil(n * 0.6) && kw) return 10;
      if (n >= 3 && withStatus >= Math.ceil(n * 0.6)) return 8;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return emptyBody();
      return `<table class="pv-table"><thead><tr><th>タスク</th><th style="width:16%;">担当</th><th style="width:16%;">状況</th><th style="width:18%;">期限</th></tr></thead><tbody>
        ${items
          .map((it) => {
            const f = splitFields(it.value);
            const status = f[1] || '';
            return `<tr><td>${esc(it.key)}</td><td>${esc(f[0] || '')}</td><td><span class="pv-badge ${statusCls(status)}">${esc(status)}</span></td><td>${esc(f[2] || '')}</td></tr>`;
          })
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return;
      const head = { bold: true, color: theme.white, fill: { color: theme.primary } };
      const rows = [
        [
          { text: 'タスク', options: head },
          { text: '担当', options: head },
          { text: '状況', options: head },
          { text: '期限', options: head },
        ],
        ...items.map((it, i) => {
          const f = splitFields(it.value);
          const status = f[1] || '';
          const fill = { color: i % 2 ? theme.lighter : theme.white };
          return [
            { text: it.key, options: { fill } },
            { text: f[0] || '', options: { fill, align: 'center' } },
            { text: status, options: { fill: { color: statusTone(status, theme) }, color: theme.white, bold: true, align: 'center' } },
            { text: f[2] || '', options: { fill, align: 'center' } },
          ];
        }),
      ];
      slide.addTable(rows, {
        x: box.x, y: box.y, w: box.w,
        colW: [box.w * 0.46, box.w * 0.16, box.w * 0.18, box.w * 0.2],
        fontSize: 11, color: theme.text,
        border: { type: 'solid', color: theme.white, pt: 1.5 },
        valign: 'middle', autoPage: false,
      });
    },
  };

  // ---------- 課題・リスク一覧 ----------
  // 入力形式： - 要件の認識齟齬：部門間で解釈が異なる｜高｜3/5に合同レビュー
  const issueRiskList = {
    id: 'issue-risk-list',
    name: '課題・リスク一覧',
    category: 'アクションプラン',
    description: '課題・リスクを内容／影響度／対応方針の3列で一覧化する。影響度「高」だけがレッドで強調される。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['課題', 'リスク', '懸念', 'Issue', 'Risk']);
      const items = A.extractItems(bullets);
      const withSeverity = items.filter((it) => splitFields(it.value).length >= 2).length;
      const n = bullets.length;
      if (n >= 2 && kw && withSeverity >= Math.ceil(n * 0.6)) return 10;
      if (n >= 2 && kw) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (!items.length) return emptyBody();
      return `<table class="pv-table"><thead><tr><th style="width:24%;">課題・リスク</th><th>内容</th><th style="width:12%;">影響度</th><th style="width:26%;">対応方針</th></tr></thead><tbody>
        ${items
          .map((it) => {
            const f = splitFields(it.value);
            const sev = f[1] || '';
            return `<tr><td>${esc(it.key)}</td><td>${esc(f[0] || '')}</td><td><span class="pv-badge ${severityCls(sev)}">${esc(sev)}</span></td><td>${esc(f[2] || '')}</td></tr>`;
          })
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (!items.length) return;
      const head = { bold: true, color: theme.white, fill: { color: theme.primary } };
      const rows = [
        [
          { text: '課題・リスク', options: head },
          { text: '内容', options: head },
          { text: '影響度', options: head },
          { text: '対応方針', options: head },
        ],
        ...items.map((it, i) => {
          const f = splitFields(it.value);
          const sev = f[1] || '';
          const fill = { color: i % 2 ? theme.lighter : theme.white };
          return [
            { text: it.key, options: { fill, bold: true } },
            { text: f[0] || '', options: { fill } },
            { text: sev, options: { fill: { color: severityTone(sev, theme) }, color: theme.white, bold: true, align: 'center' } },
            { text: f[2] || '', options: { fill } },
          ];
        }),
      ];
      slide.addTable(rows, {
        x: box.x, y: box.y, w: box.w,
        colW: [box.w * 0.24, box.w * 0.38, box.w * 0.12, box.w * 0.26],
        fontSize: 10.5, color: theme.text,
        border: { type: 'solid', color: theme.white, pt: 1.5 },
        valign: 'middle', autoPage: false,
      });
    },
  };

  // ---------- 意思決定依頼 ----------
  // 入力形式：
  //   - 依頼事項：B社CRMの採用可否をご判断いただきたい
  //   - 期限：2026年3月10日
  //   - A案：A社CRM｜低コストだが拡張性に課題
  //   - B案：★推奨 B社CRM｜コストと機能のバランスが良い
  function parseDecisionRequest(bullets) {
    const items = A.extractItems(bullets);
    const ask = items.find((it) => /依頼|決定|判断|論点|ご確認/.test(it.key));
    const due = items.find((it) => /期限|期日|回答|Deadline/i.test(it.key));
    const options = items.filter((it) => it !== ask && it !== due).slice(0, 3);
    return {
      ask: ask ? ask.value : '',
      due: due ? due.value : '',
      options: options.map((it) => {
        const f = splitFields(it.value);
        return { label: it.key, title: f[0] || it.value, desc: f[1] || '', highlight: it.highlight };
      }),
    };
  }

  const decisionRequest = {
    id: 'decision-request',
    name: '意思決定依頼',
    category: '比較',
    description: '依頼事項と回答期限を上部の帯で示し、選択肢を並べて意思決定を促す。推奨案は「★推奨」でハイライトされる。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.countAny(text, ['ご判断', '意思決定', '依頼事項', 'ご承認', '決裁', 'ご確認いただきたい']);
      const d = parseDecisionRequest(section.bullets || []);
      if (kw >= 1 && d.ask && d.options.length >= 2) return 10;
      if (kw >= 1 && d.options.length >= 2) return 8;
      if (kw >= 1) return 6;
      return 0;
    },
    renderBody(bullets) {
      const d = parseDecisionRequest(bullets);
      if (!d.ask && !d.options.length) return emptyBody();
      return `<div class="pv-decision">
        ${
          d.ask
            ? `<div class="pv-decision-ask"><span class="pv-decision-ask-text">${esc(d.ask)}</span>${
                d.due ? `<span class="pv-decision-due">期限：${esc(d.due)}</span>` : ''
              }</div>`
            : ''
        }
        <div class="pv-decision-options">${d.options
          .map(
            (o) => `<div class="pv-decision-option${o.highlight ? ' is-highlight' : ''}">
            <div class="pv-decision-option-head">${o.highlight ? '★ ' : ''}${esc(o.label)}</div>
            <div class="pv-decision-option-title">${esc(o.title)}</div>
            ${o.desc ? `<div class="pv-decision-option-desc">${esc(o.desc)}</div>` : ''}
          </div>`
          )
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const d = parseDecisionRequest(bullets);
      if (!d.ask && !d.options.length) return;
      let y = box.y;
      if (d.ask) {
        const askH = 0.62;
        slide.addShape('rect', { x: box.x, y, w: box.w, h: askH, fill: { color: theme.primary }, line: { type: 'none' } });
        const dueW = d.due ? Math.min(2.5, box.w * 0.26) : 0;
        slide.addText(d.ask, {
          x: box.x + 0.18, y, w: box.w - dueW - 0.36, h: askH,
          fontSize: 13, bold: true, color: theme.white, valign: 'middle',
        });
        if (d.due) {
          slide.addShape('rect', { x: box.x + box.w - dueW - 0.12, y: y + 0.13, w: dueW, h: askH - 0.26, fill: { color: theme.white }, line: { type: 'none' } });
          slide.addText(`期限：${d.due}`, {
            x: box.x + box.w - dueW - 0.12, y: y + 0.13, w: dueW, h: askH - 0.26,
            fontSize: 10, bold: true, color: theme.primary, align: 'center', valign: 'middle',
          });
        }
        y += askH + 0.2;
      }
      const opts = d.options;
      if (!opts.length) return;
      const n = opts.length;
      const gap = 0.2;
      const w = (box.w - gap * (n - 1)) / n;
      const h = box.y + box.h - y;
      opts.forEach((o, i) => {
        const x = box.x + i * (w + gap);
        const tone = o.highlight ? theme.highlight : theme.primaryLight;
        slide.addShape('rect', { x, y, w, h: 0.42, fill: { color: tone }, line: { type: 'none' } });
        slide.addText(`${o.highlight ? '★ ' : ''}${o.label}`, {
          x, y, w, h: 0.42, fontSize: 11.5, bold: true, color: theme.white, align: 'center', valign: 'middle',
        });
        slide.addShape('rect', {
          x, y: y + 0.42, w, h: h - 0.42, fill: { color: theme.light },
          line: o.highlight ? { color: theme.highlight, width: 1.5 } : { type: 'none' },
        });
        const runs = [{ text: o.title, options: { bold: true, fontSize: 11.5, breakLine: true, paraSpaceAfter: 6 } }];
        if (o.desc) runs.push({ text: o.desc, options: { fontSize: 10 } });
        slide.addText(runs, { x: x + 0.14, y: y + 0.56, w: w - 0.28, h: h - 0.68, color: theme.text, valign: 'top' });
      });
    },
  };

  // ---------- ブロッカーと依頼 ----------
  // 入力形式： - 要件確定の遅れ｜業務部門から3/10までにご回答いただきたい
  function parseBlockerPairs(bullets) {
    return (bullets || [])
      .map((b) => {
        const stripped = A.stripEmphasis(b);
        const highlight = /★\s*(?:推奨)?/.test(stripped);
        const clean = highlight ? stripped.replace(/★\s*(?:推奨)?/, '').trim() : stripped;
        const f = splitFields(clean);
        return { blocker: f[0] || clean, request: f[1] || '', highlight };
      })
      .filter((p) => p.blocker)
      .slice(0, 5);
  }

  const blockerRequest = {
    id: 'blocker-request',
    name: 'ブロッカーと依頼',
    category: 'アクションプラン',
    description: '進行を止めている事象と、それを解消するための依頼を左右で対にして示す。「ブロッカー｜依頼内容」の形式で書く。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ブロッカー', '障害', '止まって', 'ご依頼', 'お願いしたい', 'Blocker']);
      const pairs = parseBlockerPairs(section.bullets || []);
      const withReq = pairs.filter((p) => p.request).length;
      if (kw && withReq >= 2) return 10;
      if (kw && pairs.length >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const pairs = parseBlockerPairs(bullets);
      if (!pairs.length) return emptyBody();
      return `<div class="pv-blocker">
        <div class="pv-blocker-head"><div>ブロッカー</div><div></div><div>ご依頼事項</div></div>
        ${pairs
          .map(
            (p) => `<div class="pv-blocker-row${p.highlight ? ' is-highlight' : ''}">
          <div class="pv-blocker-cell is-issue">${esc(p.blocker)}</div>
          <div class="pv-blocker-arrow">▶</div>
          <div class="pv-blocker-cell is-req">${esc(p.request)}</div>
        </div>`
          )
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const pairs = parseBlockerPairs(bullets);
      if (!pairs.length) return;
      const headH = 0.3;
      const arrowW = 0.42;
      const colW = (box.w - arrowW) / 2;
      const rightX = box.x + colW + arrowW;
      slide.addText('ブロッカー', { x: box.x, y: box.y, w: colW, h: headH, fontSize: 10, bold: true, color: theme.subtext, valign: 'middle' });
      slide.addText('ご依頼事項', { x: rightX, y: box.y, w: colW, h: headH, fontSize: 10, bold: true, color: theme.subtext, valign: 'middle' });
      const n = pairs.length;
      const gap = 0.1;
      const rowH = (box.h - headH - gap * (n - 1)) / n;
      pairs.forEach((p, i) => {
        const y = box.y + headH + i * (rowH + gap);
        const issueTone = p.highlight ? theme.critical : theme.gray;
        slide.addShape('rect', { x: box.x, y, w: 0.05, h: rowH, fill: { color: issueTone }, line: { type: 'none' } });
        slide.addShape('rect', { x: box.x + 0.05, y, w: colW - 0.05, h: rowH, fill: { color: theme.lighter }, line: { type: 'none' } });
        slide.addText(p.blocker, { x: box.x + 0.16, y, w: colW - 0.28, h: rowH, fontSize: 10.5, color: theme.text, valign: 'middle' });
        slide.addShape('rightArrow', { x: box.x + colW + 0.08, y: y + rowH / 2 - 0.1, w: arrowW - 0.16, h: 0.2, fill: { color: theme.primary }, line: { type: 'none' } });
        slide.addShape('rect', { x: rightX, y, w: colW, h: rowH, fill: { color: theme.light }, line: { type: 'none' } });
        slide.addText(p.request, { x: rightX + 0.12, y, w: colW - 0.24, h: rowH, fontSize: 10.5, bold: true, color: theme.primary, valign: 'middle' });
      });
    },
  };

  // ---------- 入れ子シェブロン ----------
  // 大きな流れ（外側の帯）と、その中の工程（内側の白いシェブロン）を二段で示す。
  // 入力形式： - フェーズ1：現行分析｜要件定義
  function parseChevronGroups(bullets) {
    return A.extractItems(bullets)
      .slice(0, 4)
      .map((it) => ({ label: it.key, steps: splitFields(it.value).slice(0, 4), highlight: it.highlight }))
      .filter((g) => g.label);
  }

  const chevronFlow = {
    id: 'chevron-flow',
    name: '入れ子シェブロン',
    category: 'フロー・プロセス',
    description: '大きな流れを外側の帯、その中の工程を内側のシェブロンで二段に示す。「フェーズ：工程｜工程」の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['フェーズ', '工程', '段階', 'ステージ']);
      const groups = parseChevronGroups(section.bullets || []);
      const nested = groups.filter((g) => g.steps.length >= 2).length;
      if (kw && nested >= 2) return 10;
      if (nested >= 2) return 8;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseChevronGroups(bullets);
      if (!groups.length) return emptyBody();
      return `<div class="pv-cf">${groups
        .map(
          (g, i) => `
        <div class="pv-cf-group">
          <div class="pv-cf-band" style="background:${g.highlight ? 'var(--highlight)' : scaleCssVar(i, groups.length)};">${esc(g.label)}</div>
          <div class="pv-cf-steps">${g.steps.map((s) => `<div class="pv-cf-step"><span>${esc(s)}</span></div>`).join('')}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseChevronGroups(bullets);
      if (!groups.length) return;
      const n = groups.length;
      const gap = 0.08;
      const groupW = (box.w - gap * (n - 1)) / n;
      const bandH = Math.min(0.66, box.h * 0.38);
      const stepY = box.y + bandH + 0.12;
      const stepH = Math.min(0.72, box.h - bandH - 0.24);
      groups.forEach((g, gi) => {
        const gx = box.x + gi * (groupW + gap);
        const bandColor = g.highlight ? theme.highlight : scaleColor(theme, gi, n);
        slide.addShape(gi === 0 ? 'homePlate' : 'chevron', {
          x: gx, y: box.y, w: groupW, h: bandH,
          fill: { color: bandColor }, line: { type: 'none' },
        });
        slide.addText(g.label, {
          x: gx + 0.22, y: box.y, w: groupW - 0.44, h: bandH,
          fontSize: 12, bold: true, color: theme.white, align: 'center', valign: 'middle',
        });
        const sn = g.steps.length;
        if (!sn) return;
        const overlap = 0.1;
        const stepW = (groupW + overlap * (sn - 1)) / sn;
        g.steps.forEach((s, si) => {
          const sx = gx + si * (stepW - overlap);
          slide.addShape(si === 0 ? 'homePlate' : 'chevron', {
            x: sx, y: stepY, w: stepW, h: stepH,
            fill: { color: theme.white }, line: { color: bandColor, width: HAIRLINE },
          });
          slide.addText(s, {
            x: sx + 0.16, y: stepY, w: stepW - 0.32, h: stepH,
            fontSize: 9.5, color: theme.text, align: 'center', valign: 'middle',
          });
        });
      });
    },
  };

  // ---------- 菱形フロー ----------
  // 段階的に濃くなる帯の上に菱形を置く。STP（Segmentation→Targeting→Positioning）のような
  // 「3〜4語で言い切る流れ」を象徴的に見せるための型。
  // 英字ラベルは頭文字だけ大きくする（参考テンプレートの表現）。
  function initialRuns(label, theme, big, small) {
    const s = String(label || '');
    if (/^[A-Za-z]/.test(s) && s.length > 1) {
      return [
        { text: s.charAt(0), options: { fontSize: big, bold: true, color: theme.primary } },
        { text: s.slice(1), options: { fontSize: small, color: theme.primary } },
      ];
    }
    return [{ text: s, options: { fontSize: small + 1, bold: true, color: theme.primary } }];
  }

  const diamondFlow = {
    id: 'diamond-flow',
    name: '菱形フロー',
    category: 'フロー・プロセス',
    description: '段階的に濃くなる帯の上に菱形を並べ、3〜4段階の流れを象徴的に示す（STPなど）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['STP', 'ステップ', '流れ', 'プロセス', '段階']);
      const n = bullets.length;
      if (kw && n >= 3 && n <= 4) return 9;
      if (n === 3) return 5;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return emptyBody();
      const n = items.length;
      return `<div class="pv-df">
        <div class="pv-df-track">
          <div class="pv-df-band">${items
            .map((_, i) => `<div class="pv-df-seg" style="background:${scaleCssVar(i, n)};"></div>`)
            .join('')}<div class="pv-df-arrow"></div></div>
          <div class="pv-df-nodes">${items
            .map((it) => `<div class="pv-df-node"><span><i>${esc(it.key)}</i></span></div>`)
            .join('')}</div>
        </div>
        <div class="pv-df-captions">${items.map((it) => `<div>${esc(it.value !== it.key ? it.value : '')}</div>`).join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return;
      const n = items.length;
      const bandH = Math.min(0.5, box.h * 0.26);
      const bandY = box.y + Math.min(1.0, box.h * 0.34);
      const arrowW = 0.55;
      const segW = (box.w - arrowW) / n;
      items.forEach((_, i) => {
        slide.addShape('rect', {
          x: box.x + i * segW, y: bandY, w: segW + 0.01, h: bandH,
          fill: { color: scaleColor(theme, i, n) }, line: { type: 'none' },
        });
      });
      slide.addShape('rightArrow', {
        x: box.x + box.w - arrowW - 0.1, y: bandY - bandH * 0.32, w: arrowW + 0.1, h: bandH * 1.64,
        fill: { color: theme.primaryLight }, line: { type: 'none' },
      });
      const dSize = Math.min(1.7, segW * 0.92, box.h * 0.62);
      items.forEach((it, i) => {
        const cx = box.x + i * segW + segW / 2;
        const dy = bandY + bandH / 2 - dSize / 2;
        slide.addShape('diamond', {
          x: cx - dSize / 2, y: dy, w: dSize, h: dSize,
          fill: { color: theme.light }, line: { color: theme.white, width: 1.5 },
        });
        slide.addText(initialRuns(it.key, theme, 17, 11), {
          x: cx - dSize / 2, y: dy, w: dSize, h: dSize,
          align: 'center', valign: 'middle',
        });
        if (it.value && it.value !== it.key) {
          slide.addText(it.value, {
            x: cx - segW / 2 + 0.06, y: dy + dSize + 0.06, w: segW - 0.12, h: 0.46,
            fontSize: 9, color: theme.subtext, align: 'center', valign: 'top',
          });
        }
      });
    },
  };

  // ---------- STEPフロー ----------
  // 「STEP1／見出し／補足」を三角形の区切りで並べる、説明の細かい手順向けの型。
  // 入力形式： - STEP1：要件定義｜関係者ヒアリング｜要件定義書の作成
  const stepArrows = {
    id: 'step-arrows',
    name: 'STEPフロー',
    category: 'フロー・プロセス',
    description: 'STEP見出しと複数行の補足を、三角形の区切りで横に並べる。手順の中身まで書きたいときに使う。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = /STEP\s*\d|ステップ\s*\d/i.test(text);
      const items = A.extractItems(bullets);
      const detailed = items.filter((it) => splitFields(it.value).length >= 2).length;
      const n = bullets.length;
      if (kw && n >= 3 && n <= 5) return 10;
      if (detailed >= 2 && n >= 3 && n <= 5) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 5);
      if (!items.length) return emptyBody();
      return `<div class="pv-sa">${items
        .map((it, i) => {
          const lines = splitFields(it.value);
          const col = `<div class="pv-sa-col">
            <div class="pv-sa-head${it.highlight ? ' is-highlight' : ''}">${esc(it.key)}</div>
            ${lines.map((l) => `<div class="pv-sa-line">${esc(l)}</div>`).join('')}
          </div>`;
          return col + (i < items.length - 1 ? '<div class="pv-sa-tri"></div>' : '');
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 5);
      if (!items.length) return;
      const n = items.length;
      const triW = 0.34;
      const colW = (box.w - triW * (n - 1)) / n;
      items.forEach((it, i) => {
        const x = box.x + i * (colW + triW);
        slide.addText(it.key, {
          x, y: box.y, w: colW, h: 0.34,
          fontSize: 12, bold: true, color: it.highlight ? theme.highlight : theme.primary,
          align: 'center', valign: 'middle',
        });
        const lines = splitFields(it.value);
        if (lines.length) {
          slide.addText(
            lines.map((l, li) => ({ text: l, options: { fontSize: 9.5, breakLine: li < lines.length - 1, paraSpaceAfter: 3 } })),
            { x: x + 0.06, y: box.y + 0.4, w: colW - 0.12, h: box.h - 0.46, color: theme.subtext, align: 'center', valign: 'top' }
          );
        }
        if (i < n - 1) {
          slide.addShape('triangle', {
            x: x + colW + 0.04, y: box.y + 0.06, w: triW - 0.08, h: 0.26,
            rotate: 90, fill: { color: theme.gray }, line: { type: 'none' },
          });
        }
      });
    },
  };

  // ---------- 4分割サイクル ----------
  // 太い円環と4象限のラベルで、繰り返すプロセスを示す。
  const cycleQuadrant = {
    id: 'cycle-quadrant',
    name: '4分割サイクル',
    category: 'フロー・プロセス',
    description: '太い円環の中を4象限に区切り、繰り返すプロセスを示す（4項目専用）。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['サイクル', 'PDCA', '循環', '繰り返し']);
      if (kw && bullets.length === 4) return 10;
      if (bullets.length === 4) return 4;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return emptyBody();
      return `<div class="pv-cq">
        <div class="pv-cq-ring"></div>
        <div class="pv-cq-grid">${items
          .map((it) => `<div class="pv-cq-cell">${esc(it.key)}</div>`)
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return;
      const size = Math.min(box.w, box.h);
      const cx = box.x + box.w / 2;
      const cy = box.y + box.h / 2;
      const R = size / 2;
      slide.addShape('donut', {
        x: cx - R, y: cy - R, w: R * 2, h: R * 2,
        fill: { color: theme.accent }, line: { type: 'none' },
      });
      // 円環の上下左右に矢じりを置いて、回る向きを示す
      const head = R * 0.2;
      [
        { x: cx, y: cy - R, rotate: 90 },
        { x: cx + R, y: cy, rotate: 180 },
        { x: cx, y: cy + R, rotate: 270 },
        { x: cx - R, y: cy, rotate: 0 },
      ].forEach((p) => {
        slide.addShape('triangle', {
          x: p.x - head / 2, y: p.y - head / 2, w: head, h: head,
          rotate: p.rotate, fill: { color: theme.accent }, line: { type: 'none' },
        });
      });
      // 内側の2×2ラベル
      const inner = R * 0.66;
      const cellW = inner;
      const cellH = inner;
      const pos = [
        [cx - inner, cy - inner],
        [cx, cy - inner],
        [cx - inner, cy],
        [cx, cy],
      ];
      items.forEach((it, i) => {
        const [x, y] = pos[i];
        slide.addShape('rect', {
          x: x + 0.03, y: y + 0.03, w: cellW - 0.06, h: cellH - 0.06,
          fill: { color: theme.white }, line: { color: theme.light, width: HAIRLINE },
        });
        slide.addText(it.key, {
          x: x + 0.08, y: y + 0.03, w: cellW - 0.16, h: cellH - 0.06,
          fontSize: 11, color: theme.text, align: 'center', valign: 'middle',
        });
      });
    },
  };

  // ---------- 三角ピラミッド ----------
  // 上に行くほど狭くなる三角形を水平に分割し、階層（ハイタッチ／ロータッチ等）を示す。
  const triangleTiers = {
    id: 'triangle-tiers',
    name: '三角ピラミッド',
    category: '構造・ロジック',
    description: '三角形を水平に分割して階層を示す。上ほど少数・高付加価値、下ほど多数という関係の表現に使う。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['階層', 'ピラミッド', 'タッチ', '層', 'セグメント']);
      const n = bullets.length;
      if (kw && n >= 3 && n <= 4) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return emptyBody();
      const n = items.length;
      return `<div class="pv-tt">${items
        .map((it, i) => {
          const wTop = ((i / n) * 100).toFixed(1);
          const wBottom = (((i + 1) / n) * 100).toFixed(1);
          return `<div class="pv-tt-tier" style="background:${scaleCssVar(i, n, true)};clip-path:polygon(${(50 - wTop / 2).toFixed(
            1
          )}% 0,${(50 + wTop / 2).toFixed(1)}% 0,${(50 + wBottom / 2).toFixed(1)}% 100%,${(50 - wBottom / 2).toFixed(1)}% 100%);">
            <span>${esc(it.key)}</span></div>`;
        })
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return;
      const n = items.length;
      const size = Math.min(box.h, box.w * 0.55);
      const baseW = Math.min(box.w * 0.62, size * 1.5);
      const cx = box.x + box.w / 2;
      const tierH = size / n;
      items.forEach((it, i) => {
        const y = box.y + (box.h - size) / 2 + i * tierH;
        const wBottom = baseW * ((i + 1) / n);
        // 最上段だけ三角、以降は台形を積む
        slide.addShape(i === 0 ? 'triangle' : 'trapezoid', {
          x: cx - wBottom / 2, y, w: wBottom, h: tierH,
          fill: { color: scaleColor(theme, i, n, true) }, line: { color: theme.white, width: 1 },
        });
        const dark = i >= n - 1;
        slide.addText(it.key, {
          x: cx - wBottom / 2, y, w: wBottom, h: tierH,
          fontSize: 10.5, bold: true, color: dark ? theme.text : theme.white,
          align: 'center', valign: 'middle',
        });
        if (it.value && it.value !== it.key) {
          slide.addText(it.value, {
            x: cx + baseW / 2 + 0.18, y, w: box.x + box.w - (cx + baseW / 2 + 0.18), h: tierH,
            fontSize: 9.5, color: theme.text, valign: 'middle',
          });
        }
      });
    },
  };

  // ---------- 三者関係図 ----------
  // 3つの立場を三角形に配置し、中央にその関係が生む価値を置く（「三方良し」型）。
  const triangleRelation = {
    id: 'triangle-relation',
    name: '三者関係図',
    category: '構造・ロジック',
    description: '3つの立場を三角形に配置し、中央にその関係が生む価値を置く。「中心：価値」＋3項目の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['三方良し', '三者', '関係', 'Win-Win', 'ステークホルダー']);
      const items = A.extractItems(section.bullets || []);
      const nodes = items.filter((it) => !/中心|中央|価値/.test(it.key));
      if (kw && nodes.length === 3) return 10;
      if (nodes.length === 3 && kw) return 8;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets);
      const center = items.find((it) => /中心|中央|価値/.test(it.key));
      const nodes = items.filter((it) => it !== center).slice(0, 3);
      if (nodes.length < 3) return emptyBody();
      return `<div class="pv-tr">
        <div class="pv-tr-line"><svg viewBox="0 0 100 100" preserveAspectRatio="none">
          <polygon points="50,16 88,84 12,84" fill="none" stroke="var(--accent)" stroke-width="2.5" vector-effect="non-scaling-stroke" stroke-linejoin="round"/>
        </svg></div>
        ${center ? `<div class="pv-tr-center">${esc(center.value)}</div>` : ''}
        <div class="pv-tr-node is-top">${esc(nodes[0].key)}</div>
        <div class="pv-tr-node is-left">${esc(nodes[1].key)}</div>
        <div class="pv-tr-node is-right">${esc(nodes[2].key)}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets);
      const center = items.find((it) => /中心|中央|価値/.test(it.key));
      const nodes = items.filter((it) => it !== center).slice(0, 3);
      if (nodes.length < 3) return;
      const boxW = Math.min(1.9, box.w * 0.26);
      const boxH = 0.44;
      const cx = box.x + box.w / 2;
      const top = box.y + 0.12;
      const bottom = box.y + box.h - boxH - 0.12;
      const spread = Math.min(box.w * 0.32, 2.5);
      const pts = [
        { x: cx - boxW / 2, y: top },
        { x: cx - spread - boxW / 2, y: bottom },
        { x: cx + spread - boxW / 2, y: bottom },
      ];
      // 3辺（頂点の中心どうしを結ぶ）
      const centers = pts.map((p) => ({ x: p.x + boxW / 2, y: p.y + boxH / 2 }));
      [[0, 1], [1, 2], [0, 2]].forEach(([a, b]) => {
        slide.addShape('line', {
          x: Math.min(centers[a].x, centers[b].x),
          y: Math.min(centers[a].y, centers[b].y),
          w: Math.abs(centers[b].x - centers[a].x),
          h: Math.abs(centers[b].y - centers[a].y),
          line: { color: theme.accent, width: 2.5 },
          flipV: (centers[b].x - centers[a].x) * (centers[b].y - centers[a].y) < 0,
        });
      });
      pts.forEach((p, i) => {
        slide.addShape('roundRect', {
          x: p.x, y: p.y, w: boxW, h: boxH, rectRadius: ROUND,
          fill: { color: theme.white }, line: { color: theme.primaryLight, width: 1 },
        });
        slide.addText(nodes[i].key, {
          x: p.x, y: p.y, w: boxW, h: boxH,
          fontSize: 11, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
      });
      if (center) {
        slide.addText(`“${center.value}”`, {
          x: cx - 1.5, y: box.y + box.h * 0.52, w: 3.0, h: 0.4,
          fontSize: 14, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
      }
    },
  };

  // ---------- 優先度マトリクス ----------
  // 4象限を淡く塗り分け、各象限に優先度ラベルと内容ボックスを置く。
  // 塗りを淡くして文字を濃くすることで、ベタ塗りより落ち着いた印象にする。
  const QUADRANT_PRIORITY = ['高', '中', '中', '低'];
  const quadrantPriority = {
    id: 'quadrant-priority',
    name: '優先度マトリクス',
    category: 'マトリクス／ポジショニング',
    description: '4象限を淡く塗り分け、象限ごとの優先度と打ち手を示す。軸で優先順位を説明したいときに使う。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['優先度', '優先順位', '象限', '重要度']);
      if (kw && bullets.length === 4) return 10;
      if (kw && bullets.length >= 2) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return emptyBody();
      return `<div class="pv-qp">
        ${items
          .map(
            (it, i) => `<div class="pv-qp-cell pv-qp-${i}">
          <span class="pv-qp-pri">優先度：<b>${QUADRANT_PRIORITY[i]}</b></span>
          <span class="pv-qp-box">${esc(it.key)}</span>
        </div>`
          )
          .join('')}
        <div class="pv-qp-axis-x"></div><div class="pv-qp-axis-y"></div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return;
      const cw = box.w / 2;
      const ch = box.h / 2;
      const fills = [theme.light, theme.lighter, theme.pale, theme.paleGray];
      const pos = [
        [box.x, box.y],
        [box.x + cw, box.y],
        [box.x, box.y + ch],
        [box.x + cw, box.y + ch],
      ];
      items.forEach((it, i) => {
        const [x, y] = pos[i];
        slide.addShape('rect', { x, y, w: cw, h: ch, fill: { color: fills[i] }, line: { type: 'none' } });
        const rightSide = i % 2 === 1;
        slide.addText(
          [
            { text: '優先度：', options: { fontSize: 9, color: theme.subtext } },
            { text: QUADRANT_PRIORITY[i], options: { fontSize: 12, bold: true, color: theme.primary } },
          ],
          {
            x: x + 0.12, y: i < 2 ? y + 0.06 : y + ch - 0.34, w: cw - 0.24, h: 0.28,
            align: rightSide ? 'right' : 'left', valign: 'middle',
          }
        );
        const bw = Math.min(2.4, cw * 0.62);
        slide.addShape('roundRect', {
          x: x + (cw - bw) / 2, y: y + ch / 2 - 0.24, w: bw, h: 0.48, rectRadius: ROUND,
          fill: { color: theme.white }, line: { color: theme.border, width: HAIRLINE },
        });
        slide.addText(it.key, {
          x: x + (cw - bw) / 2, y: y + ch / 2 - 0.24, w: bw, h: 0.48,
          fontSize: 10.5, color: theme.text, align: 'center', valign: 'middle',
        });
      });
      // 中央を通る軸（矢印つき）
      slide.addShape('rightArrow', {
        x: box.x, y: box.y + ch - 0.035, w: box.w, h: 0.07,
        fill: { color: theme.subtext }, line: { type: 'none' },
      });
      slide.addShape('upArrow', {
        x: box.x + cw - 0.035, y: box.y, w: 0.07, h: box.h,
        fill: { color: theme.subtext }, line: { type: 'none' },
      });
    },
  };

  // ---------- 年表 ----------
  // 横軸に丸を打ち、年と出来事を交互に振り分けて並べる。
  // 入力形式： - 1987年：創業
  const yearTimeline = {
    id: 'year-timeline',
    name: '年表',
    category: '時系列',
    description: '横軸に丸を打って年と出来事を並べる。沿革や制度改正の経緯など、点が多い年表向け。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const years = (text.match(/\d{4}\s*年/g) || []).length;
      const kw = A.hasAny(text, ['沿革', '年表', '経緯', '変遷', 'これまでの']);
      const n = bullets.length;
      if (n >= 5 && years >= Math.ceil(n * 0.7)) return kw ? 10 : 9;
      if (kw && years >= 3) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 12);
      if (!items.length) return emptyBody();
      return `<div class="pv-yt">
        <div class="pv-yt-axis"></div>
        <div class="pv-yt-items">${items
          .map(
            (it, i) => `<div class="pv-yt-item${i % 2 ? ' is-below' : ''}">
            <span class="pv-yt-label">${esc(it.value !== it.key ? it.value : '')}</span>
            <span class="pv-yt-dot"></span>
            <span class="pv-yt-year">${esc(it.key)}</span>
          </div>`
          )
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 12);
      if (!items.length) return;
      const n = items.length;
      const axisY = box.y + box.h * 0.52;
      slide.addShape('line', {
        x: box.x, y: axisY, w: box.w - 0.28, h: 0,
        line: { color: theme.accent, width: 1.5 },
      });
      slide.addShape('triangle', {
        x: box.x + box.w - 0.28, y: axisY - 0.09, w: 0.18, h: 0.18,
        rotate: 90, fill: { color: theme.accent }, line: { type: 'none' },
      });
      const stepW = (box.w - 0.5) / n;
      const dot = 0.17;
      items.forEach((it, i) => {
        const cx = box.x + 0.16 + stepW * i + stepW / 2;
        slide.addShape('oval', {
          x: cx - dot / 2, y: axisY - dot / 2, w: dot, h: dot,
          fill: { color: theme.white }, line: { color: theme.accent, width: 1.25 },
        });
        slide.addText(it.key, {
          x: cx - stepW / 2, y: axisY + 0.12, w: stepW, h: 0.26,
          fontSize: 8.5, color: theme.text, align: 'center', valign: 'top',
        });
        if (it.value && it.value !== it.key) {
          // 出来事は上側に、隣と重ならないよう1つおきに高さを変える
          const lift = i % 2 ? 0.86 : 0.48;
          slide.addText(it.value, {
            x: cx - stepW * 0.9, y: axisY - lift, w: stepW * 1.8, h: 0.36,
            fontSize: 8.5, color: theme.subtext, align: 'center', valign: 'bottom',
          });
          slide.addShape('line', {
            x: cx, y: axisY - lift + 0.36, w: 0, h: lift - 0.36 - dot / 2,
            line: { color: theme.border, width: HAIRLINE_THIN },
          });
        }
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
    quadrantPriority,
    positioningMap,
    processFlow,
    flowVertical,
    chevronFlow,
    diamondFlow,
    stepArrows,
    cycle,
    cycleQuadrant,
    funnel,
    waterfall,
    pyramid,
    pyramidTiered,
    triangleTiers,
    triangleRelation,
    logicTree,
    beforeAfter,
    timeline,
    timelineVertical,
    yearTimeline,
    comparisonTable,
    decisionRequest,
    actionPlanTable,
    taskStatusTable,
    issueRiskList,
    blockerRequest,
    asIsToBe,
    ganttChart,
    milestoneChevron,
    kpiSummary,
    barChart,
    lineChart,
    pieChart,
  ];

  // テンプレート選択ギャラリー（app.jsのモーダル）の絞り込み軸。
  //   scenes … どの会議・資料で使うか（1パターンが複数のシーンに属してよい）
  //   role   … 資料の中でどの役割を担うスライドか
  // パターン定義そのものには持たせず、ここで一括して付与する。軸を増やしたいときも
  // この表だけを直せばよく、24種を超えるパターン定義に手を入れずに済む。
  const SCENES = ['提案書', '進捗MTG', '報告書', 'キックオフ', '経営・定例報告', '構成図', '汎用ロジック図解', 'データ提示', '研修・説明会', '会社・採用'];
  const ROLES = ['目次', '本文', '比較', '計画', 'KPI', '表', '図解'];

  const PATTERN_META = {
    'title-message': { scenes: ['提案書', '報告書', '進捗MTG', 'キックオフ', '経営・定例報告'], role: '本文' },
    agenda: { scenes: ['提案書', '報告書', 'キックオフ', '進捗MTG'], role: '目次' },
    'box-compare': { scenes: ['提案書', '報告書'], role: '比較' },
    'compare-vertical': { scenes: ['提案書', '報告書'], role: '比較' },
    'comparison-table': { scenes: ['提案書', '報告書'], role: '表' },
    'decision-request': { scenes: ['進捗MTG', '提案書', '経営・定例報告'], role: '比較' },
    'before-after': { scenes: ['提案書', '報告書'], role: '比較' },
    'as-is-to-be': { scenes: ['提案書', '構成図', 'キックオフ'], role: '図解' },
    swot: { scenes: ['提案書', '経営・定例報告'], role: '図解' },
    'matrix-2x2': { scenes: ['提案書', '汎用ロジック図解'], role: '図解' },
    'positioning-map': { scenes: ['提案書', 'データ提示'], role: '図解' },
    'process-flow': { scenes: ['提案書', 'キックオフ', '構成図'], role: '図解' },
    'flow-vertical': { scenes: ['提案書', 'キックオフ', '構成図'], role: '図解' },
    'chevron-flow': { scenes: ['提案書', 'キックオフ', '構成図'], role: '図解' },
    'diamond-flow': { scenes: ['提案書', '汎用ロジック図解'], role: '図解' },
    'step-arrows': { scenes: ['提案書', 'キックオフ', '研修・説明会'], role: '図解' },
    cycle: { scenes: ['提案書', '汎用ロジック図解'], role: '図解' },
    'cycle-quadrant': { scenes: ['提案書', '汎用ロジック図解'], role: '図解' },
    'triangle-tiers': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'triangle-relation': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'quadrant-priority': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'year-timeline': { scenes: ['報告書', '経営・定例報告', '会社・採用'], role: '計画' },
    funnel: { scenes: ['データ提示', '経営・定例報告'], role: '図解' },
    waterfall: { scenes: ['データ提示', '経営・定例報告'], role: 'KPI' },
    pyramid: { scenes: ['提案書', '報告書', '汎用ロジック図解'], role: '図解' },
    'pyramid-tiered': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'logic-tree': { scenes: ['汎用ロジック図解', '報告書'], role: '図解' },
    timeline: { scenes: ['提案書', 'キックオフ', '進捗MTG'], role: '計画' },
    'timeline-vertical': { scenes: ['提案書', 'キックオフ', '進捗MTG'], role: '計画' },
    'gantt-chart': { scenes: ['進捗MTG', 'キックオフ', '報告書'], role: '計画' },
    'milestone-chevron': { scenes: ['進捗MTG', 'キックオフ', '経営・定例報告'], role: '計画' },
    'action-plan-table': { scenes: ['提案書', 'キックオフ', '進捗MTG'], role: '計画' },
    'task-status-table': { scenes: ['進捗MTG', '報告書'], role: '表' },
    'issue-risk-list': { scenes: ['進捗MTG', '報告書'], role: '表' },
    'blocker-request': { scenes: ['進捗MTG'], role: '表' },
    'kpi-summary': { scenes: ['経営・定例報告', 'データ提示', '進捗MTG'], role: 'KPI' },
    'bar-chart': { scenes: ['データ提示', '経営・定例報告'], role: 'KPI' },
    'line-chart': { scenes: ['データ提示', '経営・定例報告'], role: 'KPI' },
    'pie-chart': { scenes: ['データ提示', '経営・定例報告'], role: 'KPI' },
  };

  // テンプレート選択ギャラリーのサムネイル用サンプル。特殊な記法を要求するパターン
  // （ガント・移行図・状況表など）は、その記法どおりのサンプルを持たせないと
  // サムネイルが空になってしまうため個別に定義する。それ以外は共通サンプルを使う。
  const DEFAULT_SAMPLE = ['項目A：説明テキストが入ります', '項目B：説明テキストが入ります', '項目C：説明テキストが入ります'];
  const SAMPLES = {
    'title-message': ['**重要な語句**を含む説明が入ります', '補足の説明テキストが入ります', '3点目の説明が入ります'],
    agenda: ['本日の論点', '現状と課題', 'ご提案', '今後の進め方'],
    'box-compare': ['A案：低コストだが拡張性に課題', 'B案：★推奨 バランスが最も良い', 'C案：高機能だが過剰投資'],
    'compare-vertical': ['初期費用：A社30万円、B社50万円', '運用コスト：A社1万円、B社3万円', '拡張性：★推奨 B社はオプションで拡張可能', 'サポート：B社は電話・チャット対応'],
    'comparison-table': ['初期費用：50万円', '運用コスト：月額3万円', '導入期間：2ヶ月', '拡張性：オプションで拡張可能', 'サポート：電話・チャット対応'],
    'decision-request': ['依頼事項：導入ベンダーをB社に確定することをご承認いただきたい', '期限：2026年3月10日', 'A案：A社CRM｜低コストだが拡張性に課題', 'B案：★推奨 B社CRM｜バランスが最も良い'],
    'before-after': ['現状：手作業で平均15分を要している', '導入後：自動化により平均2分に短縮される'],
    'as-is-to-be': ['As-Is：基幹システム（オンプレミス）／顧客DB（Oracle）', 'To-Be：SaaS基幹システム（クラウド）／顧客DB（PostgreSQL）', '移行ステップ：現行分析｜データ移行｜並行稼働'],
    swot: ['強み：ブランド認知度が高い', '弱み：デジタル対応が遅れている', '機会：市場が拡大している', '脅威：新規参入が増えている'],
    'matrix-2x2': ['重要度高・緊急度高：即時対応', '重要度高・緊急度低：計画的に対応', '重要度低・緊急度高：委任する', '重要度低・緊急度低：対応しない'],
    'positioning-map': ['A社：高価格・高機能', 'B社：中価格・標準機能', 'C社：低価格・限定機能', '自社：中価格・高機能'],
    'process-flow': ['ステップ1：現状分析', 'ステップ2：要件定義', 'ステップ3：設計・開発', 'ステップ4：全社展開'],
    'flow-vertical': ['ステップ1：現状分析', 'ステップ2：要件定義', 'ステップ3：設計・開発', 'ステップ4：試験運用', 'ステップ5：全社展開'],
    cycle: ['Plan：計画を立てる', 'Do：実行する', 'Check：効果を測定する', 'Action：改善する'],
    'cycle-quadrant': ['Plan：計画を立てる', 'Do：実行する', 'Check：効果を測定する', 'Action：改善する'],
    'chevron-flow': ['第1フェーズ：現行分析｜課題整理', '第2フェーズ：要件定義｜設計', '第3フェーズ：開発｜移行'],
    'diamond-flow': ['Segmentation：市場を切り分ける', 'Targeting：狙う市場を選ぶ', 'Positioning：立ち位置を定める'],
    'step-arrows': ['STEP1：現状把握｜関係者ヒアリング｜業務量の実測', 'STEP2：課題設定｜論点の構造化｜優先順位づけ', 'STEP3：施策立案｜打ち手の比較｜実行計画の策定'],
    'triangle-tiers': ['ハイタッチ：個別訪問による深耕', 'ロータッチ：セミナー・研修による支援', 'テックタッチ：オンラインでの自己解決'],
    'triangle-relation': ['中心：三方良し', '企業：持続的な成長', '従業員：働きがいの向上', '顧客：課題解決の実現'],
    'quadrant-priority': ['即時着手する施策：効果が大きく実行も容易', '計画的に進める施策：効果は大きいが時間を要する', '余力があれば行う施策：容易だが効果は限定的', '当面見送る施策：効果が小さく負荷も大きい'],
    'year-timeline': ['1987年：創業', '1996年：株式上場', '2001年：海外進出', '2010年：事業再編', '2015年：新規事業を開始', '2020年：デジタル基盤を刷新'],
    funnel: ['認知：10000件', '検討：4000件', '商談：1200件', '受注：300件'],
    waterfall: ['期首：100', '新規獲得：+40', '解約：-15', '期末：125'],
    pyramid: ['結論：市場拡大を優先すべきである', '根拠1：市場が年10%成長している', '根拠2：自社シェアに伸びしろがある', '根拠3：競合の本格参入前である'],
    'pyramid-tiered': ['ビジョン：業界No.1のサービス品質', '戦略：デジタル基盤への集中投資', '施策：CRM導入と業務標準化'],
    'logic-tree': ['売上：既存顧客と新規顧客に分解', 'コスト：固定費と変動費に分解', '人員：営業部門と管理部門に分解'],
    timeline: ['2026年9月：要件定義', '2026年10月：データ移行', '2026年11月：試験運用', '2026年12月：全社展開'],
    'timeline-vertical': ['2026年9月：要件定義とベンダー選定', '2026年10月：データ移行と初期設定', '2026年11月：一部部署での試験運用', '2026年12月：全社展開と定着支援'],
    'gantt-chart': ['期間：9月｜10月｜11月｜12月', '要件定義：1-2｜1-3', 'データ移行：2-3｜2-3', '試験運用：3-3', '全社展開：3-4'],
    'milestone-chevron': ['M1：要件確定｜9月末', 'M2：契約締結｜10月末', 'M3：試験運用｜11月末', 'M4：全社展開｜12月末'],
    'action-plan-table': ['要件定義を完了する｜山田｜2026年9月末', 'データ移行を実施する｜佐藤｜2026年10月末', '全社展開を行う｜鈴木｜2026年12月末'],
    'task-status-table': ['要件定義書の作成：山田｜完了｜9/10', '業務要件の確定：業務部門｜遅延｜9/30', 'データ移行設計：鈴木｜進行中｜10/15', 'テスト計画策定：田中｜未着手｜10/31'],
    'issue-risk-list': ['要件の認識齟齬：部門間で解釈が異なる｜高｜合同レビューを実施', 'データ移行の品質：重複レコードが存在する｜中｜事前にクレンジング', '現場の習熟度：一時的な生産性低下｜低｜研修で対応'],
    'blocker-request': ['業務要件が確定せず設計に着手できない｜業務部門より3/10までにご回答いただきたい', 'テスト環境が未整備｜情報システム部にて3/15までに構築をお願いしたい'],
    'kpi-summary': ['対応リードタイム：50%短縮', '解約率：20%改善', '商談数：15%増加'],
    'bar-chart': ['A地域：120', 'B地域：85', 'C地域：60', 'D地域：45'],
    'line-chart': ['2023年：100', '2024年：130', '2025年：160', '2026年：190'],
    'pie-chart': ['直販：45%', '代理店：30%', 'オンライン：25%'],
  };

  DocAssist.patternById = {};
  DocAssist.patterns.forEach((p) => {
    const meta = PATTERN_META[p.id] || {};
    p.scenes = meta.scenes || [];
    p.role = meta.role || '本文';
    p.sample = SAMPLES[p.id] || DEFAULT_SAMPLE;
    DocAssist.patternById[p.id] = p;
  });
  DocAssist.patternScenes = SCENES;
  DocAssist.patternRoles = ROLES;
  DocAssist.patternCategories = ['汎用', '比較', 'マトリクス／ポジショニング', 'フロー・プロセス', '構造・ロジック', '変化・対比', '時系列', 'アクションプラン', '数値・グラフ'];

  // ユーザー定義テンプレート（js/templateSpec.js）が同じ描画部品を使えるように、
  // ここで組み立てに使っている共通ヘルパーを公開する。
  // 組み込みパターンと自作テンプレートで見た目の作法を揃えるための共有点。
  DocAssist.patternHelpers = {
    esc,
    emptyBody,
    emphasisHtml,
    bulletTextRuns,
    SQUARE_BULLET,
    splitGroupItems,
    splitFields,
    splitItemSub,
    guessIcon,
    iconSvg,
    addIcon,
    statusTone,
    severityTone,
    statusCls,
    severityCls,
    ICON_KEYS: Object.keys(ICONS),
  };
})();
