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

  // 丸囲み数字。自動採番のラベルとして複数パターンで共有する（9個を超える場合は
  // 呼び出し側で「10.」のような通常の数字にフォールバックする）。
  const CIRCLED_DIGITS = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨'];
  function autoLabel(i) {
    return CIRCLED_DIGITS[i] || `${i + 1}.`;
  }

  // ■ を箇条書きマーカーに使う（NRI報告書の本文で多用される記号）。
  // DocAssist.theme.bulletChar を単一ソースとする（テンプレート読み込み機能が
  // 取り込んだ記号に差し替えられるよう、固定値ではなく毎回テーマから読み直す）。
  function charToBulletCode(ch) {
    const cp = String(ch || '■').codePointAt(0);
    return cp.toString(16).toUpperCase().padStart(4, '0');
  }
  function SQUARE_BULLET() {
    return { code: charToBulletCode(DocAssist.theme && DocAssist.theme.bulletChar), indent: 14 };
  }

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

  // 表の共通スタイル。参考パーツ集の表は「縦罫線を引かず、横だけを細い点線で区切る」
  // 作法で、罫線でマス目を作らずに視線を横方向へ流す。表側（1列目）は淡いブルーで
  // 塗り、見出し行はネイビーのベタ塗りにする。
  // 表を描くパターンはすべてこの関数を通し、スタイルを1箇所で決められるようにする。
  //   opts: { x, y, w, colW, header: [文字列], rows: [[セル]], firstColAccent, zebra, fontSize }
  //   セルは文字列か { text, options } のどちらでもよい（options はそのまま尊重する）。
  const TABLE_BORDER = [
    { type: 'none' },
    { type: 'none' },
    { type: 'dash', color: 'C9CDD6', pt: 0.75 },
    { type: 'none' },
  ];

  function addStyledTable(slide, theme, opts) {
    const o = opts || {};
    const rows = [];
    if (o.header && o.header.length) {
      rows.push(
        o.header.map((h) => ({
          text: typeof h === 'string' ? h : h.text,
          options: Object.assign(
            { bold: true, color: theme.white, fill: { color: theme.primary }, align: 'center', border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] },
            (h && h.options) || {}
          ),
        }))
      );
    }
    (o.rows || []).forEach((row, ri) => {
      rows.push(
        row.map((cell, ci) => {
          const isObj = cell && typeof cell === 'object';
          const text = isObj ? cell.text : cell;
          const given = (isObj && cell.options) || {};
          const base = { border: TABLE_BORDER, valign: 'middle' };
          if (o.firstColAccent && ci === 0) {
            base.fill = { color: theme.light };
            base.color = theme.primary;
            base.bold = true;
          } else if (o.zebra && ri % 2) {
            base.fill = { color: 'F4F5F7' };
          }
          return { text, options: Object.assign(base, given) };
        })
      );
    });
    slide.addTable(rows, {
      x: o.x, y: o.y, w: o.w,
      colW: o.colW,
      fontSize: o.fontSize || 11,
      color: theme.text,
      valign: 'middle',
      autoPage: false,
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
    const bulletChar = (opts && opts.bulletChar) || SQUARE_BULLET();
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

  // 番号付き箇条書き用。PowerPointのネイティブ自動採番（a:buAutoNum）を使うため、
  // 書き出し後にPowerPoint上で行を並べ替えても番号が自動でつき直る。
  function numberedTextRuns(bullets, theme, opts) {
    const fontSize = (opts && opts.fontSize) || 14;
    const spaceAfter = (opts && opts.spaceAfter) || 12;
    const runs = [];
    (bullets || []).forEach((b) => {
      const tokens = A.parseEmphasisTokens(b);
      tokens.forEach((t, ti) => {
        const isLast = ti === tokens.length - 1;
        const runOpts = {
          fontSize,
          color: t.bold ? theme.primary : theme.text,
        };
        if (ti === 0) runOpts.bullet = { type: 'number', style: 'arabicPeriod' };
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

  // 箇条書きの1行を「ラベル：本文」に分割する。A.extractItemsは強調記法（**bold**）を
  // 除去してしまうため、ラベルにも本文にも太字を残せる生の行に対してA.splitKVを直接使う。
  // コロンが無い行はラベル無し（本文のみ）として扱う。
  function splitLabelValue(raw) {
    const kv = A.splitKV(raw);
    if (kv) return { label: kv.key, body: kv.value };
    return { label: null, body: raw };
  }

  // 「文章を並べただけの箇条書き」でも、本文がスライドいっぱいの無地の点リストに
  // ならないよう、行ごとに必ず何かしらの見出し・ラベルを添える共通レイアウト。
  // 「ラベル：本文」の行はラベルをそのまま太字見出しにし、ラベルの無い行（コロン無し）
  // には丸囲み数字を自動で振って見出し代わりにする。タイトル＋メッセージ
  // （フォールバック）と2列箇条書きの両方から使う。
  function labeledListRuns(bullets, theme, opts) {
    const fontSize = (opts && opts.fontSize) || 12;
    const labelFontSize = (opts && opts.labelFontSize) || fontSize + 1;
    const spaceAfter = (opts && opts.spaceAfter) || 10;
    const runs = [];
    (bullets || []).forEach((raw, i) => {
      const { label, body } = splitLabelValue(raw);
      runs.push({
        text: A.stripEmphasis(label || autoLabel(i)),
        options: { fontSize: labelFontSize, bold: true, color: theme.primary, breakLine: true, paraSpaceAfter: 2 },
      });
      const bodyTokens = A.parseEmphasisTokens(body);
      bodyTokens.forEach((t, ti) => {
        const isLast = ti === bodyTokens.length - 1;
        const runOpts = { fontSize, color: t.bold ? theme.primary : theme.text };
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

  function labeledListHtml(bullets) {
    if (!bullets || !bullets.length) return emptyBody();
    return `<div class="pv-lbl-list">${bullets
      .map((raw, i) => {
        const { label, body } = splitLabelValue(raw);
        return `<div class="pv-lbl-item">
          <div class="pv-lbl-label">${esc(A.stripEmphasis(label || autoLabel(i)))}</div>
          <div class="pv-lbl-body">${emphasisHtml(body)}</div>
        </div>`;
      })
      .join('')}</div>`;
  }

  // labeledListRunsの軽量版。見出し・本文を別行にせず、同じ行の先頭に太字のラベルを
  // 置くだけにする（項目数が多く密度を保ちたい場面向け。2列箇条書きから使う）。
  // startIndexは列分割時に丸囲み数字を通し番号にするためのオフセット。
  function inlineLabeledBulletRuns(bullets, theme, opts) {
    const fontSize = (opts && opts.fontSize) || 12;
    const spaceAfter = (opts && opts.spaceAfter) || 10;
    const startIndex = (opts && opts.startIndex) || 0;
    const runs = [];
    (bullets || []).forEach((raw, i) => {
      const { label, body } = splitLabelValue(raw);
      const labelText = A.stripEmphasis(label || autoLabel(startIndex + i));
      runs.push({
        text: labelText + (label ? '：' : ' '),
        options: { fontSize, bold: true, color: theme.primary, bullet: SQUARE_BULLET() },
      });
      const bodyTokens = A.parseEmphasisTokens(body);
      bodyTokens.forEach((t, ti) => {
        const isLast = ti === bodyTokens.length - 1;
        const runOpts = { fontSize, color: t.bold ? theme.primary : theme.text };
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

  function inlineLabeledBulletHtml(bullets, startIndex) {
    if (!bullets || !bullets.length) return '';
    const offset = startIndex || 0;
    return `<ul class="pv-bullets pv-bullets-square">${bullets
      .map((raw, i) => {
        const { label, body } = splitLabelValue(raw);
        const labelText = A.stripEmphasis(label || autoLabel(offset + i));
        return `<li><b class="em">${esc(labelText)}${label ? '：' : ' '}</b>${emphasisHtml(body)}</li>`;
      })
      .join('')}</ul>`;
  }

  // 「見出し：」（コロンの後に何も続かない行）を新しいブロックの区切りとして扱い、
  // 箇条書きを「小見出し＋その配下の項目」の配列に分解する。
  // 小見出し行が無ければ、全体を見出し無しの1ブロックとして返す
  // （＝小見出し記法を使っていない普通の箇条書きでもスコア判定・描画とも壊れない）。
  const HEADING_ONLY_RE = /^(.{1,24})[：:]\s*$/;
  function parseHeadingGroups(bullets) {
    const groups = [];
    let current = null;
    (bullets || []).forEach((b) => {
      const stripped = A.stripEmphasis(b).trim();
      const m = stripped.match(HEADING_ONLY_RE);
      if (m) {
        current = { heading: m[1].trim(), items: [] };
        groups.push(current);
      } else {
        if (!current) {
          current = { heading: '', items: [] };
          groups.push(current);
        }
        current.items.push(b);
      }
    });
    return groups.filter((g) => g.heading || g.items.length);
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
    description: '補足の箇条書きのみのシンプルな本文。どんな内容にも使える汎用のフォールバック。「ラベル：本文」の行はラベルを見出しにし、ラベルの無い行には丸囲み数字を自動で振る。',
    score(section) {
      const n = (section.bullets || []).length;
      return n <= 4 ? 3 : 2;
    },
    renderBody(bullets) {
      return labeledListHtml(bullets);
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      slide.addText(labeledListRuns(bullets, theme), { x: box.x, y: box.y, w: box.w, h: box.h, valign: 'top' });
    },
  };

  // ---------- 番号付き箇条書き ----------
  // 手順ではなく「条項・論点を順序立てて列挙する」ための、報告書本文でよく使う番号リスト。
  // PowerPointのネイティブ自動採番を使うため、書き出し後に行を並べ替えても番号がずれない。
  const bulletNumbered = {
    id: 'bullet-numbered',
    name: '番号付き箇条書き',
    category: '汎用',
    description: '箇条書きに1.2.3…の番号を振る。手順ではなく、論点や条項を順序立てて列挙したいときに使う。',
    score(section) {
      const bullets = section.bullets || [];
      const numbered = A.numberedBulletCount(bullets);
      if (bullets.length >= 2 && numbered >= Math.ceil(bullets.length * 0.6)) return 8;
      return 0;
    },
    renderBody(bullets) {
      if (!bullets.length) return emptyBody();
      return `<ol class="pv-bullets pv-bullets-numbered">${bullets.map((b) => `<li>${emphasisHtml(b)}</li>`).join('')}</ol>`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      slide.addText(numberedTextRuns(bullets, theme), { x: box.x, y: box.y, w: box.w, h: box.h, valign: 'top' });
    },
  };

  // ---------- 小見出し＋箇条書き（複数ブロック） ----------
  // 1つの論点の中に「背景」「課題」「対策」のような複数の小節がある場合に、
  // 小見出しごとにブロックを縦に積んで示す。行末が全角/半角コロンだけで終わる行
  // （例：「背景：」）を小見出しとして扱い、それ以降の行を配下の箇条書きとする。
  const bulletHeadingGroups = {
    id: 'bullet-heading-groups',
    name: '小見出し＋箇条書き（複数ブロック）',
    category: '汎用',
    description: '1枚の中を「背景：」のような小見出しで複数ブロックに分け、それぞれに箇条書きを添えて縦に並べる。小見出し行は行末をコロンだけにして書く。',
    score(section) {
      const groups = parseHeadingGroups(section.bullets || []);
      const valid = groups.filter((g) => g.heading && g.items.length);
      if (valid.length >= 2) return 9;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return emptyBody();
      return `<div class="pv-hg">${groups
        .map(
          (g) => `<div class="pv-hg-group">
          <div class="pv-hg-heading">${esc(g.heading)}</div>
          <ul class="pv-bullets pv-bullets-square">${g.items.map((it) => `<li>${emphasisHtml(it)}</li>`).join('')}</ul>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return;
      const n = groups.length;
      const gap = 0.12;
      const rowH = (box.h - gap * (n - 1)) / n;
      groups.forEach((g, i) => {
        const y = box.y + i * (rowH + gap);
        slide.addText(g.heading, {
          x: box.x, y, w: box.w, h: 0.3,
          fontSize: 12, bold: true, color: theme.primary,
        });
        slide.addText(bulletTextRuns(g.items, theme, { fontSize: 11, spaceAfter: 4 }), {
          x: box.x + 0.1, y: y + 0.3, w: box.w - 0.1, h: rowH - 0.3, valign: 'top',
        });
      });
    },
  };

  // ---------- 小見出し＋箇条書き（2カラム） ----------
  // 上と同じ小見出し記法を使うが、ちょうど2ブロックのときに左右へ並べる版。
  // 縦積み版より1ブロックあたりの縦スペースを確保できるため、各ブロックの
  // 項目数が多いときに読みやすい。
  const headingBullets2Col = {
    id: 'heading-bullets-2col',
    name: '小見出し＋箇条書き（2カラム）',
    category: '汎用',
    description: '「背景：」のような小見出しで2ブロックに分け、左右に並べて箇条書きを添える。ちょうど2つの観点を並べたいときに使う。',
    score(section) {
      const groups = parseHeadingGroups(section.bullets || []);
      const valid = groups.filter((g) => g.heading && g.items.length);
      if (valid.length === 2) return 8;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseHeadingGroups(bullets)
        .filter((g) => g.heading && g.items.length)
        .slice(0, 2);
      if (groups.length < 2) return emptyBody();
      return `<div class="pv-hg2">${groups
        .map(
          (g) => `<div class="pv-hg2-col">
          <div class="pv-hg-heading">${esc(g.heading)}</div>
          <ul class="pv-bullets pv-bullets-square">${g.items.map((it) => `<li>${emphasisHtml(it)}</li>`).join('')}</ul>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseHeadingGroups(bullets)
        .filter((g) => g.heading && g.items.length)
        .slice(0, 2);
      if (groups.length < 2) return;
      const gap = 0.3;
      const colW = (box.w - gap) / 2;
      groups.forEach((g, i) => {
        const x = box.x + i * (colW + gap);
        slide.addText(g.heading, {
          x, y: box.y, w: colW, h: 0.3,
          fontSize: 12, bold: true, color: theme.primary,
        });
        slide.addText(bulletTextRuns(g.items, theme, { fontSize: 11, spaceAfter: 6 }), {
          x: x + 0.05, y: box.y + 0.34, w: colW - 0.05, h: box.h - 0.34, valign: 'top',
        });
      });
    },
  };

  // ---------- 2列箇条書き ----------
  // 見出しの無い、フラットな箇条書きが8件以上あるときに、左右2列に均等分割して
  // 密度を上げる。項目数が多いだけで意味的なグルーピングは無い場合に使う。
  const bulletTwoCol = {
    id: 'bullet-two-col',
    name: '2列箇条書き',
    category: '汎用',
    description: '見出しの無いフラットな箇条書きを、項目数が多いときに左右2列に均等分割して密度を上げる。「ラベル：本文」の行はラベルを太字にし、ラベルの無い行には丸囲み数字を自動で振る。',
    score(section) {
      const bullets = section.bullets || [];
      const n = bullets.length;
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading);
      // 小見出し記法を使っている場合はそちら（bullet-heading-groups等）を優先する
      if (groups.length) return 0;
      if (n >= 8 && n <= 16) return 6;
      return 0;
    },
    renderBody(bullets) {
      if (!bullets.length) return emptyBody();
      const mid = Math.ceil(bullets.length / 2);
      const cols = [bullets.slice(0, mid), bullets.slice(mid)];
      return `<div class="pv-bc2">${cols.map((col, i) => inlineLabeledBulletHtml(col, i * mid)).join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      if (!bullets.length) return;
      const mid = Math.ceil(bullets.length / 2);
      const cols = [bullets.slice(0, mid), bullets.slice(mid)];
      const gap = 0.3;
      const colW = (box.w - gap) / 2;
      cols.forEach((col, i) => {
        if (!col.length) return;
        const x = box.x + i * (colW + gap);
        slide.addText(inlineLabeledBulletRuns(col, theme, { fontSize: 12, spaceAfter: 10, startIndex: i * mid }), {
          x, y: box.y, w: colW, h: box.h, valign: 'top',
        });
      });
    },
  };

  // ---------- チェックリスト ----------
  // 確認事項・合意事項を四角いチェックボックス付きで並べる。行頭に「★」を付けた項目は
  // 「対応済み／確認済み」として塗りつぶしたチェックボックスで示す（他の比較パターンと
  // 同じ★記法の再利用。ここでは「推奨」ではなく「完了」の意味で使う）。
  const checklist = {
    id: 'checklist',
    name: 'チェックリスト',
    category: '汎用',
    description: '確認事項・合意事項をチェックボックス付きで並べる。行頭に「★」を付けた項目はチェック済みとして示す。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['チェックリスト', '確認事項', '確認事項一覧', 'ご確認', 'チェック項目']);
      const n = (section.bullets || []).length;
      if (kw && n >= 2 && n <= 8) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return emptyBody();
      return `<div class="pv-cl">${items
        .map(
          (it) => `<div class="pv-cl-row${it.highlight ? ' is-checked' : ''}">
          <span class="pv-cl-box">${it.highlight ? '✓' : ''}</span>
          <span class="pv-cl-text">${esc(it.value !== it.key ? `${it.key}：${it.value}` : it.key)}</span>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 8);
      if (!items.length) return;
      const n = items.length;
      const gap = 0.1;
      const rowH = (box.h - gap * (n - 1)) / n;
      const boxSize = Math.min(0.3, rowH * 0.6);
      items.forEach((it, i) => {
        const y = box.y + i * (rowH + gap) + (rowH - boxSize) / 2;
        slide.addShape('rect', {
          x: box.x, y, w: boxSize, h: boxSize,
          fill: { color: it.highlight ? theme.primary : theme.white },
          line: { color: theme.primary, width: 1.25 },
        });
        if (it.highlight) {
          slide.addText('✓', {
            x: box.x, y, w: boxSize, h: boxSize,
            fontSize: 12, bold: true, color: theme.white, align: 'center', valign: 'middle',
          });
        }
        const label = it.value !== it.key ? `${it.key}：${it.value}` : it.key;
        slide.addText(label, {
          x: box.x + boxSize + 0.14, y: box.y + i * (rowH + gap), w: box.w - boxSize - 0.14, h: rowH,
          fontSize: 12, color: theme.text, valign: 'middle',
        });
      });
    },
  };

  // ---------- ラベルボックス＋箇条書き（横並び） ----------
  // 「背景：」のような小見出し記法を、見出しをネイビーの塗りボックスにして左に置き、
  // 右側に箇条書きを添える横並びの行として縦に積む。小見出し＋箇条書き（複数ブロック）の
  // 見た目違いバリエーション（parseHeadingGroupsをそのまま再利用）。
  const labelBoxBulletRows = {
    id: 'label-box-bullet-rows',
    name: 'ラベルボックス＋箇条書き（横並び）',
    category: '汎用',
    description: '「背景：」のような小見出しを塗りボックスにして左に置き、右側に箇条書きを添えて横並びの行として積む。小見出し＋箇条書きの見た目違いバリエーション。',
    score(section) {
      const groups = parseHeadingGroups(section.bullets || []);
      const valid = groups.filter((g) => g.heading && g.items.length);
      if (valid.length >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return emptyBody();
      return `<div class="pv-lbr">${groups
        .map(
          (g) => `<div class="pv-lbr-row">
          <div class="pv-lbr-box">${esc(g.heading)}</div>
          <ul class="pv-lbr-content pv-bullets pv-bullets-square">${g.items.map((it) => `<li>${emphasisHtml(it)}</li>`).join('')}</ul>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return;
      const n = groups.length;
      const gap = 0.22;
      const rowH = (box.h - gap * (n - 1)) / n;
      const boxW = Math.min(1.9, box.w * 0.22);
      groups.forEach((g, i) => {
        const y = box.y + i * (rowH + gap);
        slide.addShape('rect', { x: box.x, y, w: boxW, h: rowH, fill: { color: theme.primary }, line: { type: 'none' } });
        slide.addText(g.heading, {
          x: box.x + 0.1, y, w: boxW - 0.2, h: rowH,
          fontSize: 13, bold: true, color: theme.white, align: 'center', valign: 'middle',
        });
        slide.addText(bulletTextRuns(g.items, theme, { fontSize: 11.5, spaceAfter: 6 }), {
          x: box.x + boxW + 0.24, y, w: box.w - boxW - 0.24, h: rowH, valign: 'middle',
        });
      });
    },
  };

  // ---------- ノート風見出し＋箇条書き（サブ項目付き） ----------
  // 外枠付きのノート風コンテナに、プレーンな見出し文字＋箇条書きを並べる。行頭が
  // 「-」「－」「–」で始まる行は、直前の箇条書きに対するサブ項目として一段深く
  // インデントして表示する（メモ書きのような細かい内訳を残したいときに使う）。
  const SUB_ITEM_RE = /^[-－–]\s*(.+)/;
  function parseNestedHeadingGroups(bullets) {
    return parseHeadingGroups(bullets).map((g) => {
      const items = [];
      g.items.forEach((raw) => {
        const m = A.stripEmphasis(raw).trim().match(SUB_ITEM_RE);
        if (m && items.length) {
          const rawMatch = raw.match(SUB_ITEM_RE);
          items[items.length - 1].subs.push(rawMatch ? rawMatch[1] : m[1]);
        } else {
          items.push({ text: raw, subs: [] });
        }
      });
      return { heading: g.heading, items };
    });
  }
  function hasNestedSubItems(groups) {
    return groups.some((g) => g.items.some((it) => it.subs.length));
  }

  const notebookHeadingBullets = {
    id: 'notebook-heading-bullets',
    name: 'ノート風見出し＋箇条書き（サブ項目付き）',
    category: '汎用',
    description: '外枠付きのノート風レイアウトで、見出し文字＋箇条書きを並べる。行頭を「-」で始めると直前の項目のサブ項目として一段深くインデントされる。',
    score(section) {
      const groups = parseNestedHeadingGroups(section.bullets || []);
      const valid = groups.filter((g) => g.heading && g.items.length);
      if (valid.length >= 2 && hasNestedSubItems(valid)) return 9;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseNestedHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return emptyBody();
      return `<div class="pv-nb">${groups
        .map(
          (g) => `<div class="pv-nb-group">
          <div class="pv-nb-heading">${esc(g.heading)}</div>
          <ul class="pv-nb-list">${g.items
            .map(
              (it) => `<li>${emphasisHtml(it.text)}${
                it.subs.length ? `<ul class="pv-nb-sublist">${it.subs.map((s) => `<li>${emphasisHtml(s)}</li>`).join('')}</ul>` : ''
              }</li>`
            )
            .join('')}</ul>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseNestedHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return;
      slide.addShape('rect', {
        x: box.x, y: box.y, w: box.w, h: box.h, fill: { type: 'none' }, line: { color: theme.border, width: HAIRLINE },
      });
      const pad = 0.18;
      const innerX = box.x + pad;
      const innerW = box.w - pad * 2;
      const n = groups.length;
      const gap = 0.14;
      const groupH = (box.h - pad * 2 - gap * (n - 1)) / n;
      groups.forEach((g, i) => {
        const y = box.y + pad + i * (groupH + gap);
        slide.addText(g.heading, { x: innerX, y, w: innerW, h: 0.32, fontSize: 14, bold: true, color: theme.primary });
        const runs = [];
        g.items.forEach((it) => {
          const tokens = A.parseEmphasisTokens(it.text);
          tokens.forEach((t, ti) => {
            const isLast = ti === tokens.length - 1;
            const runOpts = { fontSize: 12, bullet: SQUARE_BULLET(), color: t.bold ? theme.primary : theme.text };
            if (t.bold) runOpts.bold = true;
            if (isLast) { runOpts.breakLine = true; runOpts.paraSpaceAfter = 3; }
            runs.push({ text: t.text, options: runOpts });
          });
          it.subs.forEach((s, si) => {
            const subTokens = A.parseEmphasisTokens(s);
            subTokens.forEach((t, ti) => {
              const isLast = ti === subTokens.length - 1;
              const runOpts = { fontSize: 12, indentLevel: 1, bullet: { code: '2013', indent: 14 }, color: t.bold ? theme.primary : theme.text };
              if (t.bold) runOpts.bold = true;
              if (isLast) { runOpts.breakLine = true; runOpts.paraSpaceAfter = si === it.subs.length - 1 ? 5 : 2; }
              runs.push({ text: t.text, options: runOpts });
            });
          });
        });
        slide.addText(runs, { x: innerX, y: y + 0.34, w: innerW, h: groupH - 0.34, valign: 'top' });
      });
    },
  };

  // ---------- フラッグ見出し＋枠線ボックス ----------
  // 「背景：」のような小見出しを、右端がとがったフラッグ型の見出しバーにし、下に
  // 枠線だけのボックスで箇条書きを添える。ボックスごとに独立したカード感を出したい
  // ときに使う、小見出し＋箇条書きの見た目違いバリエーション。
  const flagHeadingBoxGroups = {
    id: 'flag-heading-box-groups',
    name: 'フラッグ見出し＋枠線ボックス',
    category: '汎用',
    description: '「背景：」のような小見出しを右端のとがったフラッグ型バーにし、下に枠線ボックスで箇条書きを添えて縦に積む。小見出し＋箇条書きの見た目違いバリエーション。',
    score(section) {
      const groups = parseHeadingGroups(section.bullets || []);
      const valid = groups.filter((g) => g.heading && g.items.length);
      if (valid.length >= 2) return 6;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return emptyBody();
      return `<div class="pv-fb">${groups
        .map(
          (g) => `<div class="pv-fb-group">
          <div class="pv-fb-flag">${esc(g.heading)}</div>
          <div class="pv-fb-box">
            <ul class="pv-bullets pv-bullets-square">${g.items.map((it) => `<li>${emphasisHtml(it)}</li>`).join('')}</ul>
          </div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length);
      if (!groups.length) return;
      const n = groups.length;
      const gap = 0.18;
      const rowH = (box.h - gap * (n - 1)) / n;
      const flagH = 0.36;
      groups.forEach((g, i) => {
        const y = box.y + i * (rowH + gap);
        const flagW = Math.min(box.w * 0.32, 2.6);
        slide.addShape('homePlate', { x: box.x, y, w: flagW, h: flagH, fill: { color: theme.primary }, line: { type: 'none' } });
        slide.addText(g.heading, {
          x: box.x + 0.12, y, w: flagW - 0.34, h: flagH, fontSize: 13, bold: true, color: theme.white, valign: 'middle',
        });
        slide.addShape('rect', {
          x: box.x, y: y + flagH, w: box.w, h: rowH - flagH,
          fill: { type: 'none' }, line: { color: theme.primary, width: HAIRLINE },
        });
        slide.addText(bulletTextRuns(g.items, theme, { fontSize: 11.5, spaceAfter: 5 }), {
          x: box.x + 0.16, y: y + flagH + 0.08, w: box.w - 0.32, h: rowH - flagH - 0.16, valign: 'top',
        });
      });
    },
  };

  // ---------- 背景→課題→対策（ボックス＋矢印、横並び） ----------
  // 「背景：」のような小見出しを2〜4個のボックス見出しにして横に並べ、間を矢印で
  // つなぐ。各ボックスの下にチェックマーク付きの箇条書きを添える。段階を追って
  // 読ませたい「背景→課題→対策」のような流れに使う、横型バリエーション。
  const chevronBoxGroups = {
    id: 'chevron-box-groups',
    name: '背景→課題→対策（ボックス＋矢印、横並び）',
    category: 'フロー・プロセス',
    description: '「背景：」のような小見出しを2〜4個のボックス見出しにして横に並べ、矢印でつなぐ。各ボックスの下にチェックマーク付きの箇条書きを添える。',
    score(section) {
      const groups = parseHeadingGroups(section.bullets || []);
      const valid = groups.filter((g) => g.heading && g.items.length);
      if (valid.length >= 2 && valid.length <= 4) return 5;
      return 0;
    },
    renderBody(bullets) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length).slice(0, 4);
      if (!groups.length) return emptyBody();
      return `<div class="pv-cbg">${groups
        .map(
          (g, i) => `${i > 0 ? '<div class="pv-cbg-arrow"></div>' : ''}<div class="pv-cbg-col">
          <div class="pv-cbg-head">${esc(g.heading)}</div>
          <ul class="pv-bullets pv-bullets-check">${g.items.map((it) => `<li>${emphasisHtml(it)}</li>`).join('')}</ul>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const groups = parseHeadingGroups(bullets).filter((g) => g.heading && g.items.length).slice(0, 4);
      if (!groups.length) return;
      const n = groups.length;
      const arrowW = 0.3;
      const colW = (box.w - arrowW * (n - 1)) / n;
      const headH = 0.5;
      groups.forEach((g, i) => {
        const x = box.x + i * (colW + arrowW);
        slide.addShape('rect', { x, y: box.y, w: colW, h: headH, fill: { color: theme.primary }, line: { type: 'none' } });
        slide.addText(g.heading, {
          x, y: box.y, w: colW, h: headH, fontSize: 13, bold: true, color: theme.white, align: 'center', valign: 'middle',
        });
        const checkRuns = [];
        g.items.forEach((it) => {
          const tokens = A.parseEmphasisTokens(it);
          tokens.forEach((t, ti) => {
            const isLast = ti === tokens.length - 1;
            const runOpts = { fontSize: 12, bullet: { code: '2713', indent: 16 }, color: t.bold ? theme.primary : theme.text };
            if (t.bold) runOpts.bold = true;
            if (isLast) { runOpts.breakLine = true; runOpts.paraSpaceAfter = 6; }
            checkRuns.push({ text: t.text, options: runOpts });
          });
        });
        slide.addText(checkRuns, { x, y: box.y + headH + 0.14, w: colW, h: box.h - headH - 0.14, valign: 'top' });
        if (i < n - 1) {
          slide.addShape('triangle', {
            x: x + colW + arrowW / 2 - 0.09, y: box.y + headH / 2 - 0.09, w: 0.18, h: 0.18,
            rotate: 90, fill: { color: theme.primary }, line: { type: 'none' },
          });
        }
      });
    },
  };

  // ---------- 前回の振り返り×本日の討議内容 ----------
  // 左に前回までの振り返り（見出し＋ネスト箇条書き）、右に本日の討議内容
  // （番号付きボックスの一覧）を並べる、MTG冒頭でよく使う2カラム構成。
  // 箇条書きの中に区切り行（「---」「＝＝＝」など3文字以上の記号だけの行）を
  // 1つ入れると、それより前が左側、後が右側として扱われる。各側の1行目は
  // 小見出しとして使う。左側は行頭を「-」で始めると一段深くインデントされる
  // （ノート風見出し＋箇条書きと同じサブ項目記法）。右側の番号は自動で振られる。
  const RECAP_DIVIDER_RE = /^[＝=\-－―ー]{3,}$/;
  function parseRecapAgenda(bullets) {
    const list = bullets || [];
    const dividerIdx = list.findIndex((b) => RECAP_DIVIDER_RE.test(A.stripEmphasis(b).trim()));
    if (dividerIdx < 1 || dividerIdx >= list.length - 1) return null;
    const leftLines = list.slice(0, dividerIdx);
    const rightLines = list.slice(dividerIdx + 1);
    if (leftLines.length < 2 || rightLines.length < 2) return null;
    const leftHeading = leftLines[0];
    const leftItems = [];
    leftLines.slice(1).forEach((raw) => {
      const m = A.stripEmphasis(raw).trim().match(SUB_ITEM_RE);
      if (m && leftItems.length) {
        const rawMatch = raw.match(SUB_ITEM_RE);
        leftItems[leftItems.length - 1].subs.push(rawMatch ? rawMatch[1] : m[1]);
      } else {
        leftItems.push({ text: raw, subs: [] });
      }
    });
    const rightHeading = rightLines[0];
    const rightGroups = parseHeadingGroups(rightLines.slice(1)).filter((g) => g.heading && g.items.length);
    if (!leftItems.length || !rightGroups.length) return null;
    return { leftHeading, leftItems, rightHeading, rightGroups };
  }

  const recapAgendaSplit = {
    id: 'recap-agenda-split',
    name: '前回の振り返り×本日の討議内容',
    category: '汎用',
    description: '左に前回までの振り返り（見出し＋ネスト箇条書き）、右に本日の討議内容（番号付きボックス）を並べる。箇条書きの中に「---」のような区切り行を1つ入れると、それより前が左側、後が右側になる。各側の1行目は小見出しとして使う。左側は行頭を「-」で始めると一段深くインデントされる。',
    score(section) {
      const parsed = parseRecapAgenda(section.bullets || []);
      if (!parsed) return 0;
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['振り返り', '前回', '討議内容', 'アジェンダ']);
      return kw ? 10 : 7;
    },
    renderBody(bullets) {
      const parsed = parseRecapAgenda(bullets);
      if (!parsed) return emptyBody();
      const { leftHeading, leftItems, rightHeading, rightGroups } = parsed;
      const groups = rightGroups.slice(0, CIRCLED_DIGITS.length);
      return `<div class="pv-ra">
        <div class="pv-ra-col pv-ra-left">
          <div class="pv-ra-subhead">${esc(leftHeading)}</div>
          <ul class="pv-ra-outline">${leftItems
            .map(
              (it) => `<li>${emphasisHtml(it.text)}${
                it.subs.length ? `<ul class="pv-ra-sub">${it.subs.map((s) => `<li>${emphasisHtml(s)}</li>`).join('')}</ul>` : ''
              }</li>`
            )
            .join('')}</ul>
        </div>
        <div class="pv-ra-col pv-ra-right">
          <div class="pv-ra-subhead pv-ra-subhead-center">${esc(rightHeading)}</div>
          ${groups
            .map(
              (g, i) => `<div class="pv-ra-box">
            <div class="pv-ra-box-head">${CIRCLED_DIGITS[i]}${esc(g.heading)}</div>
            <ul class="pv-bullets pv-bullets-square">${g.items.map((it) => `<li>${emphasisHtml(it)}</li>`).join('')}</ul>
          </div>`
            )
            .join('')}
        </div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const parsed = parseRecapAgenda(bullets);
      if (!parsed) return;
      const { leftHeading, leftItems, rightHeading, rightGroups } = parsed;
      const gap = 0.3;
      const colW = (box.w - gap) / 2;

      slide.addText(leftHeading, { x: box.x, y: box.y, w: colW, h: 0.3, fontSize: 13, bold: true, color: theme.text });
      slide.addShape('line', { x: box.x, y: box.y + 0.32, w: colW, h: 0, line: { color: theme.border, width: HAIRLINE } });
      const runs = [];
      leftItems.forEach((it) => {
        const tokens = A.parseEmphasisTokens(it.text);
        tokens.forEach((t, ti) => {
          const isLast = ti === tokens.length - 1;
          const runOpts = { fontSize: 12, bullet: SQUARE_BULLET(), color: t.bold ? theme.primary : theme.text };
          if (t.bold) runOpts.bold = true;
          if (isLast) { runOpts.breakLine = true; runOpts.paraSpaceAfter = 4; }
          runs.push({ text: t.text, options: runOpts });
        });
        it.subs.forEach((s, si) => {
          const subTokens = A.parseEmphasisTokens(s);
          subTokens.forEach((t, ti) => {
            const isLast = ti === subTokens.length - 1;
            const runOpts = { fontSize: 12, indentLevel: 1, bullet: { code: '2013', indent: 14 }, color: t.bold ? theme.primary : theme.text };
            if (t.bold) runOpts.bold = true;
            if (isLast) { runOpts.breakLine = true; runOpts.paraSpaceAfter = si === it.subs.length - 1 ? 6 : 2; }
            runs.push({ text: t.text, options: runOpts });
          });
        });
      });
      slide.addText(runs, { x: box.x, y: box.y + 0.42, w: colW, h: box.h - 0.42, valign: 'top' });

      const rx = box.x + colW + gap;
      slide.addText(rightHeading, { x: rx, y: box.y, w: colW, h: 0.3, fontSize: 13, bold: true, color: theme.text, align: 'center' });
      const groups = rightGroups.slice(0, CIRCLED_DIGITS.length);
      const n = groups.length;
      const boxGap = 0.14;
      const boxY0 = box.y + 0.42;
      const boxH = (box.h - 0.42 - boxGap * (n - 1)) / n;
      groups.forEach((g, i) => {
        const y = boxY0 + i * (boxH + boxGap);
        slide.addShape('rect', { x: rx, y, w: colW, h: boxH, fill: { color: theme.lighter }, line: { type: 'none' } });
        slide.addShape('rect', { x: rx, y, w: colW, h: 0.34, fill: { color: theme.primaryLight }, line: { type: 'none' } });
        slide.addText(`${CIRCLED_DIGITS[i]}${g.heading}`, {
          x: rx + 0.12, y, w: colW - 0.24, h: 0.34, fontSize: 12, bold: true, color: theme.white, valign: 'middle',
        });
        slide.addText(bulletTextRuns(g.items, theme, { fontSize: 11, spaceAfter: 4 }), {
          x: rx + 0.14, y: y + 0.4, w: colW - 0.28, h: boxH - 0.46, valign: 'top',
        });
      });
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
          [{ text: it.value, options: { bullet: SQUARE_BULLET() } }],
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
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [box.w * 0.3, box.w * 0.7],
        fontSize: 12,
        header: ['項目', '内容'],
        firstColAccent: true,
        rows: items.map((it) => [
          {
            text: `${it.highlight ? '★ ' : ''}${it.key}`,
            options: it.highlight ? { color: theme.highlight } : {},
          },
          { text: it.value, options: it.highlight ? { bold: true, fill: { color: theme.lighter } } : {} },
        ]),
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
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [box.w * 0.08, box.w * 0.5, box.w * 0.2, box.w * 0.22],
        fontSize: 11.5,
        header: ['No', 'タスク', '担当', '期限'],
        firstColAccent: true,
        zebra: true,
        rows: items.map((it, i) => [
          { text: String(i + 1), options: { align: 'center' } },
          { text: it.task },
          { text: it.owner, options: { align: 'center' } },
          { text: it.due, options: { align: 'center' } },
        ]),
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
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [box.w * 0.46, box.w * 0.16, box.w * 0.18, box.w * 0.2],
        fontSize: 11,
        header: ['タスク', '担当', '状況', '期限'],
        firstColAccent: true,
        zebra: true,
        rows: items.map((it) => {
          const f = splitFields(it.value);
          const status = f[1] || '';
          return [
            { text: it.key },
            { text: f[0] || '', options: { align: 'center' } },
            { text: status, options: { fill: { color: statusTone(status, theme) }, color: theme.white, bold: true, align: 'center' } },
            { text: f[2] || '', options: { align: 'center' } },
          ];
        }),
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
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [box.w * 0.24, box.w * 0.38, box.w * 0.12, box.w * 0.26],
        fontSize: 10.5,
        header: ['課題・リスク', '内容', '影響度', '対応方針'],
        firstColAccent: true,
        zebra: true,
        rows: items.map((it) => {
          const f = splitFields(it.value);
          const sev = f[1] || '';
          return [
            { text: it.key },
            { text: f[0] || '' },
            { text: sev, options: { fill: { color: severityTone(sev, theme) }, color: theme.white, bold: true, align: 'center' } },
            { text: f[2] || '' },
          ];
        }),
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

  // ---------- 課題分類×内容→対応案（4列表） ----------
  // 課題をカテゴリー別にグループ化し、「課題」「課題内容」から「対応案」への
  // 流れを矢印で示す4列の表。1行目に「列：カテゴリー｜課題｜課題内容｜対応案」、
  // 以降を「カテゴリー：課題｜内容1／内容2｜対応1／対応2」で書く（parseCrossTableを
  // 再利用。「／」区切りで1セル内に複数の箇条書きを入れられる）。同じカテゴリーが
  // 連続する行は自動的に縦結合される。
  const categorizedIssueResponseTable = {
    id: 'categorized-issue-response-table',
    name: '課題分類×内容→対応案（4列表）',
    category: 'アクションプラン',
    description: '課題をカテゴリー別にグループ化し、「課題内容」から「対応案」への流れを矢印で示す4列の表。1行目に「列：カテゴリー｜課題｜課題内容｜対応案」、以降を「カテゴリー：課題｜内容1／内容2｜対応1／対応2」で書く。同じカテゴリーが連続する行は自動的に結合される。',
    score(section) {
      const t = parseCrossTable(section.bullets || []);
      const rows = t.rows.filter((r) => r.cells.length >= 2);
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['課題内容', '対応案', '対応方針']);
      if (t.cols.length >= 3 && rows.length >= 2 && kw) return 10;
      return 0;
    },
    renderBody(bullets) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return emptyBody();
      const headers = t.cols.length ? t.cols : ['カテゴリー', '課題', '課題内容', '対応案'];
      const rows = t.rows.map((r) => rowCells(r, headers.length));
      const spans = rows.map((cells, i) => {
        if (i > 0 && rows[i - 1][0] === cells[0]) return 0;
        let span = 1;
        for (let j = i + 1; j < rows.length && rows[j][0] === cells[0]; j++) span++;
        return span;
      });
      return `<table class="pv-table pv-cirt"><thead><tr><th>${esc(headers[0])}</th><th>${esc(headers[1])}</th><th>${esc(
        headers[2]
      )}</th><th class="pv-cirt-arrow-th"></th><th>${esc(headers[3] || '対応案')}</th></tr></thead><tbody>
        ${rows
          .map((cells, i) => {
            const [cat, issue, content, response] = cells;
            const contentItems = splitGroupItems(content);
            const responseItems = splitGroupItems(response);
            return `<tr>
            ${spans[i] ? `<td class="pv-cirt-cat" rowspan="${spans[i]}">${esc(cat)}</td>` : ''}
            <td class="pv-cirt-issue">${esc(issue)}</td>
            <td><ul class="pv-bullets pv-bullets-square">${contentItems.map((c) => `<li>${esc(c)}</li>`).join('')}</ul></td>
            ${i === 0 ? `<td class="pv-cirt-arrow-td" rowspan="${rows.length}"></td>` : ''}
            <td><ul class="pv-bullets pv-bullets-square">${responseItems.map((c) => `<li>${esc(c)}</li>`).join('')}</ul></td>
          </tr>`;
          })
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return;
      const headers = t.cols.length ? t.cols : ['カテゴリー', '課題', '課題内容', '対応案'];
      const rows = t.rows.map((r) => rowCells(r, headers.length));
      const spans = rows.map((cells, i) => {
        if (i > 0 && rows[i - 1][0] === cells[0]) return 0;
        let span = 1;
        for (let j = i + 1; j < rows.length && rows[j][0] === cells[0]; j++) span++;
        return span;
      });
      const arrowW = 0.34;
      const catW = box.w * 0.14;
      const issueW = box.w * 0.2;
      const contentW = (box.w - catW - issueW - arrowW) / 2;
      const bulletJoin = (s) => splitGroupItems(s).map((v) => `${(theme.bulletChar || '■')} ${v}`).join('\n');

      const tableRows = rows.map((cells, i) => {
        const [cat, issue, content, response] = cells;
        const row = [];
        if (spans[i]) {
          row.push({ text: cat, options: { rowspan: spans[i], fill: { color: theme.primary }, color: theme.white, bold: true, align: 'center', valign: 'middle' } });
        }
        row.push({ text: issue, options: { fill: { color: theme.primaryLight }, color: theme.white, bold: true, valign: 'middle' } });
        row.push({ text: bulletJoin(content), options: { valign: 'top' } });
        if (i === 0) {
          row.push({
            text: '', options: {
              rowspan: rows.length, fill: { type: 'none' },
              border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }],
            },
          });
        }
        row.push({ text: bulletJoin(response), options: { valign: 'top' } });
        return row;
      });

      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [catW, issueW, contentW, arrowW, contentW],
        fontSize: 11,
        header: [headers[0] || 'カテゴリー', headers[1] || '課題', headers[2] || '課題内容', '', headers[3] || '対応案'],
        rows: tableRows,
      });

      const arrowX = box.x + catW + issueW + contentW;
      slide.addShape('triangle', {
        x: arrowX + arrowW / 2 - 0.12, y: box.y + box.h / 2 - 0.16, w: 0.24, h: 0.32,
        rotate: 90, fill: { color: theme.gray }, line: { type: 'none' },
      });
    },
  };

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

  // ---------- クロス表（表頭×表側） ----------
  // 1行目で列見出しを定義し、以降の行を「表側の見出し＋各列の値」として並べる。
  // 入力形式：
  //   - 列：初期費用｜運用コスト｜拡張性
  //   - A社：30万円｜月1万円｜限定的
  //   - B社：★推奨 50万円｜月3万円｜拡張可能
  function parseCrossTable(bullets) {
    const items = A.extractItems(bullets);
    let cols = [];
    const rows = [];
    items.forEach((it) => {
      if (!cols.length && /^(列|項目|軸|観点|評価軸|比較軸)$/.test(it.key)) {
        cols = splitFields(it.value);
        return;
      }
      rows.push({ label: it.key, cells: splitFields(it.value), highlight: it.highlight });
    });
    const width = rows.reduce((m, r) => Math.max(m, r.cells.length), 0);
    if (!cols.length) cols = Array.from({ length: width }, (_, i) => `項目${i + 1}`);
    while (cols.length < width) cols.push('');
    return { cols: cols.slice(0, 5), rows: rows.slice(0, 7) };
  }

  // 行を「ラベル＋値セル」から実際の列数（見出し行の列数）に合わせて過不足なく揃える。
  // 「列：」宣言の項目数とデータ行のセル数が食い違っていても、表がずれずに描画できる。
  function rowCells(row, colCount) {
    const cells = [row.label].concat(row.cells);
    while (cells.length < colCount) cells.push('');
    return cells.slice(0, colCount);
  }

  const crossTable = {
    id: 'cross-table',
    name: 'クロス表（表頭×表側）',
    category: '比較',
    description: '複数の対象を複数の評価軸で交差させて比較する表。1行目に「列：軸1｜軸2｜軸3」を書き、以降を「対象：値｜値｜値」で書く。',
    score(section) {
      const t = parseCrossTable(section.bullets || []);
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['比較', '評価', '観点', '軸']);
      const multi = t.rows.filter((r) => r.cells.length >= 2).length;
      if (t.cols.length >= 2 && multi >= 2 && kw) return 10;
      if (t.cols.length >= 2 && multi >= 2) return 9;
      return 0;
    },
    renderBody(bullets) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return emptyBody();
      return `<table class="pv-table pv-xtable"><thead><tr><th></th>${t.cols
        .map((c) => `<th>${esc(c)}</th>`)
        .join('')}</tr></thead><tbody>
        ${t.rows
          .map(
            (r) => `<tr${r.highlight ? ' class="pv-table-highlight"' : ''}>
          <th class="pv-xtable-side">${r.highlight ? '★ ' : ''}${esc(r.label)}</th>
          ${t.cols.map((_, ci) => `<td>${esc(r.cells[ci] || '')}</td>`).join('')}
        </tr>`
          )
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return;
      const sideW = box.w * 0.2;
      const colW = (box.w - sideW) / t.cols.length;
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [sideW].concat(t.cols.map(() => colW)),
        fontSize: 11,
        header: [''].concat(t.cols),
        firstColAccent: true,
        rows: t.rows.map((r) =>
          [{ text: `${r.highlight ? '★ ' : ''}${r.label}`, options: r.highlight ? { color: theme.highlight } : {} }].concat(
            t.cols.map((_, ci) => ({
              text: r.cells[ci] || '',
              options: Object.assign({ align: 'center' }, r.highlight ? { bold: true, fill: { color: theme.lighter } } : {}),
            }))
          )
        ),
      });
    },
  };

  // ---------- 汎用表（自由列） ----------
  // クロス表と同じ「1行目で列を定義」形式だが、1列目を特別扱いしない素の表。
  // 生産の現場では「対象×評価軸」より単純な「項目を並べた一覧表」の方が頻度が高いため、
  // 表側の色分けを前提にしないニュートラルな表として別パターンにしている。
  const genericTable = {
    id: 'generic-table',
    name: '汎用表（自由列）',
    category: '汎用',
    description: '列を自由に定義できる、ニュートラルな一覧表。1行目に「列：1列目の見出し｜2列目の見出し｜…」（1列目の見出しも含める）、以降を「行の1列目：2列目以降の値｜値｜…」で書く。',
    score(section) {
      const t = parseCrossTable(section.bullets || []);
      const multi = t.rows.filter((r) => r.cells.length >= 1).length;
      // クロス表（比較・評価の文脈）と競合しないよう、比較色の強いキーワードが
      // 無いときのニュートラルな表としてこちらを優先させる。
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['比較', '評価', '観点', '軸']);
      if (t.cols.length >= 2 && multi >= 3 && !kw) return 9;
      if (t.cols.length >= 2 && multi >= 3) return 6;
      return 0;
    },
    renderBody(bullets) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return emptyBody();
      const headers = t.cols.length ? t.cols : ['項目'];
      return `<table class="pv-table"><thead><tr>${headers.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>
        ${t.rows
          .map((r) => {
            const cells = rowCells(r, headers.length);
            return `<tr${r.highlight ? ' class="pv-table-highlight"' : ''}>${cells
              .map((v, ci) => `<td>${ci === 0 && r.highlight ? '★ ' : ''}${esc(v || '')}</td>`)
              .join('')}</tr>`;
          })
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return;
      const headers = t.cols.length ? t.cols : ['項目'];
      const colW = box.w / headers.length;
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: headers.map(() => colW),
        fontSize: 11,
        header: headers,
        zebra: true,
        rows: t.rows.map((r) =>
          rowCells(r, headers.length).map((v, ci) => ({
            text: `${ci === 0 && r.highlight ? '★ ' : ''}${v || ''}`,
            options: r.highlight ? { bold: true, color: ci === 0 ? theme.highlight : theme.text } : {},
          }))
        ),
      });
    },
  };

  // ---------- 中扉（セクション区切り） ----------
  // 章の切り替わりを示すスライド。全セクションを並べ、★を付けた1つを現在地として強調する。
  // 入力形式： - 01：現状の課題 ／ - 02：★ ご提案 …
  const sectionDivider = {
    id: 'section-divider',
    name: '中扉（セクション区切り）',
    category: '汎用',
    description: '章の切り替わりを示す扉ページ。全セクションを並べ、「★」を付けた1つを現在地として大きく示す。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['中扉', 'セクション', '章', '本章', 'ここから']);
      const items = A.extractItems(section.bullets || []);
      const marked = items.filter((it) => it.highlight).length;
      if (kw && marked === 1) return 10;
      if (marked === 1 && items.length >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return emptyBody();
      const cur = items.find((it) => it.highlight) || items[0];
      return `<div class="pv-sd">
        <div class="pv-sd-main">
          <div class="pv-sd-num">${esc(cur.key)}</div>
          <div class="pv-sd-rule"></div>
          <div class="pv-sd-title">${esc(cur.value)}</div>
        </div>
        <div class="pv-sd-list">${items
          .map(
            (it) => `<div class="pv-sd-item${it === cur ? ' is-current' : ''}"><span>${esc(it.key)}</span>${esc(it.value)}</div>`
          )
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return;
      const cur = items.find((it) => it.highlight) || items[0];
      const leftW = box.w * 0.56;
      slide.addText(cur.key, {
        x: box.x, y: box.y + box.h * 0.16, w: leftW, h: 0.9,
        fontSize: 40, bold: true, color: theme.light, valign: 'middle',
      });
      slide.addShape('rect', {
        x: box.x, y: box.y + box.h * 0.16 + 0.98, w: Math.min(leftW, 3.2), h: 0.045,
        fill: { color: theme.primary }, line: { type: 'none' },
      });
      slide.addText(cur.value, {
        x: box.x, y: box.y + box.h * 0.16 + 1.12, w: leftW, h: 0.7,
        fontSize: 22, bold: true, color: theme.primary, valign: 'top',
      });
      // 右側に全セクションを並べ、現在地以外は淡く表示する
      const listX = box.x + leftW + 0.3;
      const listW = box.x + box.w - listX;
      const n = items.length;
      const rowH = Math.min(0.48, box.h / n);
      const startY = box.y + (box.h - rowH * n) / 2;
      items.forEach((it, i) => {
        const y = startY + i * rowH;
        const isCur = it === cur;
        if (isCur) {
          slide.addShape('rect', { x: listX - 0.12, y, w: 0.045, h: rowH, fill: { color: theme.primary }, line: { type: 'none' } });
        }
        slide.addText(
          [
            { text: it.key + '   ', options: { bold: true, color: isCur ? theme.primary : theme.accent } },
            { text: it.value, options: { color: isCur ? theme.text : theme.subtext } },
          ],
          { x: listX, y, w: listW, h: rowH, fontSize: isCur ? 12 : 11, bold: isCur, valign: 'middle' }
        );
      });
    },
  };

  // ---------- アイコン付き見出しボックス ----------
  // 観点ごとに「アイコン＋淡い見出しバー＋枠線の本文ボックス」を並べる。
  // 入力形式： - 経営：意思決定の迅速化｜投資対効果の可視化
  const iconHeaderBox = {
    id: 'icon-header-box',
    name: 'アイコン付き見出しボックス',
    category: '比較',
    description: '観点ごとにアイコン付きの見出しバーと枠線ボックスを並べる。立場・部門別の論点整理に使う。',
    score(section) {
      const items = A.extractItems(section.bullets || []);
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['観点', '立場', '部門', '目線', 'ごと']);
      const n = items.length;
      if (kw && n >= 2 && n <= 3) return 9;
      if (n >= 2 && n <= 3 && items.filter((it) => splitFields(it.value).length >= 2).length >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 3);
      if (!items.length) return emptyBody();
      return `<div class="pv-ihb">${items
        .map(
          (it) => `<div class="pv-ihb-col">
          <div class="pv-ihb-head">
            <span class="pv-ihb-icon">${iconSvg(guessIcon(it.key))}</span>
            <span class="pv-ihb-bar">${esc(it.key)}</span>
          </div>
          <div class="pv-ihb-body">${splitFields(it.value)
            .map((l) => `<div class="pv-ihb-line">${esc(l)}</div>`)
            .join('')}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 3);
      if (!items.length) return;
      const n = items.length;
      const gap = 0.24;
      const colW = (box.w - gap * (n - 1)) / n;
      const icon = 0.52;
      const headH = 0.4;
      const headY = box.y + icon / 2 - headH / 2;
      items.forEach((it, i) => {
        const x = box.x + i * (colW + gap);
        // 見出しバー（アイコンの分だけ右にずらす）
        slide.addShape('rect', {
          x: x + icon * 0.6, y: headY, w: colW - icon * 0.6, h: headH,
          fill: { color: theme.light }, line: { type: 'none' },
        });
        slide.addText(it.key, {
          x: x + icon * 0.6 + 0.34, y: headY, w: colW - icon * 0.6 - 0.42, h: headH,
          fontSize: 11.5, bold: true, color: theme.primary, valign: 'middle',
        });
        // 本文の枠線ボックス
        const bodyY = box.y + icon;
        const bodyH = box.h - icon;
        slide.addShape('rect', {
          x: x + icon * 0.3, y: bodyY, w: colW - icon * 0.3, h: bodyH,
          fill: { color: theme.white }, line: { color: theme.primaryLight, width: 1 },
        });
        const lines = splitFields(it.value);
        if (lines.length) {
          slide.addText(bulletTextRuns(lines, theme, { fontSize: 10, spaceAfter: 6 }), {
            x: x + icon * 0.3 + 0.16, y: bodyY + 0.14, w: colW - icon * 0.3 - 0.32, h: bodyH - 0.28,
            valign: 'top',
          });
        }
        // アイコンの丸（左上に重ねる）
        slide.addShape('oval', {
          x, y: box.y, w: icon, h: icon,
          fill: { color: theme.white }, line: { color: theme.primary, width: 1.25 },
        });
        addIcon(slide, guessIcon(it.key), { x: x + icon * 0.24, y: box.y + icon * 0.24, w: icon * 0.52, h: icon * 0.52 }, theme);
      });
    },
  };

  // ---------- Point強調 ----------
  // 「Point 1」のタグを付けた強調ボックスを積む。押さえてほしい要点を数点だけ示す型。
  const pointCallout = {
    id: 'point-callout',
    name: 'Point強調',
    category: '汎用',
    description: '「Point」タグ付きの強調ボックスを積んで、押さえてほしい要点を2〜4点示す。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ポイント', 'Point', '要点', '押さえ']);
      const n = (section.bullets || []).length;
      if (kw && n >= 2 && n <= 4) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return emptyBody();
      return `<div class="pv-pc">${items
        .map(
          (it, i) => `<div class="pv-pc-row${it.highlight ? ' is-highlight' : ''}">
          <span class="pv-pc-tag">Point ${i + 1}</span>
          <div class="pv-pc-box"><b>${esc(it.key)}</b>${it.value !== it.key ? `<span>${esc(it.value)}</span>` : ''}</div>
        </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (!items.length) return;
      const n = items.length;
      const gap = 0.18;
      const rowH = (box.h - gap * (n - 1)) / n;
      const tagW = 0.92;
      const tagH = 0.28;
      items.forEach((it, i) => {
        const y = box.y + i * (rowH + gap);
        const tone = it.highlight ? theme.highlight : theme.primary;
        // 本文ボックス（タグの分だけ下げて、タグが上辺に載るようにする）
        slide.addShape('rect', {
          x: box.x, y: y + tagH * 0.5, w: box.w, h: rowH - tagH * 0.5,
          fill: { color: theme.lighter }, line: { type: 'none' },
        });
        slide.addShape('rect', {
          x: box.x, y: y + tagH * 0.5, w: 0.05, h: rowH - tagH * 0.5,
          fill: { color: tone }, line: { type: 'none' },
        });
        slide.addShape('rect', { x: box.x + 0.22, y, w: tagW, h: tagH, fill: { color: tone }, line: { type: 'none' } });
        slide.addText(`Point ${i + 1}`, {
          x: box.x + 0.22, y, w: tagW, h: tagH,
          fontSize: 9.5, bold: true, color: theme.white, align: 'center', valign: 'middle',
        });
        const runs = [{ text: it.key, options: { bold: true, fontSize: 12, color: theme.primary } }];
        if (it.value !== it.key) runs.push({ text: '　' + it.value, options: { fontSize: 11, color: theme.text } });
        slide.addText(runs, {
          x: box.x + 0.26, y: y + tagH + 0.04, w: box.w - 0.52, h: rowH - tagH - 0.1,
          valign: 'middle',
        });
      });
    },
  };

  // ---------- 前提→検討→結論 ----------
  // 左に前提、中央に検討した観点、右に結論を置き、矢印で結論へ導く。
  // 入力形式： - 前提：… / - 検討：… （複数可） / - 結論：…
  function parseConclusionFlow(bullets) {
    const items = A.extractItems(bullets);
    const premise = items.find((it) => /前提|背景|現状/.test(it.key));
    const conclusion = items.find((it) => /結論|示唆|so ?what|提案/i.test(it.key));
    let middle = items.filter((it) => it !== premise && it !== conclusion);
    if (!middle.length) middle = items.slice(1, -1);
    // 「検討：コスト｜拡張性｜移行負荷」のように1件へまとめて書かれた場合は、
    // 観点ごとに別々のボックスへ展開する（横並びで見比べられるようにするため）
    if (middle.length === 1) {
      const fields = splitFields(middle[0].value);
      if (fields.length >= 2) middle = fields.map((f) => ({ key: f, value: '' }));
    }
    return {
      premise: premise || (items.length >= 3 ? items[0] : null),
      middle: middle.slice(0, 4),
      conclusion: conclusion || (items.length >= 2 ? items[items.length - 1] : null),
    };
  }

  const conclusionFlow = {
    id: 'conclusion-flow',
    name: '前提→検討→結論',
    category: '構造・ロジック',
    description: '左に前提、中央に検討した観点、右に結論を置き、矢印で結論へ導く。「前提：」「検討：」「結論：」の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.countAny(text, ['前提', '結論', '示唆', 'したがって', 'So What']);
      const d = parseConclusionFlow(section.bullets || []);
      if (kw >= 2 && d.premise && d.conclusion && d.middle.length) return 10;
      if (kw >= 1 && d.premise && d.conclusion) return 7;
      return 0;
    },
    renderBody(bullets) {
      const d = parseConclusionFlow(bullets);
      if (!d.premise || !d.conclusion) return emptyBody();
      return `<div class="pv-cfl">
        <div class="pv-cfl-premise"><b>${esc(d.premise.key)}</b><span>${esc(d.premise.value)}</span></div>
        <div class="pv-cfl-middle">${d.middle
          .map((m) => `<div class="pv-cfl-item"><b>${esc(m.key)}</b>${m.value ? `<span>${esc(m.value)}</span>` : ''}</div>`)
          .join('')}</div>
        <div class="pv-cfl-arrow"></div>
        <div class="pv-cfl-conclusion"><b>${esc(d.conclusion.key)}</b><span>${esc(d.conclusion.value)}</span></div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const d = parseConclusionFlow(bullets);
      if (!d.premise || !d.conclusion) return;
      const arrowW = 0.42;
      const premiseW = box.w * 0.19;
      const conclusionW = box.w * 0.24;
      const midW = box.w - premiseW - conclusionW - arrowW - 0.32;
      // 前提（ネイビーのベタ塗り）
      slide.addShape('rect', { x: box.x, y: box.y, w: premiseW, h: box.h, fill: { color: theme.primary }, line: { type: 'none' } });
      slide.addText(
        [
          { text: d.premise.key + '\n', options: { bold: true, fontSize: 11.5 } },
          { text: d.premise.value, options: { fontSize: 10 } },
        ],
        { x: box.x + 0.14, y: box.y, w: premiseW - 0.28, h: box.h, color: theme.white, align: 'center', valign: 'middle' }
      );
      // 検討（白地＋枠線）
      const midX = box.x + premiseW + 0.16;
      const mn = d.middle.length || 1;
      const mGap = 0.12;
      const mW = (midW - mGap * (mn - 1)) / mn;
      d.middle.forEach((m, i) => {
        const x = midX + i * (mW + mGap);
        slide.addShape('rect', { x, y: box.y, w: mW, h: box.h, fill: { color: theme.white }, line: { color: theme.border, width: HAIRLINE } });
        const runs = [{ text: m.key, options: { bold: true, fontSize: 10.5, color: theme.primary, breakLine: !!m.value } }];
        if (m.value) runs.push({ text: m.value, options: { fontSize: 9.5, color: theme.text } });
        slide.addText(runs, {
          x: x + 0.12, y: box.y + 0.12, w: mW - 0.24, h: box.h - 0.24,
          align: m.value ? 'left' : 'center', valign: m.value ? 'top' : 'middle',
        });
      });
      // 矢印（グレーの太い矢羽根）
      slide.addShape('rightArrow', {
        x: midX + midW + 0.06, y: box.y + box.h / 2 - 0.22, w: arrowW, h: 0.44,
        fill: { color: theme.gray }, line: { type: 'none' },
      });
      // 結論（淡いブルー）
      const cx = midX + midW + arrowW + 0.16;
      slide.addShape('rect', { x: cx, y: box.y, w: conclusionW, h: box.h, fill: { color: theme.light }, line: { type: 'none' } });
      slide.addText(
        [
          { text: d.conclusion.key + '\n', options: { bold: true, fontSize: 11.5, color: theme.primary } },
          { text: d.conclusion.value, options: { fontSize: 10.5, color: theme.text } },
        ],
        { x: cx + 0.14, y: box.y, w: conclusionW - 0.28, h: box.h, align: 'center', valign: 'middle' }
      );
    },
  };

  // ---------- レーダーチャート ----------
  // 複数の評価軸でのバランスを多角形で示す。1系列でも、比較対象を並べた2系列でも描ける。
  // 入力形式：
  //   - 系列：自社｜競合A          ← 任意。系列名を付けたいときだけ書く
  //   - 提案力：4｜3
  //   - 技術力：5｜4
  function parseRadar(bullets) {
    const items = A.extractItems(bullets);
    let names = [];
    const axes = [];
    items.forEach((it) => {
      if (!axes.length && /^(系列|凡例|対象|比較)$/.test(it.key)) {
        names = splitFields(it.value);
        return;
      }
      const nums = splitFields(it.value)
        .map((v) => firstNumber(v))
        .filter((v) => v !== null);
      if (nums.length) axes.push({ label: it.key, values: nums });
    });
    const seriesCount = axes.reduce((m, a) => Math.max(m, a.values.length), 0);
    return { names, axes: axes.slice(0, 8), seriesCount: Math.min(seriesCount, 3) };
  }

  const radarChart = {
    id: 'radar-chart',
    name: 'レーダーチャート',
    category: '数値・グラフ',
    description: '複数の評価軸でのバランスを多角形で示す（PowerPointのネイティブグラフ）。「軸：値」または「軸：自社の値｜競合の値」の形式。',
    score(section) {
      const r = parseRadar(section.bullets || []);
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['レーダー', 'バランス', '評価', 'スコア', '強み']);
      if (r.axes.length >= 4 && kw) return 10;
      if (r.axes.length >= 5) return 8;
      return 0;
    },
    renderBody(bullets) {
      const r = parseRadar(bullets);
      if (r.axes.length < 3) return emptyBody();
      const n = r.axes.length;
      const maxV = Math.max(...r.axes.map((a) => Math.max(...a.values)), 1);
      const R = 42;
      const pt = (i, v) => {
        const ang = ((-90 + (360 / n) * i) * Math.PI) / 180;
        const rr = (v / maxV) * R;
        return `${(50 + rr * Math.cos(ang)).toFixed(1)},${(50 + rr * Math.sin(ang)).toFixed(1)}`;
      };
      const grid = [0.25, 0.5, 0.75, 1]
        .map((f) => `<polygon points="${r.axes.map((_, i) => pt(i, maxV * f)).join(' ')}" fill="none" stroke="#dfe3ea" stroke-width="0.4"/>`)
        .join('');
      const spokes = r.axes
        .map((_, i) => `<line x1="50" y1="50" x2="${pt(i, maxV).split(',')[0]}" y2="${pt(i, maxV).split(',')[1]}" stroke="#dfe3ea" stroke-width="0.4"/>`)
        .join('');
      const seriesColors = ['var(--primary)', 'var(--accent)', 'var(--gray)'];
      const polys = Array.from({ length: r.seriesCount }, (_, si) => {
        const pts = r.axes.map((a, i) => pt(i, a.values[si] != null ? a.values[si] : 0)).join(' ');
        return `<polygon points="${pts}" fill="${seriesColors[si]}" fill-opacity="0.18" stroke="${seriesColors[si]}" stroke-width="1.2"/>`;
      }).join('');
      const labels = r.axes
        .map((a, i) => {
          const [x, y] = pt(i, maxV * 1.16).split(',');
          return `<text x="${x}" y="${y}" font-size="3.4" fill="#595959" text-anchor="middle" dominant-baseline="middle">${esc(a.label)}</text>`;
        })
        .join('');
      return `<div class="pv-radar">
        <svg viewBox="0 0 100 100">${grid}${spokes}${polys}${labels}</svg>
        ${
          r.names.length
            ? `<div class="pv-radar-legend">${r.names
                .slice(0, r.seriesCount)
                .map((nm, i) => `<span><i style="background:${seriesColors[i]}"></i>${esc(nm)}</span>`)
                .join('')}</div>`
            : ''
        }
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const r = parseRadar(bullets);
      if (r.axes.length < 3) return;
      const labels = r.axes.map((a) => a.label);
      const colors = [theme.primary, theme.accent, theme.gray];
      const data = Array.from({ length: r.seriesCount }, (_, si) => ({
        name: r.names[si] || `系列${si + 1}`,
        labels,
        values: r.axes.map((a) => (a.values[si] != null ? a.values[si] : 0)),
      }));
      slide.addChart('radar', data, {
        x: box.x, y: box.y, w: box.w, h: box.h,
        radarStyle: 'marker',
        chartColors: colors.slice(0, r.seriesCount),
        showLegend: r.names.length > 0,
        legendPos: 'b',
        legendColor: theme.text,
        legendFontSize: 10,
        catAxisLabelColor: theme.text,
        catAxisLabelFontSize: 9,
        valAxisLabelColor: theme.subtext,
        valAxisLabelFontSize: 8,
        lineSize: 2,
      });
    },
  };

  // ---------- ベン図 ----------
  // 2〜3つの領域の重なりで、共通部分や掛け合わせの価値を示す。
  // 入力形式： - 営業の知見：… / - 技術の知見：… / - 重なり：両者を掛け合わせた提案力
  function parseVenn(bullets) {
    const items = A.extractItems(bullets);
    const overlap = items.find((it) => /重な|共通|交差|掛け合わ|両者|中心/.test(it.key));
    const circles = items.filter((it) => it !== overlap).slice(0, 3);
    return { circles, overlap };
  }

  const vennDiagram = {
    id: 'venn-diagram',
    name: 'ベン図',
    category: '構造・ロジック',
    description: '2〜3つの領域の重なりで共通部分や掛け合わせの価値を示す。「重なり：〜」を書くと交差部分の説明になる。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ベン図', '重なり', '共通', '掛け合わせ', '交差']);
      const v = parseVenn(section.bullets || []);
      if (kw && v.circles.length >= 2 && v.overlap) return 10;
      if (kw && v.circles.length >= 2) return 8;
      return 0;
    },
    renderBody(bullets) {
      const v = parseVenn(bullets);
      if (v.circles.length < 2) return emptyBody();
      const three = v.circles.length >= 3;
      const pos = three
        ? [{ cx: 38, cy: 34 }, { cx: 62, cy: 34 }, { cx: 50, cy: 58 }]
        : [{ cx: 37, cy: 46 }, { cx: 63, cy: 46 }];
      return `<div class="pv-venn">
        <svg viewBox="0 0 100 92">
          ${pos
            .map((p) => `<circle cx="${p.cx}" cy="${p.cy}" r="24" fill="var(--primary-light)" fill-opacity="0.18" stroke="var(--primary-light)" stroke-width="0.7"/>`)
            .join('')}
          ${
            v.overlap
              ? `<text x="50" y="${three ? 44 : 46}" font-size="3.6" font-weight="bold" fill="var(--primary)" text-anchor="middle" dominant-baseline="middle">${esc(
                  v.overlap.key
                )}</text>`
              : ''
          }
          ${pos
            .map(
              (p, i) =>
                `<text x="${p.cx + (three && i === 2 ? 0 : i === 0 ? -9 : 9)}" y="${
                  three && i === 2 ? p.cy + 11 : p.cy - 10
                }" font-size="3.6" font-weight="bold" fill="var(--primary)" text-anchor="middle">${esc(v.circles[i].key)}</text>`
            )
            .join('')}
        </svg>
        <div class="pv-venn-notes">${v.circles
          .map((c) => `<div><b>${esc(c.key)}</b>${esc(c.value !== c.key ? c.value : '')}</div>`)
          .join('')}${v.overlap ? `<div class="is-overlap"><b>${esc(v.overlap.key)}</b>${esc(v.overlap.value)}</div>` : ''}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const v = parseVenn(bullets);
      if (v.circles.length < 2) return;
      const three = v.circles.length >= 3;
      const diagW = box.w * 0.46;
      const size = Math.min(diagW, box.h) * (three ? 0.62 : 0.7);
      const cx = box.x + diagW / 2;
      const cy = box.y + box.h / 2;
      const off = size * 0.34;
      const centers = three
        ? [{ x: cx - off, y: cy - off * 0.6 }, { x: cx + off, y: cy - off * 0.6 }, { x: cx, y: cy + off * 0.72 }]
        : [{ x: cx - off, y: cy }, { x: cx + off, y: cy }];
      centers.forEach((c) => {
        slide.addShape('ellipse', {
          x: c.x - size / 2, y: c.y - size / 2, w: size, h: size,
          fill: { color: theme.primaryLight, transparency: 82 },
          line: { color: theme.primaryLight, width: 1 },
        });
      });
      centers.forEach((c, i) => {
        const dx = three && i === 2 ? 0 : i === 0 ? -size * 0.3 : size * 0.3;
        const dy = three && i === 2 ? size * 0.34 : -size * 0.36;
        slide.addText(v.circles[i].key, {
          x: c.x + dx - 0.9, y: c.y + dy - 0.16, w: 1.8, h: 0.32,
          fontSize: 10.5, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
      });
      if (v.overlap) {
        slide.addText(v.overlap.key, {
          x: cx - 0.85, y: cy - (three ? 0.02 : 0.16), w: 1.7, h: 0.32,
          fontSize: 10, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
      }
      // 右側に各領域の説明を並べる
      const notes = v.circles.concat(v.overlap ? [v.overlap] : []);
      const nx = box.x + diagW + 0.3;
      const nw = box.x + box.w - nx;
      const nh = box.h / notes.length;
      notes.forEach((it, i) => {
        const y = box.y + i * nh;
        const isOverlap = it === v.overlap;
        slide.addShape('rect', {
          x: nx, y: y + nh * 0.12, w: 0.045, h: nh * 0.76,
          fill: { color: isOverlap ? theme.primary : theme.accent }, line: { type: 'none' },
        });
        slide.addText(
          [
            { text: it.key + '\n', options: { bold: true, fontSize: 10.5, color: theme.primary } },
            { text: it.value !== it.key ? it.value : '', options: { fontSize: 9.5, color: theme.text } },
          ],
          { x: nx + 0.16, y, w: nw - 0.16, h: nh, valign: 'middle' }
        );
      });
    },
  };

  // ---------- 組織図 ----------
  // 1行目を頂点、以降を配下の部門として並べ、各部門の下にさらに配下を積む。
  // 入力形式：
  //   - 代表取締役社長：山田太郎
  //   - 営業本部：第一営業部｜第二営業部
  //   - 技術本部：開発部｜品質保証部
  function parseOrgChart(bullets) {
    const items = A.extractItems(bullets);
    if (!items.length) return { root: null, branches: [] };
    return {
      root: items[0],
      branches: items.slice(1, 6).map((it) => ({ label: it.key, children: splitFields(it.value).slice(0, 4) })),
    };
  }

  const orgChart = {
    id: 'org-chart',
    name: '組織図',
    category: '構造・ロジック',
    description: '頂点と配下の部門を階層で示す。1行目が頂点、以降が「部門：配下｜配下」の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['組織', '体制', '本部', '部門', '推進体制', 'チーム編成']);
      const o = parseOrgChart(section.bullets || []);
      if (kw && o.branches.length >= 2) return 10;
      if (o.branches.length >= 3 && o.branches.filter((b) => b.children.length).length >= 2) return 7;
      return 0;
    },
    renderBody(bullets) {
      const o = parseOrgChart(bullets);
      if (!o.root || !o.branches.length) return emptyBody();
      return `<div class="pv-org">
        <div class="pv-org-root">${esc(o.root.key)}${o.root.value !== o.root.key ? `<span>${esc(o.root.value)}</span>` : ''}</div>
        <div class="pv-org-stem"></div>
        <div class="pv-org-branches">${o.branches
          .map(
            (b) => `<div class="pv-org-branch">
          <div class="pv-org-node">${esc(b.label)}</div>
          ${b.children.map((c) => `<div class="pv-org-child">${esc(c)}</div>`).join('')}
        </div>`
          )
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const o = parseOrgChart(bullets);
      if (!o.root || !o.branches.length) return;
      const n = o.branches.length;
      const rootW = Math.min(2.6, box.w * 0.3);
      const rootH = 0.44;
      const cx = box.x + box.w / 2;
      slide.addShape('rect', { x: cx - rootW / 2, y: box.y, w: rootW, h: rootH, fill: { color: theme.primary }, line: { type: 'none' } });
      slide.addText(o.root.value !== o.root.key ? `${o.root.key}　${o.root.value}` : o.root.key, {
        x: cx - rootW / 2, y: box.y, w: rootW, h: rootH,
        fontSize: 11, bold: true, color: theme.white, align: 'center', valign: 'middle',
      });
      // 縦のつなぎ線と横のバス線
      const stemY = box.y + rootH;
      const busY = stemY + 0.24;
      slide.addShape('line', { x: cx, y: stemY, w: 0, h: 0.24, line: { color: theme.accent, width: 1 } });
      const gap = 0.14;
      const colW = (box.w - gap * (n - 1)) / n;
      const firstCx = box.x + colW / 2;
      const lastCx = box.x + (n - 1) * (colW + gap) + colW / 2;
      slide.addShape('line', { x: firstCx, y: busY, w: lastCx - firstCx, h: 0, line: { color: theme.accent, width: 1 } });
      const nodeY = busY + 0.2;
      const nodeH = 0.4;
      o.branches.forEach((b, i) => {
        const x = box.x + i * (colW + gap);
        const bcx = x + colW / 2;
        slide.addShape('line', { x: bcx, y: busY, w: 0, h: 0.2, line: { color: theme.accent, width: 1 } });
        slide.addShape('rect', { x, y: nodeY, w: colW, h: nodeH, fill: { color: theme.light }, line: { type: 'none' } });
        slide.addText(b.label, {
          x: x + 0.06, y: nodeY, w: colW - 0.12, h: nodeH,
          fontSize: 10.5, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
        const childH = 0.34;
        b.children.forEach((c, ci) => {
          const y = nodeY + nodeH + 0.1 + ci * (childH + 0.07);
          if (y + childH > box.y + box.h) return;
          slide.addShape('rect', {
            x: x + 0.1, y, w: colW - 0.2, h: childH,
            fill: { color: theme.white }, line: { color: theme.border, width: HAIRLINE },
          });
          slide.addText(c, {
            x: x + 0.14, y, w: colW - 0.28, h: childH,
            fontSize: 9.5, color: theme.text, align: 'center', valign: 'middle',
          });
        });
      });
    },
  };

  // ---------- 同心円（中核要素） ----------
  // 中核から外側へ広がる層を同心円で示す。1行目が最も内側。
  const concentricCircles = {
    id: 'concentric-circles',
    name: '同心円（中核要素）',
    category: '構造・ロジック',
    description: '中核から外側へ広がる層を同心円で示す。1行目が最も内側の中核。3〜4層向け。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['中核', '核', '同心円', '層', '広がり', 'コア']);
      const n = (section.bullets || []).length;
      if (kw && n >= 3 && n <= 4) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return emptyBody();
      const n = items.length;
      return `<div class="pv-cc">
        <div class="pv-cc-circles">${items
          .map((it, i) => {
            // 内側ほど小さく、かつ内側ほど手前（z-index を大きく）にする
            const size = 100 - (n - 1 - i) * (60 / n);
            const onDark = i <= n - 2;
            return `<div class="pv-cc-ring" style="width:${size}%;height:${size}%;background:${scaleCssVar(
              i,
              n,
              true
            )};z-index:${n - i};color:${onDark ? '#fff' : 'var(--primary)'};"><span>${esc(it.key)}</span></div>`;
          })
          .join('')}</div>
        <div class="pv-cc-notes">${items
          .map((it) => `<div><b>${esc(it.key)}</b>${esc(it.value !== it.key ? it.value : '')}</div>`)
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 4);
      if (items.length < 2) return;
      const n = items.length;
      const diagW = box.w * 0.44;
      const maxD = Math.min(diagW, box.h) * 0.94;
      const cx = box.x + diagW / 2;
      const cy = box.y + box.h / 2;
      // 外側（＝最後の項目）から描いて内側を上に重ねる
      for (let i = n - 1; i >= 0; i--) {
        const d = maxD * (1 - (n - 1 - i) * (0.62 / n));
        slide.addShape('ellipse', {
          x: cx - d / 2, y: cy - d / 2, w: d, h: d,
          fill: { color: scaleColor(theme, i, n, true) }, line: { color: theme.white, width: 1.5 },
        });
      }
      // ラベルは各層の上端付近に置く（内側の円に隠れないように）
      for (let i = n - 1; i >= 0; i--) {
        const d = maxD * (1 - (n - 1 - i) * (0.62 / n));
        const inner = i > 0 ? maxD * (1 - (n - i) * (0.62 / n)) : 0;
        const bandTop = cy - d / 2;
        const labelY = i === 0 ? cy - 0.16 : bandTop + (d / 2 - inner / 2) / 2 - 0.14;
        slide.addText(items[i].key, {
          x: cx - d / 2, y: labelY, w: d, h: 0.3,
          fontSize: i === 0 ? 11 : 9.5, bold: true,
          color: i <= n - 2 ? theme.white : theme.primary,
          align: 'center', valign: 'middle',
        });
      }
      const nx = box.x + diagW + 0.3;
      const nw = box.x + box.w - nx;
      const nh = box.h / n;
      items.forEach((it, i) => {
        const y = box.y + i * nh;
        slide.addShape('rect', {
          x: nx, y: y + nh * 0.14, w: 0.045, h: nh * 0.72,
          fill: { color: scaleColor(theme, i, n, true) }, line: { type: 'none' },
        });
        slide.addText(
          [
            { text: it.key + '\n', options: { bold: true, fontSize: 10.5, color: theme.primary } },
            { text: it.value !== it.key ? it.value : '', options: { fontSize: 9.5, color: theme.text } },
          ],
          { x: nx + 0.16, y, w: nw - 0.16, h: nh, valign: 'middle' }
        );
      });
    },
  };

  // ---------- 決定木 ----------
  // 1行目の問いから枝分かれして結論に至る流れを示す。
  // 入力形式： - 判断：内製で対応できるか / - はい：自社開発で進める｜3ヶ月 / - いいえ：外部委託｜6ヶ月
  function parseDecisionTree(bullets) {
    const items = A.extractItems(bullets);
    if (items.length < 2) return { root: null, branches: [] };
    return {
      root: items[0],
      branches: items.slice(1, 5).map((it) => {
        const f = splitFields(it.value);
        return { label: it.key, result: f[0] || it.value, note: f[1] || '', highlight: it.highlight };
      }),
    };
  }

  const decisionTree = {
    id: 'decision-tree',
    name: '決定木',
    category: '構造・ロジック',
    description: '問いから枝分かれして結論に至る流れを示す。1行目が問い、以降が「分岐：結論｜補足」の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['決定木', '分岐', '判断', '場合分け', 'はい', 'いいえ', 'Yes', 'No']);
      const d = parseDecisionTree(section.bullets || []);
      if (kw && d.branches.length >= 2) return 9;
      return 0;
    },
    renderBody(bullets) {
      const d = parseDecisionTree(bullets);
      if (!d.root || !d.branches.length) return emptyBody();
      return `<div class="pv-dt">
        <div class="pv-dt-root">${esc(d.root.value !== d.root.key ? d.root.value : d.root.key)}</div>
        <div class="pv-dt-branches">${d.branches
          .map(
            (b) => `<div class="pv-dt-branch">
          <span class="pv-dt-label">${esc(b.label)}</span>
          <div class="pv-dt-result${b.highlight ? ' is-highlight' : ''}"><b>${esc(b.result)}</b>${
              b.note ? `<span>${esc(b.note)}</span>` : ''
            }</div>
        </div>`
          )
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const d = parseDecisionTree(bullets);
      if (!d.root || !d.branches.length) return;
      const n = d.branches.length;
      const rootW = box.w * 0.24;
      const rootH = Math.min(0.9, box.h * 0.4);
      const rootY = box.y + box.h / 2 - rootH / 2;
      slide.addShape('rect', { x: box.x, y: rootY, w: rootW, h: rootH, fill: { color: theme.primary }, line: { type: 'none' } });
      slide.addText(d.root.value !== d.root.key ? d.root.value : d.root.key, {
        x: box.x + 0.12, y: rootY, w: rootW - 0.24, h: rootH,
        fontSize: 11, bold: true, color: theme.white, align: 'center', valign: 'middle',
      });
      const gap = 0.14;
      const bh = (box.h - gap * (n - 1)) / n;
      const bx = box.x + rootW + 0.75;
      const bw = box.x + box.w - bx;
      const midX = box.x + rootW + 0.34;
      slide.addShape('line', { x: midX, y: box.y + bh / 2, w: 0, h: box.h - bh, line: { color: theme.accent, width: 1 } });
      slide.addShape('line', { x: box.x + rootW, y: box.y + box.h / 2, w: midX - (box.x + rootW), h: 0, line: { color: theme.accent, width: 1 } });
      d.branches.forEach((b, i) => {
        const y = box.y + i * (bh + gap);
        const cy = y + bh / 2;
        slide.addShape('line', { x: midX, y: cy, w: bx - midX, h: 0, line: { color: theme.accent, width: 1 } });
        slide.addText(b.label, {
          x: midX + 0.06, y: cy - 0.28, w: bx - midX - 0.12, h: 0.24,
          fontSize: 8.5, bold: true, color: theme.subtext, align: 'center', valign: 'bottom',
        });
        slide.addShape('rect', {
          x: bx, y, w: bw, h: bh,
          fill: { color: b.highlight ? theme.light : theme.white },
          line: { color: b.highlight ? theme.highlight : theme.border, width: b.highlight ? 1.5 : HAIRLINE },
        });
        const runs = [{ text: b.result, options: { bold: true, fontSize: 10.5, color: theme.primary, breakLine: !!b.note } }];
        if (b.note) runs.push({ text: b.note, options: { fontSize: 9.5, color: theme.subtext } });
        slide.addText(runs, { x: bx + 0.14, y, w: bw - 0.28, h: bh, valign: 'middle' });
      });
    },
  };

  // ---------- タイムテーブル ----------
  // 時刻と内容を並べる時間割。研修・ワークショップ・当日進行の共有向け。
  // 入力形式： - 9:00-10:00：オープニング｜事務局
  const scheduleTable = {
    id: 'schedule-table',
    name: 'タイムテーブル',
    category: '時系列',
    description: '時刻と内容を並べる当日進行表。「9:00-10:00：内容｜担当」の形式。研修・ワークショップの進行共有に使う。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const times = (text.match(/\d{1,2}[:：]\d{2}/g) || []).length;
      const kw = A.hasAny(text, ['タイムテーブル', '進行', '当日', '時間割', 'アジェンダ']);
      const n = bullets.length;
      if (n >= 4 && times >= Math.ceil(n * 0.7)) return kw ? 10 : 9;
      if (kw && times >= 3) return 7;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 10);
      if (!items.length) return emptyBody();
      const hasOwner = items.some((it) => splitFields(it.value).length >= 2);
      return `<table class="pv-table"><thead><tr><th style="width:22%;">時刻</th><th>内容</th>${
        hasOwner ? '<th style="width:20%;">担当</th>' : ''
      }</tr></thead><tbody>
        ${items
          .map((it) => {
            const f = splitFields(it.value);
            return `<tr${it.highlight ? ' class="pv-table-highlight"' : ''}><td>${esc(it.key)}</td><td>${esc(f[0] || '')}</td>${
              hasOwner ? `<td>${esc(f[1] || '')}</td>` : ''
            }</tr>`;
          })
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 10);
      if (!items.length) return;
      const hasOwner = items.some((it) => splitFields(it.value).length >= 2);
      const colW = hasOwner ? [box.w * 0.22, box.w * 0.58, box.w * 0.2] : [box.w * 0.22, box.w * 0.78];
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW,
        fontSize: 10.5,
        header: hasOwner ? ['時刻', '内容', '担当'] : ['時刻', '内容'],
        firstColAccent: true,
        rows: items.map((it) => {
          const f = splitFields(it.value);
          const row = [
            { text: it.key, options: { align: 'center' } },
            { text: f[0] || '', options: it.highlight ? { bold: true, color: theme.primary } : {} },
          ];
          if (hasOwner) row.push({ text: f[1] || '', options: { align: 'center' } });
          return row;
        }),
      });
    },
  };

  // ---------- 横型ファネル ----------
  // 左から右へ絞り込まれる流れを台形で示す。段階ごとの数値を下に添える。
  const funnelHorizontal = {
    id: 'funnel-horizontal',
    name: '横型ファネル',
    category: '変化・対比',
    description: '左から右へ段階的に絞り込まれる流れを横向きの台形で示す。縦型より段階数が多い場合に読みやすい。',
    score(section) {
      const bullets = section.bullets || [];
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ファネル', '歩留まり', 'コンバージョン', '絞り込み']);
      const items = A.extractItems(bullets);
      const dec = bullets.length >= 4 && isDecreasingSeq(items);
      if (kw && dec && bullets.length >= 5) return 10;
      if (dec && bullets.length >= 5) return 8;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (items.length < 2) return emptyBody();
      const n = items.length;
      return `<div class="pv-fh">
        <div class="pv-fh-body">${items
          .map((it, i) => {
            const h = (100 - i * (58 / (n - 1 || 1))).toFixed(1);
            return `<div class="pv-fh-seg" style="background:${scaleCssVar(i, n)};height:${h}%;"><b>${esc(it.value)}</b></div>`;
          })
          .join('')}</div>
        <div class="pv-fh-labels">${items.map((it) => `<div>${esc(it.key)}</div>`).join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 7);
      if (items.length < 2) return;
      const n = items.length;
      const labelH = 0.36;
      const chartH = box.h - labelH - 0.1;
      const segW = box.w / n;
      items.forEach((it, i) => {
        const x = box.x + i * segW;
        // 段ごとに高さを減らして絞り込みを表す。図形の回転で台形を寝かせると
        // 外接枠ごと回ってしまい横長のコマに収まらないため、矩形の高さで表現する。
        const h = chartH * (1 - i * (0.58 / (n - 1 || 1)));
        slide.addShape('rect', {
          x, y: box.y + (chartH - h) / 2, w: segW + 0.01, h,
          fill: { color: scaleColor(theme, i, n) }, line: { color: theme.white, width: 1 },
        });
        slide.addText(it.value, {
          x, y: box.y + chartH / 2 - 0.18, w: segW, h: 0.36,
          fontSize: 10, bold: true, color: i >= n - 3 ? theme.white : theme.text,
          align: 'center', valign: 'middle',
        });
        slide.addText(it.key, {
          x, y: box.y + chartH + 0.06, w: segW, h: labelH,
          fontSize: 9, color: theme.subtext, align: 'center', valign: 'top',
        });
      });
    },
  };

  // ---------- 番号付きアジェンダ（カード型） ----------
  // 番号・見出し・補足の3段をカードにしてグリッドで並べる。項目数が多い目次向け。
  const agendaCards = {
    id: 'agenda-cards',
    name: '番号付きアジェンダ（カード型）',
    category: '汎用',
    description: '番号・見出し・補足の3段をカードにしてグリッドで並べる。項目数の多い目次やアジェンダ向け。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['アジェンダ', '目次', 'もくじ', 'Agenda', '本日']);
      const items = A.extractItems(section.bullets || []);
      const n = items.length;
      if (kw && n >= 4 && n <= 6) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return emptyBody();
      return `<div class="pv-ac">${items
        .map(
          (it, i) => `<div class="pv-ac-card${it.highlight ? ' is-highlight' : ''}">
        <div class="pv-ac-num">${String(i + 1).padStart(2, '0')}.</div>
        <div class="pv-ac-title">${esc(it.key)}</div>
        ${it.value !== it.key ? `<div class="pv-ac-desc">${esc(it.value)}</div>` : ''}
      </div>`
        )
        .join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 6);
      if (!items.length) return;
      const n = items.length;
      const cols = n <= 3 ? n : 3;
      const rowsN = Math.ceil(n / cols);
      const gapX = 0.3;
      const gapY = 0.24;
      const cw = (box.w - gapX * (cols - 1)) / cols;
      const ch = (box.h - gapY * (rowsN - 1)) / rowsN;
      items.forEach((it, i) => {
        const x = box.x + (i % cols) * (cw + gapX);
        const y = box.y + Math.floor(i / cols) * (ch + gapY);
        const tone = it.highlight ? theme.highlight : theme.primary;
        slide.addText(`${String(i + 1).padStart(2, '0')}.`, {
          x, y, w: cw, h: 0.42,
          fontSize: 19, bold: true, color: tone, valign: 'middle',
        });
        slide.addShape('line', { x, y: y + 0.46, w: cw * 0.32, h: 0, line: { color: tone, width: 1.5 } });
        slide.addText(it.key, {
          x, y: y + 0.54, w: cw, h: 0.34,
          fontSize: 12, bold: true, color: theme.text, valign: 'top',
        });
        if (it.value !== it.key) {
          slide.addText(it.value, {
            x, y: y + 0.9, w: cw, h: ch - 0.94,
            fontSize: 9.5, color: theme.subtext, valign: 'top',
          });
        }
      });
    },
  };

  // ---------- 評価マトリクス（◎○△×） ----------
  // 対象×評価軸のマスに記号を置いて優劣を一覧する、日本の提案書で定番の表。
  // 入力形式：
  //   - 列：コスト｜機能｜サポート｜拡張性
  //   - A社：◎｜○｜△｜×
  //   - B社：★推奨 ○｜◎｜◎｜○
  const MARK_RE = /^[◎○◯〇△▲✕×✖xX\-−]$/;
  function markTone(mark, theme) {
    const m = String(mark || '').trim();
    if (/^[◎]/.test(m)) return theme.primary;
    if (/^[○◯〇]/.test(m)) return theme.primaryLight;
    if (/^[△▲]/.test(m)) return theme.gray;
    if (/^[✕×✖xX]/.test(m)) return theme.critical;
    return theme.text;
  }
  function markCls(mark) {
    const m = String(mark || '').trim();
    if (/^[◎]/.test(m)) return 'is-best';
    if (/^[○◯〇]/.test(m)) return 'is-good';
    if (/^[△▲]/.test(m)) return 'is-fair';
    if (/^[✕×✖xX]/.test(m)) return 'is-poor';
    return '';
  }

  const dotMatrix = {
    id: 'dot-matrix',
    name: '評価マトリクス（◎○△×）',
    category: '比較',
    description: '対象×評価軸のマスに◎○△×を置いて優劣を一覧する。1行目に「列：軸1｜軸2」、以降を「対象：◎｜○｜△」で書く。',
    score(section) {
      const t = parseCrossTable(section.bullets || []);
      const marks = t.rows.reduce((n, r) => n + r.cells.filter((c) => MARK_RE.test(String(c).trim())).length, 0);
      const cells = t.rows.reduce((n, r) => n + r.cells.length, 0);
      if (t.cols.length >= 2 && cells && marks >= cells * 0.6) return 10;
      return 0;
    },
    renderBody(bullets) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return emptyBody();
      return `<table class="pv-table pv-xtable pv-dotm"><thead><tr><th></th>${t.cols
        .map((c) => `<th>${esc(c)}</th>`)
        .join('')}</tr></thead><tbody>
        ${t.rows
          .map(
            (r) => `<tr${r.highlight ? ' class="pv-table-highlight"' : ''}>
          <th class="pv-xtable-side">${r.highlight ? '★ ' : ''}${esc(r.label)}</th>
          ${t.cols
            .map((_, ci) => {
              const m = r.cells[ci] || '';
              return `<td><span class="pv-mark ${markCls(m)}">${esc(m)}</span></td>`;
            })
            .join('')}
        </tr>`
          )
          .join('')}
      </tbody></table>`;
    },
    buildBody(slide, bullets, theme, box) {
      const t = parseCrossTable(bullets);
      if (!t.rows.length) return;
      const sideW = box.w * 0.24;
      const colW = (box.w - sideW) / t.cols.length;
      addStyledTable(slide, theme, {
        x: box.x, y: box.y, w: box.w,
        colW: [sideW].concat(t.cols.map(() => colW)),
        fontSize: 11,
        header: [''].concat(t.cols),
        firstColAccent: true,
        rows: t.rows.map((r) =>
          [{ text: `${r.highlight ? '★ ' : ''}${r.label}`, options: r.highlight ? { color: theme.highlight } : {} }].concat(
            t.cols.map((_, ci) => {
              const m = r.cells[ci] || '';
              return {
                text: m,
                options: Object.assign(
                  { align: 'center', bold: true, fontSize: 13, color: markTone(m, theme) },
                  r.highlight ? { fill: { color: theme.lighter } } : {}
                ),
              };
            })
          )
        ),
      });
    },
  };

  // ---------- スイムレーン ----------
  // 役割（レーン）×工程のマスに作業を置き、誰が何をいつ行うかを一枚で示す。
  // 入力形式：
  //   - 列：申請｜承認｜処理｜完了
  //   - 営業部：申請書を作成｜｜｜受領を確認
  //   - 上長：｜内容を承認｜｜
  function parseSwimLane(bullets) {
    const items = A.extractItems(bullets);
    let cols = [];
    const lanes = [];
    items.forEach((it) => {
      if (!lanes.length && /^(列|工程|フェーズ|ステップ|段階)$/.test(it.key)) {
        cols = splitFields(it.value);
        return;
      }
      // 空欄を保持したいので splitFields は使わず、そのまま分割する
      const cells = String(it.value == null ? '' : it.value).split(/[｜|]/).map((s) => s.trim());
      lanes.push({ label: it.key, cells, highlight: it.highlight });
    });
    const width = lanes.reduce((m, l) => Math.max(m, l.cells.length), 0);
    if (!cols.length) cols = Array.from({ length: width }, (_, i) => `工程${i + 1}`);
    while (cols.length < width) cols.push('');
    return { cols: cols.slice(0, 5), lanes: lanes.slice(0, 5) };
  }

  const swimLane = {
    id: 'swim-lane',
    name: 'スイムレーン（役割別プロセス）',
    category: 'フロー・プロセス',
    description: '役割（レーン）×工程のマスに作業を置き、誰が何をいつ行うかを示す。1行目に「列：工程1｜工程2」、以降を「役割：作業｜作業」で書く（空欄可）。',
    score(section) {
      const s = parseSwimLane(section.bullets || []);
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['スイムレーン', '役割', '担当', '業務フロー', '誰が']);
      const filled = s.lanes.reduce((n, l) => n + l.cells.filter(Boolean).length, 0);
      if (s.cols.length >= 2 && s.lanes.length >= 2 && kw) return 10;
      if (s.cols.length >= 3 && s.lanes.length >= 2 && filled >= 4) return 8;
      return 0;
    },
    renderBody(bullets) {
      const s = parseSwimLane(bullets);
      if (!s.lanes.length) return emptyBody();
      return `<div class="pv-sl">
        <div class="pv-sl-head"><div class="pv-sl-lane-label"></div>${s.cols
          .map((c) => `<div class="pv-sl-col">${esc(c)}</div>`)
          .join('')}</div>
        ${s.lanes
          .map(
            (l) => `<div class="pv-sl-row">
          <div class="pv-sl-lane-label${l.highlight ? ' is-highlight' : ''}">${esc(l.label)}</div>
          ${s.cols
            .map((_, ci) => {
              const v = l.cells[ci] || '';
              return `<div class="pv-sl-cell">${v ? `<span>${esc(v)}</span>` : ''}</div>`;
            })
            .join('')}
        </div>`
          )
          .join('')}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const s = parseSwimLane(bullets);
      if (!s.lanes.length) return;
      const laneW = box.w * 0.16;
      const headH = 0.32;
      const colW = (box.w - laneW) / s.cols.length;
      const rowH = (box.h - headH) / s.lanes.length;
      s.cols.forEach((c, i) => {
        const x = box.x + laneW + i * colW;
        slide.addShape('rect', { x, y: box.y, w: colW - 0.03, h: headH, fill: { color: theme.primary }, line: { type: 'none' } });
        slide.addText(c, {
          x, y: box.y, w: colW - 0.03, h: headH,
          fontSize: 9.5, bold: true, color: theme.white, align: 'center', valign: 'middle',
        });
      });
      s.lanes.forEach((l, ri) => {
        const y = box.y + headH + ri * rowH;
        slide.addShape('rect', {
          x: box.x, y: y + 0.02, w: laneW - 0.05, h: rowH - 0.04,
          fill: { color: l.highlight ? theme.light : theme.lighter }, line: { type: 'none' },
        });
        slide.addText(l.label, {
          x: box.x + 0.05, y: y + 0.02, w: laneW - 0.15, h: rowH - 0.04,
          fontSize: 9.5, bold: true, color: theme.primary, valign: 'middle',
        });
        s.cols.forEach((_, ci) => {
          const v = l.cells[ci] || '';
          const x = box.x + laneW + ci * colW;
          if (!v) return;
          slide.addShape('rect', {
            x, y: y + 0.05, w: colW - 0.03, h: rowH - 0.1,
            fill: { color: theme.white }, line: { color: theme.primaryLight, width: 1 },
          });
          slide.addText(v, {
            x: x + 0.08, y: y + 0.05, w: colW - 0.19, h: rowH - 0.1,
            fontSize: 9, color: theme.text, align: 'center', valign: 'middle',
          });
        });
        slide.addShape('line', {
          x: box.x, y: y + rowH, w: box.w, h: 0,
          line: { color: theme.border, width: HAIRLINE_THIN },
        });
      });
    },
  };

  // ---------- ビジネスモデルキャンバス ----------
  // 9つの構成要素を定位置に配置する定番フレーム。項目名から自動でマスに割り当てる。
  const BMC_SLOTS = [
    { key: 'partners', label: 'パートナー', re: /パートナー|提携|Partner/i, col: 0, row: 0, rowSpan: 2 },
    { key: 'activities', label: '主要活動', re: /主要活動|活動|Activit/i, col: 1, row: 0, rowSpan: 1 },
    { key: 'resources', label: 'リソース', re: /リソース|資源|Resource/i, col: 1, row: 1, rowSpan: 1 },
    { key: 'value', label: '価値提案', re: /価値|提供価値|Value/i, col: 2, row: 0, rowSpan: 2 },
    { key: 'relations', label: '顧客との関係', re: /関係|Relation/i, col: 3, row: 0, rowSpan: 1 },
    { key: 'channels', label: 'チャネル', re: /チャネル|販路|Channel/i, col: 3, row: 1, rowSpan: 1 },
    { key: 'segments', label: '顧客セグメント', re: /顧客|セグメント|Segment/i, col: 4, row: 0, rowSpan: 2 },
    { key: 'cost', label: 'コスト構造', re: /コスト|費用|Cost/i, col: 0, row: 2, rowSpan: 1, wide: 'left' },
    { key: 'revenue', label: '収益の流れ', re: /収益|売上|Revenue/i, col: 0, row: 2, rowSpan: 1, wide: 'right' },
  ];

  function parseBmc(bullets) {
    const items = A.extractItems(bullets);
    const used = new Set();
    const map = {};
    BMC_SLOTS.forEach((s) => {
      const hit = items.find((it) => !used.has(it) && s.re.test(it.key));
      if (hit) {
        used.add(hit);
        map[s.key] = hit;
      }
    });
    return map;
  }

  const businessModelCanvas = {
    id: 'business-model-canvas',
    name: 'ビジネスモデルキャンバス',
    category: 'マトリクス／ポジショニング',
    description: '9つの構成要素を定位置に配置する定番フレーム。「価値提案：〜」「顧客セグメント：〜」のように項目名を書くと自動でマスに入る。',
    score(section) {
      const map = parseBmc(section.bullets || []);
      const filled = Object.keys(map).length;
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ビジネスモデル', 'キャンバス', 'BMC']);
      if (filled >= 5 && kw) return 10;
      if (filled >= 6) return 8;
      return 0;
    },
    renderBody(bullets) {
      const map = parseBmc(bullets);
      if (!Object.keys(map).length) return emptyBody();
      const cell = (s) => {
        const it = map[s.key];
        return `<div class="pv-bmc-cell" style="grid-column:${s.col + 1};grid-row:${s.row + 1} / span ${s.rowSpan};">
          <div class="pv-bmc-label">${esc(s.label)}</div>
          <div class="pv-bmc-text">${it ? esc(it.value) : ''}</div>
        </div>`;
      };
      const top = BMC_SLOTS.filter((s) => s.row < 2).map(cell).join('');
      const bottom = BMC_SLOTS.filter((s) => s.row === 2)
        .map((s) => {
          const it = map[s.key];
          return `<div class="pv-bmc-cell">
            <div class="pv-bmc-label">${esc(s.label)}</div>
            <div class="pv-bmc-text">${it ? esc(it.value) : ''}</div>
          </div>`;
        })
        .join('');
      return `<div class="pv-bmc"><div class="pv-bmc-top">${top}</div><div class="pv-bmc-bottom">${bottom}</div></div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const map = parseBmc(bullets);
      if (!Object.keys(map).length) return;
      const gap = 0.04;
      const bottomH = box.h * 0.24;
      const topH = box.h - bottomH - gap;
      const colW = box.w / 5;
      const rowH = topH / 2;
      const draw = (x, y, w, h, label, item) => {
        slide.addShape('rect', { x, y, w: w - gap, h: h - gap, fill: { color: theme.white }, line: { color: theme.border, width: HAIRLINE } });
        slide.addShape('rect', { x, y, w: w - gap, h: 0.24, fill: { color: theme.light }, line: { type: 'none' } });
        slide.addText(label, {
          x: x + 0.06, y, w: w - gap - 0.12, h: 0.24,
          fontSize: 8.5, bold: true, color: theme.primary, valign: 'middle',
        });
        if (item) {
          slide.addText(item.value, {
            x: x + 0.07, y: y + 0.28, w: w - gap - 0.14, h: h - gap - 0.32,
            fontSize: 8.5, color: theme.text, valign: 'top',
          });
        }
      };
      BMC_SLOTS.filter((s) => s.row < 2).forEach((s) => {
        draw(box.x + s.col * colW, box.y + s.row * rowH, colW, rowH * s.rowSpan, s.label, map[s.key]);
      });
      const by = box.y + topH + gap;
      const halfW = box.w / 2;
      BMC_SLOTS.filter((s) => s.row === 2).forEach((s, i) => {
        draw(box.x + i * halfW, by, halfW, bottomH, s.label, map[s.key]);
      });
    },
  };

  // ---------- カスタマージャーニー ----------
  // 各段階の接点と満足度を折れ線で示し、落ち込んでいる段階を可視化する。
  // 入力形式： - 認知：Web広告で知る｜4    （段階：接点｜満足度1〜5）
  function parseJourney(bullets) {
    return A.extractItems(bullets)
      .slice(0, 6)
      .map((it) => {
        const f = splitFields(it.value);
        const score = firstNumber(f[1]) ;
        return { stage: it.key, touch: f[0] || it.value, score: score == null ? 3 : Math.max(1, Math.min(5, score)), highlight: it.highlight };
      });
  }

  const customerJourney = {
    id: 'customer-journey',
    name: 'カスタマージャーニー',
    category: '時系列',
    description: '各段階の接点と満足度を折れ線で示し、体験が落ち込む段階を可視化する。「段階：接点｜満足度（1〜5）」の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ジャーニー', '顧客体験', '接点', 'タッチポイント', 'CX']);
      const j = parseJourney(section.bullets || []);
      if (kw && j.length >= 4) return 10;
      if (kw && j.length >= 3) return 7;
      return 0;
    },
    renderBody(bullets) {
      const j = parseJourney(bullets);
      if (j.length < 2) return emptyBody();
      const n = j.length;
      const pt = (i, s) => `${((i + 0.5) / n) * 100},${(100 - ((s - 1) / 4) * 90 - 5).toFixed(1)}`;
      return `<div class="pv-cj">
        <div class="pv-cj-stages">${j
          .map((s) => `<div class="pv-cj-stage${s.highlight ? ' is-highlight' : ''}">${esc(s.stage)}</div>`)
          .join('')}</div>
        <div class="pv-cj-chart">
          <svg viewBox="0 0 100 100" preserveAspectRatio="none">
            <polyline points="${j.map((s, i) => pt(i, s.score)).join(' ')}" fill="none" stroke="var(--primary)" stroke-width="1" vector-effect="non-scaling-stroke"/>
          </svg>
          ${j
            .map(
              (s, i) =>
                `<span class="pv-cj-dot${s.score <= 2 ? ' is-low' : ''}" style="left:${((i + 0.5) / n) * 100}%;bottom:${(((s.score - 1) / 4) * 90 + 5).toFixed(1)}%;"></span>`
            )
            .join('')}
        </div>
        <div class="pv-cj-touches">${j.map((s) => `<div>${esc(s.touch)}</div>`).join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const j = parseJourney(bullets);
      if (j.length < 2) return;
      const n = j.length;
      const colW = box.w / n;
      const stageH = 0.34;
      const touchH = 0.62;
      const chartY = box.y + stageH + 0.08;
      const chartH = box.h - stageH - touchH - 0.16;
      j.forEach((s, i) => {
        const x = box.x + i * colW;
        slide.addShape('rect', {
          x: x + 0.02, y: box.y, w: colW - 0.04, h: stageH,
          fill: { color: s.highlight ? theme.highlight : theme.light }, line: { type: 'none' },
        });
        slide.addText(s.stage, {
          x: x + 0.02, y: box.y, w: colW - 0.04, h: stageH,
          fontSize: 9.5, bold: true, color: s.highlight ? theme.white : theme.primary,
          align: 'center', valign: 'middle',
        });
      });
      // 満足度の折れ線（隣り合う点を線分でつなぐ）
      const py = (s) => chartY + chartH - ((s - 1) / 4) * chartH;
      for (let i = 0; i < n - 1; i++) {
        const x1 = box.x + i * colW + colW / 2;
        const x2 = box.x + (i + 1) * colW + colW / 2;
        const y1 = py(j[i].score);
        const y2 = py(j[i + 1].score);
        slide.addShape('line', {
          x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
          line: { color: theme.primaryLight, width: 1.75 },
          flipV: y2 < y1,
        });
      }
      j.forEach((s, i) => {
        const cx = box.x + i * colW + colW / 2;
        const cy = py(s.score);
        const low = s.score <= 2;
        slide.addShape('oval', {
          x: cx - 0.09, y: cy - 0.09, w: 0.18, h: 0.18,
          fill: { color: low ? theme.critical : theme.primary }, line: { color: theme.white, width: 1 },
        });
        slide.addText(s.touch, {
          x: box.x + i * colW + 0.04, y: box.y + box.h - touchH, w: colW - 0.08, h: touchH,
          fontSize: 8.5, color: theme.text, align: 'center', valign: 'top',
        });
      });
    },
  };

  // ---------- バリューチェーン ----------
  // 支援活動を上に積み、主活動を矢羽根で並べる定番フレーム。
  // 入力形式：
  //   - 支援活動：全社インフラ｜人事管理｜技術開発
  //   - 主活動：購買物流｜製造｜出荷物流｜販売｜サービス
  function parseValueChain(bullets) {
    const items = A.extractItems(bullets);
    const support = items.find((it) => /支援|サポート|Support/i.test(it.key));
    const primary = items.find((it) => /主活動|主要活動|Primary/i.test(it.key));
    return {
      support: splitFields(support ? support.value : '').slice(0, 4),
      primary: splitFields(primary ? primary.value : (items[items.length - 1] || {}).value).slice(0, 6),
    };
  }

  const valueChain = {
    id: 'value-chain',
    name: 'バリューチェーン',
    category: '構造・ロジック',
    description: '支援活動を上に積み、主活動を矢羽根で並べる定番フレーム。「支援活動：〜｜〜」「主活動：〜｜〜」の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['バリューチェーン', '価値連鎖', '主活動', '支援活動']);
      const v = parseValueChain(section.bullets || []);
      if (kw && v.primary.length >= 3) return 10;
      return 0;
    },
    renderBody(bullets) {
      const v = parseValueChain(bullets);
      if (!v.primary.length) return emptyBody();
      return `<div class="pv-vc">
        ${
          v.support.length
            ? `<div class="pv-vc-support">${v.support
                .map((s) => `<div class="pv-vc-support-row">${esc(s)}</div>`)
                .join('')}</div>`
            : ''
        }
        <div class="pv-vc-primary">${v.primary
          .map((p, i) => `<div class="pv-vc-seg" style="background:${scaleCssVar(i, v.primary.length)};">${esc(p)}</div>`)
          .join('')}<div class="pv-vc-margin">利益</div></div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const v = parseValueChain(bullets);
      if (!v.primary.length) return;
      const gap = 0.05;
      const supportH = v.support.length ? Math.min(box.h * 0.5, v.support.length * 0.42) : 0;
      const rowH = v.support.length ? supportH / v.support.length : 0;
      v.support.forEach((s, i) => {
        const y = box.y + i * rowH;
        slide.addShape('rect', {
          x: box.x, y, w: box.w * 0.88, h: rowH - gap,
          fill: { color: theme.lighter }, line: { color: theme.border, width: HAIRLINE },
        });
        slide.addText(s, {
          x: box.x + 0.12, y, w: box.w * 0.88 - 0.24, h: rowH - gap,
          fontSize: 9.5, color: theme.text, valign: 'middle',
        });
      });
      const py = box.y + supportH + (v.support.length ? 0.1 : 0);
      const ph = box.y + box.h - py;
      const n = v.primary.length;
      const marginW = box.w * 0.12;
      const overlap = 0.1;
      const segW = (box.w - marginW + overlap * (n - 1)) / n;
      v.primary.forEach((p, i) => {
        const x = box.x + i * (segW - overlap);
        slide.addShape(i === 0 ? 'homePlate' : 'chevron', {
          x, y: py, w: segW, h: ph,
          fill: { color: scaleColor(theme, i, n) }, line: { color: theme.white, width: 1 },
        });
        slide.addText(p, {
          x: x + 0.18, y: py, w: segW - 0.36, h: ph,
          fontSize: 9.5, bold: true, color: i >= n - 3 ? theme.white : theme.text,
          align: 'center', valign: 'middle',
        });
      });
      slide.addShape('chevron', {
        x: box.x + box.w - marginW, y: py, w: marginW, h: ph,
        fill: { color: theme.primary }, line: { type: 'none' },
      });
      slide.addText('利益', {
        x: box.x + box.w - marginW, y: py, w: marginW, h: ph,
        fontSize: 9.5, bold: true, color: theme.white, align: 'center', valign: 'middle',
      });
    },
  };

  // ---------- ハニカム ----------
  // 中心テーマの周りに六角形で構成要素を配置する。
  const honeycomb = {
    id: 'honeycomb',
    name: 'ハニカム（六角形配置）',
    category: '構造・ロジック',
    description: '中心テーマの周りに六角形で構成要素を配置する。「中心：テーマ」＋3〜6項目の形式。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['ハニカム', '構成要素', '要素', '六角']);
      const items = A.extractItems(section.bullets || []);
      const center = items.find((it) => /中心|中央|テーマ/.test(it.key));
      if (kw && center && items.length >= 4) return 9;
      if (kw && items.length >= 4) return 6;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets);
      const center = items.find((it) => /中心|中央|テーマ/.test(it.key));
      const nodes = items.filter((it) => it !== center).slice(0, 6);
      if (nodes.length < 3) return emptyBody();
      const n = nodes.length;
      return `<div class="pv-hc">
        <div class="pv-hc-ring">
          ${nodes
            .map((nd, i) => {
              const ang = ((-90 + (360 / n) * i) * Math.PI) / 180;
              const x = 50 + 33 * Math.cos(ang);
              const y = 50 + 33 * Math.sin(ang);
              return `<div class="pv-hc-cell" style="left:${x}%;top:${y}%;"><span>${String(i + 1).padStart(2, '0')}</span></div>`;
            })
            .join('')}
          <div class="pv-hc-cell is-center"><span>${center ? esc(center.value) : ''}</span></div>
        </div>
        <div class="pv-hc-notes">${nodes
          .map((nd, i) => `<div><b>${String(i + 1).padStart(2, '0')}　${esc(nd.key)}</b>${esc(nd.value !== nd.key ? nd.value : '')}</div>`)
          .join('')}</div>
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets);
      const center = items.find((it) => /中心|中央|テーマ/.test(it.key));
      const nodes = items.filter((it) => it !== center).slice(0, 6);
      if (nodes.length < 3) return;
      const n = nodes.length;
      const diagW = box.w * 0.42;
      const R = Math.min(diagW, box.h) * 0.33;
      const cx = box.x + diagW / 2;
      const cy = box.y + box.h / 2;
      const hex = R * 0.92;
      slide.addShape('hexagon', {
        x: cx - hex / 2, y: cy - hex / 2, w: hex, h: hex,
        fill: { color: theme.light }, line: { color: theme.white, width: 1.5 },
      });
      if (center) {
        slide.addText(center.value, {
          x: cx - hex / 2 + 0.06, y: cy - hex / 2, w: hex - 0.12, h: hex,
          fontSize: 9, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
      }
      nodes.forEach((nd, i) => {
        const ang = ((-90 + (360 / n) * i) * Math.PI) / 180;
        const x = cx + R * Math.cos(ang) - hex / 2;
        const y = cy + R * Math.sin(ang) - hex / 2;
        slide.addShape('hexagon', {
          x, y, w: hex, h: hex,
          fill: { color: theme.lighter }, line: { color: theme.white, width: 1.5 },
        });
        slide.addText(String(i + 1).padStart(2, '0'), {
          x, y, w: hex, h: hex,
          fontSize: 11, bold: true, color: theme.primary, align: 'center', valign: 'middle',
        });
      });
      const nx = box.x + diagW + 0.26;
      const nw = box.x + box.w - nx;
      const nh = box.h / n;
      nodes.forEach((nd, i) => {
        const y = box.y + i * nh;
        slide.addText(
          [
            { text: `${String(i + 1).padStart(2, '0')}　${nd.key}`, options: { bold: true, fontSize: 10, color: theme.primary, breakLine: nd.value !== nd.key } },
            { text: nd.value !== nd.key ? nd.value : '', options: { fontSize: 9, color: theme.text } },
          ],
          { x: nx, y, w: nw, h: nh, valign: 'middle' }
        );
      });
    },
  };

  // ---------- メリット・デメリット ----------
  // 賛否を左右に分けて示す。「メリット：〜｜〜」「デメリット：〜｜〜」の形式。
  function parseProsCons(bullets) {
    const items = A.extractItems(bullets);
    const pros = items.find((it) => /メリット|利点|良い点|長所|Pros/i.test(it.key));
    const cons = items.find((it) => /デメリット|欠点|課題|短所|懸念|Cons/i.test(it.key));
    return {
      pros: splitFields(pros ? pros.value : '').slice(0, 5),
      cons: splitFields(cons ? cons.value : '').slice(0, 5),
      prosLabel: pros ? pros.key : 'メリット',
      consLabel: cons ? cons.key : 'デメリット',
    };
  }

  const prosCons = {
    id: 'pros-cons',
    name: 'メリット・デメリット',
    category: '比較',
    description: '賛否を左右に分けて示す。「メリット：〜｜〜」「デメリット：〜｜〜」の形式。',
    score(section) {
      const p = parseProsCons(section.bullets || []);
      if (p.pros.length && p.cons.length) return 10;
      return 0;
    },
    renderBody(bullets) {
      const p = parseProsCons(bullets);
      if (!p.pros.length && !p.cons.length) return emptyBody();
      const col = (label, list, cls, mark) => `
        <div class="pv-pcn-col ${cls}">
          <div class="pv-pcn-head"><span class="pv-pcn-mark">${mark}</span>${esc(label)}</div>
          <div class="pv-pcn-list">${list.map((v) => `<div>${esc(v)}</div>`).join('')}</div>
        </div>`;
      return `<div class="pv-pcn">${col(p.prosLabel, p.pros, 'is-pros', '✓')}${col(p.consLabel, p.cons, 'is-cons', '✕')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const p = parseProsCons(bullets);
      if (!p.pros.length && !p.cons.length) return;
      const gap = 0.28;
      const colW = (box.w - gap) / 2;
      const headH = 0.44;
      [
        { label: p.prosLabel, list: p.pros, tone: theme.primary, mark: '✓' },
        { label: p.consLabel, list: p.cons, tone: theme.gray, mark: '✕' },
      ].forEach((c, i) => {
        const x = box.x + i * (colW + gap);
        slide.addShape('oval', {
          x, y: box.y, w: headH, h: headH,
          fill: { color: theme.white }, line: { color: c.tone, width: 1.5 },
        });
        slide.addText(c.mark, {
          x, y: box.y, w: headH, h: headH,
          fontSize: 14, bold: true, color: c.tone, align: 'center', valign: 'middle',
        });
        slide.addText(c.label, {
          x: x + headH + 0.12, y: box.y, w: colW - headH - 0.12, h: headH,
          fontSize: 13, bold: true, color: c.tone, valign: 'middle',
        });
        slide.addShape('line', {
          x, y: box.y + headH + 0.08, w: colW, h: 0,
          line: { color: c.tone, width: 1.25 },
        });
        if (c.list.length) {
          slide.addText(bulletTextRuns(c.list, theme, { fontSize: 11, spaceAfter: 9 }), {
            x: x + 0.06, y: box.y + headH + 0.2, w: colW - 0.12, h: box.h - headH - 0.28,
            valign: 'top',
          });
        }
      });
    },
  };

  // ---------- 引用 ----------
  // 顧客の声や調査コメントを大きく引用する。1行目が引用文、2行目が出典。
  const quoteSlide = {
    id: 'quote',
    name: '引用',
    category: '汎用',
    description: '顧客の声や調査コメントを大きく引用する。1行目が引用文、2行目が発言者・出典。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['引用', '顧客の声', 'コメント', 'インタビュー', 'ヒアリング']);
      const n = (section.bullets || []).length;
      if (kw && n >= 1 && n <= 2) return 9;
      return 0;
    },
    renderBody(bullets) {
      const list = (bullets || []).map((b) => A.stripEmphasis(b)).filter(Boolean);
      if (!list.length) return emptyBody();
      return `<div class="pv-quote">
        <div class="pv-quote-mark">“</div>
        <div class="pv-quote-text">${esc(list[0])}</div>
        ${list[1] ? `<div class="pv-quote-by">— ${esc(list[1])}</div>` : ''}
      </div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const list = (bullets || []).map((b) => A.stripEmphasis(b)).filter(Boolean);
      if (!list.length) return;
      slide.addText('“', {
        x: box.x, y: box.y, w: 0.7, h: 0.7,
        fontSize: 44, bold: true, color: theme.light, valign: 'middle',
      });
      slide.addText(list[0], {
        x: box.x + 0.66, y: box.y + 0.06, w: box.w - 0.8, h: box.h * 0.6,
        fontSize: 17, italic: true, color: theme.primary, valign: 'top', lineSpacingMultiple: 1.25,
      });
      if (list[1]) {
        slide.addText(`— ${list[1]}`, {
          x: box.x + 0.66, y: box.y + box.h * 0.68, w: box.w - 0.8, h: 0.36,
          fontSize: 11, color: theme.subtext, valign: 'top',
        });
      }
    },
  };

  // ---------- 3×3マトリクス ----------
  // 2軸を高・中・低の3段階で切り、9マスに分類する。
  const MATRIX3_AXIS = ['高', '中', '低'];
  const matrix3x3 = {
    id: 'matrix-3x3',
    name: '3×3マトリクス',
    category: 'マトリクス／ポジショニング',
    description: '2軸を高・中・低の3段階で切って9マスに分類する。項目は左上から右下の順に9個まで。',
    score(section) {
      const text = A.fullText(section);
      const kw = A.hasAny(text, ['9マス', '九', '3×3', 'マトリクス', '象限']);
      const n = (section.bullets || []).length;
      if (kw && n >= 6 && n <= 9) return 9;
      return 0;
    },
    renderBody(bullets) {
      const items = A.extractItems(bullets).slice(0, 9);
      if (items.length < 4) return emptyBody();
      return `<div class="pv-m3">${Array.from({ length: 9 }, (_, i) => {
        const it = items[i];
        const row = Math.floor(i / 3);
        return `<div class="pv-m3-cell pv-m3-r${row}">${
          it ? `<b>${esc(it.key)}</b>${it.value !== it.key ? `<span>${esc(it.value)}</span>` : ''}` : ''
        }</div>`;
      }).join('')}</div>`;
    },
    buildBody(slide, bullets, theme, box) {
      const items = A.extractItems(bullets).slice(0, 9);
      if (items.length < 4) return;
      const axisW = 0.28;
      const gridX = box.x + axisW;
      const gridW = box.w - axisW;
      const gridH = box.h - axisW;
      const cw = gridW / 3;
      const ch = gridH / 3;
      // 行が上ほど濃くなるように塗り分ける（上＝高）
      const rowFills = [theme.light, theme.lighter, theme.pale];
      for (let r = 0; r < 3; r++) {
        for (let c = 0; c < 3; c++) {
          const i = r * 3 + c;
          const x = gridX + c * cw;
          const y = box.y + r * ch;
          slide.addShape('rect', {
            x, y, w: cw - 0.03, h: ch - 0.03,
            fill: { color: rowFills[r] }, line: { type: 'none' },
          });
          const it = items[i];
          if (!it) continue;
          slide.addText(
            [
              { text: it.key, options: { bold: true, fontSize: 10, color: theme.primary, breakLine: it.value !== it.key } },
              { text: it.value !== it.key ? it.value : '', options: { fontSize: 9, color: theme.text } },
            ],
            { x: x + 0.1, y: y + 0.06, w: cw - 0.23, h: ch - 0.15, valign: 'top' }
          );
        }
        slide.addText(MATRIX3_AXIS[r], {
          x: box.x, y: box.y + r * ch, w: axisW - 0.06, h: ch,
          fontSize: 9, color: theme.subtext, align: 'center', valign: 'middle',
        });
      }
      MATRIX3_AXIS.slice().reverse().forEach((label, c) => {
        slide.addText(label, {
          x: gridX + c * cw, y: box.y + gridH, w: cw, h: axisW,
          fontSize: 9, color: theme.subtext, align: 'center', valign: 'middle',
        });
      });
    },
  };

  DocAssist.patterns = [
    titleMessage,
    bulletNumbered,
    bulletHeadingGroups,
    headingBullets2Col,
    bulletTwoCol,
    checklist,
    genericTable,
    labelBoxBulletRows,
    notebookHeadingBullets,
    flagHeadingBoxGroups,
    chevronBoxGroups,
    recapAgendaSplit,
    agenda,
    sectionDivider,
    pointCallout,
    agendaCards,
    quoteSlide,
    boxCompare,
    compareVertical,
    swot,
    matrix2x2,
    quadrantPriority,
    matrix3x3,
    businessModelCanvas,
    positioningMap,
    processFlow,
    flowVertical,
    chevronFlow,
    diamondFlow,
    stepArrows,
    swimLane,
    cycle,
    cycleQuadrant,
    funnel,
    funnelHorizontal,
    waterfall,
    pyramid,
    pyramidTiered,
    triangleTiers,
    triangleRelation,
    conclusionFlow,
    vennDiagram,
    concentricCircles,
    honeycomb,
    valueChain,
    orgChart,
    decisionTree,
    logicTree,
    beforeAfter,
    timeline,
    timelineVertical,
    yearTimeline,
    customerJourney,
    scheduleTable,
    comparisonTable,
    crossTable,
    dotMatrix,
    prosCons,
    iconHeaderBox,
    decisionRequest,
    actionPlanTable,
    taskStatusTable,
    issueRiskList,
    categorizedIssueResponseTable,
    blockerRequest,
    asIsToBe,
    ganttChart,
    milestoneChevron,
    kpiSummary,
    barChart,
    lineChart,
    pieChart,
    radarChart,
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
    'bullet-numbered': { scenes: ['提案書', '報告書', '経営・定例報告', '研修・説明会'], role: '本文' },
    'bullet-heading-groups': { scenes: ['提案書', '報告書', '経営・定例報告'], role: '本文' },
    'heading-bullets-2col': { scenes: ['提案書', '報告書', '経営・定例報告'], role: '本文' },
    'bullet-two-col': { scenes: ['提案書', '報告書', '経営・定例報告', 'データ提示'], role: '本文' },
    checklist: { scenes: ['進捗MTG', 'キックオフ', '報告書'], role: '表' },
    'generic-table': { scenes: ['提案書', '報告書', 'データ提示', '経営・定例報告'], role: '表' },
    'label-box-bullet-rows': { scenes: ['提案書', '報告書', '経営・定例報告'], role: '本文' },
    'notebook-heading-bullets': { scenes: ['提案書', '報告書'], role: '本文' },
    'flag-heading-box-groups': { scenes: ['提案書', '報告書', '経営・定例報告'], role: '本文' },
    'chevron-box-groups': { scenes: ['提案書', 'キックオフ', '構成図'], role: '図解' },
    'recap-agenda-split': { scenes: ['進捗MTG', 'キックオフ', '経営・定例報告'], role: '目次' },
    agenda: { scenes: ['提案書', '報告書', 'キックオフ', '進捗MTG'], role: '目次' },
    'section-divider': { scenes: ['提案書', '報告書', '研修・説明会', '経営・定例報告'], role: '目次' },
    'point-callout': { scenes: ['提案書', '報告書', '経営・定例報告'], role: '本文' },
    quote: { scenes: ['提案書', '報告書', '経営・定例報告'], role: '本文' },
    'dot-matrix': { scenes: ['提案書', '報告書', 'データ提示'], role: '表' },
    'pros-cons': { scenes: ['提案書', '報告書', '経営・定例報告'], role: '比較' },
    'matrix-3x3': { scenes: ['提案書', '汎用ロジック図解', '経営・定例報告'], role: '図解' },
    'business-model-canvas': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'swim-lane': { scenes: ['構成図', 'キックオフ', '報告書'], role: '図解' },
    honeycomb: { scenes: ['提案書', '汎用ロジック図解'], role: '図解' },
    'value-chain': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'customer-journey': { scenes: ['提案書', 'データ提示', '経営・定例報告'], role: '計画' },
    'agenda-cards': { scenes: ['提案書', '報告書', 'キックオフ', '研修・説明会'], role: '目次' },
    'venn-diagram': { scenes: ['提案書', '汎用ロジック図解', '経営・定例報告'], role: '図解' },
    'concentric-circles': { scenes: ['提案書', '経営・定例報告', '汎用ロジック図解'], role: '図解' },
    'org-chart': { scenes: ['キックオフ', '報告書', '会社・採用'], role: '図解' },
    'decision-tree': { scenes: ['提案書', '汎用ロジック図解', '報告書'], role: '図解' },
    'schedule-table': { scenes: ['研修・説明会', 'キックオフ', '進捗MTG'], role: '計画' },
    'funnel-horizontal': { scenes: ['データ提示', '経営・定例報告'], role: '図解' },
    'radar-chart': { scenes: ['データ提示', '経営・定例報告', '提案書'], role: 'KPI' },
    'cross-table': { scenes: ['提案書', '報告書', 'データ提示'], role: '表' },
    'icon-header-box': { scenes: ['提案書', '報告書', '汎用ロジック図解'], role: '比較' },
    'conclusion-flow': { scenes: ['提案書', '報告書', '汎用ロジック図解'], role: '図解' },
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
    'categorized-issue-response-table': { scenes: ['提案書', '報告書', '進捗MTG'], role: '表' },
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
    'title-message': ['背景：これまでの検討経緯を踏まえた**重要な語句**を含む説明が入ります', '論点：補足の説明テキストが入ります', '3点目の説明文がラベル無しで入ります（自動で番号が振られます）'],
    'bullet-numbered': ['1つ目の論点：現状の業務プロセスに関する課題', '2つ目の論点：システム間のデータ連携における課題', '3つ目の論点：運用体制における課題'],
    'bullet-heading-groups': ['背景：', '既存システムの保守期限が2027年に到来する', '運用コストが年々増加している', '課題：', '担当者ごとに対応品質にばらつきがある', '対策：', 'クラウド型システムへの移行を検討する'],
    'heading-bullets-2col': ['現状の課題：', '対応品質にばらつきがある', '情報共有に時間がかかる', '期待される効果：', '対応品質を平準化できる', '情報共有の時間を半減できる'],
    'bullet-two-col': ['項目1：説明テキストが入ります', '項目2：説明テキストが入ります', '項目3：説明テキストが入ります', 'ラベルの無い説明文もここに入ります', '項目5：説明テキストが入ります', '項目6：説明テキストが入ります', '項目7：説明テキストが入ります', '項目8：説明テキストが入ります'],
    checklist: ['確認事項：★導入範囲について合意済み', '確認事項：予算枠の確保', '確認事項：プロジェクトオーナーの任命', '確認事項：★キックオフ日程の確定'],
    'generic-table': ['列：拠点｜担当者｜連絡先', '東京本社：山田｜03-xxxx-xxxx', '大阪支社：佐藤｜06-xxxx-xxxx', '名古屋支社：鈴木｜052-xxxx-xxxx'],
    'label-box-bullet-rows': ['背景：', '顧客情報が部門ごとに分散管理されている', '対応品質のばらつきと解約率上昇が発生している', '課題認識：', '担当者ごとに検索・共有の手間が発生している', '本活動の目的：', 'クラウド型CRMによる一元化を実現する'],
    'notebook-heading-bullets': ['背景：', '既存システムの保守期限が2027年に到来する', '運用コストが年々増加している、特にライセンス費用が課題である', '-サーバー保守費用', '-ライセンス更新費用', '-運用委託費用', '目的：', '以上の背景を踏まえ、クラウド型システムへの移行を提案する', '-移行スケジュールの策定', '-移行対象システムの選定'],
    'flag-heading-box-groups': ['背景：', '顧客情報が部門ごとに分散管理されている', '検索・共有に平均15分を要している', '課題認識：', '対応品質にばらつきが生じている', '解約率が前年比で上昇している', '本活動の目的：', 'クラウド型CRMによる一元管理を実現する'],
    'chevron-box-groups': ['背景：', '顧客情報が部門ごとに分散している', '検索・共有に時間を要している', '課題：', '対応品質にばらつきが生じている', '解約率が上昇している', '対策：', 'クラウド型CRMを導入する', '全社の対応履歴を一元管理する'],
    'recap-agenda-split': [
      '5/21(火) 第2回検討WGのサマリ',
      'WGでの検討対象施策の棚卸',
      '-対象範囲は管理課及び人事課の施策とすることで合意',
      '-議論が深まった段階で給与課へのアプローチを検討する',
      'タレマネ選定の進め方',
      '-試験導入検証の候補選定、評価のロジック案を確認',
      '-同案をもとに評価資料を作成して整理を進める',
      '---',
      '本日の討議内容',
      'WGでの検討対象施策の棚卸：',
      '管理課及び人事課施策の棚卸状況と概要の確認を行いたい',
      '同各施策の全体スケジュールへの反映方法の確認を行いたい',
      'タレマネシステム選定の進め方の整理：',
      'デジタル統括室で整理いただいた評価資料内容の確認を行いたい',
      '試験導入検証の進め方について認識共有を行いたい',
    ],
    agenda: ['本日の論点', '現状と課題', 'ご提案', '今後の進め方'],
    'box-compare': ['A案：低コストだが拡張性に課題', 'B案：★推奨 バランスが最も良い', 'C案：高機能だが過剰投資'],
    'compare-vertical': ['初期費用：A社30万円、B社50万円', '運用コスト：A社1万円、B社3万円', '拡張性：★推奨 B社はオプションで拡張可能', 'サポート：B社は電話・チャット対応'],
    'section-divider': ['01：現状と課題', '02：★ ご提案', '03：実行計画'],
    'point-callout': ['意思決定の迅速化：稟議の所要日数を半減できる', '運用負荷の軽減：手作業の帳票出力が不要になる', '拡張性の確保：将来の機能追加に対応できる'],
    'cross-table': ['列：初期費用｜運用コスト｜拡張性', 'A社：30万円｜月1万円｜限定的', 'B社：★推奨 50万円｜月3万円｜拡張可能', 'C社：120万円｜月8万円｜充実'],
    'icon-header-box': ['経営：意思決定の迅速化｜投資対効果の可視化', '現場：入力作業の削減｜情報共有の円滑化', 'システム：運用保守の負荷軽減｜セキュリティの担保'],
    'conclusion-flow': ['前提：保守期限が2027年に到来する', '検討：コスト｜拡張性｜移行負荷', '結論：SaaS基盤への移行を推奨する'],
    'agenda-cards': ['背景と目的：なぜ今この論点を扱うのか', '現状の課題：定量・定性の両面から整理', 'ご提案：3案の比較と推奨', '導入効果：定量効果の見込み', '実行計画：体制とスケジュール', '今後の進め方：次回までの宿題'],
    quote: ['現場で本当に必要なのは、機能の多さではなく毎日確実に使えることだ', '導入企業A社　営業部長（2026年1月ヒアリング）'],
    'dot-matrix': ['列：コスト｜機能｜サポート｜拡張性', 'A社：◎｜△｜○｜×', 'B社：★推奨 ○｜◎｜◎｜○', 'C社：×｜◎｜○｜◎'],
    'pros-cons': ['メリット：初期費用を抑えられる｜短期間で導入できる｜既存の運用を変えずに済む', 'デメリット：将来の機能追加に制約がある｜他システムとの連携が限定的'],
    'matrix-3x3': ['最優先：即時に着手する', '優先：今期中に着手する', '検討：来期の候補とする', '注視：状況を見て判断する', '保留：条件が整えば着手', '見送り：当面は対応しない', '要検討：費用対効果を精査', '低優先：余力があれば', '対象外：実施しない'],
    'business-model-canvas': ['パートナー：システムベンダー、業務委託先', '主要活動：システム開発、運用保守', 'リソース：技術者、顧客データ', '価値提案：業務時間を半減する仕組み', '顧客との関係：専任担当による伴走支援', 'チャネル：直販とパートナー経由', '顧客セグメント：中堅製造業の情報システム部門', 'コスト構造：人件費、インフラ費用', '収益の流れ：月額利用料、導入支援費'],
    'swim-lane': ['列：申請｜承認｜処理｜完了', '営業部：申請書を作成｜｜｜受領を確認', '営業部長：｜内容を承認｜｜', '業務部：｜｜システムに登録｜通知を送付'],
    honeycomb: ['中心：顧客体験の向上', '認知：接点を増やす', '検討：情報を届ける', '購入：手続きを簡素化', '利用：定着を支援する', '継続：関係を深める'],
    'value-chain': ['支援活動：全社インフラ｜人事・労務管理｜技術開発｜調達', '主活動：購買物流｜製造｜出荷物流｜販売・マーケティング｜サービス'],
    'customer-journey': ['認知：Web広告・展示会で知る｜4', '情報収集：資料請求・比較サイト｜3', '検討：営業への問い合わせ｜2', '導入：契約・初期設定｜3', '定着：問い合わせ対応｜4'],
    'venn-diagram': ['営業部門の知見：顧客の業務実態を把握している', '技術部門の知見：実現可能性を判断できる', '重なり：実行可能で顧客に刺さる提案力'],
    'concentric-circles': ['理念：社会課題の解決に貢献する', '戦略：デジタル基盤への集中投資', '施策：CRM導入と業務標準化'],
    'org-chart': ['プロジェクトオーナー：情報システム部長', '業務要件チーム：営業部｜カスタマーサポート部', 'システム構築チーム：情報システム部｜開発ベンダー', '推進事務局：経営企画部'],
    'decision-tree': ['判断：既存システムを活かせるか', 'はい：改修で対応する｜期間3ヶ月・費用50万円', 'いいえ：★推奨 新規に導入する｜期間6ヶ月・費用200万円'],
    'schedule-table': ['9:30-10:00：オープニング・本日のゴール共有｜事務局', '10:00-11:30：現状課題の共有｜業務部門', '11:30-12:30：休憩', '12:30-14:30：解決案のディスカッション｜全員', '14:30-15:00：次回までのアクション確認｜事務局'],
    'funnel-horizontal': ['認知：10000件', '興味：6000件', '比較検討：3000件', '商談：1200件', '受注：300件'],
    'radar-chart': ['系列：自社｜競合A', '提案力：4｜3', '技術力：5｜4', '価格競争力：3｜5', 'サポート体制：4｜3', '実績：3｜5'],
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
    'categorized-issue-response-table': [
      '列：カテゴリー｜課題｜課題内容｜対応案',
      'オンライン化手法：【課題1】eメールが選ばれているもののオンライン化の余地がある｜オンライン化手法にeメールが多く選ばれている／オンライン化のメリットが理解されていない可能性がある｜オンライン化のメリットを説明し、既存プラットフォームの利点を伝える',
      'オンライン化手法：【課題2】新規システムは開発コストが大きくなる可能性がある｜新規システム構築を検討する手続が多い／既存システムを活用することで開発コストを削減できる可能性がある｜具体的なオンライン化手法を紹介し、既存プラットフォームの活用を再考いただく',
      'オンライン化時期：【課題3】回答期限が遅い手続が多い｜期限までの回答が多く、十分な検討に着手できていない可能性がある｜検討を具体的に進めていただくために、説明会や検討支援を実施',
    ],
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
    addStyledTable,
    addNumberCircle,
    scaleColor,
    ROUND,
    HAIRLINE,
    ICON_KEYS: Object.keys(ICONS),
  };
})();
