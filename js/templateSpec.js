// ユーザーが自分で「型（テンプレート）」を追加できるようにするための、宣言的なレイアウト仕様。
//
// これまでのデザインパターンは js/patterns.js にJavaScriptで直接書く必要があったが、
// 普段の業務のなかで「この型を追加したい」と思ったときにコードを書かずに追加できるよう、
// レイアウトをJSONで表現し、それを解釈して描画する汎用レンダラーをここに置く。
//
// 1つの仕様（spec）から、プレビュー用HTML（renderBody）とPowerPoint描画（buildBody）の
// 両方を生成する。組み込みパターンと同じ形のオブジェクトを返すので、生成されたテンプレートは
// ギャラリー・自動選択・PPTX書き出しのすべてでまったく同じように扱われる。
//
// ---- 仕様の形 ----
// {
//   id, name, category, description,
//   scenes: ['提案書'],        // ギャラリーの「用途」チップ
//   role: '比較',              // ギャラリーの「役割」チップ
//   keywords: ['比較', '検討'], // 自動選択に使うキーワード（空なら自動選択の候補にしない）
//   minItems: 2, maxItems: 5,  // 自動選択が働く箇条書きの件数レンジ
//   sample: ['A案：説明', ...], // ギャラリーのサムネイル用サンプル
//   layout: {
//     kind: 'cards' | 'table',
//
//     // kind: 'cards' — 箇条書き1件を1枚のカードとして、横または縦に並べる
//     direction: 'row' | 'column',
//     shape: 'rect' | 'roundRect' | 'chevron',
//     numbered: false,          // 連番バッジを付ける
//     icon: false,              // 項目名からアイコンを自動判定して表示する
//     connector: 'none' | 'arrow',
//     head: { show: true, from: 'key', fill: 'primary' },
//     body: { from: 'value', fill: 'light', bullet: false },
//
//     // kind: 'table' — 箇条書き1件を1行として表に並べる
//     zebra: true,
//     columns: [{ label: '項目', from: 'key', width: 0.3, badge: '', bold: false }],
//   }
// }
//
// from に指定できる値：
//   'key'      … 「見出し：内容」の見出し部分
//   'value'    … 同じく内容部分
//   'field:N'  … 内容を「｜」で区切ったN番目（0始まり）
//   'index'    … 1始まりの連番
window.DocAssist = window.DocAssist || {};

(function () {
  const A = DocAssist.analyze;
  const H = DocAssist.patternHelpers;

  const SPEC_VERSION = 1;

  // 塗りに使える色トークン。テーマ（js/theme.js）の色名にそのまま対応する。
  const COLOR_TOKENS = ['primary', 'primaryLight', 'accent', 'accent2', 'gray', 'light', 'lighter', 'white', 'highlight', 'critical', 'none'];
  // 濃い塗りの上は白文字、淡い塗りの上は本文色にする。
  const DARK_FILLS = ['primary', 'primaryLight', 'accent', 'accent2', 'gray', 'highlight', 'critical'];
  const CSS_VARS = {
    primary: 'var(--primary)',
    primaryLight: 'var(--primary-light)',
    accent: 'var(--accent)',
    accent2: 'var(--accent2)',
    gray: 'var(--gray)',
    light: 'var(--light)',
    lighter: 'var(--lighter)',
    highlight: 'var(--highlight)',
    critical: 'var(--critical)',
    white: '#fff',
    none: 'transparent',
  };

  function fillCss(token) {
    return CSS_VARS[token] || CSS_VARS.primary;
  }
  function fillColor(token, theme) {
    return theme[token] || theme.primary;
  }
  function textOnCss(token) {
    return DARK_FILLS.indexOf(token) !== -1 ? '#fff' : 'var(--text)';
  }
  function textOnColor(token, theme) {
    return DARK_FILLS.indexOf(token) !== -1 ? theme.white : theme.text;
  }

  // 「★推奨」が付いた項目は、塗りを highlight に差し替えて1件だけ目立たせる。
  function withHighlight(token, item) {
    return item && item.highlight ? 'highlight' : token;
  }

  // spec の from 指定を、実際の項目データから解決する。
  function resolveFrom(from, item, index) {
    const key = String(from == null ? 'value' : from);
    if (key === 'index') return String(index + 1);
    if (key === 'key') return item.key || '';
    if (key === 'value') return item.value || '';
    const m = key.match(/^field:(\d+)$/);
    if (m) {
      const fields = H.splitFields(item.value);
      return fields[parseInt(m[1], 10)] || '';
    }
    return '';
  }

  // 未指定の項目を既定値で埋め、値域を検証した spec を返す。
  // 手書きJSONを読み込む場合もあるので、壊れた値が来ても描画が落ちないようにする。
  function normalizeSpec(raw) {
    const spec = Object.assign({}, raw || {});
    const layout = Object.assign({}, spec.layout || {});
    layout.kind = layout.kind === 'table' ? 'table' : 'cards';

    if (layout.kind === 'cards') {
      layout.direction = layout.direction === 'column' ? 'column' : 'row';
      layout.shape = ['rect', 'roundRect', 'chevron'].indexOf(layout.shape) !== -1 ? layout.shape : 'rect';
      layout.connector = layout.connector === 'arrow' ? 'arrow' : 'none';
      layout.numbered = !!layout.numbered;
      layout.icon = !!layout.icon;
      const head = Object.assign({ show: true, from: 'key', fill: 'primary' }, layout.head || {});
      const body = Object.assign({ from: 'value', fill: 'light', bullet: false }, layout.body || {});
      head.fill = COLOR_TOKENS.indexOf(head.fill) !== -1 ? head.fill : 'primary';
      body.fill = COLOR_TOKENS.indexOf(body.fill) !== -1 ? body.fill : 'light';
      head.show = !!head.show;
      body.bullet = !!body.bullet;
      layout.head = head;
      layout.body = body;
      // シェブロンは横並びでしか意味を成さない
      if (layout.shape === 'chevron') layout.direction = 'row';
      // 表レイアウト専用の項目が残っていると書き出したJSONが紛らわしくなるので落とす
      delete layout.columns;
      delete layout.zebra;
    } else {
      layout.zebra = layout.zebra !== false;
      let cols = Array.isArray(layout.columns) ? layout.columns : [];
      if (!cols.length) cols = [{ label: '項目', from: 'key', width: 0.3 }, { label: '内容', from: 'value', width: 0.7 }];
      cols = cols.slice(0, 6).map((c) => ({
        label: String(c.label == null ? '' : c.label),
        from: c.from || 'value',
        width: typeof c.width === 'number' && c.width > 0 ? c.width : 1 / cols.length,
        badge: c.badge === 'status' || c.badge === 'severity' ? c.badge : '',
        bold: !!c.bold,
      }));
      // 幅の合計を1に正規化する
      const sum = cols.reduce((a, c) => a + c.width, 0) || 1;
      cols.forEach((c) => {
        c.width = c.width / sum;
      });
      layout.columns = cols;
      // カードレイアウト専用の項目を落とす（上と対称）
      ['direction', 'shape', 'connector', 'numbered', 'icon', 'head', 'body'].forEach((k) => delete layout[k]);
    }

    spec.layout = layout;
    spec.specVersion = SPEC_VERSION;
    spec.id = spec.id || 'custom-' + Math.random().toString(36).slice(2, 9);
    spec.name = spec.name || '無題のテンプレート';
    spec.category = spec.category || '汎用';
    spec.description = spec.description || 'ユーザー定義テンプレート';
    spec.scenes = Array.isArray(spec.scenes) ? spec.scenes : [];
    spec.role = spec.role || '本文';
    spec.keywords = Array.isArray(spec.keywords) ? spec.keywords.filter(Boolean) : [];
    spec.minItems = typeof spec.minItems === 'number' ? spec.minItems : 2;
    spec.maxItems = typeof spec.maxItems === 'number' ? spec.maxItems : 6;
    spec.sample = Array.isArray(spec.sample) && spec.sample.length
      ? spec.sample
      : ['項目A：説明テキストが入ります', '項目B：説明テキストが入ります', '項目C：説明テキストが入ります'];
    return spec;
  }

  // ---------------- カードレイアウト ----------------
  function cardsHtml(spec, items) {
    const L = spec.layout;
    const dirCls = L.direction === 'column' ? 'is-column' : 'is-row';
    const shapeCls = L.shape === 'chevron' ? ' is-chevron' : L.shape === 'roundRect' ? ' is-round' : '';
    const parts = items.map((it, i) => {
      const headFill = withHighlight(L.head.fill, it);
      const bodyFill = L.body.fill;
      const headText = resolveFrom(L.head.from, it, i);
      const bodyText = resolveFrom(L.body.from, it, i);
      const head = L.head.show
        ? `<div class="ts-card-head" style="background:${fillCss(headFill)};color:${textOnCss(headFill)};">${
            L.numbered ? `<span class="ts-num">${i + 1}</span>` : ''
          }${it.highlight ? '★ ' : ''}${H.esc(headText)}</div>`
        : '';
      const icon = L.icon ? H.iconSvg(H.guessIcon(headText || bodyText)) : '';
      const body = `<div class="ts-card-body${L.body.bullet ? ' is-bullet' : ''}" style="background:${fillCss(bodyFill)};color:${textOnCss(
        bodyFill
      )};">${icon}<span>${H.esc(bodyText)}</span></div>`;
      const card = `<div class="ts-card${shapeCls}${it.highlight ? ' is-highlight' : ''}">${head}${body}</div>`;
      const needsArrow = L.connector === 'arrow' && i < items.length - 1;
      return card + (needsArrow ? `<div class="ts-arrow">${L.direction === 'column' ? '▼' : '▶'}</div>` : '');
    });
    return `<div class="ts-cards ${dirCls}">${parts.join('')}</div>`;
  }

  function cardsPptx(spec, items, slide, theme, box) {
    const L = spec.layout;
    const n = items.length;
    if (!n) return;
    const isRow = L.direction === 'row';
    const arrowSize = L.connector === 'arrow' ? (isRow ? 0.3 : 0.24) : 0;
    const gap = L.shape === 'chevron' ? 0 : 0.16;
    const overlap = L.shape === 'chevron' ? 0.12 : 0;

    const span = isRow ? box.w : box.h;
    const cell = (span - gap * (n - 1) - arrowSize * (n - 1) + overlap * (n - 1)) / n;

    items.forEach((it, i) => {
      const offset = i * (cell + gap + arrowSize - overlap);
      const x = isRow ? box.x + offset : box.x;
      const y = isRow ? box.y : box.y + offset;
      const w = isRow ? cell : box.w;
      const h = isRow ? box.h : cell;

      const headFill = withHighlight(L.head.fill, it);
      const headText = resolveFrom(L.head.from, it, i);
      const bodyText = resolveFrom(L.body.from, it, i);

      if (L.shape === 'chevron') {
        // シェブロンは1枚の図形の中に見出しと本文をまとめて置く
        const runs = [];
        if (L.head.show && headText) runs.push({ text: `${it.highlight ? '★ ' : ''}${headText}`, options: { bold: true, fontSize: 11, breakLine: true } });
        if (bodyText) runs.push({ text: bodyText, options: { fontSize: 9.5 } });
        const fill = withHighlight(i % 2 === 0 ? 'primary' : 'primaryLight', it);
        slide.addShape(i === 0 ? 'homePlate' : 'chevron', { x, y, w, h, fill: { color: fillColor(fill, theme) }, line: { type: 'none' } });
        if (runs.length) {
          slide.addText(runs, {
            x: x + 0.18, y, w: w - 0.36, h,
            color: theme.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.15,
          });
        }
        return;
      }

      const shapeName = L.shape === 'roundRect' ? 'roundRect' : 'rect';
      const rectOpts = L.shape === 'roundRect' ? { rectRadius: 0.06 } : {};

      // 横並びは上に見出し帯、縦並びは左に見出し帯を置く
      const headSize = L.head.show ? (isRow ? Math.min(0.5, h * 0.32) : Math.min(w * 0.28, 2.4)) : 0;
      const bodyX = isRow ? x : x + headSize;
      const bodyY = isRow ? y + headSize : y;
      const bodyW = isRow ? w : w - headSize;
      const bodyH = isRow ? h - headSize : h;

      slide.addShape(shapeName, Object.assign({
        x: bodyX, y: bodyY, w: bodyW, h: bodyH,
        fill: { color: fillColor(L.body.fill, theme) },
        line: it.highlight ? { color: theme.highlight, width: 1.5 } : { type: 'none' },
      }, rectOpts));

      if (L.head.show) {
        slide.addShape(shapeName, Object.assign({
          x, y, w: isRow ? w : headSize, h: isRow ? headSize : h,
          fill: { color: fillColor(headFill, theme) }, line: { type: 'none' },
        }, rectOpts));
        slide.addText(`${it.highlight ? '★ ' : ''}${headText}`, {
          x: x + 0.06, y, w: (isRow ? w : headSize) - 0.12, h: isRow ? headSize : h,
          fontSize: isRow ? 12 : 11, bold: true, color: textOnColor(headFill, theme),
          align: isRow ? 'center' : 'left', valign: 'middle',
        });
      }

      let textX = bodyX + 0.12;
      let textW = bodyW - 0.24;
      if (L.icon) {
        const size = Math.min(0.38, bodyH * 0.4, bodyW * 0.3);
        H.addIcon(slide, H.guessIcon(headText || bodyText), { x: bodyX + 0.12, y: bodyY + 0.1, w: size, h: size }, theme);
        if (isRow) {
          // 横並びはアイコンを上、本文をその下に置く
          slide.addText(bodyText, {
            x: textX, y: bodyY + 0.1 + size + 0.08, w: textW, h: bodyH - size - 0.28,
            fontSize: 10.5, color: textOnColor(L.body.fill, theme), valign: 'top',
          });
          if (L.numbered && !L.head.show) addNumberBadge(slide, i, theme, x, y);
          return;
        }
        textX = bodyX + 0.12 + size + 0.12;
        textW = bodyW - (size + 0.36);
      }

      const bodyOpts = {
        x: textX, y: bodyY + (isRow ? 0.08 : 0), w: textW, h: bodyH - (isRow ? 0.16 : 0),
        fontSize: 11, color: textOnColor(L.body.fill, theme), valign: isRow ? 'top' : 'middle',
      };
      if (L.body.bullet) {
        slide.addText([{ text: bodyText, options: { bullet: H.SQUARE_BULLET() } }], bodyOpts);
      } else {
        slide.addText(bodyText, bodyOpts);
      }
      if (L.numbered && !L.head.show) addNumberBadge(slide, i, theme, x, y);
    });

    // コネクタ（矢印）
    if (L.connector === 'arrow') {
      for (let i = 0; i < n - 1; i++) {
        const offset = i * (cell + gap + arrowSize) + cell + gap / 2;
        if (isRow) {
          slide.addShape('rightArrow', {
            x: box.x + offset, y: box.y + box.h / 2 - 0.1, w: arrowSize - 0.06, h: 0.2,
            fill: { color: theme.primary }, line: { type: 'none' },
          });
        } else {
          slide.addShape('downArrow', {
            x: box.x + box.w / 2 - 0.1, y: box.y + offset, w: 0.2, h: arrowSize - 0.05,
            fill: { color: theme.primary }, line: { type: 'none' },
          });
        }
      }
    }
  }

  function addNumberBadge(slide, i, theme, x, y) {
    slide.addShape('oval', { x: x + 0.06, y: y + 0.06, w: 0.34, h: 0.34, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addText(String(i + 1), {
      x: x + 0.06, y: y + 0.06, w: 0.34, h: 0.34,
      fontSize: 11, bold: true, color: theme.white, align: 'center', valign: 'middle',
    });
  }

  // ---------------- 表レイアウト ----------------
  function badgeCls(kind, text) {
    if (kind === 'status') return H.statusCls(text);
    if (kind === 'severity') return H.severityCls(text);
    return '';
  }
  function badgeTone(kind, text, theme) {
    if (kind === 'status') return H.statusTone(text, theme);
    if (kind === 'severity') return H.severityTone(text, theme);
    return null;
  }

  function tableHtml(spec, items) {
    const cols = spec.layout.columns;
    const head = cols.map((c) => `<th style="width:${(c.width * 100).toFixed(1)}%;">${H.esc(c.label)}</th>`).join('');
    const rows = items
      .map((it, i) => {
        const cells = cols
          .map((c) => {
            const raw = resolveFrom(c.from, it, i);
            if (c.badge) return `<td><span class="pv-badge ${badgeCls(c.badge, raw)}">${H.esc(raw)}</span></td>`;
            const b = c.bold || it.highlight;
            return `<td>${b ? '<b>' : ''}${it.highlight && c === cols[0] ? '★ ' : ''}${H.esc(raw)}${b ? '</b>' : ''}</td>`;
          })
          .join('');
        return `<tr${it.highlight ? ' class="pv-table-highlight"' : ''}>${cells}</tr>`;
      })
      .join('');
    return `<table class="pv-table"><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table>`;
  }

  // 組み込みパターンと同じ表スタイル（縦罫線なし・横は点線・表側は淡いブルー）を使う。
  function tablePptx(spec, items, slide, theme, box) {
    const L = spec.layout;
    const cols = L.columns;
    H.addStyledTable(slide, theme, {
      x: box.x, y: box.y, w: box.w,
      colW: cols.map((c) => box.w * c.width),
      fontSize: 11,
      header: cols.map((c) => c.label),
      firstColAccent: true,
      zebra: !!L.zebra,
      rows: items.map((it, i) =>
        cols.map((c, ci) => {
          const raw = resolveFrom(c.from, it, i);
          const tone = badgeTone(c.badge, raw, theme);
          if (tone) return { text: raw, options: { fill: { color: tone }, color: theme.white, bold: true, align: 'center' } };
          return {
            text: `${it.highlight && ci === 0 ? '★ ' : ''}${raw}`,
            options: Object.assign(
              c.bold ? { bold: true } : {},
              it.highlight ? { bold: true, color: ci === 0 ? theme.highlight : theme.text } : {}
            ),
          };
        })
      ),
    });
  }

  // ---------------- spec → パターンオブジェクト ----------------
  // 組み込みパターンとまったく同じ形（score / renderBody / buildBody）を返すので、
  // 呼び出し側は自作テンプレートかどうかを意識しなくてよい。
  function createPatternFromSpec(rawSpec) {
    const spec = normalizeSpec(rawSpec);
    const limit = Math.max(1, spec.maxItems || 6);

    return {
      id: spec.id,
      name: spec.name,
      category: spec.category,
      description: spec.description,
      scenes: spec.scenes,
      role: spec.role,
      sample: spec.sample,
      isCustom: true,
      spec,

      score(section) {
        const bullets = section.bullets || [];
        const n = bullets.length;
        if (n < spec.minItems || n > spec.maxItems) return 0;
        if (!spec.keywords.length) return 0; // キーワード未設定なら自動選択の候補にしない（手動選択専用）
        const text = A.fullText(section);
        const hits = spec.keywords.filter((k) => text.indexOf(k) !== -1).length;
        if (!hits) return 0;
        return Math.min(10, 6 + hits * 2);
      },

      renderBody(bullets) {
        const items = A.extractItems(bullets).slice(0, limit);
        if (!items.length) return H.emptyBody();
        return spec.layout.kind === 'table' ? tableHtml(spec, items) : cardsHtml(spec, items);
      },

      buildBody(slide, bullets, theme, box) {
        const items = A.extractItems(bullets).slice(0, limit);
        if (!items.length) return;
        if (spec.layout.kind === 'table') tablePptx(spec, items, slide, theme, box);
        else cardsPptx(spec, items, slide, theme, box);
      },
    };
  }

  DocAssist.templateSpec = {
    SPEC_VERSION,
    COLOR_TOKENS,
    normalizeSpec,
    createPatternFromSpec,
    resolveFrom,
  };
})();
