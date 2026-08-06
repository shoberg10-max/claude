// 既存のPowerPointファイルをテンプレートとして読み込み、スライドマスターのテーマ色・
// フォント・箇条書き記号をこのアプリの配色（js/theme.js）に反映する機能。
//
// .pptx は実体がZIPファイルで、中に ppt/theme/theme1.xml（配色・フォントの定義）と
// ppt/slideMasters/slideMaster1.xml（見出し・本文の既定スタイル、箇条書き記号）を
// XMLとして持っている。この2つを読み、このアプリの4層構成（ヘッダー／メッセージ／
// ボディ／フッター）はそのままに、色とフォントと箇条書き記号だけを差し替える。
//
// 意図的にやらないこと：
//   - スライドの配置（プレースホルダーの位置・サイズ）をそのまま複製すること。
//     テンプレートごとにレイアウトはバラバラで、これをやると本アプリの一貫した
//     4層構成が崩れてしまう。取り込むのは「ブランドを識別する要素」（色・フォント・
//     箇条書き記号）に絞る。
//   - accent2〜accent6 をそのまま取り込むこと。テーマの副次アクセント色は赤やオレンジを
//     含むことが珍しくなく、それを取り込むと「オレンジ・レッドは★推奨等の予約色」という
//     このアプリの方針が崩れる。取り込むのは主要ブランド色（accent1）だけとし、そこから
//     青の階調を機械的に生成する（js/theme.jsの7段階と同じ考え方）。
window.DocAssist = window.DocAssist || {};

(function () {
  const STORAGE_KEY = 'docassist.importedTheme.v1';

  // ---- 色のユーティリティ ----
  function hexToRgb(hex) {
    const h = String(hex || '000000').replace('#', '');
    return [parseInt(h.slice(0, 2), 16) || 0, parseInt(h.slice(2, 4), 16) || 0, parseInt(h.slice(4, 6), 16) || 0];
  }
  function rgbToHex([r, g, b]) {
    return [r, g, b].map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  // 白と混ぜて明るくする（pct=0で元の色、1で白）。淡いブルーの階調を機械的に作るための関数。
  function tint(hex, pct) {
    const [r, g, b] = hexToRgb(hex);
    return rgbToHex([r + (255 - r) * pct, g + (255 - g) * pct, b + (255 - b) * pct]);
  }

  // ---- XML読み取りのユーティリティ ----
  // <a:srgbClr val="RRGGBB"/> または <a:sysClr val="windowText" lastClr="RRGGBB"/> を
  // 子要素に持つノードから、実際の色（16進6桁）を取り出す。
  function readDirectColor(node) {
    if (!node) return null;
    const srgb = node.getElementsByTagName('a:srgbClr')[0];
    if (srgb) return srgb.getAttribute('val');
    const sys = node.getElementsByTagName('a:sysClr')[0];
    if (sys) return sys.getAttribute('lastClr') || null;
    return null;
  }

  function parseClrScheme(themeDoc) {
    const scheme = themeDoc.getElementsByTagName('a:clrScheme')[0];
    if (!scheme) return null;
    const slots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const out = {};
    slots.forEach((slot) => {
      const el = scheme.getElementsByTagName('a:' + slot)[0];
      out[slot] = el ? readDirectColor(el) : null;
    });
    return out;
  }

  // <p:clrMap bg1="lt1" tx1="dk1" .../> の属性を読む。無ければPowerPointの既定対応
  // （bg1→lt1, tx1→dk1, bg2→lt2, tx2→dk2, accentN→accentN, hlink/folHlink→そのまま）を使う。
  function parseClrMap(masterDoc) {
    const DEFAULT = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink' };
    if (!masterDoc) return DEFAULT;
    const el = masterDoc.getElementsByTagName('p:clrMap')[0];
    if (!el) return DEFAULT;
    const out = Object.assign({}, DEFAULT);
    Object.keys(DEFAULT).forEach((k) => {
      const v = el.getAttribute(k);
      if (v) out[k] = v;
    });
    return out;
  }

  // schemeClr val="tx1" のような参照を、clrMapを介して実際の色（16進6桁）に解決する。
  function resolveSchemeClr(name, clrScheme, clrMap) {
    if (!name || !clrScheme) return null;
    const slot = clrMap[name] || name; // clrMapに無ければ直接そのスロット名として試す
    return clrScheme[slot] || null;
  }

  function parseFontScheme(themeDoc) {
    const scheme = themeDoc.getElementsByTagName('a:fontScheme')[0];
    if (!scheme) return null;
    function readFont(tagName) {
      const el = scheme.getElementsByTagName(tagName)[0];
      if (!el) return { latin: '', ea: '' };
      const latin = el.getElementsByTagName('a:latin')[0];
      const ea = el.getElementsByTagName('a:ea')[0];
      return {
        latin: (latin && latin.getAttribute('typeface')) || '',
        ea: (ea && ea.getAttribute('typeface')) || '',
      };
    }
    return { major: readFont('a:majorFont'), minor: readFont('a:minorFont') };
  }

  // Wingdings等の記号フォントで定義された箇条書き文字は、そのフォント無しでは
  // 意図した記号にならない（例：Wingdingsの"§"は本来は四角のアイコン）ため、
  // 安全側に倒してそのまま使わず既定の■にフォールバックする。
  const SYMBOL_FONTS = ['wingdings', 'wingdings 2', 'wingdings 3', 'webdings', 'symbol', 'marlett'];
  function isSymbolFont(name) {
    return SYMBOL_FONTS.indexOf(String(name || '').toLowerCase().trim()) !== -1;
  }

  // ppt/slideMasters/slideMaster1.xml の p:txStyles/p:bodyStyle/a:lvl1pPr から
  // 箇条書き記号（a:buChar）を取り出す。記号フォントの場合はnullを返す。
  function parseBodyBullet(masterDoc) {
    if (!masterDoc) return null;
    const bodyStyle = masterDoc.getElementsByTagName('p:bodyStyle')[0];
    if (!bodyStyle) return null;
    const lvl1 = bodyStyle.getElementsByTagName('a:lvl1pPr')[0];
    if (!lvl1) return null;
    const buChar = lvl1.getElementsByTagName('a:buChar')[0];
    if (!buChar) return null;
    const buFont = lvl1.getElementsByTagName('a:buFont')[0];
    const fontName = buFont ? buFont.getAttribute('typeface') : '';
    if (isSymbolFont(fontName)) return null;
    const ch = buChar.getAttribute('char');
    if (!ch || !ch.trim()) return null;
    return ch.trim().charAt(0);
  }

  // zip内のファイル名から、既定名（theme1.xml等）を優先しつつ最初に見つかったものを返す。
  function findZipFile(zip, dir, prefix) {
    const preferred = `ppt/${dir}/${prefix}1.xml`;
    if (zip.files[preferred]) return preferred;
    const re = new RegExp(`^ppt/${dir}/${prefix}\\d+\\.xml$`);
    const names = Object.keys(zip.files).filter((n) => re.test(n)).sort();
    return names[0] || null;
  }

  function parseXml(text) {
    return new DOMParser().parseFromString(text, 'application/xml');
  }

  // 取り込んだ値を、js/theme.jsが持つ7段階の青の階調にあわせて機械的に展開する。
  // primaryだけを本物のテンプレート色にし、残りはそこからの明度違いとして生成する
  // （accent2〜6をそのまま使わない理由はファイル冒頭のコメントを参照）。
  function buildThemePatch(primaryHex) {
    return {
      primary: primaryHex,
      primaryLight: tint(primaryHex, 0.25),
      mid: tint(primaryHex, 0.42),
      accent: tint(primaryHex, 0.58),
      softBlue: tint(primaryHex, 0.7),
      light: tint(primaryHex, 0.82),
      lighter: tint(primaryHex, 0.91),
      pale: tint(primaryHex, 0.95),
      positive: primaryHex, // theme.js の positive は元々 primary と同じ値の運用
    };
  }

  // ---- メイン ----
  // file: <input type="file"> から得た File オブジェクト。
  // 戻り値：プレビュー表示と適用の両方に使う要約オブジェクト。
  async function parseFile(file) {
    if (typeof window.JSZip === 'undefined') {
      throw new Error('ZIP読み込み機能（JSZip）が見つかりません。js/vendor/pptxgen.bundle.js が読み込まれているか確認してください。');
    }
    const buf = await file.arrayBuffer();
    let zip;
    try {
      zip = await window.JSZip.loadAsync(buf);
    } catch (e) {
      throw new Error('ZIPとして読み込めませんでした。.pptxファイルを選んでください。');
    }

    const themePath = findZipFile(zip, 'theme', 'theme');
    if (!themePath) {
      throw new Error('テーマ情報（ppt/theme/theme*.xml）が見つかりませんでした。PowerPointで保存された.pptxファイルか確認してください。');
    }
    const themeXml = parseXml(await zip.files[themePath].async('string'));
    const clrScheme = parseClrScheme(themeXml);
    if (!clrScheme) {
      throw new Error('配色情報を読み取れませんでした。');
    }
    const fontScheme = parseFontScheme(themeXml);

    const masterPath = findZipFile(zip, 'slideMasters', 'slideMaster');
    let masterXml = null;
    if (masterPath) {
      try {
        masterXml = parseXml(await zip.files[masterPath].async('string'));
      } catch (e) {
        masterXml = null; // マスターが読めなくても配色・フォントだけで進める
      }
    }
    const clrMap = parseClrMap(masterXml);

    // ブランドの主要色：accent1を採用する（多くのテーマで見出し・強調に使われる色）。
    // 取得できない・白に近すぎる場合はdk2（濃い方の予備色）にフォールバックする。
    let primaryHex = clrScheme.accent1;
    const isNearWhite = (hex) => {
      if (!hex) return true;
      const [r, g, b] = hexToRgb(hex);
      return r > 235 && g > 235 && b > 235;
    };
    if (isNearWhite(primaryHex)) primaryHex = clrScheme.dk2 || clrScheme.accent1;
    if (isNearWhite(primaryHex)) primaryHex = '000F78'; // 最終フォールバック（このアプリの既定ネイビー）

    // フォント：本文（minor）を優先し、和文（ea）があればそちらを使う
    // （見出しmajorではなく本文minorを優先するのは、このアプリが全文字を単一フォントに
    // 統一する設計のため。分量の多い本文側の指定を優先したほうが実態に合う）。
    let fontFace = '';
    if (fontScheme) {
      fontFace = fontScheme.minor.ea || fontScheme.minor.latin || fontScheme.major.ea || fontScheme.major.latin || '';
    }
    // "+mn-lt" のようなプレースホルダー自己参照は実フォント名ではないので無視する。
    if (/^\+m[nj]-(lt|ea|cs)$/.test(fontFace)) fontFace = '';

    const bulletChar = parseBodyBullet(masterXml);

    return {
      sourceFileName: file.name,
      primary: primaryHex,
      fontFace: fontFace || null,
      bulletChar: bulletChar || null,
      // 参考表示用（適用はしない）：抽出できたテーマ色の生データ
      rawAccents: [1, 2, 3, 4, 5, 6].map((n) => clrScheme['accent' + n]).filter(Boolean),
      rawDark: clrScheme.dk1,
      rawLight: clrScheme.lt1,
    };
  }

  // summary（parseFileの戻り値）を DocAssist.theme に反映し、CSS変数・localStorageも更新する。
  function applyTheme(summary) {
    const patch = buildThemePatch(summary.primary);
    Object.assign(DocAssist.theme, patch);
    if (summary.fontFace) DocAssist.theme.fontFace = summary.fontFace;
    if (summary.bulletChar) DocAssist.theme.bulletChar = summary.bulletChar;
    applyCssVars();
    persist(summary);
  }

  // 元のNRI配色に戻す。
  function resetTheme() {
    Object.assign(DocAssist.theme, DocAssist.defaultTheme);
    applyCssVars();
    try {
      window.localStorage.removeItem(STORAGE_KEY);
    } catch (e) {
      /* localStorageが使えない環境では何もしない */
    }
  }

  // DocAssist.theme の現在値をCSS変数へ反映する（HTMLプレビュー用）。
  const CSS_VAR_MAP = {
    primary: '--primary', primaryLight: '--primary-light', accent: '--accent', accent2: '--accent2',
    gray: '--gray', highlight: '--highlight', critical: '--critical', positive: '--positive', negative: '--negative',
    light: '--light', lighter: '--lighter', mid: '--mid', softBlue: '--soft-blue', pale: '--pale',
    paleGray: '--pale-gray', border: '--border', text: '--text', subtext: '--subtext',
  };
  function applyCssVars() {
    const root = document.documentElement;
    const t = DocAssist.theme;
    Object.keys(CSS_VAR_MAP).forEach((key) => {
      if (t[key]) root.style.setProperty(CSS_VAR_MAP[key], '#' + t[key]);
    });
    if (t.fontFace) {
      root.style.setProperty('--preview-font', `"${t.fontFace}", "Hiragino Kaku Gothic ProN", "Meiryo", "Yu Gothic", sans-serif`);
    }
    if (t.bulletChar) {
      root.style.setProperty('--bullet-char', `'${t.bulletChar}'`);
    }
  }

  function persist(summary) {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(summary));
    } catch (e) {
      /* 保存できなくても致命的ではない（このセッション内では適用済み） */
    }
  }

  function getSaved() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  // 起動時：前回読み込んだテンプレートがあれば自動的に反映する。
  function restoreOnLoad() {
    const saved = getSaved();
    if (saved && saved.primary) {
      const patch = buildThemePatch(saved.primary);
      Object.assign(DocAssist.theme, patch);
      if (saved.fontFace) DocAssist.theme.fontFace = saved.fontFace;
      if (saved.bulletChar) DocAssist.theme.bulletChar = saved.bulletChar;
    }
    applyCssVars(); // 保存が無くても、既定テーマの値をCSS変数へ反映しておく
    return saved;
  }

  DocAssist.pptxImport = {
    parseFile,
    applyTheme,
    resetTheme,
    getSaved,
    tint,
  };

  const restored = restoreOnLoad();
  DocAssist.pptxImport.restoredSummary = restored;
})();
