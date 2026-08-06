// 構成案（アウトライン）から実際の PowerPoint (.pptx) ファイルを生成する。
//
// NRIが経済産業省向けに作成した報告書PDF（2024FY/000700, 2023FY/000283）のデザインを
// 実物から色・レイアウトを抽出して再現している。スライド1枚ごとに
//   1. ヘッダー：紺の縦棒＋パンくず形式の小見出し（「｜」区切り）
//   2. メッセージ：結論を一文で言い切る大きな太字ネイビーの見出し文＋細いヘアライン罫線
//   3. ボディ：パターンごとの図・表・チャートで、メッセージの根拠を示す
//   4. フッター：出所表記（あれば）とページ番号
// という構成を共通で描いたあと、選択されたパターンの buildBody() を呼び出して
// ボディ部分のレイアウトを構築する。
window.DocAssist = window.DocAssist || {};

(function () {
  // A4横向き（297mm×210mm）。
  const SLIDE_W = 11.69;
  const SLIDE_H = 8.27;

  // フォントは DocAssist.theme.fontFace を単一ソースとする（既定は Yu Gothic UI）。
  // テンプレート読み込み機能（js/pptxImport.js）がここを書き換えることで、
  // 取り込んだテンプレートの本文フォントを書き出しに反映できる。
  function currentFontFace() {
    return (DocAssist.theme && DocAssist.theme.fontFace) || 'Yu Gothic UI';
  }

  // 本文として読める最小サイズ。表のセルやバッジなど、パターン側の個別実装が
  // 詰め込み目的で9〜11pt程度を指定している箇所が多いため、書き出しの直前に
  // 一箇所でまとめて底上げする（グラフの軸ラベル等、addChartを使う箇所は対象外）。
  const MIN_FONT_SIZE = 12;
  function clampFontSize(opts) {
    if (opts && typeof opts.fontSize === 'number' && opts.fontSize < MIN_FONT_SIZE) {
      opts.fontSize = MIN_FONT_SIZE;
    }
    return opts;
  }

  // slide.addText / addTable に渡す options に既定の fontFace を差し込み、
  // fontSize が最小値を下回っていれば底上げする。
  // 個々のパターン実装（40箇所以上）に手を入れず、フォントを一箇所で統一するための薄いラッパー。
  // 呼び出し側が明示的に fontFace を指定した場合はそちらを優先する。
  function withDefaultFont(slide) {
    const origAddText = slide.addText.bind(slide);
    const origAddTable = slide.addTable.bind(slide);
    function withFont(opts) {
      const merged = opts && opts.fontFace ? opts : Object.assign({ fontFace: currentFontFace() }, opts || {});
      return clampFontSize(merged);
    }
    slide.addText = (text, options) => {
      // 配列（テキストラン）で渡す場合、各ランが自前の fontSize を持つことが多く、
      // トップレベルの options だけを底上げしても効かないため、ラン単位でも底上げする。
      if (Array.isArray(text)) {
        text.forEach((run) => {
          if (run && run.options) clampFontSize(run.options);
        });
      }
      return origAddText(text, withFont(options));
    };
    slide.addTable = (rows, options) => {
      // セル単位で fontSize を上書きしていることがあるため、行・セルも底上げする。
      (rows || []).forEach((row) => {
        (row || []).forEach((cell) => {
          if (cell && typeof cell === 'object' && cell.options) clampFontSize(cell.options);
        });
      });
      return origAddTable(rows, withFont(options));
    };
    return slide;
  }

  const MARGIN_X = 0.55;
  const BODY_X = MARGIN_X;
  const BODY_W = SLIDE_W - MARGIN_X * 2;

  const HEADER_Y = 0.28;
  const BAR_W = 0.07;
  const CRUMB_H = 0.26;
  const HEADLINE_Y = HEADER_Y + CRUMB_H + 0.04;
  const HEADLINE_H_1LINE = 0.42;
  const HEADLINE_H_2LINE = 0.72;
  const RULE_GAP = 0.1;

  const BODY_GAP = 0.16;
  const FOOTER_H = 0.3;
  const BODY_PADDING = 0.05;

  // 取り込んだスライドマスターのタイトルプレースホルダー位置（js/pptxImport.jsが
  // theme.titleLayoutにEMU由来の比率で格納）を、実寸（インチ）に変換する。
  // 未取り込みの場合はnullを返し、呼び出し側は既定のヘッダー位置にフォールバックする。
  function importedTitleBox(theme) {
    const tl = theme && theme.titleLayout;
    if (!tl) return null;
    return {
      x: tl.xFrac * SLIDE_W,
      y: tl.yFrac * SLIDE_H,
      w: tl.wFrac * SLIDE_W,
      h: tl.hFrac * SLIDE_H,
    };
  }

  // ヘッダー：紺の縦棒＋パンくず（heading内に「｜」があれば複数セグメントとして扱う）＋
  // 結論を言い切る大きな太字ネイビーの見出し文（メッセージ）＋直下の細いヘアライン罫線。
  // 見出し文の長さに応じて1行/2行を判定し、罫線の位置を調整する。
  // テンプレートからタイトルレイアウトを取り込んでいる場合は、その位置・サイズ・
  // フォントサイズをそのまま使う（パンくずは取り込んだタイトル枠の上段に収める）。
  function addHeader(slide, heading, message, subMessage, theme) {
    const A = DocAssist.analyze;
    const crumbText = (heading || '(見出し未設定)').trim();
    const titleBox = importedTitleBox(theme);

    const groupX = titleBox ? titleBox.x : BODY_X;
    const groupW = titleBox ? titleBox.w : BODY_W;
    const groupY = titleBox ? titleBox.y : HEADER_Y;

    slide.addShape('rect', { x: groupX, y: groupY + 0.02, w: BAR_W, h: CRUMB_H - 0.04, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addText(crumbText, {
      x: groupX + BAR_W + 0.1, y: groupY, w: groupW - BAR_W - 0.1, h: CRUMB_H,
      fontSize: 12, bold: true, color: theme.text, valign: 'middle',
    });

    const msg = message || '';
    const isLong = A.stripEmphasis(msg).length > 34;
    const headlineY = groupY + CRUMB_H + 0.04;
    const headlineH = titleBox
      ? Math.max(0.32, titleBox.h - CRUMB_H - 0.04)
      : (isLong ? HEADLINE_H_2LINE : HEADLINE_H_1LINE);
    const headlineFontSize = (theme && theme.titleFontSize) || 20;
    slide.addText(A.emphasisRuns(msg, theme.primary, { bold: true }), {
      x: groupX, y: headlineY, w: groupW, h: headlineH,
      fontSize: headlineFontSize, color: theme.primary, valign: 'top',
      lineSpacingMultiple: 1.15,
    });

    let y = headlineY + headlineH;
    const hasSub = !!(subMessage && subMessage.trim());
    if (hasSub) {
      slide.addText(A.emphasisRuns(subMessage.trim(), theme.primary), {
        x: groupX, y, w: groupW, h: 0.3,
        fontSize: 12, color: theme.subtext, valign: 'top',
      });
      y += 0.3;
    }

    const ruleY = y + RULE_GAP;
    slide.addShape('line', {
      x: 0, y: ruleY, w: SLIDE_W, h: 0,
      line: { color: theme.border, width: 0.75 },
    });
    return ruleY;
  }

  function bodyBoxAfterHeader(ruleY) {
    const y = ruleY + BODY_GAP;
    const h = SLIDE_H - y - FOOTER_H - 0.12;
    return { x: BODY_X, y, w: BODY_W, h };
  }

  // フッター：出所表記（左下、あれば）とページ番号（右下）。
  function addFooter(slide, sourceNote, pageNum, theme) {
    const y = SLIDE_H - FOOTER_H;
    if (sourceNote && sourceNote.trim()) {
      slide.addText(`出所：${sourceNote.trim()}`, {
        x: BODY_X, y, w: BODY_W - 0.6, h: FOOTER_H,
        fontSize: 8, color: theme.subtext, valign: 'bottom',
      });
    }
    if (pageNum) {
      slide.addText(String(pageNum), {
        x: SLIDE_W - MARGIN_X - 0.5, y, w: 0.5, h: FOOTER_H,
        fontSize: 9, color: theme.subtext, align: 'right', valign: 'bottom',
      });
    }
  }

  function addCoverSlide(pptx, title, theme) {
    const slide = withDefaultFont(pptx.addSlide());
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addShape('rect', { x: MARGIN_X, y: SLIDE_H / 2 + 0.62, w: 1.1, h: 0.045, fill: { color: theme.highlight }, line: { type: 'none' } });
    slide.addText(title || '資料構成案', {
      x: MARGIN_X, y: SLIDE_H / 2 - 1.0, w: SLIDE_W - MARGIN_X * 2, h: 1.6,
      fontSize: 30, bold: true, color: theme.white,
    });
  }

  // outline: { title, slides: [{ heading, message, subMessage, sourceNote, bullets, patternId }] }
  // options.fileName: 保存ファイル名（省略時は資料タイトルから自動生成）
  function buildPptx(outline) {
    if (typeof PptxGenJS === 'undefined') {
      throw new Error('PptxGenJS が読み込まれていません（js/vendor/pptxgen.bundle.js を確認してください）');
    }
    const theme = DocAssist.theme;
    const A = DocAssist.analyze;
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'DOC_ASSIST_A4', width: SLIDE_W, height: SLIDE_H });
    pptx.layout = 'DOC_ASSIST_A4';
    pptx.author = '資料作成支援アプリ';
    pptx.title = outline.title || '資料構成案';

    addCoverSlide(pptx, outline.title, theme);

    (outline.slides || []).forEach((s, idx) => {
      const pattern = DocAssist.patternById[s.patternId] || DocAssist.patternById['title-message'];
      const slide = withDefaultFont(pptx.addSlide());
      const ruleY = addHeader(slide, s.heading, A.effectiveMessage(s), s.subMessage, theme);
      const outerBox = bodyBoxAfterHeader(ruleY);
      const innerBox = {
        x: outerBox.x + BODY_PADDING,
        y: outerBox.y + BODY_PADDING,
        w: outerBox.w - BODY_PADDING * 2,
        h: outerBox.h - BODY_PADDING * 2,
      };
      try {
        pattern.buildBody(slide, s.bullets || [], theme, innerBox);
      } catch (e) {
        console.error('スライド生成エラー:', s.heading, e);
        slide.addText('このスライドの生成中にエラーが発生しました: ' + e.message, {
          x: innerBox.x, y: innerBox.y, w: innerBox.w, h: 1, fontSize: 12, color: 'C00000',
        });
      }
      addFooter(slide, s.sourceNote, idx + 1, theme);
    });

    return pptx;
  }

  function exportPptx(outline, fileName) {
    const pptx = buildPptx(outline);
    const name = fileName || (outline.title || '資料構成案') + '.pptx';
    return pptx.writeFile({ fileName: name });
  }

  // タイトル取り込み機能（js/pptxImport.js）がプレビューCSSの余白計算に使う、
  // このファイルが実際に使っているスライド寸法・余白の単一ソース。
  DocAssist.slideGeometry = {
    width: SLIDE_W, height: SLIDE_H, marginX: MARGIN_X,
    headerY: HEADER_Y, footerH: FOOTER_H, bodyGap: BODY_GAP, crumbH: CRUMB_H,
  };

  DocAssist.buildPptx = buildPptx;
  DocAssist.exportPptx = exportPptx;
})();
