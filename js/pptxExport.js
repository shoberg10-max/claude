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
  const SLIDE_W = 13.33;
  const SLIDE_H = 7.5;

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

  // ヘッダー：紺の縦棒＋パンくず（heading内に「｜」があれば複数セグメントとして扱う）＋
  // 結論を言い切る大きな太字ネイビーの見出し文（メッセージ）＋直下の細いヘアライン罫線。
  // 見出し文の長さに応じて1行/2行を判定し、罫線の位置を調整する。
  function addHeader(slide, heading, message, subMessage, theme) {
    const A = DocAssist.analyze;
    const crumbText = (heading || '(見出し未設定)').trim();
    slide.addShape('rect', { x: BODY_X, y: HEADER_Y + 0.02, w: BAR_W, h: CRUMB_H - 0.04, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addText(crumbText, {
      x: BODY_X + BAR_W + 0.1, y: HEADER_Y, w: BODY_W - BAR_W - 0.1, h: CRUMB_H,
      fontSize: 11, bold: true, color: theme.text, valign: 'middle', fontFace: 'Meiryo UI',
    });

    const msg = message || '';
    const isLong = A.stripEmphasis(msg).length > 34;
    const headlineH = isLong ? HEADLINE_H_2LINE : HEADLINE_H_1LINE;
    slide.addText(A.emphasisRuns(msg, theme.primary, { bold: true }), {
      x: BODY_X, y: HEADLINE_Y, w: BODY_W, h: headlineH,
      fontSize: 20, color: theme.primary, valign: 'top',
      fontFace: 'Meiryo UI', lineSpacingMultiple: 1.15,
    });

    let y = HEADLINE_Y + headlineH;
    const hasSub = !!(subMessage && subMessage.trim());
    if (hasSub) {
      slide.addText(A.emphasisRuns(subMessage.trim(), theme.primary), {
        x: BODY_X, y, w: BODY_W, h: 0.3,
        fontSize: 11, color: theme.subtext, valign: 'top',
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
    const slide = pptx.addSlide();
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addShape('rect', { x: MARGIN_X, y: SLIDE_H / 2 + 0.62, w: 1.1, h: 0.045, fill: { color: theme.accent }, line: { type: 'none' } });
    slide.addText(title || '資料構成案', {
      x: MARGIN_X, y: SLIDE_H / 2 - 1.0, w: SLIDE_W - MARGIN_X * 2, h: 1.6,
      fontSize: 30, bold: true, color: theme.white, fontFace: 'Meiryo UI',
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
    pptx.defineLayout({ name: 'DOC_ASSIST_WIDE', width: SLIDE_W, height: SLIDE_H });
    pptx.layout = 'DOC_ASSIST_WIDE';
    pptx.author = '資料作成支援アプリ';
    pptx.title = outline.title || '資料構成案';

    addCoverSlide(pptx, outline.title, theme);

    (outline.slides || []).forEach((s, idx) => {
      const pattern = DocAssist.patternById[s.patternId] || DocAssist.patternById['title-message'];
      const slide = pptx.addSlide();
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

  DocAssist.buildPptx = buildPptx;
  DocAssist.exportPptx = exportPptx;
})();
