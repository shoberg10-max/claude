// 構成案（アウトライン）から実際の PowerPoint (.pptx) ファイルを生成する。
//
// 官公庁向け報告書・提案書を意識し、スライド1枚ごとに
//   1. タイトル（トピックラベル・小さめ）
//   2. メッセージ帯（リード文＋サブメッセージ。結論を先に示す「ピラミッド原則」型）
//   3. ボディ（パターンごとの図・表・チャートで、メッセージの根拠を示す）
// という3層構成を共通で描いたあと、選択されたパターンの buildBody() を呼び出して
// ボディ部分のレイアウトを構築する。
window.DocAssist = window.DocAssist || {};

(function () {
  const SLIDE_W = 13.33;
  const SLIDE_H = 7.5;

  const HEADER_Y = 0;
  const HEADER_H = 0.52;
  const MESSAGE_Y = HEADER_H + 0.1;
  const MESSAGE_H_SHORT = 0.62;
  const MESSAGE_H_TALL = 1.05;
  const BODY_GAP = 0.14;
  const BODY_BOTTOM_MARGIN = 0.32;
  const BODY_X = 0.5;
  const BODY_W = SLIDE_W - 1.0;
  const BODY_PADDING = 0.16;

  // タイトル（トピックラベル）とページ番号だけの控えめなヘッダー。
  // 官公庁向け資料でよく見る「太字の帯」ではなく、細い罫線＋小さな見出しにしている。
  function addHeader(slide, heading, pageNum, theme) {
    slide.addShape('rect', { x: 0, y: HEADER_Y, w: SLIDE_W, h: 0.035, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addShape('rect', { x: 0.5, y: 0.16, w: 0.12, h: 0.12, fill: { color: theme.accent }, line: { type: 'none' } });
    slide.addText(heading || '(見出し未設定)', {
      x: 0.74, y: 0.06, w: SLIDE_W - 2.2, h: 0.36,
      fontSize: 14, bold: true, color: theme.primary, valign: 'middle', fontFace: 'Meiryo UI',
    });
    if (pageNum) {
      slide.addText(String(pageNum), {
        x: SLIDE_W - 1.2, y: 0.06, w: 0.7, h: 0.36,
        fontSize: 10, color: theme.subtext, align: 'right', valign: 'middle',
      });
    }
    slide.addShape('line', { x: 0, y: HEADER_H, w: SLIDE_W, h: 0, line: { color: theme.border, width: 0.75 } });
  }

  // リード文（結論）＋サブメッセージの帯。左に太いアクセントバーを置く
  // 「メッセージ・ボックス」型で、コンサルティングファームの報告書に多い形式。
  function addMessageBlock(slide, message, subMessage, theme) {
    const hasSub = !!(subMessage && subMessage.trim());
    const h = hasSub ? MESSAGE_H_TALL : MESSAGE_H_SHORT;
    slide.addShape('rect', { x: BODY_X, y: MESSAGE_Y, w: BODY_W, h, fill: { color: theme.light }, line: { type: 'none' } });
    slide.addShape('rect', { x: BODY_X, y: MESSAGE_Y, w: 0.07, h, fill: { color: theme.primaryLight }, line: { type: 'none' } });
    slide.addText(message || '', {
      x: BODY_X + 0.28, y: MESSAGE_Y + (hasSub ? 0.08 : 0.05),
      w: BODY_W - 0.5, h: hasSub ? 0.5 : h - 0.1,
      fontSize: 15, bold: true, color: theme.primary, valign: hasSub ? 'top' : 'middle', fontFace: 'Meiryo UI',
    });
    if (hasSub) {
      slide.addText(subMessage, {
        x: BODY_X + 0.28, y: MESSAGE_Y + 0.58, w: BODY_W - 0.5, h: 0.4,
        fontSize: 11, color: theme.subtext, valign: 'top',
      });
    }
    return h;
  }

  function bodyBoxAfterMessage(messageH) {
    const y = MESSAGE_Y + messageH + BODY_GAP;
    const h = SLIDE_H - y - BODY_BOTTOM_MARGIN;
    return { x: BODY_X, y, w: BODY_W, h };
  }

  function addCoverSlide(pptx, title, theme) {
    const slide = pptx.addSlide();
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addShape('rect', { x: 0, y: SLIDE_H / 2 + 0.6, w: 2.2, h: 0.08, fill: { color: theme.accent }, line: { type: 'none' } });
    slide.addText(title || '資料構成案', {
      x: 1.0, y: SLIDE_H / 2 - 1.0, w: SLIDE_W - 2.0, h: 1.6,
      fontSize: 32, bold: true, color: theme.white, fontFace: 'Meiryo UI',
    });
  }

  // outline: { title, slides: [{ heading, message, subMessage, bullets, patternId }] }
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
      addHeader(slide, s.heading, idx + 1, theme);
      const messageH = addMessageBlock(slide, A.effectiveMessage(s), s.subMessage, theme);
      const outerBox = bodyBoxAfterMessage(messageH);
      slide.addShape('rect', {
        x: outerBox.x, y: outerBox.y, w: outerBox.w, h: outerBox.h,
        fill: { type: 'none' }, line: { color: theme.border, width: 0.75 },
      });
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
