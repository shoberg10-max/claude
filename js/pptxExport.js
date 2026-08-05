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

  // MBB/Big4系の報告書に典型的な、装飾を抑えたレイアウト。
  // タイポグラフィと細い罫線（ヘアライン）だけで階層を作り、色つきボックスや
  // 角丸カード、太い枠線は使わない（「テンプレート感」の最大の原因になるため）。
  const MARGIN_X = 0.6;
  const BODY_X = MARGIN_X;
  const BODY_W = SLIDE_W - MARGIN_X * 2;

  const HEADER_Y = 0.32;
  const HEADER_H = 0.34;
  const RULE_GAP = 0.1;
  const MESSAGE_Y = HEADER_Y + HEADER_H + RULE_GAP + 0.14;
  const MESSAGE_H_SHORT = 0.46;
  const MESSAGE_H_TALL = 0.86;
  const MESSAGE_RULE_GAP = 0.16;
  const BODY_GAP = 0.14;
  const BODY_BOTTOM_MARGIN = 0.34;
  const BODY_PADDING = 0.06;

  // タイトル（トピックラベル）とページ番号だけの控えめなヘッダー。
  // 色つきの帯やアイコンは使わず、細い罫線と控えめな文字だけで構成する。
  function addHeader(slide, heading, pageNum, theme) {
    slide.addText((heading || '(見出し未設定)').toUpperCase(), {
      x: BODY_X, y: HEADER_Y, w: SLIDE_W - MARGIN_X - 1.2, h: HEADER_H,
      fontSize: 10.5, bold: true, color: theme.subtext, charSpacing: 1, valign: 'bottom', fontFace: 'Meiryo UI',
    });
    if (pageNum) {
      slide.addText(String(pageNum), {
        x: SLIDE_W - MARGIN_X - 0.6, y: HEADER_Y, w: 0.6, h: HEADER_H,
        fontSize: 9, color: theme.subtext, align: 'right', valign: 'bottom',
      });
    }
    slide.addShape('line', {
      x: BODY_X, y: HEADER_Y + HEADER_H + RULE_GAP, w: BODY_W, h: 0,
      line: { color: theme.primary, width: 1.25 },
    });
  }

  // リード文（結論）＋サブメッセージ。色つきの箱には入れず、太字の地の文として
  // 大きく置き、下に細いヘアラインを1本引いてボディと区切るだけにする。
  function addMessageBlock(slide, message, subMessage, theme) {
    const hasSub = !!(subMessage && subMessage.trim());
    const h = hasSub ? MESSAGE_H_TALL : MESSAGE_H_SHORT;
    slide.addText(message || '', {
      x: BODY_X, y: MESSAGE_Y, w: BODY_W, h: hasSub ? 0.5 : h,
      fontSize: 17, bold: true, color: theme.text, valign: hasSub ? 'top' : 'middle',
      fontFace: 'Meiryo UI', lineSpacingMultiple: 1.12,
    });
    if (hasSub) {
      slide.addText(subMessage, {
        x: BODY_X, y: MESSAGE_Y + 0.52, w: BODY_W, h: 0.34,
        fontSize: 11, color: theme.subtext, valign: 'top',
      });
    }
    slide.addShape('line', {
      x: BODY_X, y: MESSAGE_Y + h + MESSAGE_RULE_GAP, w: BODY_W, h: 0,
      line: { color: theme.border, width: 0.75 },
    });
    return h + MESSAGE_RULE_GAP;
  }

  function bodyBoxAfterMessage(messageH) {
    const y = MESSAGE_Y + messageH + BODY_GAP;
    const h = SLIDE_H - y - BODY_BOTTOM_MARGIN;
    return { x: BODY_X, y, w: BODY_W, h };
  }

  function addCoverSlide(pptx, title, theme) {
    const slide = pptx.addSlide();
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addShape('rect', { x: MARGIN_X, y: SLIDE_H / 2 + 0.62, w: 1.1, h: 0.045, fill: { color: theme.white }, line: { type: 'none' } });
    slide.addText(title || '資料構成案', {
      x: MARGIN_X, y: SLIDE_H / 2 - 1.0, w: SLIDE_W - MARGIN_X * 2, h: 1.6,
      fontSize: 30, bold: true, color: theme.white, fontFace: 'Meiryo UI',
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
