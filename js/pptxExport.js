// 構成案（アウトライン）から実際の PowerPoint (.pptx) ファイルを生成する。
// スライド1枚ごとに、共通のタイトルバーを描いたあと、選択されたパターンの
// buildBody() を呼び出して本文レイアウトを構築する。
window.DocAssist = window.DocAssist || {};

(function () {
  const SLIDE_W = 13.33;
  const SLIDE_H = 7.5;
  const CONTENT_BOX = { x: 0.5, y: 1.25, w: 12.33, h: 5.85 };

  function addTitleBar(slide, heading, patternName, theme) {
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: 0.95, fill: { color: theme.primary }, line: { type: 'none' } });
    slide.addShape('rect', { x: 0, y: 0.95, w: SLIDE_W, h: 0.05, fill: { color: theme.accent }, line: { type: 'none' } });
    slide.addText(heading || '(見出し未設定)', {
      x: 0.5, y: 0, w: SLIDE_W - 3.0, h: 0.95,
      fontSize: 22, bold: true, color: theme.white, valign: 'middle', fontFace: 'Meiryo UI',
    });
    if (patternName) {
      slide.addText(patternName, {
        x: SLIDE_W - 2.9, y: 0, w: 2.6, h: 0.95,
        fontSize: 10, color: theme.white, align: 'right', valign: 'middle', italic: true,
      });
    }
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

  // outline: { title, slides: [{ heading, bullets, patternId }] }
  // options.fileName: 保存ファイル名（省略時は資料タイトルから自動生成）
  function buildPptx(outline) {
    if (typeof PptxGenJS === 'undefined') {
      throw new Error('PptxGenJS が読み込まれていません（js/vendor/pptxgen.bundle.js を確認してください）');
    }
    const theme = DocAssist.theme;
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'DOC_ASSIST_WIDE', width: SLIDE_W, height: SLIDE_H });
    pptx.layout = 'DOC_ASSIST_WIDE';
    pptx.author = '資料作成支援アプリ';
    pptx.title = outline.title || '資料構成案';

    addCoverSlide(pptx, outline.title, theme);

    (outline.slides || []).forEach((s) => {
      const pattern = DocAssist.patternById[s.patternId] || DocAssist.patternById['title-message'];
      const slide = pptx.addSlide();
      addTitleBar(slide, s.heading, pattern.name, theme);
      try {
        pattern.buildBody(slide, s.bullets || [], theme, CONTENT_BOX);
      } catch (e) {
        console.error('スライド生成エラー:', s.heading, e);
        slide.addText('このスライドの生成中にエラーが発生しました: ' + e.message, {
          x: CONTENT_BOX.x, y: CONTENT_BOX.y, w: CONTENT_BOX.w, h: 1, fontSize: 12, color: 'C00000',
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
