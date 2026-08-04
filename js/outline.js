// 会議メモ等のフリーテキストから「構成案」(タイトル＋スライド見出し＋箇条書き) を組み立てる。
//
// 将来、社内の隔離LLM環境に接続する場合は、この parseNotes をそのまま使う代わりに
// DocAssist.outlineProvider を上書きして、LLMが返した構成案JSONを
// { title, slides: [{ heading, bullets }] } の形で返すようにすればよい。
// UI側 (app.js) は DocAssist.outlineProvider(text) を呼ぶだけなので、
// ルールベース版とLLM版を差し替え可能にしてある。
window.DocAssist = window.DocAssist || {};

(function () {
  function stripBullet(line) {
    return line.replace(/^\s*(?:[-*・●○◆■]|\d+[\.\)]|[①②③④⑤⑥⑦⑧⑨⑩])\s*/, '').trim();
  }

  function isBulletLine(line) {
    return /^\s*(?:[-*・●○◆■]|\d+[\.\)]|[①②③④⑤⑥⑦⑧⑨⑩])\s+/.test(line);
  }

  function isHeadingLine(line) {
    return /^\s*#{1,3}\s+/.test(line);
  }

  function stripHeading(line) {
    return line.replace(/^\s*#{1,3}\s+/, '').trim();
  }

  function parseNotes(rawText) {
    const text = (rawText || '').replace(/\r\n?/g, '\n').trim();
    if (!text) return { title: '資料構成案', slides: [] };

    // 先頭行が「# タイトル」形式なら、直後に空行があるかどうかに関わらず
    // 文書タイトルとして取り出しておく（残りは通常のブロック分割にかける）。
    let title = '';
    let bodyText = text;
    const firstLine = text.split('\n', 1)[0].trim();
    if (/^#\s+/.test(firstLine)) {
      title = stripHeading(firstLine);
      bodyText = text.slice(text.indexOf('\n') + 1).trim();
      if (text.indexOf('\n') === -1) bodyText = '';
    }

    const blocks = bodyText
      .split(/\n\s*\n+/)
      .map((b) => b.split('\n').map((l) => l.trim()).filter((l) => l !== ''))
      .filter((b) => b.length > 0);

    const slides = [];
    for (let b = 0; b < blocks.length; b++) {
      const lines = blocks[b];
      const allBullets = lines.every(isBulletLine);

      if (allBullets && slides.length) {
        // 見出しなしの箇条書きだけのブロックは、直前のスライドへの追記とみなす
        slides[slides.length - 1].bullets.push(...lines.map(stripBullet));
        continue;
      }

      let heading;
      let bodyLines;
      if (isHeadingLine(lines[0])) {
        heading = stripHeading(lines[0]);
        bodyLines = lines.slice(1);
      } else if (lines.length > 1) {
        heading = lines[0];
        bodyLines = lines.slice(1);
      } else {
        heading = lines[0];
        bodyLines = [];
      }

      slides.push({
        heading: heading || '(見出し未設定)',
        bullets: bodyLines.map(stripBullet).filter(Boolean),
      });
    }

    slides.forEach((s, idx) => {
      s.id = 'slide-' + (idx + 1);
    });

    return { title: title || '資料構成案', slides };
  }

  DocAssist.parseNotes = parseNotes;
  // デフォルトはルールベース版。差し替え可能にするためプロバイダとして公開する。
  DocAssist.outlineProvider = parseNotes;
})();
