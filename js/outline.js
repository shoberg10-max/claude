// 会議メモ等のフリーテキストから「構成案」(タイトル＋スライド見出し＋箇条書き) を組み立てる。
//
// 将来、社内の隔離LLM環境に接続する場合は、この parseNotes をそのまま使う代わりに
// DocAssist.outlineProvider を上書きして、LLMが返した構成案JSONを
// { title, slides: [{ heading, bullets }] } の形で返すようにすればよい。
// UI側 (app.js) は DocAssist.outlineProvider(text) を呼ぶだけなので、
// ルールベース版とLLM版を差し替え可能にしてある。
//
// 見出し行に "## [pattern: パターンID] 見出し文" の形式でタグが付いている場合、
// タグ部分を取り除いた見出し文と一緒に slide.taggedPatternId としてIDを返す。
// これは js/promptBuilder.js が生成するプロンプトの出力フォーマットで、
// 社内LLMにこの形式で構成案を作らせることで、パターン自動選択のヒントを
// APIを使わずに（コピペで）渡せるようにするためのもの。
//
// 見出しの直後に "> リード文（メッセージ）" と ">> サブメッセージ" の行があれば、
// それぞれ slide.message / slide.subMessage として取り出す。官公庁向け報告書に
// 典型的な「タイトル→メッセージ→ボディ」の3層構成をこのアプリ全体で共通化するための
// 入力形式で、js/app.js と js/pptxExport.js の両方がこの2項目を使って
// スライド上部のメッセージ帯を描画する。
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

  function extractPatternTag(headingText) {
    const m = headingText.match(/^\[pattern:\s*([a-z0-9-]+)\]\s*(.*)$/i);
    if (!m) return { patternId: null, heading: headingText };
    return { patternId: m[1].toLowerCase(), heading: m[2].trim() || headingText };
  }

  function isSubMessageLine(line) {
    return /^>>/.test(line);
  }
  function isMessageLine(line) {
    return /^>/.test(line) && !isSubMessageLine(line);
  }
  function stripMessage(line) {
    return line.replace(/^>\s*/, '').trim();
  }
  function stripSubMessage(line) {
    return line.replace(/^>>\s*/, '').trim();
  }

  // 「出所：」「出典：」で始まる行を出所表記として取り出し、箇条書きから除外する。
  function extractSourceNote(lines) {
    const idx = lines.findIndex((l) => /^(出所|出典)[：:]/.test(l));
    if (idx === -1) return { sourceNote: '', rest: lines };
    const sourceNote = lines[idx].replace(/^(出所|出典)[：:]\s*/, '').trim();
    const rest = lines.slice(0, idx).concat(lines.slice(idx + 1));
    return { sourceNote, rest };
  }

  // 見出しの直後に並ぶ ">" / ">>" 行をリード文・サブメッセージとして取り出し、
  // 残りを本文（箇条書き候補の行）として返す。
  function extractMessage(bodyLines) {
    let i = 0;
    let message = '';
    let subMessage = '';
    if (bodyLines[i] && isSubMessageLine(bodyLines[i])) {
      subMessage = stripSubMessage(bodyLines[i]);
      i++;
    } else if (bodyLines[i] && isMessageLine(bodyLines[i])) {
      message = stripMessage(bodyLines[i]);
      i++;
      if (bodyLines[i] && isSubMessageLine(bodyLines[i])) {
        subMessage = stripSubMessage(bodyLines[i]);
        i++;
      }
    }
    return { message, subMessage, rest: bodyLines.slice(i) };
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

      let rawHeading;
      let bodyLines;
      if (isHeadingLine(lines[0])) {
        rawHeading = stripHeading(lines[0]);
        bodyLines = lines.slice(1);
      } else if (lines.length > 1) {
        rawHeading = lines[0];
        bodyLines = lines.slice(1);
      } else {
        rawHeading = lines[0];
        bodyLines = [];
      }

      const { patternId, heading } = extractPatternTag(rawHeading);
      const { message, subMessage, rest } = extractMessage(bodyLines);
      const { sourceNote, rest: bulletLines } = extractSourceNote(rest);
      const slide = {
        heading: heading || '(見出し未設定)',
        message,
        subMessage,
        sourceNote,
        bullets: bulletLines.map(stripBullet).filter(Boolean),
      };
      if (patternId) slide.taggedPatternId = patternId;
      slides.push(slide);
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
