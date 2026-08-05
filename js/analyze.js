// 構成案の各スライド（見出し＋箇条書き）を解析するための共通ヘルパー。
// パターン自動選択（patternSelector.js）と、プレビュー／PPTX書き出し（patterns.js）の両方から使う。
window.DocAssist = window.DocAssist || {};

(function () {
  function fullText(section) {
    return [section.heading || '', ...(section.bullets || [])].join('\n');
  }

  function hasAny(text, words) {
    return words.some((w) => text.includes(w));
  }

  function countAny(text, words) {
    return words.reduce((n, w) => n + (text.includes(w) ? 1 : 0), 0);
  }

  // 「見出し：本文」「見出し: 本文」のような1行を項目名／内容に分割する。
  function splitKV(line) {
    const m = line.match(/^\s*(.{1,24}?)\s*[：:]\s*(.+)\s*$/);
    if (!m) return null;
    if (!m[1] || !m[2]) return null;
    return { key: m[1].trim(), value: m[2].trim() };
  }

  function kvRatio(bullets) {
    if (!bullets || !bullets.length) return 0;
    const hits = bullets.filter((b) => splitKV(b)).length;
    return hits / bullets.length;
  }

  // 「★」または「★推奨」が含まれる行を「この項目を特に強調する（推奨案など）」の
  // 印として扱う。マーカー自体は取り除き、item.highlight = true を立てる。
  // 比較系パターン（ボックス比較・縦型比較・比較表など）がこれを見て、その項目だけ
  // オレンジ等の予約色でハイライトする（多用しないための唯一のトリガー）。
  const HIGHLIGHT_RE = /★\s*(?:推奨)?/;

  // 箇条書きを {key, value, highlight} の配列に変換する。KV形式でない行は
  // 「項目N」を仮の見出しとして割り当てる。
  // 「**強調**」構文には対応していない図表系パターンで使われるため、
  // ここで一律にマーカーを取り除いておく（生の ** が出力に漏れるのを防ぐ安全策）。
  function extractItems(bullets) {
    return (bullets || []).map((b, i) => {
      const stripped = stripEmphasis(b);
      const highlight = HIGHLIGHT_RE.test(stripped);
      const clean = highlight ? stripped.replace(HIGHLIGHT_RE, '').trim() : stripped;
      const kv = splitKV(clean);
      if (kv) return Object.assign({}, kv, { highlight });
      return { key: `項目${i + 1}`, value: clean, highlight };
    });
  }

  const NUMBER_RE = /\d+(?:\.\d+)?\s*(?:%|％|億|万|千|件|人|円|pt|ポイント)/g;
  function numberTokenCount(text) {
    const m = text.match(NUMBER_RE);
    return m ? m.length : 0;
  }

  const DATE_RE = /(\d{4}\s*年|\d{1,2}\s*月|Q[1-4]|[1-4]\s*Q|フェーズ\s*\d|第\s*\d\s*四半期)/g;
  function dateTokenCount(text) {
    const m = text.match(DATE_RE);
    return m ? m.length : 0;
  }

  const SEQUENCE_WORDS = ['まず', '次に', 'つぎに', 'その後', '最後に', 'はじめに', 'ステップ', 'Step', 'STEP'];
  function sequenceWordCount(text) {
    return countAny(text, SEQUENCE_WORDS);
  }

  const NUMBERED_BULLET_RE = /^\s*(?:\d+[\.\)]|[①②③④⑤⑥⑦⑧⑨⑩])/;
  function numberedBulletCount(bullets) {
    return (bullets || []).filter((b) => NUMBERED_BULLET_RE.test(b)).length;
  }

  // 「**強調したい語句**」の形式で書かれた文章を、太字部分とそれ以外に分解する。
  // NRIの報告書に見られる「文中の重要な語句だけ太字＋ネイビーにする」表現を
  // 再現するための共通パーサー。プレビュー用HTML描画（<b>タグ）とPPTX書き出し
  // （text runの配列）の両方でこの分解結果を使う。
  function parseEmphasisTokens(text) {
    const parts = String(text == null ? '' : text).split(/(\*\*[^*]+\*\*)/g).filter((p) => p !== '');
    if (!parts.length) return [{ text: '', bold: false }];
    return parts.map((part) => {
      const m = part.match(/^\*\*([^*]+)\*\*$/);
      return m ? { text: m[1], bold: true } : { text: part, bold: false };
    });
  }

  function stripEmphasis(text) {
    return String(text == null ? '' : text).replace(/\*\*([^*]+)\*\*/g, '$1');
  }

  // parseEmphasisTokens を PptxGenJS の text run 配列に変換する共通ヘルパー。
  // 強調部分には bold:true と boldColor（指定時）を追加する。
  function emphasisRuns(text, boldColor, baseOptions) {
    const base = baseOptions || {};
    return parseEmphasisTokens(text).map((t) => {
      const opts = Object.assign({}, base);
      if (t.bold) {
        opts.bold = true;
        if (boldColor) opts.color = boldColor;
      }
      return { text: t.text, options: opts };
    });
  }

  // スライドに明示的なリード文（message）が無い場合のフォールバック。
  // 官公庁向け報告書の「タイトル→メッセージ→ボディ」の型を崩さないよう、
  // 未入力でも空欄のまま出力しないためのもの（先頭の箇条書き、それも無ければ見出しを流用）。
  function effectiveMessage(slide) {
    if (slide.message && slide.message.trim()) return slide.message.trim();
    if (slide.bullets && slide.bullets.length) return slide.bullets[0];
    return slide.heading || '';
  }

  DocAssist.analyze = {
    fullText,
    hasAny,
    countAny,
    splitKV,
    kvRatio,
    extractItems,
    numberTokenCount,
    dateTokenCount,
    sequenceWordCount,
    numberedBulletCount,
    effectiveMessage,
    parseEmphasisTokens,
    stripEmphasis,
    emphasisRuns,
  };
})();
