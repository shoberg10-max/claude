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

  // 箇条書きを {key, value} の配列に変換する。KV形式でない行は
  // 「項目N」を仮の見出しとして割り当てる。
  function extractItems(bullets) {
    return (bullets || []).map((b, i) => {
      const kv = splitKV(b);
      if (kv) return kv;
      return { key: `項目${i + 1}`, value: b };
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
  };
})();
