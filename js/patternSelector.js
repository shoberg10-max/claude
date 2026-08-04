// 構成案の各スライド（見出し＋箇条書き）から、最もふさわしいデザインパターンを自動選択する。
window.DocAssist = window.DocAssist || {};

(function () {
  // section: {heading, bullets}. patterns省略時は DocAssist.patterns を使う。
  // 戻り値: { pattern, score, ranked: [{pattern, score}, ...] }（scoreの降順）
  function selectPattern(section, patterns) {
    const list = patterns || DocAssist.patterns;
    const ranked = list
      .map((pattern) => ({ pattern, score: safeScore(pattern, section) }))
      .sort((a, b) => b.score - a.score);
    return { pattern: ranked[0].pattern, score: ranked[0].score, ranked };
  }

  function safeScore(pattern, section) {
    try {
      return pattern.score(section) || 0;
    } catch (e) {
      console.error('パターン採点でエラー:', pattern.id, e);
      return 0;
    }
  }

  DocAssist.selectPattern = selectPattern;
})();
