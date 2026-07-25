// さんすうの問題を作る
const MathTopics = (() => {

  function randInt(min, max) { // inclusive
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function genAddSub() {
    const qs = [];
    for (let i = 0; i < 10; i++) {
      const isAdd = Math.random() < 0.5;
      const digits = randInt(2, 4); // 2〜4けた
      const max = Math.pow(10, digits) - 1;
      const min = Math.pow(10, digits - 1);
      let a = randInt(min, max);
      let b = randInt(min, max);
      if (!isAdd && a < b) [a, b] = [b, a];
      qs.push({
        kind: 'number',
        text: `${a} ${isAdd ? '+' : '−'} ${b} = ?`,
        answer: isAdd ? a + b : a - b
      });
    }
    return qs;
  }

  function genMul() {
    const qs = [];
    for (let i = 0; i < 10; i++) {
      let a, b;
      if (i < 4) {
        // 九九のふくしゅう
        a = randInt(1, 9);
        b = randInt(1, 9);
      } else {
        // 2〜3けた × 1けた
        a = randInt(10, randInt(0, 1) ? 99 : 999);
        b = randInt(2, 9);
      }
      qs.push({
        kind: 'number',
        text: `${a} × ${b} = ?`,
        answer: a * b
      });
    }
    return qs;
  }

  function genDiv() {
    const qs = [];
    for (let i = 0; i < 10; i++) {
      const divisor = randInt(2, 9);
      const quotient = randInt(2, 9);
      const dividend = divisor * quotient;
      qs.push({
        kind: 'number',
        text: `${dividend} ÷ ${divisor} = ?`,
        answer: quotient
      });
    }
    return qs;
  }

  function genDivRem() {
    const qs = [];
    for (let i = 0; i < 10; i++) {
      const divisor = randInt(2, 9);
      const quotient = randInt(1, 9);
      const remainder = randInt(1, divisor - 1);
      const dividend = divisor * quotient + remainder;
      qs.push({
        kind: 'div_rem',
        text: `${dividend} ÷ ${divisor} = ? あまり ?`,
        answer: { q: quotient, r: remainder }
      });
    }
    return qs;
  }

  function genCompare() {
    const qs = [];
    for (let i = 0; i < 10; i++) {
      const digits = randInt(3, 6);
      const min = Math.pow(10, digits - 1);
      const max = Math.pow(10, digits) - 1;
      let a = randInt(min, max);
      let b;
      if (i % 4 === 0) {
        b = a; // たまに同じ数にする
      } else {
        b = randInt(min, max);
      }
      const answer = a > b ? '>' : (a < b ? '<' : '=');
      qs.push({
        kind: 'mc',
        text: `${a} 〇 ${b}`,
        choices: ['<', '>', '='],
        answer
      });
    }
    return qs;
  }

  const PLACE_NAMES = ['一', '十', '百', '千', '一万', '十万'];

  function genPlace() {
    const qs = [];
    for (let i = 0; i < 10; i++) {
      const digits = randInt(4, 6);
      const min = Math.pow(10, digits - 1);
      const max = Math.pow(10, digits) - 1;
      const n = randInt(min, max);
      const placeIdx = randInt(0, digits - 1); // 0=一の位
      const s = String(n);
      const digitAtPlace = Number(s[s.length - 1 - placeIdx]);
      const label = PLACE_NAMES[placeIdx] + 'の位';
      const choicesSet = new Set([digitAtPlace]);
      while (choicesSet.size < 4) {
        choicesSet.add(randInt(0, 9));
      }
      qs.push({
        kind: 'mc',
        text: `${n} の ${label} の数字は？`,
        choices: shuffle([...choicesSet]).map(String),
        answer: String(digitAtPlace)
      });
    }
    return qs;
  }

  const topics = [
    { id: 'add_sub', title: 'たし算・ひき算', emoji: '➕', gen: genAddSub },
    { id: 'mul', title: 'かけ算', emoji: '✖️', gen: genMul },
    { id: 'div', title: 'わり算', emoji: '➗', gen: genDiv },
    { id: 'div_rem', title: 'あまりのあるわり算', emoji: '🔢', gen: genDivRem },
    { id: 'compare', title: '大きい数のくらべっこ', emoji: '⚖️', gen: genCompare },
    { id: 'place', title: '大きい数の位', emoji: '🔟', gen: genPlace }
  ];

  return {
    list: topics,
    getTopic(id) { return topics.find(t => t.id === id); }
  };
})();
