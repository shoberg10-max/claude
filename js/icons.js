// UI用のドット絵アイコン（マップタイル・ナビ・ステータス）
// キャラと同じ方針で、12x12 の文字グリッド＋アイコンごとのパレット。
//   X=主色  Y=副色  Z=第3色  W=白  G=緑  .=透明
const Icons = (() => {
  const SIZE = 12;

  const G = {
    // ---- 島マップのタイル ----
    tree: {
      colors: { X: '#5EA84C', Y: '#8A5A2B' },
      rows: [
        '....XXXX....',
        '..XXXXXXXX..',
        '.XXXXXXXXXX.',
        '.XXXXXXXXXX.',
        '..XXXXXXXX..',
        '...XXXXXX...',
        '....XXXX....',
        '.....YY.....',
        '.....YY.....',
        '.....YY.....',
        '....YYYY....',
        '............'
      ]
    },
    blossom: {
      colors: { X: '#FF9FBB', Y: '#FFD84D', Z: '#5EA84C' },
      rows: [
        '............',
        '...XX..XX...',
        '..XXXXXXXX..',
        '..XXXYYXXX..',
        '..XXXYYXXX..',
        '..XXXXXXXX..',
        '...XX..XX...',
        '.....ZZ.....',
        '.....ZZ.....',
        '....ZZZZ....',
        '............',
        '............'
      ]
    },
    palm: {
      colors: { X: '#4C9E58', Y: '#A9743F', Z: '#F0DFA8' },
      rows: [
        '..XX....XX..',
        '.XXXXXXXXXX.',
        '..XX.YY.XX..',
        '.....YY.....',
        '.....YY.....',
        '.....YY.....',
        '....YY......',
        '....YY......',
        '...YY.......',
        '..ZZZZZZZZ..',
        '............',
        '............'
      ]
    },
    mountain: {
      colors: { X: '#7A8B99', W: '#FFFFFF' },
      rows: [
        '............',
        '.....XX.....',
        '....XXXX....',
        '....XWWX....',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '.XXXXXXXXXX.',
        'XXXXXXXXXXXX',
        '............',
        '............',
        '............',
        '............'
      ]
    },
    wave: {
      colors: { X: '#4FA3E3' },
      rows: [
        '............',
        '............',
        '............',
        '.XX..XX..XX.',
        'XXXXXXXXXXXX',
        '............',
        '.XX..XX..XX.',
        'XXXXXXXXXXXX',
        '............',
        '.XX..XX..XX.',
        'XXXXXXXXXXXX',
        '............'
      ]
    },
    house: {
      colors: { X: '#E36B5A', Y: '#F9F1D9', Z: '#8A5A2B' },
      rows: [
        '............',
        '.....XX.....',
        '....XXXX....',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '.XXXXXXXXXX.',
        '..YYYYYYYY..',
        '..YYYZZYYY..',
        '..YYYZZYYY..',
        '..YYYYYYYY..',
        '............',
        '............'
      ]
    },
    sunflower: {
      colors: { X: '#FFD84D', Y: '#8A5A2B', Z: '#5EA84C' },
      rows: [
        '............',
        '...XX..XX...',
        '..XXXXXXXX..',
        '..XXYYYYXX..',
        '..XXYYYYXX..',
        '..XXXXXXXX..',
        '...XX..XX...',
        '.....ZZ.....',
        '.....ZZ.....',
        '....ZZZZ....',
        '............',
        '............'
      ]
    },
    castle: {
      colors: { X: '#4A6B8A', Y: '#F9F1D9', Z: '#8A5A2B' },
      rows: [
        '............',
        '.....XX.....',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '....YYYY....',
        '..XXXXXXXX..',
        '.XXXXXXXXXX.',
        '..YYYYYYYY..',
        '..YYYZZYYY..',
        '..YYYZZYYY..',
        '..YYYYYYYY..',
        '............'
      ]
    },
    cactus: {
      colors: { X: '#4C9E58', Z: '#F0DFA8' },
      rows: [
        '............',
        '.....XX.....',
        '..X..XX..X..',
        '..XXXXXXXX..',
        '.....XX.....',
        '.....XX.....',
        '.....XX.....',
        '....XXXX....',
        '..ZZZZZZZZ..',
        '............',
        '............',
        '............'
      ]
    },
    shell: {
      colors: { X: '#FFC9C0', Y: '#E3908A' },
      rows: [
        '............',
        '............',
        '.....XX.....',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '.XXYXXXXYXX.',
        '.XXXXXXXXXX.',
        '.XYXXXXXXYX.',
        '.XXXXXXXXXX.',
        '............',
        '............',
        '............'
      ]
    },
    bluebell: {
      colors: { X: '#A8CFF0', Y: '#FFD84D', Z: '#5EA84C' },
      rows: [
        '............',
        '...XX..XX...',
        '..XXXXXXXX..',
        '..XXXYYXXX..',
        '..XXXYYXXX..',
        '..XXXXXXXX..',
        '...XX..XX...',
        '.....ZZ.....',
        '.....ZZ.....',
        '....ZZZZ....',
        '............',
        '............'
      ]
    },
    star: {
      colors: { X: '#FFD84D' },
      rows: [
        '............',
        '.....XX.....',
        '.....XX.....',
        '.XXXXXXXXXX.',
        '..XXXXXXXX..',
        '...XXXXXX...',
        '...XXXXXX...',
        '..XXX..XXX..',
        '..XX....XX..',
        '............',
        '............',
        '............'
      ]
    },
    fog: {
      colors: { X: '#D3E3F0' },
      rows: [
        '............',
        '............',
        '............',
        '....XXXX....',
        '..XXXXXXXX..',
        '.XXXXXXXXXX.',
        '.XXXXXXXXXX.',
        '..XXXXXXXX..',
        '............',
        '............',
        '............',
        '............'
      ]
    },

    // ---- ナビ・メニュー ----
    book: {
      colors: { X: '#E8834A', Y: '#FFF7E8' },
      rows: [
        '............',
        '............',
        '............',
        '.XXXXXXXXXX.',
        'XYYYYXXYYYYX',
        'XYYYYXXYYYYX',
        'XYYYYXXYYYYX',
        'XYYYYXXYYYYX',
        '.XXXXXXXXXX.',
        '............',
        '............',
        '............'
      ]
    },
    books: {
      colors: { X: '#E8556B', Y: '#4FA3E3', Z: '#5EA84C', W: '#FFF7E8' },
      rows: [
        '............',
        '............',
        '.XXXXXXXXXX.',
        '.XWXXXXXXXX.',
        '.XXXXXXXXXX.',
        '.YYYYYYYYYY.',
        '.YWYYYYYYYY.',
        '.YYYYYYYYYY.',
        '.ZZZZZZZZZZ.',
        '.ZWZZZZZZZZ.',
        '.ZZZZZZZZZZ.',
        '............'
      ]
    },
    gift: {
      colors: { X: '#E8556B', Z: '#FFB3C0' },
      rows: [
        '............',
        '............',
        '..X......X..',
        '..XX....XX..',
        'XXXXXXXXXXXX',
        '.ZZZZXXZZZZ.',
        '.ZZZZXXZZZZ.',
        '.ZZZZXXZZZZ.',
        '.ZZZZXXZZZZ.',
        '.ZZZZXXZZZZ.',
        '............',
        '............'
      ]
    },
    face: {
      colors: { X: '#F9F1D9', Y: '#3A2A1E', G: '#7CC15C', K: '#6B4A2F' },
      rows: [
        '............',
        '...G....G...',
        '....GGGG....',
        '...KKKKKK...',
        '..KXXXXXXK..',
        '.KXXXXXXXXK.',
        '.KXYXXXXYXK.',
        '.KXXXXXXXXK.',
        '..KXXXXXXK..',
        '...KKKKKK...',
        '............',
        '............'
      ]
    },
    trophy: {
      colors: { X: '#FFC93C' },
      rows: [
        '............',
        '.XXXXXXXXXX.',
        'X.XXXXXXXX.X',
        'X.XXXXXXXX.X',
        '..XXXXXXXX..',
        '...XXXXXX...',
        '....XXXX....',
        '.....XX.....',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '............',
        '............'
      ]
    },
    family: {
      colors: { X: '#8A9BB8' },
      rows: [
        '............',
        '............',
        '..XX....XX..',
        '..XX....XX..',
        '.XXXX..XXXX.',
        '.XXXX..XXXX.',
        '.XXXX..XXXX.',
        '.XXXX..XXXX.',
        '............',
        '............',
        '............',
        '............'
      ]
    },
    review: {
      colors: { X: '#8FCDE8', Y: '#FFFFFF', K: '#3C6E8A' },
      rows: [
        '............',
        '............',
        '...KKKKKK...',
        '..KXXXXXXK..',
        '.KXXXXXXXXK.',
        '.KXYXXXXYXK.',
        '.KXXXXXXXXK.',
        '.KXXXXXXXXK.',
        '..KXXXXXXK..',
        '...KKKKKK...',
        '............',
        '............'
      ]
    },

    // ---- ステータスバー ----
    coin: {
      colors: { X: '#FFC93C', Y: '#E8A317' },
      rows: [
        '............',
        '............',
        '............',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '.XXXYYYYXXX.',
        '.XXXYYYYXXX.',
        '..XXXXXXXX..',
        '...XXXXXX...',
        '............',
        '............',
        '............'
      ]
    },
    ticket: {
      colors: { X: '#FFD84D', Y: '#B8860B' },
      rows: [
        '............',
        '............',
        '............',
        '.XXXXXXXXXX.',
        '.XXXXXXXXXX.',
        '.XXXYYYXXXX.',
        '.XXXXXXXXXX.',
        '.XXXXXXXXXX.',
        '............',
        '............',
        '............',
        '............'
      ]
    },
    fire: {
      colors: { X: '#FF7043', Y: '#FFD84D' },
      rows: [
        '............',
        '............',
        '.....XX.....',
        '....XXXX....',
        '...XXXXXX...',
        '..XXXXXXXX..',
        '..XXXYYXXX..',
        '..XXXYYXXX..',
        '...XXXXXX...',
        '............',
        '............',
        '............'
      ]
    }
  };

  function render(name, opts = {}) {
    const icon = G[name];
    if (!icon) return '';
    const size = opts.size || 24;
    const pal = icon.colors;
    let out = '';
    for (let y = 0; y < icon.rows.length; y++) {
      const row = icon.rows[y];
      let x = 0;
      while (x < row.length) {
        const ch = row[x];
        let run = 1;
        while (x + run < row.length && row[x + run] === ch) run++;
        const color = pal[ch];
        if (color) out += `<rect x="${x}" y="${y}" width="${run}" height="1" fill="${color}"/>`;
        x += run;
      }
    }
    const cls = 'pxIcon' + (opts.className ? ' ' + opts.className : '');
    return `<svg class="${cls}" width="${size}" height="${size}" viewBox="0 0 ${SIZE} ${SIZE}"
      shape-rendering="crispEdges" role="img" aria-label="${opts.label || ''}">${out}</svg>`;
  }

  function has(name) { return Object.prototype.hasOwnProperty.call(G, name); }

  return { render, has, names: Object.keys(G) };
})();
