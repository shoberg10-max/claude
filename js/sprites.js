// キャラクターのSVGスプライト
// 体・顔・かぶりものを別レイヤーで組み立てる。
// 顔はローカル座標（原点＝顔の中心、目は x=±9）で定義し、
// キャラごとの faceAnchor に translate して載せるので、全キャラで使い回せる。
const Sprites = (() => {
  // ---- 全キャラ共通の統一ルール ----
  const OUTLINE = '#6B4A2F';   // 輪郭は黒ではなく濃い茶
  const OW = 3;                // 輪郭の太さ（全キャラ固定）
  const CREAM = '#F9F1D9';
  const BLUSH = '#FFA9BE';
  const EYE = '#3A2A1E';
  const MOUTH = '#C4576B';

  const line = `fill="none" stroke="${OUTLINE}" stroke-linecap="round"`;
  const solid = `fill="${CREAM}" stroke="${OUTLINE}" stroke-width="${OW}"`;

  // 手足：太い丸線を輪郭色→クリームの順に重ねてカプセル状にする。
  // 体より先に描くので、付け根は体に隠れる。
  function limb(x1, y1, x2, y2, w = 8) {
    const d = `M${x1} ${y1} L${x2} ${y2}`;
    return `<path d="${d}" fill="none" stroke="${OUTLINE}" stroke-width="${w + OW * 2}" stroke-linecap="round"/>
            <path d="${d}" fill="none" stroke="${CREAM}" stroke-width="${w}" stroke-linecap="round"/>`;
  }

  function blush() {
    return `<ellipse cx="-19" cy="7" rx="5.5" ry="3.4" fill="${BLUSH}" opacity=".7"/>
            <ellipse cx="19" cy="7" rx="5.5" ry="3.4" fill="${BLUSH}" opacity=".7"/>`;
  }

  function eyeDot(cx, r) {
    return `<circle cx="${cx}" cy="0" r="${r}" fill="${EYE}"/>
            <circle cx="${cx + 1.5}" cy="-1.6" r="${r * 0.34}" fill="#fff"/>`;
  }

  // ---- 表情差分（顔だけ差し替える） ----
  const FACES = {
    normal: () => `${blush()}
      ${eyeDot(-9, 4.3)}${eyeDot(9, 4.3)}
      <path d="M-4 8 Q0 12.5 4 8" ${line} stroke-width="2.4"/>`,

    happy: () => `${blush()}
      <path d="M-13.5 1.5 Q-9 -5 -4.5 1.5" ${line} stroke-width="3"/>
      <path d="M4.5 1.5 Q9 -5 13.5 1.5" ${line} stroke-width="3"/>
      <path d="M-6.5 6 Q0 15.5 6.5 6 Z" fill="${MOUTH}" stroke="${OUTLINE}" stroke-width="2.2" stroke-linejoin="round"/>`,

    surprised: () => `${blush()}
      ${eyeDot(-9, 5.4)}${eyeDot(9, 5.4)}
      <ellipse cx="0" cy="9.5" rx="3.4" ry="4.3" fill="${MOUTH}" stroke="${OUTLINE}" stroke-width="2.2"/>`,

    sad: () => `${blush()}
      ${eyeDot(-9, 4)}${eyeDot(9, 4)}
      <path d="M-14 -7 Q-9.5 -9.5 -5 -7.5" ${line} stroke-width="2.2"/>
      <path d="M5 -7.5 Q9.5 -9.5 14 -7" ${line} stroke-width="2.2"/>
      <path d="M-4.5 11.5 Q0 6.5 4.5 11.5" ${line} stroke-width="2.4"/>`,

    wink: () => `${blush()}
      ${eyeDot(-9, 4.3)}
      <path d="M4.5 1 Q9 -5 13.5 1" ${line} stroke-width="3"/>
      <path d="M-5 7.5 Q0 13 5 7.5 Z" fill="${MOUTH}" stroke="${OUTLINE}" stroke-width="2.2" stroke-linejoin="round"/>`
  };

  // ---- キャラクター（体のみ。顔とかぶりものは別レイヤー） ----
  const CHARS = {
    mojimaru: {
      faceAnchor: { x: 50, y: 57 },
      hatAnchor: { x: 50, y: 33 },   // 芽が上に見えるよう、帽子はおでこ寄りに載せる
      body: () => `
        <ellipse cx="50" cy="93" rx="24" ry="4.5" fill="#000" opacity=".10"/>
        <ellipse cx="39" cy="87" rx="8.5" ry="5" ${solid}/>
        <ellipse cx="61" cy="87" rx="8.5" ry="5" ${solid}/>
        ${limb(30, 70, 20, 80)}
        ${limb(70, 70, 80, 80)}
        <path d="M50 31 C50 26 50 23 50 20" fill="none" stroke="#5E9E48" stroke-width="3.4" stroke-linecap="round"/>
        <ellipse cx="50" cy="57" rx="28" ry="27" ${solid}/>
        <ellipse cx="50" cy="76" rx="19" ry="7" fill="#EADFBE" opacity=".45"/>
        <ellipse cx="40" cy="46" rx="10" ry="6.5" fill="#fff" opacity=".38"/>
        <path d="M50 21 C45 14 36 13 32 18 C36 25 45 26 50 21 Z" fill="#7CC15C" stroke="${OUTLINE}" stroke-width="2.4" stroke-linejoin="round"/>
        <path d="M50 21 C55 14 64 13 68 18 C64 25 55 26 50 21 Z" fill="#8FD06D" stroke="${OUTLINE}" stroke-width="2.4" stroke-linejoin="round"/>`
    }
  };

  // かぶりものは今のところ絵文字を装着位置に載せる（アイテム絵の実装までのつなぎ）
  function hatLayer(char, hatEmoji) {
    if (!hatEmoji) return '';
    const a = char.hatAnchor;
    return `<text x="${a.x}" y="${a.y}" font-size="30" text-anchor="middle" dominant-baseline="central">${hatEmoji}</text>`;
  }

  function render(id, opts = {}) {
    const char = CHARS[id];
    if (!char) return '';
    const face = FACES[opts.face] || FACES.normal;
    const size = opts.size || 96;
    const a = char.faceAnchor;
    const cls = 'charSprite' + (opts.className ? ' ' + opts.className : '');
    return `<svg class="${cls}" width="${size}" height="${size}" viewBox="0 0 100 100" role="img" aria-label="${opts.label || ''}">
      <g class="charSpriteInner">
        ${char.body()}
        <g transform="translate(${a.x} ${a.y})">${face()}</g>
        ${hatLayer(char, opts.hat)}
      </g>
    </svg>`;
  }


  // =========================================================
  // ドット絵スタイル（昔のたまごっち風）
  // =========================================================
  // 20x18 のドットを文字グリッドで持つ。1マス＝1ドット。
  //   K=輪郭  B=体の塗り  A=アクセント  E=目  M=口  W=ハイライト
  //   P=ほっぺ  G=葉  O=くちばし・足  C/D/J=かぶりもの用  .=透明
  //
  // 体・顔・かぶりものは別レイヤーで、この順に重ねる。
  // 顔は全キャラ共通の1セットだけ持ち、どのキャラも頭の位置と大きさを
  // 揃えてあるので、そのまま載せられる（キャラごとの表情差分は不要）。
  const W_ = 20;
  const H_ = 18;

  // 色は記号→色の対応表。キャラ側の colors で上書きできるので、
  // 「1キャラ＝主役の色1つ＋共通の輪郭色」という規格を保てる。
  const BASE_PALETTE = {
    K: OUTLINE, B: CREAM, A: '#EADFBE', E: EYE, M: OUTLINE,
    W: '#FFFFFF', P: BLUSH, G: '#7CC15C', O: '#FFA53C'
  };
  const LCD_PALETTE = {
    K: '#2E3A22', B: '#2E3A22', A: '#2E3A22', E: null, M: null,
    W: null, P: '#2E3A22', G: '#2E3A22', O: '#2E3A22'
  };
  const LCD_BG = '#A8BC84';

  // 全キャラ共通の体。頭の丸み・腕・足の位置をここで固定する。
  const BASE_BODY = {
    3: '.......KKKKKK.......',
    4: '.....KKBBBBBBKK.....',
    5: '....KBBBBBBBBBBK....',
    6: '...KBBBBBBBBBBBBK...',
    7: '..KBBBBBBBBBBBBBBK..',
    8: '..KBBBBBBBBBBBBBBK..',
    9: '..KBBBBBBBBBBBBBBK..',
    10: '..KBBBBBBBBBBBBBBK..',
    11: '..KBBBBBBBBBBBBBBK..',
    12: '...KBBBBBBBBBBBBK...',
    13: '..KKKBBBBBBBBBBKKK..',
    14: '.....KKBBBBBBKK.....',
    15: '.......KKKKKK.......',
    16: '.....KKK....KKK.....',
    17: '.....KKK....KKK.....'
  };

  // 表情（全キャラ共通）。目は行8-10、ほっぺは行11、口は行12。
  // 閉じた目は横一直線だと寝顔に見えるので ^ ^ の形にしている。
  const PIXEL_FACES = {
    normal: {
      8: '......WEE..WEE......',
      9: '......EEE..EEE......',
      10: '......EEE..EEE......',
      11: '....PP........PP....',
      12: '.........MM.........'
    },
    happy: {
      8: '.......E....E.......',
      9: '......E.E..E.E......',
      11: '....PP........PP....',
      12: '........MMMM........'
    },
    surprised: {
      8: '.....EEEE..EEEE.....',
      9: '.....EEEE..EEEE.....',
      10: '.....EEEE..EEEE.....',
      11: '....PP........PP....',
      12: '.........MM.........'
    },
    sad: {
      8: '......WEE..WEE......',
      9: '......EEE..EEE......',
      11: '....PP........PP....',
      12: '........M..M........'
    },
    wink: {
      8: '.......E...WEE......',
      9: '......E.E..EEE......',
      10: '...........EEE......',
      11: '....PP........PP....',
      12: '........MMMM........'
    },
    // わすれんぼう用：中が抜けた目でぼんやりした顔にする
    dizzy: {
      8: '......EEE..EEE......',
      9: '......E.E..E.E......',
      10: '......EEE..EEE......',
      12: '........M..M........'
    },
    // 敵キャラ用：つり上がった眉とギザギザの口
    angry: {
      7: '.....EE......EE.....',
      8: '......EEE..EEE......',
      9: '......EEE..EEE......',
      12: '........MMMM........'
    }
  };

  const PIXEL_CHARS = {
    // 主人公：芽
    mojimaru: {
      rows: {
        0: '......GG....GG......',
        1: '.....GGG.GG.GGG.....',
        2: '.........GG.........'
      },
      colors: { B: CREAM, G: '#7CC15C' }
    },

    // うさぎ：長い耳
    hanapyon: {
      rows: {
        0: '......KKK..KKK......',
        1: '......KAK..KAK......',
        2: '......KAK..KAK......'
      },
      colors: { B: '#FFD8E2', A: '#FF9FBB' }
    },

    // ペンギン：濃い頭と背中、オレンジのくちばしと足
    pentan: {
      rows: {
        1: '.........KK.........',
        2: '.........KK.........',
        4: '.....KKAAAAAAKK.....',
        5: '....KAAAAAAAAAAK....',
        6: '...KAAAAAAAAAAAAK...',
        7: '..KAABBBBBBBBBBAAK..',
        8: '..KABBBBBBBBBBBBAK..',
        9: '..KABBBBBBBBBBBBAK..',
        10: '..KABBBBBBBBBBBBAK..',
        11: '..KABBBBBBBBBBBBAK..',
        12: '...KBBBBBBBBBBBBK...',
        13: '..KKKABBBBBBBBAKKK..',
        16: '.....OOO....OOO.....',
        17: '.....OOO....OOO.....'
      },
      colors: { B: '#F2F7FC', A: '#5E7A9E', M: '#FF9E2C' }
    },

    // ひよこ：黄色、小さな冠羽、オレンジの足
    koganecchi: {
      rows: {
        0: '........A..A........',
        1: '.........AA.........',
        2: '.........AA.........',
        16: '.....OOO....OOO.....',
        17: '.....OOO....OOO.....'
      },
      colors: { B: '#FFE082', A: '#FFC93C', M: '#FF9E2C' }
    },

    // ねこ：とがった耳
    nekojiro: {
      rows: {
        0: '.....KK......KK.....',
        1: '....KAAK....KAAK....',
        2: '...KAAAK....KAAAK...'
      },
      colors: { B: '#C9B4E8', A: '#A88BD0' }
    },

    // 敵：まちがえ鬼（角つき、ほっぺなし）
    machigaeoni: {
      rows: {
        0: '...K............K...',
        1: '...KA..........AK...',
        2: '....KA........AK....'
      },
      colors: { B: '#8C7BB8', A: '#5F5185', K: '#3B3157', P: null }
    },

    // 敵：わすれんぼう（形のないブヨブヨ、ほっぺなし）
    wasurenbo: {
      colors: { B: '#8FCDE8', A: '#5AA3CC', K: '#3C6E8A', P: null }
    }
  };

  // かぶりもの。キャラと同じ20列の座標系で、頭に重なる行だけ持つ。
  const PIXEL_HATS = {
    hat_dongurl: {
      rows: {
        1: '.........DD.........',
        2: '.......CCCCCC.......',
        3: '......CCCCCCCC......',
        4: '.....CCCCCCCCCC.....'
      },
      colors: { C: '#A9743F', D: '#7A5230' }
    },
    hat_leaf: {
      rows: {
        1: '..........CCC.......',
        2: '........CCCCCD......',
        3: '.......CCCCC........'
      },
      colors: { C: '#7CC15C', D: '#5E9E48' }
    },
    hat_ribbon: {
      rows: {
        1: '......CC....CC......',
        2: '......CCCDDCCC......',
        3: '......CC....CC......'
      },
      colors: { C: '#FF7BA0', D: '#E04E7B' }
    },
    hat_cap: {
      rows: {
        2: '.......CCCCCC.......',
        3: '......CCCCCCCC......',
        4: '.DDDDCCCCCCCCCC.....'
      },
      colors: { C: '#4FA3E3', D: '#2E7BB8' }
    },
    hat_headband: {
      rows: {
        5: '...CCCCCCCCCCCCCC...',
        6: '..DD................'
      },
      colors: { C: '#FF5A5A', D: '#D63333' }
    },
    hat_crown: {
      rows: {
        1: '......C..CC..C......',
        2: '......CCCCCCCC......',
        3: '......CCCJJCCC......'
      },
      colors: { C: '#FFC93C', J: '#FF5A5A' }
    },
    hat_helmet: {
      rows: {
        2: '.....CCCCCCCCCC.....',
        3: '....CCCCCCCCCCCC....',
        4: '....DDDDDDDDDDDD....'
      },
      colors: { C: '#6B8E4E', D: '#4C6838' }
    },
    hat_wizard: {
      rows: {
        0: '.........CC.........',
        1: '........CCCC........',
        2: '.......CCCCCC.......',
        3: '.....CCCCCCCCCC.....',
        4: '....DDDDDDDDDDDD....'
      },
      colors: { C: '#6C4CB8', D: '#4A3080' }
    },
    hat_space: {
      rows: {
        2: '......CCCCCCCC......',
        3: '....CC........CC....',
        4: '...C............C...'
      },
      colors: { C: '#9FD8F0' }
    },
    hat_star: {
      rows: {
        0: '.........CC.........',
        1: '.........CC.........',
        2: '.....CCCCCCCCCC.....',
        3: '.......CCCCCC.......',
        4: '......CC....CC......'
      },
      colors: { C: '#FFD84D' }
    }
  };

  const EMPTY_ROW = '.'.repeat(W_);

  // 体 → 顔 → かぶりもの の順に、透明でないドットだけ上書きしていく
  function overlay(grid, rows) {
    if (!rows) return grid;
    Object.keys(rows).forEach(k => {
      const y = Number(k);
      if (y < 0 || y >= H_) return;
      const src = rows[k];
      const dst = grid[y].split('');
      for (let x = 0; x < W_ && x < src.length; x++) {
        if (src[x] !== '.') dst[x] = src[x];
      }
      grid[y] = dst.join('');
    });
    return grid;
  }

  function buildGrid(char, faceName, hat) {
    const grid = [];
    for (let y = 0; y < H_; y++) grid.push(EMPTY_ROW);
    overlay(grid, BASE_BODY);
    overlay(grid, char.rows);
    if (faceName !== null) overlay(grid, PIXEL_FACES[faceName] || PIXEL_FACES.normal);
    if (hat) overlay(grid, hat.rows);
    return grid;
  }

  function paletteFor(char, hat, mode) {
    const base = mode === 'lcd' ? LCD_PALETTE : BASE_PALETTE;
    const pal = Object.assign({}, base);
    // 1bit では体を1色に潰すので、キャラ固有の色は反映しない
    if (mode !== 'lcd') {
      Object.assign(pal, char.colors || {}, (hat && hat.colors) || {});
    } else if (hat) {
      Object.keys(hat.colors || {}).forEach(k => { pal[k] = LCD_PALETTE.K; });
    }
    return pal;
  }

  // 横に連続する同色ドットは1つの rect にまとめる
  function gridToRects(grid, pal) {
    let out = '';
    for (let y = 0; y < grid.length; y++) {
      const row = grid[y];
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
    return out;
  }

  function renderPixel(id, opts = {}) {
    const char = PIXEL_CHARS[id];
    if (!char) return '';
    const hat = opts.hat ? PIXEL_HATS[opts.hat] : null;
    const mode = opts.palette === 'lcd' ? 'lcd' : 'color';
    const grid = buildGrid(char, opts.face || 'normal', hat);
    const pal = paletteFor(char, hat, mode);
    const size = opts.size || 96;
    const pad = opts.lcd ? 2 : 0;

    let out = '';
    if (opts.lcd) {
      out += `<rect x="${-pad}" y="${-pad}" width="${W_ + pad * 2}" height="${H_ + pad * 2}" rx="1.5" fill="${LCD_BG}"/>`;
    }
    out += gridToRects(grid, pal);

    const cls = 'charSprite charSpritePixel' + (opts.className ? ' ' + opts.className : '');
    return `<svg class="${cls}" width="${size}" height="${size}"
      viewBox="${-pad} ${-pad} ${W_ + pad * 2} ${H_ + pad * 2}"
      shape-rendering="crispEdges" role="img" aria-label="${opts.label || ''}">
      <g class="charSpriteInner">${out}</g>
    </svg>`;
  }

  // 着せ替え一覧用：かぶりもの単体を、その絵の範囲だけ切り出して描く
  function renderHatIcon(hatId, opts = {}) {
    const hat = PIXEL_HATS[hatId];
    if (!hat) return '';
    const grid = [];
    for (let y = 0; y < H_; y++) grid.push(EMPTY_ROW);
    overlay(grid, hat.rows);

    let minX = W_, maxX = -1, minY = H_, maxY = -1;
    for (let y = 0; y < H_; y++) {
      for (let x = 0; x < W_; x++) {
        if (grid[y][x] !== '.') {
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }
    if (maxX < 0) return '';
    const pal = Object.assign({}, BASE_PALETTE, hat.colors);
    const size = opts.size || 48;
    return `<svg class="charSprite charSpritePixel hatIcon" width="${size}" height="${size}"
      viewBox="${minX - 0.5} ${minY - 0.5} ${maxX - minX + 2} ${maxY - minY + 2}"
      shape-rendering="crispEdges" role="img" aria-label="${opts.label || ''}">
      ${gridToRects(grid, pal)}
    </svg>`;
  }

  // 'vector' か 'pixel' を切りかえる。既存の呼び出し側は render() のまま。
  let style = 'pixel';
  function setStyle(s) { style = s; }

  function renderAuto(id, opts = {}) {
    return style === 'pixel' ? renderPixel(id, opts) : render(id, opts);
  }

  function has(id) {
    return style === 'pixel'
      ? Object.prototype.hasOwnProperty.call(PIXEL_CHARS, id)
      : Object.prototype.hasOwnProperty.call(CHARS, id);
  }

  function hasHat(hatId) { return Object.prototype.hasOwnProperty.call(PIXEL_HATS, hatId); }

  return {
    render: renderAuto, renderVector: render, renderPixel, renderHatIcon,
    setStyle, has, hasHat
  };
})();
