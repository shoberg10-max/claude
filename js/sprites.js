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
  // 1マス＝1ドットのグリッドを文字で持ち、同じデータをパレット違いで描く。
  //   K=輪郭  B=体の塗り  E=目  M=口  W=ハイライト  G=葉  P=ほっぺ  .=透明
  // カラー版と1bit LCD版の違いはパレットだけで、グリッドは共通。
  // 口はカラーでは濃い線、1bitでは「穴」にしたいので B/K とは別の記号にしている。
  const PALETTES = {
    color: { K: OUTLINE, B: CREAM, E: EYE, M: OUTLINE, W: '#FFFFFF', G: '#7CC15C', P: BLUSH },
    // 1bit：体を1色で塗りつぶし、目と口はパネルを透かす穴にする
    lcd: { K: '#2E3A22', B: '#2E3A22', E: null, M: null, W: null, G: '#2E3A22', P: '#2E3A22' }
  };
  const LCD_BG = '#A8BC84';

  const PIXEL_CHARS = {
    mojimaru: {
      grid: [
        '......GG....GG......',
        '.....GGG.GG.GGG.....',
        '.........GG.........',
        '.......KKKKKK.......',
        '.....KKBBBBBBKK.....',
        '....KBBBBBBBBBBK....',
        '...KBBBBBBBBBBBBK...',
        '..KBBBBBBBBBBBBBBK..',
        '..KBBBWEEBBWEEBBBK..',
        '..KBBBEEEBBEEEBBBK..',
        '..KBBBEEEBBEEEBBBK..',
        '..KBPPBBBBBBBBPPBK..',
        '...KBBBBBMMBBBBBK...',
        '..KKKBBBBBBBBBBKKK..',
        '.....KKBBBBBBKK.....',
        '.......KKKKKK.......',
        '.....KKK....KKK.....',
        '.....KKK....KKK.....'
      ],
      // 表情は目（行8-10）と口（行12）の差し替えだけ
      faces: {
        normal: {},
        // 閉じた目は横一直線だと寝顔に見えるので ^ ^ の形にする
        happy: {
          8: '..KBBBBEBBBBEBBBBK..',
          9: '..KBBBEBEBBEBEBBBK..',
          10: '..KBBBBBBBBBBBBBBK..',
          12: '...KBBBBMMMMBBBBK...'
        },
        surprised: {
          8: '..KBBEEEEBBEEEEBBK..',
          9: '..KBBEEEEBBEEEEBBK..',
          10: '..KBBEEEEBBEEEEBBK..'
        },
        sad: {
          10: '..KBBBBBBBBBBBBBBK..',
          12: '...KBBBBMBBMBBBBK...'
        },
        wink: {
          8: '..KBBBBEBBBWEEBBBK..',
          9: '..KBBBEBEBBEEEBBBK..',
          10: '..KBBBBBBBBEEEBBBK..',
          12: '...KBBBBMMMMBBBBK...'
        }
      },
      hatAnchor: { x: 9.5, y: 3 }
    }
  };

  function pixelGrid(char, faceName) {
    const override = char.faces[faceName] || char.faces.normal;
    return char.grid.map((row, y) => (override[y] !== undefined ? override[y] : row));
  }

  function renderPixel(id, opts = {}) {
    const char = PIXEL_CHARS[id];
    if (!char) return '';
    const grid = pixelGrid(char, opts.face);
    const pal = PALETTES[opts.palette] || PALETTES.color;
    const w = grid[0].length;
    const h = grid.length;
    const pad = opts.lcd ? 2 : 0;
    const size = opts.size || 96;

    let out = '';
    if (opts.lcd) {
      out += `<rect x="${-pad}" y="${-pad}" width="${w + pad * 2}" height="${h + pad * 2}" rx="1.5" fill="${LCD_BG}"/>`;
    }
    // 横に連続する同色ドットは1つの rect にまとめる
    for (let y = 0; y < h; y++) {
      const row = grid[y];
      let x = 0;
      while (x < w) {
        const ch = row[x];
        let run = 1;
        while (x + run < w && row[x + run] === ch) run++;
        const color = pal[ch];
        if (color) out += `<rect x="${x}" y="${y}" width="${run}" height="1" fill="${color}"/>`;
        x += run;
      }
    }
    if (opts.hat) {
      const a = char.hatAnchor;
      out += `<text x="${a.x}" y="${a.y}" font-size="7" text-anchor="middle" dominant-baseline="central">${opts.hat}</text>`;
    }

    const cls = 'charSprite charSpritePixel' + (opts.className ? ' ' + opts.className : '');
    return `<svg class="${cls}" width="${size}" height="${size}"
      viewBox="${-pad} ${-pad} ${w + pad * 2} ${h + pad * 2}"
      shape-rendering="crispEdges" role="img" aria-label="${opts.label || ''}">
      <g class="charSpriteInner">${out}</g>
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

  return { render: renderAuto, renderVector: render, renderPixel, setStyle, has };
})();
