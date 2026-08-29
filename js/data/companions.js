// ゆるもじ島のなかまと、たからばこアイテムの定義
const COMPANIONS = [
  {
    id: 'mojimaru',
    emoji: '🌱',
    sprite: 'mojimaru',
    name: 'もじまる',
    desc: '主人公。好奇心いっぱいのたんけん隊長。漢字が大好き！',
    always: true
  },
  {
    id: 'hanapyon',
    emoji: '🌸',
    name: 'はなぴょん',
    desc: 'やさしくておしゃれな女の子。花のようにやさしく教えてくれるよ。',
    need: 5,
    hint: 'かんじを 5こ おぼえると なかまになるよ'
  },
  {
    id: 'pentan',
    emoji: '🐧',
    name: 'ぺんたん',
    desc: 'まじめでがんばりやさん。ノートとえんぴつが手放せない。',
    need: 20,
    hint: 'かんじを 20こ おぼえると なかまになるよ'
  },
  {
    id: 'koganecchi',
    emoji: '⭐',
    name: 'こがねっち',
    desc: 'キラキラが大好きなひよこ。宝がいっぱいあると、レアなアイテムをくれるよ！',
    need: 50,
    hint: 'かんじを 50こ おぼえると なかまになるよ'
  },
  {
    id: 'nekojiro',
    emoji: '🐱',
    name: 'ねこじろう',
    desc: 'ちょっとミステリアスなねこ。クイズや なぞなぞが得意。',
    need: 90,
    hint: 'かんじを 90こ おぼえると なかまになるよ'
  }
];

// たからばこから出るアイテム（かぶりもの）
const TREASURE_ITEMS = [
  { id: 'hat_dongurl', emoji: '🌰', name: 'どんぐり帽', rarity: 'common' },
  { id: 'hat_leaf', emoji: '🍃', name: 'はっぱ帽', rarity: 'common' },
  { id: 'hat_ribbon', emoji: '🎀', name: 'リボン', rarity: 'common' },
  { id: 'hat_cap', emoji: '🧢', name: 'たんけん帽', rarity: 'common' },
  { id: 'hat_headband', emoji: '🎽', name: 'はちまき', rarity: 'common' },
  { id: 'hat_crown', emoji: '👑', name: 'おうかん', rarity: 'rare' },
  { id: 'hat_helmet', emoji: '🪖', name: 'ぼうけんヘルメット', rarity: 'rare' },
  { id: 'hat_wizard', emoji: '🎩', name: 'まほうのぼうし', rarity: 'rare' },
  { id: 'hat_space', emoji: '👽', name: 'うちゅうヘルメット', rarity: 'epic' },
  { id: 'hat_star', emoji: '🌟', name: 'ほしのかんむり', rarity: 'epic' }
];

function rollTreasureItem() {
  const r = Math.random();
  let pool;
  if (r < 0.08) pool = TREASURE_ITEMS.filter(i => i.rarity === 'epic');
  else if (r < 0.33) pool = TREASURE_ITEMS.filter(i => i.rarity === 'rare');
  else pool = TREASURE_ITEMS.filter(i => i.rarity === 'common');
  return pool[Math.floor(Math.random() * pool.length)];
}

function charStageForLevel(level) {
  if (level >= 30) return { title: '漢字マスター', emoji: '👑' };
  if (level >= 20) return { title: 'すごうで探検家', emoji: '🧭' };
  if (level >= 10) return { title: 'たんけんもじまる', emoji: '🎒' };
  return { title: 'ちびもじまる', emoji: '🌱' };
}
