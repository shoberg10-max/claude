// ゆるもじ島 漢字たんけん隊 - 漢字データ
// 9級＝小学2年配当・8級＝小学3年配当 から、あそびやすい代表24字ずつを収録（プロトタイプ版）
const KANJI_DATA = [
  // ===== 9級（漢検9級 相当） =====
  { k:"引", level:9, on:["イン"], kun:["ひく"], reading:"ひく", meaning:"ひっぱる", word:"引力", wordReading:"いんりょく" },
  { k:"羽", level:9, on:["ウ"], kun:["はね"], reading:"はね", meaning:"とりのはね", word:"羽毛", wordReading:"うもう" },
  { k:"雲", level:9, on:["ウン"], kun:["くも"], reading:"くも", meaning:"そらのくも", word:"雲海", wordReading:"うんかい" },
  { k:"園", level:9, on:["エン"], kun:["その"], reading:"その", meaning:"にわ・はたけ", word:"公園", wordReading:"こうえん" },
  { k:"遠", level:9, on:["エン"], kun:["とおい"], reading:"とおい", meaning:"きょりがある", word:"遠足", wordReading:"えんそく" },
  { k:"何", level:9, on:["カ"], kun:["なに"], reading:"なに", meaning:"たずねることば", word:"何色", wordReading:"なにいろ" },
  { k:"夏", level:9, on:["カ"], kun:["なつ"], reading:"なつ", meaning:"あつい季節", word:"夏休み", wordReading:"なつやすみ" },
  { k:"家", level:9, on:["カ"], kun:["いえ"], reading:"いえ", meaning:"すむところ", word:"家族", wordReading:"かぞく" },
  { k:"歌", level:9, on:["カ"], kun:["うた"], reading:"うた", meaning:"うたうもの", word:"歌手", wordReading:"かしゅ" },
  { k:"回", level:9, on:["カイ"], kun:["まわる"], reading:"まわる", meaning:"ぐるぐるまわる", word:"回転", wordReading:"かいてん" },
  { k:"会", level:9, on:["カイ"], kun:["あう"], reading:"あう", meaning:"ひとにあう", word:"会話", wordReading:"かいわ" },
  { k:"海", level:9, on:["カイ"], kun:["うみ"], reading:"うみ", meaning:"おおきな水", word:"海岸", wordReading:"かいがん" },
  { k:"絵", level:9, on:["エ"], kun:[], reading:"え", meaning:"かいたもの", word:"絵本", wordReading:"えほん" },
  { k:"外", level:9, on:["ガイ"], kun:["そと"], reading:"そと", meaning:"うちではないところ", word:"外国", wordReading:"がいこく" },
  { k:"角", level:9, on:["カク"], kun:["かど"], reading:"かど", meaning:"かど・つの", word:"三角", wordReading:"さんかく" },
  { k:"楽", level:9, on:["ガク"], kun:["たのしい"], reading:"たのしい", meaning:"たのしい気もち", word:"音楽", wordReading:"おんがく" },
  { k:"間", level:9, on:["カン"], kun:["あいだ"], reading:"あいだ", meaning:"ふたつのあいだ", word:"時間", wordReading:"じかん" },
  { k:"丸", level:9, on:["ガン"], kun:["まる"], reading:"まる", meaning:"まるいかたち", word:"丸太", wordReading:"まるた" },
  { k:"岩", level:9, on:["ガン"], kun:["いわ"], reading:"いわ", meaning:"おおきな石", word:"岩石", wordReading:"がんせき" },
  { k:"顔", level:9, on:["ガン"], kun:["かお"], reading:"かお", meaning:"かお", word:"顔色", wordReading:"かおいろ" },
  { k:"汽", level:9, on:["キ"], kun:[], reading:"き", meaning:"じょうき", word:"汽車", wordReading:"きしゃ" },
  { k:"京", level:9, on:["キョウ"], kun:[], reading:"きょう", meaning:"みやこ", word:"京都", wordReading:"きょうと" },
  { k:"強", level:9, on:["キョウ"], kun:["つよい"], reading:"つよい", meaning:"ちからがある", word:"強力", wordReading:"きょうりょく" },
  { k:"近", level:9, on:["キン"], kun:["ちかい"], reading:"ちかい", meaning:"きょりがみじかい", word:"近所", wordReading:"きんじょ" },

  // ===== 8級（漢検8級 相当） =====
  { k:"全", level:8, on:["ゼン"], kun:["まったく"], reading:"まったく", meaning:"すべて", word:"全部", wordReading:"ぜんぶ" },
  { k:"有", level:8, on:["ユウ"], kun:["ある"], reading:"ある", meaning:"もっている", word:"有名", wordReading:"ゆうめい" },
  { k:"住", level:8, on:["ジュウ"], kun:["すむ"], reading:"すむ", meaning:"くらす", word:"住所", wordReading:"じゅうしょ" },
  { k:"医", level:8, on:["イ"], kun:[], reading:"い", meaning:"びょうきをなおす人", word:"医者", wordReading:"いしゃ" },
  { k:"対", level:8, on:["タイ"], kun:[], reading:"たい", meaning:"むきあう", word:"反対", wordReading:"はんたい" },
  { k:"決", level:8, on:["ケツ"], kun:["きめる"], reading:"きめる", meaning:"きめる", word:"決心", wordReading:"けっしん" },
  { k:"事", level:8, on:["ジ"], kun:["こと"], reading:"こと", meaning:"できごと", word:"仕事", wordReading:"しごと" },
  { k:"使", level:8, on:["シ"], kun:["つかう"], reading:"つかう", meaning:"つかう", word:"使用", wordReading:"しよう" },
  { k:"味", level:8, on:["ミ"], kun:["あじ"], reading:"あじ", meaning:"あじ", word:"意味", wordReading:"いみ" },
  { k:"命", level:8, on:["メイ"], kun:["いのち"], reading:"いのち", meaning:"いのち", word:"命令", wordReading:"めいれい" },
  { k:"始", level:8, on:["シ"], kun:["はじめる"], reading:"はじめる", meaning:"スタートする", word:"開始", wordReading:"かいし" },
  { k:"実", level:8, on:["ジツ"], kun:["み"], reading:"み", meaning:"みのなるところ", word:"実力", wordReading:"じつりょく" },
  { k:"幸", level:8, on:["コウ"], kun:["しあわせ"], reading:"しあわせ", meaning:"しあわせ", word:"幸福", wordReading:"こうふく" },
  { k:"所", level:8, on:["ショ"], kun:["ところ"], reading:"ところ", meaning:"ばしょ", word:"場所", wordReading:"ばしょ" },
  { k:"泳", level:8, on:["エイ"], kun:["およぐ"], reading:"およぐ", meaning:"みずをおよぐ", word:"水泳", wordReading:"すいえい" },
  { k:"物", level:8, on:["ブツ"], kun:["もの"], reading:"もの", meaning:"もの", word:"動物", wordReading:"どうぶつ" },
  { k:"育", level:8, on:["イク"], kun:["そだつ"], reading:"そだつ", meaning:"そだつ", word:"教育", wordReading:"きょういく" },
  { k:"表", level:8, on:["ヒョウ"], kun:["おもて"], reading:"おもて", meaning:"あらわす・おもて", word:"発表", wordReading:"はっぴょう" },
  { k:"急", level:8, on:["キュウ"], kun:["いそぐ"], reading:"いそぐ", meaning:"いそぐ", word:"急行", wordReading:"きゅうこう" },
  { k:"指", level:8, on:["シ"], kun:["ゆび"], reading:"ゆび", meaning:"て・ゆび", word:"指名", wordReading:"しめい" },
  { k:"神", level:8, on:["シン"], kun:["かみ"], reading:"かみ", meaning:"かみさま", word:"神社", wordReading:"じんじゃ" },
  { k:"級", level:8, on:["キュウ"], kun:[], reading:"きゅう", meaning:"くらい・クラス", word:"学級", wordReading:"がっきゅう" },
  { k:"美", level:8, on:["ビ"], kun:["うつくしい"], reading:"うつくしい", meaning:"きれい", word:"美人", wordReading:"びじん" },
  { k:"動", level:8, on:["ドウ"], kun:["うごく"], reading:"うごく", meaning:"うごく", word:"運動", wordReading:"うんどう" },
];

const KANJI_BY_LEVEL = {
  9: KANJI_DATA.filter(e => e.level === 9),
  8: KANJI_DATA.filter(e => e.level === 8),
};

// ===== キャラクター =====
const CHARACTERS = [
  { id:"mojimaru", name:"もじまる", emoji:"🌱", role:"しゅじんこう。かんじが大すき。", stages:[
    { lv:1, name:"ちびもじまる", emoji:"🌱" },
    { lv:10, name:"たんけんもじまる", emoji:"🧭" },
    { lv:20, name:"すごうでたんけん家", emoji:"🎒" },
    { lv:30, name:"かんじマスター", emoji:"👑" },
  ]},
  { id:"hanapyon", name:"はなぴょん", emoji:"🌸", role:"書きとりたんとう。ほめじょうず。" },
  { id:"pentan", name:"ぺんたん", emoji:"📚", role:"かんじはかせ。よみかた・いみをおしえてくれる。" },
  { id:"koganecchi", name:"こがねっち", emoji:"⭐", role:"ほうしゅうたんとう。おたからだいすき。" },
  { id:"nekojiro", name:"ねこじろう", emoji:"🌙", role:"クイズたんとう。なぞめいている。" },
];

// ===== 敵キャラ（まちがえ演出） =====
const ENEMIES = {
  mistake: { name:"まちがえ鬼", emojiBefore:"😈", emojiAfter:"✨" },
  forget:  { name:"わすれんぼう", emojiBefore:"😈", emojiAfter:"✨" },
};

// ===== ガチャ・アイテム =====
const ITEMS = [
  { id:"hat_acorn", name:"どんぐりぼうし", emoji:"🌰", type:"ぼうし", rarity:"ノーマル" },
  { id:"hat_explorer", name:"たんけんぼうし", emoji:"🪖", type:"ぼうし", rarity:"ノーマル" },
  { id:"hat_crown", name:"こがねのおうかん", emoji:"👑", type:"ぼうし", rarity:"レア" },
  { id:"hat_space", name:"うちゅうヘルメット", emoji:"👨‍🚀", type:"ぼうし", rarity:"スーパーレア" },
  { id:"bag_leaf", name:"はっぱのバッグ", emoji:"🎒", type:"バッグ", rarity:"ノーマル" },
  { id:"shoes_boots", name:"たんけんブーツ", emoji:"🥾", type:"くつ", rarity:"ノーマル" },
  { id:"pet_bird", name:"ことりのなかま", emoji:"🐦", type:"ペット", rarity:"レア" },
  { id:"pet_fox", name:"きつねのなかま", emoji:"🦊", type:"ペット", rarity:"スーパーレア" },
];

// ===== 島の発しんステージ =====
const ISLAND_STAGES = [
  { min:0,  label:"きりに つつまれた島", tiles:["🌫️","🌫️","🌫️","🌫️","🌫️","🌫️🌱","🏠","🌫️","🌫️"] },
  { min:5,  label:"森が しげってきた", tiles:["🌳","🌳","🌫️","🌳","🌱","🏠","🌫️","🌳","🌫️"] },
  { min:15, label:"村と山が あらわれた", tiles:["🌳","⛰️","🌸","🌳","🏠","🌉","🌊","🌊","🌳"] },
  { min:30, label:"かんじ城が かんせい！", tiles:["🌳","⛰️","🏯","🌸","🏠","🌉","🌊","🌊","🌳"] },
];

// ===== 漢字たまご成長 =====
const TAMAGO_STAGES = [
  { min:0,  emoji:"🥚", label:"たまご" },
  { min:8,  emoji:"🥚✨", label:"ひびがはいった たまご" },
  { min:20, emoji:"🐣", label:"うまれたばかり" },
  { min:40, emoji:"⭐🐥", label:"ほしっち" },
];
