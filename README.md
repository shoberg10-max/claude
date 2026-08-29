# ３年生 さんすう・かんじ ドリル

小学校３年生向けの、算数と漢字を練習できるブラウザアプリです。ビルド不要の素の HTML / CSS / JavaScript で作られているので、`index.html` を開くだけで動きます。

> 漢検9級・8級向けの育成ゲーム版は [`yurumoji-jima/`](./yurumoji-jima/) にあります（[くわしい説明](./yurumoji-jima/README.md)）。

## あそびかた

```
python3 -m http.server 8000
```

を実行して、ブラウザで `http://localhost:8000` を開いてください（`index.html` を直接ダブルクリックしても動作します）。

## さんすう（6種類・各10問）

- たし算・ひき算（2〜4けた）
- かけ算（九九〜3けた×1けた）
- わり算（あまりなし）
- あまりのあるわり算
- 大きい数のくらべっこ（＜・＞・＝）
- 大きい数の位（一の位〜十万の位）

## かんじ

小学3年生で習う漢字200字すべてを収録（[davidluzgouveia/kanji-data](https://github.com/davidluzgouveia/kanji-data) の教育漢字データを使用）。

- よみかたクイズ：漢字を見て正しい読みを4択で選ぶ
- かんじさがしクイズ：読みを見て正しい漢字を4択で選ぶ
- かんじ一覧：200字を音読み・訓読みつきで一覧表示

まちがえた漢字は次に出やすく、3回連続で正解すると「おぼえたかんじ」としてカウントされます。

## 記録

正解数・ベストスコア・おぼえた漢字はブラウザの localStorage に保存されます（サーバーへの送信なし）。

## ファイル構成

```
index.html
css/style.css
js/app.js        画面遷移とUI
js/math.js       算数の問題生成
js/kanji.js      漢字クイズの問題生成
js/storage.js    localStorage 管理
js/data/kanji3.js  3年生配当漢字200字のデータ
```
