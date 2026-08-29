#!/usr/bin/env node
// index.html が参照する <link rel="stylesheet"> と <script src> を、
// それぞれの中身でその場に埋め込み、外部ファイルへの参照が一切ない
// 単体の .html を作る。
//
// iOS の Safari は file:// で開いた HTML から同じフォルダの他ファイル
// （js/css）を読み込めないことが多く、複数ファイル構成のままだと
// 「開いても真っ白」になりがち。1ファイルに固めれば、Files アプリで
// zip を展開してその .html をタップ→「Safariで開く」するだけで動く。
//
// 使い方: node build/bundle.js
//   → dist/index.html を生成する（ビルド成果物。git管理しない）

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC_HTML = path.join(ROOT, 'index.html');
const OUT_DIR = path.join(ROOT, 'dist');
const OUT_HTML = path.join(OUT_DIR, 'index.html');

function readRel(relPath) {
  return fs.readFileSync(path.join(ROOT, relPath), 'utf8');
}

let html = fs.readFileSync(SRC_HTML, 'utf8');

// <link rel="stylesheet" href="foo.css"> → <style>...</style>
html = html.replace(
  /<link\s+rel="stylesheet"\s+href="([^"]+)">/g,
  (match, href) => `<style>\n${readRel(href)}</style>`
);

// <script src="foo.js"></script> → <script>...</script>
html = html.replace(
  /<script\s+src="([^"]+)"><\/script>/g,
  (match, src) => `<script>\n${readRel(src)}</script>`
);

// 埋め込み漏れがないか確認（見つかったら失敗させる）
const leftoverLink = html.match(/<link\s+rel="stylesheet"\s+href="[^"]+">/);
const leftoverScript = html.match(/<script\s+src="[^"]+"><\/script>/);
if (leftoverLink || leftoverScript) {
  console.error('埋め込みに失敗した参照が残っています:', leftoverLink || leftoverScript);
  process.exit(1);
}

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(OUT_HTML, html);

const kb = (fs.statSync(OUT_HTML).size / 1024).toFixed(0);
console.log(`書き出しました: ${path.relative(ROOT, OUT_HTML)}（${kb} KB, 外部ファイル参照なし）`);
