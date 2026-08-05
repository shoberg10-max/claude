// ユーザーが作った自作テンプレート（js/templateSpec.js の仕様）の保存・読み込み・共有を担う。
//
// - 保存先はブラウザの localStorage。サーバーもAPIも使わないので、社内の隔離環境でも動く。
// - 起動時に保存済みテンプレートを DocAssist.patterns に登録するため、組み込みパターンと
//   まったく同じようにギャラリー・自動選択・PPTX書き出しから使える。
// - JSONファイルとして書き出し／読み込みできるので、作った型を同僚に配ったり、
//   バックアップしたりできる。
window.DocAssist = window.DocAssist || {};

(function () {
  const STORAGE_KEY = 'docassist.customTemplates.v1';
  const FILE_KIND = 'docassist-templates';

  // localStorage が使えない環境（プライベートモード等）でも動くよう、
  // 失敗したらメモリ上の配列にフォールバックする。
  let memoryFallback = null;

  function readStore() {
    if (memoryFallback) return memoryFallback;
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.warn('自作テンプレートの読み込みに失敗しました。メモリ上のみで動作します。', e);
      memoryFallback = [];
      return memoryFallback;
    }
  }

  function writeStore(list) {
    if (memoryFallback) {
      memoryFallback = list;
      return true;
    }
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
      return true;
    } catch (e) {
      console.warn('自作テンプレートの保存に失敗しました。メモリ上のみで保持します。', e);
      memoryFallback = list;
      return false;
    }
  }

  function listSpecs() {
    return readStore();
  }

  function uniqueId(base, takenIds) {
    const slug = String(base || 'custom')
      .toLowerCase()
      .replace(/[^a-z0-9-]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 24);
    let id = 'custom-' + (slug || 'template');
    let n = 2;
    while (takenIds.indexOf(id) !== -1) {
      id = 'custom-' + (slug || 'template') + '-' + n;
      n++;
    }
    return id;
  }

  // パターン配列への反映。更新時は同じ位置を差し替え、新規は末尾に足す。
  function registerPattern(spec) {
    const pattern = DocAssist.templateSpec.createPatternFromSpec(spec);
    const idx = DocAssist.patterns.findIndex((p) => p.id === pattern.id);
    if (idx === -1) DocAssist.patterns.push(pattern);
    else DocAssist.patterns[idx] = pattern;
    DocAssist.patternById[pattern.id] = pattern;
    return pattern;
  }

  function unregisterPattern(id) {
    const idx = DocAssist.patterns.findIndex((p) => p.id === id);
    if (idx !== -1) DocAssist.patterns.splice(idx, 1);
    delete DocAssist.patternById[id];
  }

  // 新規保存／上書き保存。id が未設定なら名前から一意なIDを作る。
  function saveSpec(rawSpec) {
    const list = readStore();
    const spec = DocAssist.templateSpec.normalizeSpec(rawSpec);
    if (!rawSpec.id) {
      const taken = DocAssist.patterns.map((p) => p.id).concat(list.map((s) => s.id));
      spec.id = uniqueId(spec.name, taken);
    }
    const idx = list.findIndex((s) => s.id === spec.id);
    if (idx === -1) list.push(spec);
    else list[idx] = spec;
    writeStore(list);
    registerPattern(spec);
    return spec;
  }

  function deleteSpec(id) {
    const list = readStore().filter((s) => s.id !== id);
    writeStore(list);
    unregisterPattern(id);
  }

  function getSpec(id) {
    return readStore().find((s) => s.id === id) || null;
  }

  // ---- 共有用のJSON入出力 ----
  function exportJson() {
    return JSON.stringify(
      {
        kind: FILE_KIND,
        version: DocAssist.templateSpec.SPEC_VERSION,
        exportedAt: new Date().toISOString(),
        templates: readStore(),
      },
      null,
      2
    );
  }

  // 取り込みは「同じIDがあれば上書き、無ければ追加」。
  // 取り込んだ件数と、無視した件数を返して呼び出し側で知らせられるようにする。
  function importJson(text) {
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error('JSONとして読み取れませんでした。ファイルの内容を確認してください。');
    }
    // 単体の仕様、配列、書き出しファイルのいずれの形でも受け付ける
    let incoming;
    if (Array.isArray(parsed)) incoming = parsed;
    else if (parsed && Array.isArray(parsed.templates)) incoming = parsed.templates;
    else if (parsed && parsed.layout) incoming = [parsed];
    else throw new Error('テンプレートが含まれていません。書き出したJSONファイルを選んでください。');

    const list = readStore();
    const builtinIds = DocAssist.patterns.filter((p) => !p.isCustom).map((p) => p.id);
    let added = 0;
    let updated = 0;
    let skipped = 0;

    incoming.forEach((raw) => {
      if (!raw || !raw.layout) {
        skipped++;
        return;
      }
      const spec = DocAssist.templateSpec.normalizeSpec(raw);
      // 組み込みパターンと同じIDは上書きさせない（組み込みを壊さないための保護）
      if (builtinIds.indexOf(spec.id) !== -1) {
        spec.id = uniqueId(spec.name, DocAssist.patterns.map((p) => p.id).concat(list.map((s) => s.id)));
      }
      const idx = list.findIndex((s) => s.id === spec.id);
      if (idx === -1) {
        list.push(spec);
        added++;
      } else {
        list[idx] = spec;
        updated++;
      }
      registerPattern(spec);
    });

    writeStore(list);
    return { added, updated, skipped };
  }

  // 起動時に保存済みテンプレートを登録する。
  // 壊れた仕様が1件あっても他が読み込めなくならないよう、1件ずつ隔離して処理する。
  function registerAll() {
    let ok = 0;
    readStore().forEach((spec) => {
      try {
        registerPattern(spec);
        ok++;
      } catch (e) {
        console.error('自作テンプレートの登録に失敗しました:', spec && spec.id, e);
      }
    });
    return ok;
  }

  DocAssist.customTemplates = {
    listSpecs,
    getSpec,
    saveSpec,
    deleteSpec,
    exportJson,
    importJson,
    registerAll,
  };

  registerAll();
})();
