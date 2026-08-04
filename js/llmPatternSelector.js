// パターン自動選択を、ルールベースではなく実際のLLM（社内の隔離LLM環境など）に
// 判断させるためのモジュール。エンドポイント／APIキー／モデル名は画面の
// 「⚙ AI設定」から入力し、ブラウザのlocalStorageにのみ保存する（サーバー送信なし）。
//
// 対応リクエスト形式:
//   - 'openai'    : OpenAI互換 Chat Completions（Azure OpenAI, LiteLLM, vLLM等の
//                    社内ゲートウェイで広く使われている形式）
//   - 'anthropic' : Anthropic Messages API互換
//
// 社内LLMのAPI仕様がこのどちらとも異なる場合は、callOpenAiCompatible /
// callAnthropicCompatible を参考に、このファイルに新しいformatを追加してください。
window.DocAssist = window.DocAssist || {};

(function () {
  const STORAGE_KEY = 'docAssist.llmConfig.v1';

  function defaultConfig() {
    return { format: 'openai', endpoint: '', apiKey: '', model: '' };
  }

  function loadLlmConfig() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return defaultConfig();
      return Object.assign(defaultConfig(), JSON.parse(raw));
    } catch (e) {
      return defaultConfig();
    }
  }

  function saveLlmConfig(cfg) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
  }

  function buildPrompt(section, patterns) {
    const list = patterns.map((p) => `- ${p.id}: ${p.name} — ${p.description}`).join('\n');
    return [
      'あなたはコンサルティング資料のデザイン担当です。',
      '以下のスライド内容に最もふさわしいデザインパターンを、候補一覧から1つだけ選んでください。',
      '',
      '# 候補パターン一覧',
      list,
      '',
      '# スライド内容',
      `見出し: ${section.heading || ''}`,
      '箇条書き:',
      (section.bullets || []).map((b) => `- ${b}`).join('\n') || '(なし)',
      '',
      '# 出力形式',
      '説明文やコードブロックなしで、次のJSON形式のみを出力してください。',
      '{"patternId": "候補一覧の中から選んだid", "reason": "選んだ理由を日本語で一文"}',
    ].join('\n');
  }

  function extractJson(text) {
    const m = String(text).match(/\{[\s\S]*\}/);
    if (!m) throw new Error('LLMの応答からJSONを取り出せませんでした: ' + String(text).slice(0, 200));
    return JSON.parse(m[0]);
  }

  async function safeText(res) {
    try {
      return (await res.text()).slice(0, 300);
    } catch (e) {
      return '';
    }
  }

  async function callOpenAiCompatible(cfg, prompt) {
    const res = await fetch(cfg.endpoint, {
      method: 'POST',
      headers: Object.assign(
        { 'Content-Type': 'application/json' },
        cfg.apiKey ? { Authorization: `Bearer ${cfg.apiKey}` } : {}
      ),
      body: JSON.stringify({
        model: cfg.model,
        temperature: 0,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    if (!res.ok) throw new Error(`LLM API エラー (HTTP ${res.status}): ${await safeText(res)}`);
    const data = await res.json();
    const text = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
    if (!text) throw new Error('LLM応答の形式が想定と異なります（choices[0].message.content が見つかりません）');
    return text;
  }

  async function callAnthropicCompatible(cfg, prompt) {
    const res = await fetch(cfg.endpoint, {
      method: 'POST',
      headers: Object.assign(
        { 'Content-Type': 'application/json', 'anthropic-version': '2023-06-01' },
        cfg.apiKey ? { 'x-api-key': cfg.apiKey } : {}
      ),
      body: JSON.stringify({
        model: cfg.model,
        max_tokens: 300,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    if (!res.ok) throw new Error(`LLM API エラー (HTTP ${res.status}): ${await safeText(res)}`);
    const data = await res.json();
    const text = data && data.content && data.content[0] && data.content[0].text;
    if (!text) throw new Error('LLM応答の形式が想定と異なります（content[0].text が見つかりません）');
    return text;
  }

  // section: {heading, bullets} / patterns省略時はDocAssist.patterns / config省略時は保存済み設定を使用
  // 戻り値: Promise<{ pattern, reason, raw }>
  async function selectPatternWithLLM(section, patterns, config) {
    const cfg = config || loadLlmConfig();
    if (!cfg.endpoint) throw new Error('LLMエンドポイントが設定されていません（右上の「⚙ AI設定」から入力してください）');
    const list = patterns || DocAssist.patterns;
    const prompt = buildPrompt(section, list);
    const caller = cfg.format === 'anthropic' ? callAnthropicCompatible : callOpenAiCompatible;
    const rawText = await caller(cfg, prompt);
    const json = extractJson(rawText);
    const pattern = DocAssist.patternById[json.patternId];
    if (!pattern) throw new Error(`LLMが未知のpatternId "${json.patternId}" を返しました`);
    return { pattern, reason: json.reason || '', raw: json };
  }

  DocAssist.loadLlmConfig = loadLlmConfig;
  DocAssist.saveLlmConfig = saveLlmConfig;
  DocAssist.selectPatternWithLLM = selectPatternWithLLM;
})();
