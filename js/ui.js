// 画面共通のヘルパー（app.js / island.js の両方から使う）
const UI = (() => {
  const screenEl = document.getElementById('screen');
  const titleEl = document.getElementById('pageTitle');
  const statBarEl = document.getElementById('statBar');
  const homeBtn = document.getElementById('homeBtn');

  function el(tag, opts = {}, children = []) {
    const e = document.createElement(tag);
    if (opts.className) e.className = opts.className;
    if (opts.text !== undefined) e.textContent = opts.text;
    if (opts.html !== undefined) e.innerHTML = opts.html;
    if (opts.attrs) Object.entries(opts.attrs).forEach(([k, v]) => e.setAttribute(k, v));
    if (opts.onClick) e.addEventListener('click', opts.onClick);
    if (opts.disabled) e.disabled = true;
    children.forEach(c => c && e.appendChild(c));
    return e;
  }

  function setTitle(t) { titleEl.textContent = t; }
  function clearScreen() { screenEl.innerHTML = ''; }
  function append(node) { screenEl.appendChild(node); }

  function setStatBar(node) {
    statBarEl.innerHTML = '';
    if (node) statBarEl.appendChild(node);
  }

  function setHomeHandler(fn) {
    homeBtn.onclick = fn;
  }

  return { el, setTitle, clearScreen, append, setStatBar, setHomeHandler, screenEl };
})();
