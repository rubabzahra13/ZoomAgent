// Runs inside the Zoom Docs tab (docs.zoom.us) to click 3-dot → Export → Word
(function() {
  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
  async function run() {
    await sleep(2000);
    const body = document.body;
    const find = (sel) => document.querySelector(sel);
    const findText = (text) => {
      const walk = document.createTreeWalker(body, NodeFilter.SHOW_ELEMENT);
      let n;
      while ((n = walk.nextNode())) {
        if (n.innerText && n.innerText.trim() === text) return n;
      }
      return null;
    };
    const buttons = document.querySelectorAll('button');
    let menuOpened = false;
    for (const btn of buttons) {
      if (btn.closest('nav')) continue;
      if (btn.querySelector('svg') && btn.offsetParent) {
        btn.click();
        await sleep(1000);
        if (document.body.innerText.includes('Export')) { menuOpened = true; break; }
      }
    }
    if (!menuOpened) return;
    const exportEl = findText('Export') || Array.from(document.querySelectorAll('span, div')).find(el => el.innerText && el.innerText.trim() === 'Export');
    if (exportEl) { exportEl.click(); await sleep(1000); }
    const wordEl = findText('Word') || Array.from(document.querySelectorAll('span, div')).find(el => el.innerText && el.innerText.trim() === 'Word');
    if (wordEl) { wordEl.click(); }
  }
  run();
})();
