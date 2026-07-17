document.querySelectorAll('.expandoButtonLink').forEach(btn => btn.click());

setTimeout(() => {
  (function() {
    const headings = [...document.querySelectorAll('h2, h3, h4')];
    
    const knownIssuesH = headings.find(el => /known issues in this update/i.test(el.textContent.trim()));
    const howToGetH = headings.find(el => /how to get this update/i.test(el.textContent.trim()));

    if (!knownIssuesH || !howToGetH) { console.warn('Başlık bulunamadı'); return; }

    const knownIssuesSection = knownIssuesH.closest('.ocpSection');
    const howToGetSection = howToGetH.closest('.ocpSection');

    const container = document.createElement('div');
    let node = knownIssuesSection.nextElementSibling;
    while (node && node !== howToGetSection) {
      container.appendChild(node.cloneNode(true));
      node = node.nextElementSibling;
    }

    const knownIssuesContent = knownIssuesSection.cloneNode(true);
    knownIssuesContent.querySelector('h2')?.remove();
    container.prepend(knownIssuesContent);

    container.querySelectorAll('.ocpExpander').forEach(el => el.remove());

    const style = document.createElement('style');
    style.textContent = `
      .popBox .ocpExpandoHeadTitleContainer {
        font-weight: bold !important;
        font-size: 1.15em !important;
      }
      .popBox .ocpExpandoHead,
      .popBox .ocpExpandoHead.opened,
      .popBox .ocpExpandoHead *,
      .popBox * {
        border: none !important;
        border-bottom: none !important;
      }
    `;
    document.head.appendChild(style);

    const overlay = document.createElement('div');
    overlay.style.cssText = `
      position:fixed; top:0; left:0; width:100%; height:100%; 
      background:rgba(0,0,0,0.85); z-index:999999; 
      display:flex; flex-direction:column; align-items:center; 
      justify-content:center; gap:12px;
    `;

    const label = document.createElement('div');
    label.textContent = '✅ Seçili — Ctrl+C ile kopyala, sonra kapat';
    label.style.cssText = 'color:white; font:bold 16px sans-serif;';

    const box = document.createElement('div');
    box.className = 'popBox ocpArticleContent';
    box.contentEditable = true;
    box.appendChild(container);
    box.style.cssText = `
      background:white; padding:24px; border-radius:8px;
      width:80%; max-height:70vh; overflow-y:auto;
    `;

    const closeBtn = document.createElement('button');
    closeBtn.textContent = '✕ Kapat';
    closeBtn.style.cssText = 'padding:8px 20px; font-size:14px; cursor:pointer; border-radius:4px; border:none; background:#e74c3c; color:white;';
    closeBtn.onclick = () => {
      document.body.removeChild(overlay);
      document.head.removeChild(style);
    };

    overlay.appendChild(label);
    overlay.appendChild(box);
    overlay.appendChild(closeBtn);
    document.body.appendChild(overlay);

    console.clear();

    // Popup açılınca tüm içeriği seçili getir
    setTimeout(() => {
      box.focus();
      const range = document.createRange();
      range.selectNodeContents(box);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }, 100);

  })();
}, 800);