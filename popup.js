const cards = document.querySelectorAll('.style-card');

// Retired themes, mapped to their successor (content.js keeps the same map).
const LEGACY_THEMES = { liquid: 'liquidlive', liquidhd: 'liquidlive', liquid2: 'liquidlive' };

chrome.storage.local.get('theme', ({ theme = 'glass' }) => {
  if (LEGACY_THEMES[theme]) {
    theme = LEGACY_THEMES[theme];
    chrome.storage.local.set({ theme });
  }
  select(document.querySelector(`[data-theme="${theme}"]`));
});

// Highlight one card; aria-pressed keeps the choice audible to screen readers.
function select(card) {
  for (const c of cards) {
    c.classList.toggle('active', c === card);
    c.setAttribute('aria-pressed', String(c === card));
  }
}

for (const card of cards) {
  card.addEventListener('click', () => {
    select(card);
    chrome.storage.local.set({ theme: card.dataset.theme });
  });
}
