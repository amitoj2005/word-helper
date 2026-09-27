const cards = document.querySelectorAll('.style-card');

// Retired themes, mapped to their successor (content.js keeps the same map).
const LEGACY_THEMES = { liquid: 'liquidlive', liquidhd: 'liquidlive', liquid2: 'liquidlive' };

chrome.storage.local.get('theme', ({ theme = 'glass' }) => {
  if (LEGACY_THEMES[theme]) {
    theme = LEGACY_THEMES[theme];
    chrome.storage.local.set({ theme });
  }
  document.querySelector(`[data-theme="${theme}"]`)?.classList.add('active');
});

for (const card of cards) {
  card.addEventListener('click', () => {
    cards.forEach(c => c.classList.remove('active'));
    card.classList.add('active');
    chrome.storage.local.set({ theme: card.dataset.theme });
  });
}
