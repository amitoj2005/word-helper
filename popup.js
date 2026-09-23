const cards = document.querySelectorAll('.style-card');

chrome.storage.local.get('theme', ({ theme = 'glass' }) => {
  document.querySelector(`[data-theme="${theme}"]`)?.classList.add('active');
});

for (const card of cards) {
  card.addEventListener('click', () => {
    cards.forEach(c => c.classList.remove('active'));
    card.classList.add('active');
    chrome.storage.local.set({ theme: card.dataset.theme });
  });
}
