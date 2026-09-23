// Main-world content script — runs in the page's JS environment so it can
// patch DataTransfer.prototype.setData before Google Docs' scripts load.
//
// When data-wh-intercept="1" is set on <html> (by the isolated-world script),
// we intercept every setData call and store the plain-text value in
// data-wh-copy-text instead of forwarding it to the native C++ DataTransfer.
// The C++ backing store stays empty, so the browser has no data to commit to
// the real system clipboard.  Combined with e.preventDefault() (called by the
// isolated-world capture handler), nothing is ever written to the clipboard.
(function () {
  const origSetData = DataTransfer.prototype.setData;
  DataTransfer.prototype.setData = function (type, data) {
    if (document.documentElement.getAttribute('data-wh-intercept') === '1') {
      if (type === 'text/plain') {
        document.documentElement.setAttribute('data-wh-copy-text', data);
      }
      // Skip origSetData → C++ DataTransfer stays empty → no clipboard write.
      return;
    }
    return origSetData.apply(this, arguments);
  };
})();
