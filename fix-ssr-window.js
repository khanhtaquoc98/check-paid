// fix-ssr-window.js
// Prevents npm packages (mbbank, xlsx) from polluting global.window on the server,
// which causes Next.js to crash with "Cannot destructure property 'protocol' of 'window.location'"

if (typeof global.window !== 'undefined' && !global.window.location) {
  delete global.window;
}

let _safeWindow;
Object.defineProperty(global, 'window', {
  get() { return _safeWindow; },
  set(val) {
    if (!val || (val.location && val.location.protocol)) {
      _safeWindow = val;
    }
  },
  configurable: true,
});
