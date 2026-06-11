// fix-ssr-window.js
// Fixes "Cannot destructure property 'protocol' of 'window.location'"
//
// Problem: mbbank sets globalThis.window = { globalThis, document: ... }
//          WITHOUT .location → Next.js reads window.location.protocol → crash
//
// Solution: Let packages set window, but ensure .location always has protocol/hostname
//           so Next.js getLocationOrigin() works correctly.

if (typeof global.window !== 'undefined' && global.window && !global.window.location) {
  delete global.window;
}

let _safeWindow;
Object.defineProperty(global, 'window', {
  get() {
    return _safeWindow;
  },
  set(val) {
    _safeWindow = val;
    // If a package sets window without location (like mbbank does),
    // add a fallback location so Next.js doesn't crash
    if (val && typeof val === 'object' && !val.location) {
      val.location = {
        protocol: 'https:',
        hostname: 'localhost',
        port: '',
        href: 'https://localhost',
        origin: 'https://localhost',
      };
    }
  },
  configurable: true,
});
