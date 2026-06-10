// server.js — Custom Next.js server for Railway
// Fixes: "Cannot destructure property 'protocol' of 'window.location'"
//
// Root cause: some npm packages (mbbank, xlsx) set global.window = {} on the
// server. Next.js then calls getLocationOrigin() → window.location.protocol → crash.
//
// Fix: 1) Block global.window pollution
//      2) Use custom server that passes parsedUrl to handler, bypassing parseUrl()

// ── Step 1: Prevent window pollution BEFORE loading anything ──
if (typeof global.window !== 'undefined' && !global.window.location) {
  delete global.window;
}

let _safeWindow;
Object.defineProperty(global, 'window', {
  get() {
    return _safeWindow;
  },
  set(val) {
    // Only allow setting if it has a real location (browser-like env) or is undefined/null
    if (!val || (val.location && val.location.protocol)) {
      _safeWindow = val;
    }
    // Otherwise silently block — some package tried to set window = {}
  },
  configurable: true,
});

// ── Step 2: Custom server with explicit URL parsing ──
const { createServer } = require('http');
const { parse } = require('url');
const next = require('next');

const dev = process.env.NODE_ENV !== 'production';
const hostname = process.env.HOSTNAME || '0.0.0.0';
const port = parseInt(process.env.PORT, 10) || 4321;

const app = next({ dev, hostname, port });
const handle = app.getRequestHandler();

app.prepare().then(() => {
  createServer(async (req, res) => {
    try {
      // Parse URL ourselves and pass it to Next.js handler.
      // This bypasses Next.js internal parseUrl() which calls getLocationOrigin()
      const parsedUrl = parse(req.url, true);
      await handle(req, res, parsedUrl);
    } catch (err) {
      console.error('Error handling', req.url, err);
      res.statusCode = 500;
      res.end('Internal Server Error');
    }
  })
    .once('error', (err) => {
      console.error(err);
      process.exit(1);
    })
    .listen(port, hostname, () => {
      console.log(`> Ready on http://${hostname}:${port}`);
    });
});
