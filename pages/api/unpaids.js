/**
 * API Route: /api/unpaids
 * Proxies requests to the external unpaids API at dat-com-ivory.vercel.app
 */

const EXTERNAL_API = 'https://dat-com-ivory.vercel.app/api/unpaids';

export default async function handler(req, res) {
  try {
    const headers = {
      'Content-Type': 'application/json',
    };

    // Forward admin passcode header if present
    if (req.headers['x-admin-passcode']) {
      headers['x-admin-passcode'] = req.headers['x-admin-passcode'];
    }

    // Forward authorization/cookie if present
    if (req.headers['cookie']) {
      headers['cookie'] = req.headers['cookie'];
    }

    const fetchOptions = {
      method: req.method,
      headers,
    };

    // Forward body for non-GET requests
    if (req.method !== 'GET' && req.body) {
      fetchOptions.body = JSON.stringify(req.body);
    }

    const response = await fetch(EXTERNAL_API, fetchOptions);
    const data = await response.json();

    return res.status(response.status).json(data);
  } catch (err) {
    console.error('[Unpaids Proxy] Error:', err.message);
    return res.status(500).json({
      success: false,
      error: 'Failed to proxy request: ' + err.message,
    });
  }
}
