/**
 * API Route: /api/mb/transactions
 *
 * Handles MB Bank login + transaction history fetching.
 * The mbbank library is Node.js only, so this runs server-side via Next.js API routes.
 */

// Increase timeout for MB Bank operations (login can take 30-40s)
export const config = {
  maxDuration: 60,
};

export default async function handler(req, res) {
  // Always return JSON
  res.setHeader('Content-Type', 'application/json');

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const { username, password, accountNumber, fromDate, toDate } = req.body || {};

    if (!username || !password || !fromDate || !toDate) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields: username, password, fromDate, toDate',
      });
    }

    let MB;
    try {
      const mbModule = await import('mbbank');
      MB = mbModule.MB;
    } catch (importErr) {
      console.error('[MB] Failed to import mbbank:', importErr.message);
      return res.status(500).json({
        success: false,
        error: 'Failed to load mbbank module: ' + importErr.message,
      });
    }

    console.log(`[MB] Logging in as ${username}...`);
    const mb = new MB({
      username,
      password,
      preferredOCRMethod: 'default',
      saveWasm: true,
    });

    try {
      await mb.login();
    } catch (loginErr) {
      console.error('[MB] Login failed:', loginErr.message);
      return res.status(401).json({
        success: false,
        error: 'Login failed: ' + loginErr.message,
      });
    }

    console.log(`[MB] Login successful. Fetching transactions...`);

    const accNum = accountNumber || username;

    let result;
    try {
      result = await mb.getTransactionsHistory({
        accountNumber: accNum,
        fromDate,
        toDate,
      });
    } catch (txnErr) {
      console.error('[MB] Fetch transactions failed:', txnErr.message);
      return res.status(500).json({
        success: false,
        error: 'Failed to fetch transactions: ' + txnErr.message,
      });
    }

    // The mbbank lib may return:
    // - An array directly
    // - An object with transactionHistoryList
    // - An object with other keys
    let transactions = [];
    if (Array.isArray(result)) {
      transactions = result;
    } else if (result?.transactionHistoryList) {
      transactions = result.transactionHistoryList;
    } else if (result && typeof result === 'object') {
      // Try to find an array in the result
      for (const key of Object.keys(result)) {
        if (Array.isArray(result[key])) {
          transactions = result[key];
          break;
        }
      }
    }
    console.log(`[MB] Got ${transactions.length} transactions`);

    return res.status(200).json({
      success: true,
      transactions,
    });
  } catch (err) {
    // Catch-all for any unexpected errors
    console.error('[MB] Unexpected error:', err);
    return res.status(500).json({
      success: false,
      error: err?.message || 'Unknown server error',
    });
  }
}
