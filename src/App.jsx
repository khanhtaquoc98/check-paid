import { useState, useCallback, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';

const API_BASE = 'https://dat-com-ivory.vercel.app/api/orders';
const ADMIN_PASSCODE = '123456';

// ========================================
// Utility: Extract name from bank transfer description
// Pattern: "CHUYEN KHOAN LUNCH {NAME}" — name may have spaces injected
// ========================================
function extractLunchName(description) {
  if (!description) return null;
  const upper = description.toUpperCase();
  const marker = 'CHUYEN KHOAN LUNCH';
  const idx = upper.indexOf(marker);
  if (idx === -1) return null;

  // Get everything after the marker
  let after = upper.substring(idx + marker.length).trim();

  // The name ends at common delimiters
  // e.g. "KANE   Ma giao dich..." or "MIC HAEL-CHUYEN TIEN..." or "ADA M. TU: ZION"
  // We stop at: dash followed by known keywords, period+space, or multiple spaces before lowercase-ish
  // Strategy: take chars until we hit a delimiter pattern
  const delimiters = [
    /\s{2,}Ma\s/i,        // "   Ma giao dich"
    /-CHUYEN/i,            // "-CHUYEN TIEN"
    /\.\s+TU:/i,           // ". TU: ZION"
    /\s+Ma\s+giao/i,       // " Ma giao"
    /-\s*$/,
  ];

  // First, try to find the earliest delimiter
  let nameStr = after;
  let endPos = after.length;

  for (const delim of delimiters) {
    const match = after.match(delim);
    if (match && match.index < endPos) {
      endPos = match.index;
    }
  }

  nameStr = after.substring(0, endPos).trim();

  // Remove extra spaces (bank may inject spaces in name: "MIC HAEL" → "MICHAEL")
  const cleanedName = nameStr.replace(/\s+/g, '');

  return cleanedName || null;
}

// ========================================
// Utility: Fuzzy match bank name to order userName
// ========================================
function matchNameToUser(bankName, userName) {
  if (!bankName || !userName) return false;
  const cleanBank = bankName.replace(/\s+/g, '').toUpperCase();
  const cleanUser = userName.replace(/\s+/g, '').toUpperCase();
  return cleanBank === cleanUser;
}

// ========================================
// Utility: Parse XLSX and extract transaction details
// ========================================
function parseTransactions(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const transactions = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    // Look for rows that have a numeric STT (column B, index 1) and transaction details (column L, index 11)
    const stt = row[1];
    const details = row[11];

    if (stt && !isNaN(Number(stt)) && details) {
      const name = extractLunchName(String(details));
      if (name) {
        transactions.push({
          stt: Number(stt),
          date: row[4] || '',
          transNo: row[6] || '',
          debit: row[9] || '0',
          credit: row[10] || '0',
          details: String(details),
          extractedName: name,
        });
      }
    }
  }

  return transactions;
}

// ========================================
// Utility: Generate avatar colors
// ========================================
function getAvatarColor(name) {
  const colors = [
    { bg: 'rgba(52, 211, 153, 0.15)', color: '#34d399' },
    { bg: 'rgba(96, 165, 250, 0.15)', color: '#60a5fa' },
    { bg: 'rgba(251, 191, 36, 0.15)', color: '#fbbf24' },
    { bg: 'rgba(167, 139, 250, 0.15)', color: '#a78bfa' },
    { bg: 'rgba(248, 113, 113, 0.15)', color: '#f87171' },
    { bg: 'rgba(244, 114, 182, 0.15)', color: '#f472b6' },
    { bg: 'rgba(45, 212, 191, 0.15)', color: '#2dd4bf' },
    { bg: 'rgba(251, 146, 60, 0.15)', color: '#fb923c' },
  ];
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return colors[Math.abs(hash) % colors.length];
}

// ========================================
// Format price
// ========================================
function formatPrice(price) {
  return new Intl.NumberFormat('vi-VN').format(price) + 'đ';
}

// ========================================
// Time string for logs
// ========================================
function timeStr() {
  return new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

// ========================================
// Main App
// ========================================
export default function App() {
  const [orders, setOrders] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [matches, setMatches] = useState([]);
  const [logs, setLogs] = useState([]);
  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState(0);
  const [loading, setLoading] = useState(false);
  const [paying, setPaying] = useState(new Set());
  const [dragging, setDragging] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [autoPayProgress, setAutoPayProgress] = useState({ current: 0, total: 0, running: false });

  const fileInputRef = useRef(null);
  const logEndRef = useRef(null);

  // Auto-scroll logs
  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [logs]);

  // Toast manager
  const addToast = useCallback((message, type = 'info') => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 4000);
  }, []);

  // Log manager
  const addLog = useCallback((msg, type = '') => {
    setLogs(prev => [...prev, { time: timeStr(), msg, type }]);
  }, []);

  // ========================================
  // Step 1: Fetch Orders
  // ========================================
  const fetchOrders = useCallback(async () => {
    setLoading(true);
    addLog('Đang tải danh sách đơn hàng...', 'info');
    try {
      const res = await fetch(API_BASE);
      const data = await res.json();
      if (data.success && data.orders) {
        setOrders(data.orders);
        addLog(`Tải thành công ${data.orders.length} đơn hàng`, 'success');
        addToast(`Đã tải ${data.orders.length} đơn hàng`, 'success');

        // Re-run matching if transactions exist
        if (transactions.length > 0) {
          const newMatches = findMatches(data.orders, transactions);
          setMatches(newMatches);
        }
      } else {
        addLog('Không có dữ liệu đơn hàng', 'warning');
      }
    } catch (err) {
      addLog(`Lỗi tải đơn hàng: ${err.message}`, 'error');
      addToast('Lỗi tải đơn hàng!', 'error');
    } finally {
      setLoading(false);
    }
  }, [transactions, addLog, addToast]);

  // Auto fetch on mount
  useEffect(() => {
    fetchOrders();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // ========================================
  // Step 2: Pay single order
  // ========================================
  const payOrder = useCallback(async (orderId) => {
    setPaying(prev => new Set(prev).add(orderId));
    addLog(`Đang thanh toán đơn ${orderId}...`, 'info');

    try {
      const res = await fetch(API_BASE, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'x-admin-passcode': ADMIN_PASSCODE,
        },
        body: JSON.stringify({ orderId, paid: true }),
      });

      const data = await res.json();

      if (res.ok) {
        // Update local state
        setOrders(prev => prev.map(o => o.id === orderId ? { ...o, paid: true } : o));
        setMatches(prev => prev.filter(m => m.order.id !== orderId));
        addLog(`✓ Đã thanh toán đơn ${orderId}`, 'success');
        addToast(`Đã thanh toán thành công!`, 'success');
      } else {
        addLog(`✗ Lỗi thanh toán đơn ${orderId}: ${JSON.stringify(data)}`, 'error');
        addToast('Lỗi thanh toán!', 'error');
      }
    } catch (err) {
      addLog(`✗ Lỗi: ${err.message}`, 'error');
      addToast('Lỗi kết nối!', 'error');
    } finally {
      setPaying(prev => {
        const next = new Set(prev);
        next.delete(orderId);
        return next;
      });
    }
  }, [addLog, addToast]);

  // ========================================
  // Auto pay a given list of matches
  // ========================================
  const autoPayMatches = useCallback(async (matchList) => {
    if (!matchList || matchList.length === 0) return;

    setAutoPayProgress({ current: 0, total: matchList.length, running: true });

    for (let i = 0; i < matchList.length; i++) {
      const { order } = matchList[i];
      setAutoPayProgress(prev => ({ ...prev, current: i + 1 }));
      await payOrder(order.id);
      // Small delay between requests
      if (i < matchList.length - 1) {
        await new Promise(r => setTimeout(r, 500));
      }
    }

    setAutoPayProgress({ current: 0, total: 0, running: false });
    addLog(`▸ Hoàn tất auto-pay ${matchList.length} đơn!`, 'success');
    addToast('Auto-pay hoàn tất!', 'success');
  }, [payOrder, addLog, addToast]);

  // ========================================
  // Find matches helper
  // ========================================
  const findMatches = (orderList, txnList) => {
    const result = [];
    for (const txn of txnList) {
      for (const order of orderList) {
        if (!order.paid && matchNameToUser(txn.extractedName, order.userName)) {
          result.push({
            transaction: txn,
            order: order,
          });
        }
      }
    }
    return result;
  };

  // ========================================
  // Step 3: Handle file upload → auto match → auto pay
  // ========================================
  const handleFile = useCallback((file) => {
    if (!file) return;
    setFileName(file.name);
    setFileSize(file.size);
    addLog(`Đang đọc file: ${file.name}`, 'info');

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const txns = parseTransactions(workbook);
        setTransactions(txns);
        addLog(`Tìm thấy ${txns.length} giao dịch "CHUYEN KHOAN LUNCH"`, 'success');

        if (txns.length === 0) {
          addToast('Không tìm thấy giao dịch LUNCH trong file', 'error');
          return;
        }

        addToast(`Đã đọc ${txns.length} giao dịch từ file`, 'success');

        // Auto match + auto pay
        if (orders.length > 0) {
          const newMatches = findMatches(orders, txns);
          setMatches(newMatches);
          addLog(`Tìm thấy ${newMatches.length} khớp`, newMatches.length > 0 ? 'success' : 'warning');

          // Automatically call paid API for all matched orders
          if (newMatches.length > 0) {
            addLog(`▸ Tự động thanh toán ${newMatches.length} đơn hàng khớp...`, 'info');
            autoPayMatches(newMatches);
          }
        }
      } catch (err) {
        addLog(`Lỗi đọc file: ${err.message}`, 'error');
        addToast('Lỗi đọc file Excel!', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }, [orders, addLog, addToast, autoPayMatches]);

  // ========================================
  // Auto pay all matches (button click)
  // ========================================
  const autoPayAll = useCallback(async () => {
    if (matches.length === 0) return;

    const toPay = [...matches];
    setAutoPayProgress({ current: 0, total: toPay.length, running: true });
    addLog(`▸ Bắt đầu auto-pay ${toPay.length} đơn hàng...`, 'info');

    for (let i = 0; i < toPay.length; i++) {
      const { order } = toPay[i];
      setAutoPayProgress(prev => ({ ...prev, current: i + 1 }));
      await payOrder(order.id);
      // Small delay between requests
      if (i < toPay.length - 1) {
        await new Promise(r => setTimeout(r, 500));
      }
    }

    setAutoPayProgress({ current: 0, total: 0, running: false });
    addLog(`▸ Hoàn tất auto-pay!`, 'success');
    addToast('Auto-pay hoàn tất!', 'success');
  }, [matches, payOrder, addLog, addToast]);

  // ========================================
  // File drop handlers
  // ========================================
  const handleDragOver = (e) => { e.preventDefault(); setDragging(true); };
  const handleDragLeave = () => setDragging(false);
  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const removeFile = () => {
    setFileName('');
    setFileSize(0);
    setTransactions([]);
    setMatches([]);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // ========================================
  // Stats
  // ========================================
  const paidCount = orders.filter(o => o.paid).length;
  const unpaidCount = orders.filter(o => !o.paid).length;

  return (
    <div className="app-container">
      {/* Header */}
      <header className="app-header slide-up">
        <h1>⚡ Auto Paid</h1>
        <p>Tự động đối chiếu giao dịch ngân hàng và thanh toán đơn hàng lunch</p>
        <div style={{ marginTop: 12 }}>
          <span className="status-badge status-badge--live">Connected</span>
        </div>
      </header>

      {/* Stats Bar */}
      {orders.length > 0 && (
        <div className="action-bar slide-up">
          <div className="action-bar-info">
            <div className="stat">
              <div className="stat-value stat-value--blue">{orders.length}</div>
              <div className="stat-label">Tổng đơn</div>
            </div>
            <div className="stat">
              <div className="stat-value stat-value--emerald">{paidCount}</div>
              <div className="stat-label">Đã trả</div>
            </div>
            <div className="stat">
              <div className="stat-value stat-value--red">{unpaidCount}</div>
              <div className="stat-label">Chưa trả</div>
            </div>
            <div className="stat">
              <div className="stat-value stat-value--amber">{matches.length}</div>
              <div className="stat-label">Khớp</div>
            </div>
          </div>
          <div className="action-bar-buttons">
            <button className="btn btn--secondary btn--sm" onClick={fetchOrders} disabled={loading}>
              {loading ? <><span className="spinner"></span> Đang tải...</> : '🔄 Refresh'}
            </button>
            <button
              className="btn btn--primary btn--sm"
              onClick={autoPayAll}
              disabled={matches.length === 0 || autoPayProgress.running}
            >
              {autoPayProgress.running
                ? <><span className="spinner"></span> {autoPayProgress.current}/{autoPayProgress.total}</>
                : `⚡ Auto Pay (${matches.length})`
              }
            </button>
          </div>
        </div>
      )}

      {autoPayProgress.running && (
        <div className="progress-bar" style={{ marginBottom: 24, marginTop: -16 }}>
          <div
            className="progress-bar-fill"
            style={{ width: `${(autoPayProgress.current / autoPayProgress.total) * 100}%` }}
          />
        </div>
      )}

      {/* Upload Section */}
      <div className="grid-2" style={{ marginBottom: 32 }}>
        <div className="section slide-up" style={{ animationDelay: '0.1s' }}>
          <div className="section-title">
            📄 Upload giao dịch ngân hàng
          </div>
          <div
            className={`upload-zone ${dragging ? 'dragging' : ''} ${fileName ? 'has-file' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => !fileName && fileInputRef.current?.click()}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(e) => handleFile(e.target.files[0])}
            />
            {fileName ? (
              <div className="file-info">
                <div className="file-info-left">
                  <div className="file-icon">📊</div>
                  <div>
                    <div className="file-name">{fileName}</div>
                    <div className="file-meta">{(fileSize / 1024).toFixed(1)} KB · {transactions.length} giao dịch</div>
                  </div>
                </div>
                <button className="file-remove" onClick={(e) => { e.stopPropagation(); removeFile(); }}>✕</button>
              </div>
            ) : (
              <>
                <div className="upload-icon">📁</div>
                <div className="upload-text">Kéo thả file Excel vào đây</div>
                <div className="upload-hint">Hoặc click để chọn file (.xlsx, .xls, .csv)</div>
              </>
            )}
          </div>
        </div>

        {/* Match Results */}
        <div className="section slide-up" style={{ animationDelay: '0.15s' }}>
          <div className="section-title">
            🔗 Kết quả khớp
            {matches.length > 0 && <span className="count">{matches.length}</span>}
          </div>
          {matches.length > 0 ? (
            <div className="match-results">
              {matches.map((m, i) => (
                <div
                  className="match-row"
                  key={`${m.order.id}-${i}`}
                  style={{ animationDelay: `${i * 0.05}s` }}
                >
                  <div className="match-bank">
                    <strong>{m.transaction.extractedName}</strong>
                    <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', marginTop: 2 }}>
                      {m.transaction.credit}đ
                    </div>
                  </div>
                  <div className="match-arrow">→</div>
                  <div className="match-order">
                    <strong>{m.order.userName}</strong>
                    <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', marginTop: 2 }}>
                      #{m.order.id}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="card">
              <div className="empty-state">
                <div className="empty-state-icon">🔍</div>
                <div className="empty-state-text">
                  {transactions.length === 0
                    ? 'Upload file để tìm kiếm khớp'
                    : 'Không tìm thấy giao dịch khớp'}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Orders List */}
      <div className="section slide-up" style={{ animationDelay: '0.2s' }}>
        <div className="section-title">
          🍽️ Danh sách đơn hàng
          {orders.length > 0 && <span className="count">{orders.length}</span>}
        </div>
        {orders.length > 0 ? (
          <>
            <div className="order-header">
              <div></div>
              <div>Tên</div>
              <div>Món ăn</div>
              <div style={{ textAlign: 'right' }}>Giá</div>
              <div style={{ textAlign: 'center' }}>Trạng thái</div>
              <div style={{ textAlign: 'right' }}>Hành động</div>
            </div>
            <div className="orders-grid">
              {orders.map((order, i) => {
                const isMatched = matches.some(m => m.order.id === order.id);
                const isPaying = paying.has(order.id);
                const avatarColor = getAvatarColor(order.userName);

                return (
                  <div
                    className={`order-row ${order.paid ? 'paid' : ''} ${isMatched ? 'matched' : ''} ${isPaying ? 'paying' : ''}`}
                    key={order.id}
                    style={{ animationDelay: `${i * 0.03}s` }}
                  >
                    <div
                      className="order-avatar"
                      style={{ background: avatarColor.bg, color: avatarColor.color }}
                    >
                      {order.userName.charAt(0)}
                    </div>
                    <div className="order-name">{order.userName}</div>
                    <div className="order-dishes">
                      {order.dishes.map(d => d.name).join(', ')}
                    </div>
                    <div className="order-price">{formatPrice(order.totalPrice)}</div>
                    <div className="order-status">
                      {order.paid ? (
                        <span className="pill pill--paid">✓ Paid</span>
                      ) : isPaying ? (
                        <span className="pill pill--paying">
                          <span className="spinner" style={{ width: 10, height: 10 }}></span> Paying
                        </span>
                      ) : isMatched ? (
                        <span className="pill pill--matched">⚡ Matched</span>
                      ) : (
                        <span className="pill pill--unpaid">✕ Unpaid</span>
                      )}
                    </div>
                    <div className="order-action">
                      {!order.paid && !isPaying && (
                        <button
                          className="btn btn--primary btn--sm"
                          onClick={() => payOrder(order.id)}
                        >
                          Pay
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </>
        ) : (
          <div className="card">
            <div className="empty-state">
              <div className="empty-state-icon">📦</div>
              <div className="empty-state-text">
                {loading ? 'Đang tải đơn hàng...' : 'Không có đơn hàng'}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Logs */}
      <div className="section slide-up" style={{ animationDelay: '0.25s' }}>
        <div className="log-panel">
          <div className="log-header">
            <div className="log-header-title">📋 Activity Log</div>
            <button
              className="btn btn--secondary btn--sm"
              onClick={() => setLogs([])}
              style={{ padding: '4px 10px', fontSize: '0.7rem' }}
            >
              Clear
            </button>
          </div>
          <div className="log-body">
            {logs.length === 0 ? (
              <div style={{ color: 'var(--text-muted)', textAlign: 'center', padding: '20px 0' }}>
                Chưa có hoạt động nào...
              </div>
            ) : (
              logs.map((log, i) => (
                <div className="log-entry" key={i}>
                  <span className="log-time">{log.time}</span>
                  <span className={`log-msg ${log.type ? `log-msg--${log.type}` : ''}`}>{log.msg}</span>
                </div>
              ))
            )}
            <div ref={logEndRef} />
          </div>
        </div>
      </div>

      {/* Toasts */}
      <div className="toast-container">
        {toasts.map(toast => (
          <div className={`toast toast--${toast.type}`} key={toast.id}>
            {toast.message}
          </div>
        ))}
      </div>
    </div>
  );
}
