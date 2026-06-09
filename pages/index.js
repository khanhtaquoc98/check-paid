import { useState, useCallback, useRef, useEffect } from 'react';
import Head from 'next/head';
import * as XLSX from 'xlsx';

const ORDERS_API = '/api/orders';
const MB_API = '/api/mb/transactions';
const ADMIN_PASSCODE = '123456';

// ========================================
// Utility: Extract name from bank transfer description
// Pattern: "CHUYEN KHOAN LUNCH {NAME}" — name may have spaces injected
// ========================================
function extractLunchName(description) {
  if (!description) return null;
  const upper = description.toUpperCase();

  const originalIndices = [];
  for (let i = 0; i < upper.length; i++) {
    if (upper[i] !== ' ') {
      originalIndices.push(i);
    }
  }
  const spaceless = upper.replace(/\s+/g, '');
  const markerStr = 'CHUYENKHOANLUNCH';
  const markerIdx = spaceless.indexOf(markerStr);
  if (markerIdx === -1) return null;

  const afterIdx = markerIdx + markerStr.length;
  if (afterIdx >= spaceless.length) return null;

  let after = upper.substring(originalIndices[afterIdx]).trim();

  const delimiters = [
    /\.\w/,
    /-\s/,
    /-CHUYEN/i,
    /\s+FT\d/i,
    /\s+CT\s/i,
    /\s+(?=[A-Z]*\d)(?=\d*[A-Z])[A-Z\d]{4,}/i,
    /\s+\d/,
    /\s{2,}Ma\s/i,
    /\.\s+TU:/i,
    /\s+Ma\s+giao/i,
    /\s+Ma\s+GD/i,
    /-\s*$/,
  ];

  let endPos = after.length;
  for (const delim of delimiters) {
    const match = after.match(delim);
    if (match && match.index < endPos) {
      endPos = match.index;
    }
  }

  const nameStr = after.substring(0, endPos).trim();
  const cleanedName = nameStr.replace(/\s+/g, '');
  return cleanedName || null;
}

// ========================================
// Utility: Remove Vietnamese diacritics
// ========================================
function removeDiacritics(str) {
  return str
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D');
}

// ========================================
// Utility: Fuzzy match bank name to order userName
// ========================================
function matchNameToUser(bankName, userName) {
  if (!bankName || !userName) return false;
  const cleanBank = bankName.replace(/\s+/g, '').toUpperCase();
  const cleanUser = removeDiacritics(userName).replace(/\s+/g, '').toUpperCase();
  return cleanBank === cleanUser;
}

// ========================================
// Utility: Parse XLSX
// ========================================
function parseTransactions(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const transactions = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
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
// Utility: Parse MB Bank API response
// ========================================
function parseMBTransactions(mbTransactions) {
  const transactions = [];
  for (let i = 0; i < mbTransactions.length; i++) {
    const txn = mbTransactions[i];
    // MB Bank API field: transactionDesc (not description)
    const desc = txn.transactionDesc || txn.description || txn.addDescription || '';
    const name = extractLunchName(desc);

    if (name) {
      transactions.push({
        stt: i + 1,
        date: txn.transactionDate || txn.postDate || '',
        transNo: txn.refNo || '',
        debit: txn.debitAmount || '0',
        credit: txn.creditAmount || '0',
        details: desc,
        extractedName: name,
        source: 'mbbank',
      });
    }
  }
  return transactions;
}

// ========================================
// Utility helpers
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

function formatPrice(price) {
  return new Intl.NumberFormat('vi-VN').format(price) + 'đ';
}

function timeStr() {
  return new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

function todayStr() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

// ========================================
// Main Page
// ========================================
export default function Home() {
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

  // Tab state
  const [activeTab, setActiveTab] = useState('mbbank');

  // MB Bank form
  const [mbUsername, setMbUsername] = useState('');
  const [mbPassword, setMbPassword] = useState('');
  const [mbAccountNumber, setMbAccountNumber] = useState('');
  const [mbStartDate, setMbStartDate] = useState('');
  const [mbEndDate, setMbEndDate] = useState('');
  const [mbLoading, setMbLoading] = useState(false);
  const [mbShowPassword, setMbShowPassword] = useState(false);
  const [mbTxnCount, setMbTxnCount] = useState(0);

  const fileInputRef = useRef(null);
  const logEndRef = useRef(null);

  // Set default dates on client side only (avoid SSR mismatch)
  useEffect(() => {
    const today = todayStr();
    setMbStartDate(today);
    setMbEndDate(today);
  }, []);

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
  // Find matches helper
  // ========================================
  const findMatches = (orderList, txnList) => {
    const result = [];
    for (const txn of txnList) {
      for (const order of orderList) {
        if (!order.paid && matchNameToUser(txn.extractedName, order.userName)) {
          result.push({ transaction: txn, order });
        }
      }
    }
    return result;
  };

  // ========================================
  // Fetch Orders
  // ========================================
  const fetchOrders = useCallback(async () => {
    setLoading(true);
    addLog('Đang tải danh sách đơn hàng...', 'info');
    try {
      const res = await fetch(ORDERS_API);
      const data = await res.json();
      if (data.success && data.orders) {
        setOrders(data.orders);
        addLog(`Tải thành công ${data.orders.length} đơn hàng`, 'success');
        addToast(`Đã tải ${data.orders.length} đơn hàng`, 'success');

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
  // Pay single order
  // ========================================
  const payOrder = useCallback(async (orderId) => {
    setPaying(prev => new Set(prev).add(orderId));
    addLog(`Đang thanh toán đơn ${orderId}...`, 'info');

    try {
      const res = await fetch(ORDERS_API, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'x-admin-passcode': ADMIN_PASSCODE,
        },
        body: JSON.stringify({ orderId, paid: true }),
      });

      const data = await res.json();

      if (res.ok) {
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
  // Auto pay matches
  // ========================================
  const autoPayMatches = useCallback(async (matchList) => {
    if (!matchList || matchList.length === 0) return;
    setAutoPayProgress({ current: 0, total: matchList.length, running: true });

    for (let i = 0; i < matchList.length; i++) {
      const { order } = matchList[i];
      setAutoPayProgress(prev => ({ ...prev, current: i + 1 }));
      await payOrder(order.id);
      if (i < matchList.length - 1) {
        await new Promise(r => setTimeout(r, 500));
      }
    }

    setAutoPayProgress({ current: 0, total: 0, running: false });
    addLog(`▸ Hoàn tất auto-pay ${matchList.length} đơn!`, 'success');
    addToast('Auto-pay hoàn tất!', 'success');
  }, [payOrder, addLog, addToast]);

  // ========================================
  // MB Bank: Fetch transactions
  // ========================================
  const fetchMBTransactions = useCallback(async () => {
    if (!mbUsername || !mbPassword) {
      addToast('Vui lòng nhập username và password', 'error');
      return;
    }
    if (!mbStartDate || !mbEndDate) {
      addToast('Vui lòng nhập ngày bắt đầu và kết thúc', 'error');
      return;
    }

    setMbLoading(true);
    addLog(`[MB Bank] Đang đăng nhập với tài khoản ${mbUsername}...`, 'info');

    try {
      const res = await fetch(MB_API, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: mbUsername,
          password: mbPassword,
          accountNumber: mbAccountNumber || mbUsername,
          fromDate: mbStartDate,
          toDate: mbEndDate,
        }),
      });

      // Safe JSON parse — handle non-JSON responses (e.g. HTML error pages)
      const text = await res.text();
      let data;
      try {
        data = JSON.parse(text);
      } catch {
        addLog(`[MB Bank] Server trả về lỗi: ${text.substring(0, 200)}`, 'error');
        addToast('Server lỗi! Kiểm tra terminal log.', 'error');
        return;
      }

      if (!data.success) {
        addLog(`[MB Bank] Lỗi: ${data.error}`, 'error');
        addToast(`Lỗi MB Bank: ${data.error}`, 'error');
        return;
      }

      const allMBTxns = data.transactions || [];
      addLog(`[MB Bank] Nhận được ${allMBTxns.length} giao dịch tổng`, 'info');

      const txns = parseMBTransactions(allMBTxns);
      setTransactions(txns);
      setMbTxnCount(allMBTxns.length);
      addLog(`[MB Bank] Tìm thấy ${txns.length} giao dịch "CHUYEN KHOAN LUNCH"`, 'success');

      if (txns.length === 0) {
        addToast(`Nhận ${allMBTxns.length} giao dịch, nhưng không có LUNCH`, 'error');
        return;
      }

      addToast(`Đã đọc ${txns.length} giao dịch LUNCH từ MB Bank`, 'success');

      if (orders.length > 0) {
        const newMatches = findMatches(orders, txns);
        setMatches(newMatches);
        addLog(`Tìm thấy ${newMatches.length} khớp`, newMatches.length > 0 ? 'success' : 'warning');

        if (newMatches.length > 0) {
          addLog(`▸ Tự động thanh toán ${newMatches.length} đơn hàng khớp...`, 'info');
          autoPayMatches(newMatches);
        }
      }
    } catch (err) {
      addLog(`[MB Bank] Lỗi kết nối: ${err.message}`, 'error');
      addToast('Lỗi kết nối MB Bank!', 'error');
    } finally {
      setMbLoading(false);
    }
  }, [mbUsername, mbPassword, mbAccountNumber, mbStartDate, mbEndDate, orders, addLog, addToast, autoPayMatches]);

  // ========================================
  // Handle file upload
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

        if (orders.length > 0) {
          const newMatches = findMatches(orders, txns);
          setMatches(newMatches);
          addLog(`Tìm thấy ${newMatches.length} khớp`, newMatches.length > 0 ? 'success' : 'warning');

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
  // Auto pay all matches (button)
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
    <>
      <Head>
        <title>Auto Paid - Lunch Order Payment</title>
        <meta name="description" content="Auto Payment Tool - Tự động đối chiếu giao dịch và thanh toán đơn hàng" />
      </Head>

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

        {/* Tabbed Input Section */}
        <div className="grid-2" style={{ marginBottom: 32 }}>
          <div className="section slide-up" style={{ animationDelay: '0.1s' }}>
            {/* Tab Headers */}
            <div className="tab-bar">
              <button
                className={`tab-btn ${activeTab === 'mbbank' ? 'tab-btn--active' : ''}`}
                onClick={() => setActiveTab('mbbank')}
              >
                <span className="tab-icon">🏦</span>
                MB Bank
              </button>
              <button
                className={`tab-btn ${activeTab === 'excel' ? 'tab-btn--active' : ''}`}
                onClick={() => setActiveTab('excel')}
              >
                <span className="tab-icon">📄</span>
                Excel
              </button>
            </div>

            {/* Tab Content */}
            <div className="tab-content">
              {/* ---- MB Bank Tab ---- */}
              {activeTab === 'mbbank' && (
                <div className="mb-form fade-in">
                  <div className="form-group">
                    <label className="form-label" htmlFor="mb-username">
                      <span className="form-label-icon">👤</span>
                      Số điện thoại / Username
                    </label>
                    <input
                      id="mb-username"
                      className="form-input"
                      type="text"
                      placeholder="0123456789"
                      value={mbUsername}
                      onChange={(e) => setMbUsername(e.target.value)}
                      autoComplete="off"
                    />
                  </div>

                  <div className="form-group">
                    <label className="form-label" htmlFor="mb-password">
                      <span className="form-label-icon">🔒</span>
                      Mật khẩu
                    </label>
                    <div className="input-with-toggle">
                      <input
                        id="mb-password"
                        className="form-input"
                        type={mbShowPassword ? 'text' : 'password'}
                        placeholder="••••••••"
                        value={mbPassword}
                        onChange={(e) => setMbPassword(e.target.value)}
                        autoComplete="off"
                      />
                      <button
                        className="input-toggle"
                        type="button"
                        onClick={() => setMbShowPassword(v => !v)}
                        title={mbShowPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
                      >
                        {mbShowPassword ? '🙈' : '👁️'}
                      </button>
                    </div>
                  </div>

                  <div className="form-group">
                    <label className="form-label" htmlFor="mb-account">
                      <span className="form-label-icon">💳</span>
                      Số tài khoản <span className="form-hint">(để trống = dùng SĐT)</span>
                    </label>
                    <input
                      id="mb-account"
                      className="form-input"
                      type="text"
                      placeholder="1234567890"
                      value={mbAccountNumber}
                      onChange={(e) => setMbAccountNumber(e.target.value)}
                      autoComplete="off"
                    />
                  </div>

                  <div className="form-row">
                    <div className="form-group">
                      <label className="form-label" htmlFor="mb-start-date">
                        <span className="form-label-icon">📅</span>
                        Từ ngày
                      </label>
                      <input
                        id="mb-start-date"
                        className="form-input"
                        type="text"
                        placeholder="dd/mm/yyyy"
                        value={mbStartDate}
                        onChange={(e) => setMbStartDate(e.target.value)}
                      />
                    </div>
                    <div className="form-group">
                      <label className="form-label" htmlFor="mb-end-date">
                        <span className="form-label-icon">📅</span>
                        Đến ngày
                      </label>
                      <input
                        id="mb-end-date"
                        className="form-input"
                        type="text"
                        placeholder="dd/mm/yyyy"
                        value={mbEndDate}
                        onChange={(e) => setMbEndDate(e.target.value)}
                      />
                    </div>
                  </div>

                  <button
                    className="btn btn--mb btn--full"
                    onClick={fetchMBTransactions}
                    disabled={mbLoading || !mbUsername || !mbPassword}
                  >
                    {mbLoading ? (
                      <><span className="spinner"></span> Đang tải giao dịch...</>
                    ) : (
                      <>🏦 Lấy giao dịch MB Bank</>
                    )}
                  </button>

                  {mbTxnCount > 0 && (
                    <div className="mb-stats fade-in">
                      <div className="mb-stat">
                        <span className="mb-stat-value">{mbTxnCount}</span>
                        <span className="mb-stat-label">Tổng GD</span>
                      </div>
                      <div className="mb-stat">
                        <span className="mb-stat-value mb-stat-value--green">{transactions.length}</span>
                        <span className="mb-stat-label">LUNCH GD</span>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* ---- Excel Tab ---- */}
              {activeTab === 'excel' && (
                <div className="excel-upload fade-in">
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
                      ? 'Upload file hoặc lấy giao dịch MB Bank để tìm kiếm khớp'
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
    </>
  );
}
