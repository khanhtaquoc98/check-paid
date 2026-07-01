import { useState, useCallback, useRef, useEffect } from 'react';
import Head from 'next/head';
import * as XLSX from 'xlsx';

const UNPAIDS_API = '/api/unpaids';
const MB_API = '/api/mb/transactions';
const ADMIN_PASSCODE = '123456';

// ========================================
// Utility: Extract name & quantity from bank transfer description
// Pattern: "PaidLunch {userName} x{SL}" or fallback "PaidLunch {userName}"
// Also supports legacy "Lunch ..." syntax.
// ========================================
function extractLunchDetails(description) {
  if (!description) return null;
  const regex = /(?:PAID)?LUNCH\s+([A-Z0-9a-zÀ-ỹ\s]+?)\s*X\s*(\d+)/i;
  const match = description.match(regex);
  if (match) {
    return {
      userName: match[1].trim(),
      quantity: parseInt(match[2], 10),
    };
  }

  // Fallback: PaidLunch {userName} without x{SL}
  const regexNoQty = /(?:PAID)?LUNCH\s+([A-Z0-9a-zÀ-ỹ\s]+)/i;
  const matchNoQty = description.match(regexNoQty);
  if (matchNoQty) {
    let namePart = matchNoQty[1].trim();
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
    let endPos = namePart.length;
    for (const delim of delimiters) {
      const match = namePart.match(delim);
      if (match && match.index < endPos) {
        endPos = match.index;
      }
    }
    const cleanedName = namePart.substring(0, endPos).trim();
    if (cleanedName) {
      return {
        userName: cleanedName,
        quantity: 1,
      };
    }
  }
  return null;
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
  const cleanBank = removeDiacritics(bankName).replace(/\s+/g, '').toUpperCase();
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
      const detailsStr = String(details);
      const detailsObj = extractLunchDetails(detailsStr);
      if (detailsObj) {
        transactions.push({
          stt: Number(stt),
          date: row[4] || '',
          transNo: row[6] || '',
          debit: row[9] || '0',
          credit: row[10] || '0',
          details: detailsStr,
          extractedName: detailsObj.userName,
          quantity: detailsObj.quantity,
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
    const desc = txn.transactionDesc || txn.description || txn.addDescription || '';
    const detailsObj = extractLunchDetails(desc);

    if (detailsObj) {
      transactions.push({
        stt: i + 1,
        date: txn.transactionDate || txn.postDate || '',
        transNo: txn.refNo || '',
        debit: txn.debitAmount || '0',
        credit: txn.creditAmount || '0',
        details: desc,
        extractedName: detailsObj.userName,
        quantity: detailsObj.quantity,
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
export default function Unpaids() {
  const [unpaids, setUnpaids] = useState([]);
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
  const findMatches = (unpaidList, txnList) => {
    const result = [];
    for (const txn of txnList) {
      for (const item of unpaidList) {
        if (!item.paid && matchNameToUser(txn.extractedName, item.userName)) {
          result.push({ transaction: txn, order: item });
        }
      }
    }
    return result;
  };

  // ========================================
  // Fetch Unpaids
  // ========================================
  const fetchUnpaids = useCallback(async () => {
    setLoading(true);
    addLog('Đang tải danh sách chưa thanh toán...', 'info');
    try {
      const res = await fetch(UNPAIDS_API);
      const data = await res.json();
      if (data.success && data.unpaids) {
        setUnpaids(data.unpaids);
        addLog(`Tải thành công ${data.unpaids.length} đơn hàng chưa thanh toán`, 'success');
        addToast(`Đã tải ${data.unpaids.length} đơn hàng chưa thanh toán`, 'success');

        if (data.config && data.config.accountNo) {
          setMbAccountNumber(prev => prev || data.config.accountNo);
          setMbUsername(prev => prev || data.config.accountNo);
        }

        if (transactions.length > 0) {
          const newMatches = findMatches(data.unpaids, transactions);
          setMatches(newMatches);
        }
      } else {
        addLog('Không có dữ liệu đơn hàng chưa thanh toán', 'warning');
      }
    } catch (err) {
      addLog(`Lỗi tải đơn hàng chưa thanh toán: ${err.message}`, 'error');
      addToast('Lỗi tải dữ liệu chưa thanh toán!', 'error');
    } finally {
      setLoading(false);
    }
  }, [transactions, addLog, addToast]);

  // Auto fetch on mount
  useEffect(() => {
    fetchUnpaids();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // ========================================
  // Pay all orders of a user
  // ========================================
  const payUser = useCallback(async (userName) => {
    setPaying(prev => new Set(prev).add(userName));
    addLog(`Đang thanh toán tất cả đơn hàng cho ${userName}...`, 'info');

    try {
      const res = await fetch(UNPAIDS_API, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'x-admin-passcode': ADMIN_PASSCODE,
        },
        body: JSON.stringify({ action: 'pay_all', userName }),
      });

      const data = await res.json();

      if (res.ok && data.success) {
        setUnpaids(prev => prev.map(o => o.userName === userName ? { ...o, paid: true } : o));
        setMatches(prev => prev.filter(m => m.order.userName !== userName));
        addLog(`✓ Đã thanh toán tất cả đơn hàng của ${userName}`, 'success');
        addToast(`Thanh toán thành công cho ${userName}!`, 'success');
      } else {
        addLog(`✗ Lỗi thanh toán cho ${userName}: ${JSON.stringify(data)}`, 'error');
        addToast('Lỗi thanh toán!', 'error');
      }
    } catch (err) {
      addLog(`✗ Lỗi: ${err.message}`, 'error');
      addToast('Lỗi kết nối!', 'error');
    } finally {
      setPaying(prev => {
        const next = new Set(prev);
        next.delete(userName);
        return next;
      });
    }
  }, [addLog, addToast]);

  // ========================================
  // Auto pay all matches (button)
  // ========================================
  const autoPayAll = useCallback(async () => {
    if (matches.length === 0) return;
    const uniqueUserNames = Array.from(new Set(matches.map(m => m.order.userName)));
    setAutoPayProgress({ current: 0, total: uniqueUserNames.length, running: true });
    addLog(`▸ Bắt đầu tự động thanh toán cho ${uniqueUserNames.length} người dùng khớp...`, 'info');

    for (let i = 0; i < uniqueUserNames.length; i++) {
      const userName = uniqueUserNames[i];
      setAutoPayProgress(prev => ({ ...prev, current: i + 1 }));
      await payUser(userName);
      if (i < uniqueUserNames.length - 1) {
        await new Promise(r => setTimeout(r, 500));
      }
    }

    setAutoPayProgress({ current: 0, total: 0, running: false });
    addLog(`▸ Hoàn tất tự động thanh toán!`, 'success');
    addToast('Hoàn tất tự động thanh toán!', 'success');
  }, [matches, payUser, addLog, addToast]);

  // ========================================
  // Auto pay matches (triggered after loading bank or excel)
  // ========================================
  const autoPayMatches = useCallback(async (matchList) => {
    if (!matchList || matchList.length === 0) return;
    const uniqueUserNames = Array.from(new Set(matchList.map(m => m.order.userName)));
    setAutoPayProgress({ current: 0, total: uniqueUserNames.length, running: true });
    addLog(`▸ Bắt đầu tự động thanh toán cho ${uniqueUserNames.length} người dùng khớp...`, 'info');

    for (let i = 0; i < uniqueUserNames.length; i++) {
      const userName = uniqueUserNames[i];
      setAutoPayProgress(prev => ({ ...prev, current: i + 1 }));
      await payUser(userName);
      if (i < uniqueUserNames.length - 1) {
        await new Promise(r => setTimeout(r, 500));
      }
    }

    setAutoPayProgress({ current: 0, total: 0, running: false });
    addLog(`▸ Hoàn tất tự động thanh toán!`, 'success');
    addToast('Hoàn tất tự động thanh toán!', 'success');
  }, [payUser, addLog, addToast]);

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
    addLog(`[MB Bank] Đang đăng nhập tài khoản ${mbUsername}...`, 'info');

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

      const text = await res.text();
      let data;
      try {
        data = JSON.parse(text);
      } catch {
        addLog(`[MB Bank] Server trả về lỗi: ${text.substring(0, 200)}`, 'error');
        addToast('Lỗi máy chủ! Vui lòng kiểm tra log.', 'error');
        return;
      }

      if (!data.success) {
        addLog(`[MB Bank] Lỗi: ${data.error}`, 'error');
        addToast(`Lỗi MB Bank: ${data.error}`, 'error');
        return;
      }

      const allMBTxns = data.transactions || [];
      addLog(`[MB Bank] Nhận được ${allMBTxns.length} giao dịch`, 'info');

      const txns = parseMBTransactions(allMBTxns);
      setTransactions(txns);
      setMbTxnCount(allMBTxns.length);
      addLog(`[MB Bank] Tìm thấy ${txns.length} giao dịch khớp "PaidLunch {userName} x{SL}"`, 'success');

      if (txns.length === 0) {
        addToast(`Đọc ${allMBTxns.length} giao dịch, không tìm thấy cú pháp PaidLunch`, 'error');
        return;
      }

      addToast(`Đã đọc ${txns.length} giao dịch PaidLunch từ MB Bank`, 'success');

      if (unpaids.length > 0) {
        const newMatches = findMatches(unpaids, txns);
        setMatches(newMatches);
        addLog(`Tìm thấy ${newMatches.length} khớp trùng tên`, newMatches.length > 0 ? 'success' : 'warning');

        if (newMatches.length > 0) {
          autoPayMatches(newMatches);
        }
      }
    } catch (err) {
      addLog(`[MB Bank] Lỗi kết nối: ${err.message}`, 'error');
      addToast('Lỗi kết nối MB Bank!', 'error');
    } finally {
      setMbLoading(false);
    }
  }, [mbUsername, mbPassword, mbAccountNumber, mbStartDate, mbEndDate, unpaids, addLog, addToast, autoPayMatches]);

  // ========================================
  // Handle file upload
  // ========================================
  const handleFile = useCallback((file) => {
    if (!file) return;
    setFileName(file.name);
    setFileSize(file.size);
    addLog(`Đang đọc file Excel: ${file.name}`, 'info');

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const txns = parseTransactions(workbook);
        setTransactions(txns);
        addLog(`Tìm thấy ${txns.length} giao dịch khớp cú pháp PaidLunch`, 'success');

        if (txns.length === 0) {
          addToast('Không tìm thấy giao dịch PaidLunch trong file Excel', 'error');
          return;
        }

        addToast(`Đã đọc ${txns.length} giao dịch từ file`, 'success');

        if (unpaids.length > 0) {
          const newMatches = findMatches(unpaids, txns);
          setMatches(newMatches);
          addLog(`Tìm thấy ${newMatches.length} khớp`, newMatches.length > 0 ? 'success' : 'warning');

          if (newMatches.length > 0) {
            autoPayMatches(newMatches);
          }
        }
      } catch (err) {
        addLog(`Lỗi đọc file: ${err.message}`, 'error');
        addToast('Lỗi đọc file Excel!', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }, [unpaids, addLog, addToast, autoPayMatches]);

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
  const paidCount = unpaids.filter(o => o.paid).length;
  const unpaidCount = unpaids.filter(o => !o.paid).length;

  return (
    <>
      <Head>
        <title>Auto Paid - Unpaids Check</title>
        <meta name="description" content="Auto Payment Tool - Check danh sách đơn hàng chưa thanh toán" />
      </Head>

      <div className="app-container">
        {/* Header */}
        <header className="app-header slide-up">
          <h1>⚡ Auto Paid Check</h1>
          <p>Tự động đối chiếu giao dịch ngân hàng và thanh toán các đơn hàng chưa thanh toán</p>
          <div className="nav-links-container">
            <a href="/" className="nav-link">Trang chủ</a>
            <a href="/unpaids" className="nav-link nav-link--active">Chưa thanh toán</a>
          </div>
        </header>

        {/* Stats Bar */}
        {unpaids.length > 0 && (
          <div className="action-bar slide-up">
            <div className="action-bar-info">
              <div className="stat">
                <div className="stat-value stat-value--blue">{unpaids.length}</div>
                <div className="stat-label">Tổng đơn</div>
              </div>
              <div className="stat">
                <div className="stat-value stat-value--emerald">{paidCount}</div>
                <div className="stat-label">Đã xử lý</div>
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
              <button className="btn btn--secondary btn--sm" onClick={fetchUnpaids} disabled={loading}>
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
                        <span className="mb-stat-label">PAIDLUNCH GD</span>
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
                      <span style={{ fontSize: '0.75rem', color: 'var(--accent-amber)', marginLeft: 8 }}>
                        (x{m.transaction.quantity})
                      </span>
                      <div style={{ fontSize: '0.72rem', color: 'var(--text-muted)', marginTop: 2 }}>
                        {m.transaction.details}
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

        {/* Unpaids List */}
        <div className="section slide-up" style={{ animationDelay: '0.2s' }}>
          <div className="section-title">
            🍽️ Danh sách chưa thanh toán
            {unpaids.length > 0 && <span className="count">{unpaids.length}</span>}
          </div>
          {unpaids.length > 0 ? (
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
                {unpaids.map((order, i) => {
                  const isMatched = matches.some(m => m.order.id === order.id);
                  const isPaying = paying.has(order.userName);
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
                        {order.dishes.map(d => `${d.name} (${d.paid ? 'Đã trả' : 'Chưa trả'})`).join(', ')}
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
                            onClick={() => payUser(order.userName)}
                          >
                            Pay All
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
                  {loading ? 'Đang tải đơn hàng...' : 'Không có đơn hàng chưa thanh toán'}
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
