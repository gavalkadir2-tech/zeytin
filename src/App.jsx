import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import {
  LayoutDashboard, Users, Scale as ScaleIcon, Wallet, FileBarChart, Warehouse, Settings as SettingsIcon,
  Plus, Printer, Bluetooth, BluetoothConnected, Search, X, Phone, ChevronRight, Download, Package, ShoppingCart,
} from 'lucide-react';

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const todayStr = () => new Date().toISOString().slice(0, 10);
const fmtTL = (n) => (Number(n) || 0).toLocaleString('tr-TR', { style: 'currency', currency: 'TRY', maximumFractionDigits: 2 });
const fmtKg = (n) => (Number(n) || 0).toLocaleString('tr-TR', { maximumFractionDigits: 1 }) + ' kg';
const fmtDate = (d) => new Date(d).toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric' });
const fmtDateShort = (d) => new Date(d).toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit' });

// Bu proje bağımsız (standalone) çalıştığı için Claude artifact ortamındaki
// window.storage yerine tarayıcının kendi localStorage'ını kullanır.
// Veriler sadece bu tarayıcıda / bu cihazda saklanır.
async function storageGet(key) {
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}
async function storageSet(key, value) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.error('Depolama hatasi:', e);
  }
}

const COLORS = {
  ink: '#2B2A25',
  inkSoft: '#5B5A50',
  paper: '#F3F0E6',
  paperCard: '#FFFFFF',
  border: '#DCD6C4',
  olive: '#3F4A2E',
  oliveLight: '#6B7A4F',
  oliveSoft: '#E7EBDC',
  gold: '#B3892B',
  goldSoft: '#F4E9D2',
  red: '#A13D2E',
  redSoft: '#F6E4DF',
  blue: '#3B5E73',
  blueSoft: '#E1EAEE',
};

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');
      :root { --font-display: 'Fraunces', Georgia, serif; --font-body: 'Inter', -apple-system, 'Segoe UI', sans-serif; }
      .zk-app { font-family: var(--font-body); color: ${COLORS.ink}; background: ${COLORS.paper}; min-height: 100vh; }
      .zk-shell { display: flex; min-height: 100vh; }
      .zk-sidebar { width: 216px; background: ${COLORS.olive}; flex-shrink: 0; padding: 22px 12px; display: flex; flex-direction: column; gap: 3px; position: relative; }
      .zk-brand-row { display: flex; align-items: center; gap: 9px; padding: 0 10px; margin-bottom: 2px; }
      .zk-brand { color: #F5F2E8; font-family: var(--font-display); font-size: 19px; font-weight: 600; letter-spacing: 0.2px; }
      .zk-brand-sub { color: #A9B896; font-size: 10.5px; padding: 0 10px; margin-bottom: 24px; letter-spacing: 0.6px; text-transform: uppercase; }
      .zk-navbtn { display: flex; align-items: center; gap: 10px; padding: 9px 11px; border-radius: 8px; background: transparent; border: none; border-left: 2px solid transparent; color: #C9D2B9; font-size: 13px; font-weight: 500; cursor: pointer; text-align: left; width: 100%; transition: background 0.12s ease, color 0.12s ease; }
      .zk-navbtn:hover { background: rgba(255,255,255,0.07); color: #F5F2E8; }
      .zk-navbtn.active { background: rgba(255,255,255,0.13); color: #fff; border-left: 2px solid ${COLORS.gold}; }
      .zk-main { flex: 1; padding: 28px 32px; max-width: 1180px; }
      .zk-h1 { font-family: var(--font-display); font-size: 23px; font-weight: 600; margin-bottom: 3px; letter-spacing: 0.1px; }
      .zk-h1-sub { font-size: 12.5px; color: ${COLORS.inkSoft}; margin-bottom: 20px; }
      .zk-card { background: ${COLORS.paperCard}; border: 1px solid ${COLORS.border}; border-radius: 12px; padding: 16px 18px; transition: box-shadow 0.15s ease, border-color 0.15s ease; }
      .zk-grid { display: grid; gap: 12px; }
      .zk-stat { background: ${COLORS.paperCard}; border: 1px solid ${COLORS.border}; border-radius: 12px; padding: 14px 16px; transition: box-shadow 0.15s ease, transform 0.15s ease; }
      .zk-stat:hover { box-shadow: 0 4px 14px rgba(43,42,37,0.07); transform: translateY(-1px); }
      .zk-stat-label { font-size: 11px; color: ${COLORS.inkSoft}; margin-bottom: 6px; display: flex; align-items: center; gap: 6px; }
      .zk-stat-icon { width: 22px; height: 22px; border-radius: 6px; display: inline-flex; align-items: center; justify-content: center; background: ${COLORS.oliveSoft}; color: ${COLORS.olive}; }
      .zk-stat-value { font-family: var(--font-display); font-size: 22px; font-weight: 600; }
      .zk-input, .zk-select { width: 100%; padding: 8px 10px; border-radius: 8px; border: 1px solid ${COLORS.border}; background: #FCFBF7; font-size: 13px; font-family: inherit; color: ${COLORS.ink}; }
      .zk-input:focus, .zk-select:focus { outline: none; border-color: ${COLORS.oliveLight}; box-shadow: 0 0 0 3px ${COLORS.oliveSoft}; }
      .zk-label { font-size: 11.5px; font-weight: 600; color: ${COLORS.inkSoft}; margin-bottom: 4px; display: block; }
      .zk-btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 8px; border: none; font-size: 12.5px; font-weight: 600; cursor: pointer; font-family: inherit; transition: background 0.12s ease, transform 0.08s ease, box-shadow 0.12s ease; }
      .zk-btn:active:not(:disabled) { transform: scale(0.97); }
      .zk-btn-primary { background: ${COLORS.olive}; color: #fff; }
      .zk-btn-primary:hover { background: #333c25; box-shadow: 0 2px 8px rgba(63,74,46,0.25); }
      .zk-btn-gold { background: ${COLORS.gold}; color: #fff; }
      .zk-btn-gold:hover { box-shadow: 0 2px 8px rgba(179,137,43,0.3); }
      .zk-btn-blue { background: ${COLORS.blue}; color: #fff; }
      .zk-btn-blue:hover { box-shadow: 0 2px 8px rgba(59,94,115,0.3); }
      .zk-btn-secondary { background: #fff; color: ${COLORS.ink}; border: 1px solid ${COLORS.border}; }
      .zk-btn-secondary:hover { background: #F7F5EE; }
      .zk-btn:disabled { opacity: 0.5; cursor: not-allowed; }
      .zk-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
      .zk-table th { text-align: left; font-size: 10.5px; letter-spacing: 0.4px; color: ${COLORS.inkSoft}; padding: 7px 9px; border-bottom: 1px solid ${COLORS.border}; font-weight: 600; }
      .zk-table td { padding: 9px 9px; border-bottom: 1px solid #EFEBDD; }
      .zk-table tr { transition: background 0.1s ease; }
      .zk-table tr:hover td { background: #FAF8F1; }
      .zk-badge { display: inline-block; padding: 2px 9px; border-radius: 20px; font-size: 10.5px; font-weight: 600; }
      .zk-badge-olive { background: ${COLORS.oliveSoft}; color: ${COLORS.olive}; }
      .zk-badge-gold { background: ${COLORS.goldSoft}; color: ${COLORS.gold}; }
      .zk-badge-red { background: ${COLORS.redSoft}; color: ${COLORS.red}; }
      .zk-badge-blue { background: ${COLORS.blueSoft}; color: ${COLORS.blue}; }
      .zk-scalebox { background: #1C2226; border-radius: 12px; padding: 15px 16px; color: #fff; }
      .zk-scale-readout { font-family: 'Courier New', monospace; font-size: 30px; font-weight: 700; color: #7FDB8F; }
      .zk-modal-overlay { position: fixed; inset: 0; background: rgba(20,20,15,0.45); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 20px; }
      .zk-modal { background: #fff; border-radius: 14px; padding: 22px; width: 100%; max-width: 440px; max-height: 90vh; overflow-y: auto; }
      .zk-close { background: none; border: none; cursor: pointer; color: ${COLORS.inkSoft}; padding: 4px; }
      .zk-empty { text-align: center; padding: 36px 20px; color: ${COLORS.inkSoft}; font-size: 12.5px; }
      .zk-farmer-row { display: flex; align-items: center; justify-content: space-between; padding: 10px 13px; border-radius: 10px; cursor: pointer; border: 1px solid transparent; transition: background 0.12s ease, border-color 0.12s ease; }
      .zk-farmer-row:hover { background: #FAF8F1; border-color: ${COLORS.border}; }
      .zk-avatar { width: 32px; height: 32px; border-radius: 50%; background: ${COLORS.oliveSoft}; color: ${COLORS.olive}; display: inline-flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; font-family: var(--font-display); flex-shrink: 0; }
      .zk-tag { font-size: 10px; padding: 1px 7px; border-radius: 5px; background: ${COLORS.oliveSoft}; color: ${COLORS.olive}; font-weight: 600; }
      .zk-checkbox-row { display: flex; align-items: center; gap: 8px; font-size: 12.5px; }
      .zk-empty-icon { color: ${COLORS.border}; margin-bottom: 8px; }
      #zk-print-area { display: none; }
      @media print {
        @page { size: 80mm auto; margin: 3mm; }
        body * { visibility: hidden; }
        #zk-print-area, #zk-print-area * { visibility: visible; }
        #zk-print-area { display: block; position: absolute; top: 0; left: 0; width: 100%; }
      }
    `}</style>
  );
}

function StatCard({ label, value, tone, icon: Icon }) {
  return (
    <div className="zk-stat">
      <div className="zk-stat-label">
        {Icon && <span className="zk-stat-icon" style={tone ? { background: tone + '1A', color: tone } : undefined}><Icon size={12} /></span>}
        {label}
      </div>
      <div className="zk-stat-value" style={{ color: tone || COLORS.ink }}>{value}</div>
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="zk-modal-overlay" onClick={onClose}>
      <div className="zk-modal" onClick={(e) => e.stopPropagation()}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <div style={{ fontSize: 15, fontWeight: 700 }}>{title}</div>
          <button className="zk-close" onClick={onClose} aria-label="Kapat"><X size={18} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

function ScaleWidget({ onWeightCapture }) {
  const [connected, setConnected] = useState(false);
  const [status, setStatus] = useState('Bağlı değil');
  const [rawLines, setRawLines] = useState([]);
  const [lastValue, setLastValue] = useState(null);
  const [baud, setBaud] = useState(9600);
  const portRef = useRef(null);
  const readerRef = useRef(null);
  const keepReadingRef = useRef(false);
  const bufferRef = useRef('');

  const extractNumber = (line) => {
    const m = line.match(/-?\d+[.,]?\d*/);
    if (!m) return null;
    return parseFloat(m[0].replace(',', '.'));
  };

  const handleIncoming = useCallback((chunk) => {
    bufferRef.current += chunk;
    let idx;
    while ((idx = bufferRef.current.search(/[\r\n]/)) >= 0) {
      const line = bufferRef.current.slice(0, idx).trim();
      bufferRef.current = bufferRef.current.slice(idx + 1);
      if (!line) continue;
      setRawLines((prev) => [...prev.slice(-4), line]);
      const num = extractNumber(line);
      if (num !== null) setLastValue(num);
    }
  }, []);

  const readLoop = useCallback(async (port) => {
    const decoder = new TextDecoderStream();
    const closed = port.readable.pipeTo(decoder.writable);
    const reader = decoder.readable.getReader();
    readerRef.current = reader;
    try {
      while (keepReadingRef.current) {
        const { value, done } = await reader.read();
        if (done) break;
        if (value) handleIncoming(value);
      }
    } catch (e) {
      console.error(e);
    } finally {
      reader.releaseLock();
      await closed.catch(() => {});
    }
  }, [handleIncoming]);

  const connect = async () => {
    if (!('serial' in navigator)) {
      alert('Bu tarayıcı Web Serial API desteklemiyor. Masaüstü Chrome veya Edge kullanın.');
      return;
    }
    try {
      setStatus('Bağlanıyor...');
      const port = await navigator.serial.requestPort();
      await port.open({ baudRate: baud });
      portRef.current = port;
      setConnected(true);
      setStatus('Bağlı');
      keepReadingRef.current = true;
      readLoop(port);
    } catch (e) {
      setStatus('Bağlanamadı');
    }
  };

  const disconnect = async () => {
    keepReadingRef.current = false;
    try {
      if (readerRef.current) await readerRef.current.cancel();
      if (portRef.current) await portRef.current.close();
    } catch (e) {}
    setConnected(false);
    setStatus('Bağlı değil');
    setLastValue(null);
  };

  return (
    <div className="zk-scalebox">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, color: '#B7C4B3' }}>
          {connected ? <BluetoothConnected size={15} /> : <Bluetooth size={15} />}
          {status}
        </div>
        {!connected && (
          <select value={baud} onChange={(e) => setBaud(Number(e.target.value))} style={{ background: '#2A3128', color: '#DCE2CC', border: '1px solid #3A4235', borderRadius: 6, fontSize: 11, padding: '4px 6px' }}>
            <option value={9600}>9600</option>
            <option value={4800}>4800</option>
            <option value={2400}>2400</option>
            <option value={1200}>1200</option>
            <option value={19200}>19200</option>
          </select>
        )}
      </div>
      <div className="zk-scale-readout">{lastValue !== null ? lastValue.toFixed(1) : '—'}</div>
      {rawLines.length > 0 && (
        <div style={{ fontSize: 10, color: '#6A7A6A', marginTop: 6, fontFamily: 'Courier New, monospace' }}>
          {rawLines[rawLines.length - 1]}
        </div>
      )}
      <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
        {!connected ? (
          <button className="zk-btn zk-btn-gold" onClick={connect} style={{ flex: 1, justifyContent: 'center' }}>Kantara bağlan</button>
        ) : (
          <>
            <button className="zk-btn zk-btn-secondary" onClick={disconnect} style={{ flex: 1, justifyContent: 'center', background: '#2A3128', color: '#DCE2CC', border: 'none' }}>Kes</button>
            <button
              className="zk-btn zk-btn-gold"
              style={{ flex: 1, justifyContent: 'center' }}
              disabled={lastValue === null}
              onClick={() => onWeightCapture(lastValue)}
            >
              Bu değeri kullan
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function stopajOraniHesapla(borsaTescilli) {
  return borsaTescilli ? 2 : 4;
}

function DashboardTab({ farmers, purchases, payments, sales, setTab }) {
  const today = todayStr();
  const todaysPurchases = purchases.filter((p) => p.date === today);
  const totalKgToday = todaysPurchases.reduce((s, p) => s + p.netKg, 0);
  const totalAmountToday = todaysPurchases.reduce((s, p) => s + p.netPayment, 0);

  const totalPurchasedKg = purchases.reduce((s, p) => s + p.netKg, 0);
  const totalSoldKg = sales.reduce((s, s2) => s + s2.kg, 0);
  const currentStock = totalPurchasedKg - totalSoldKg;

  const balances = useMemo(() => {
    const map = {};
    farmers.forEach((f) => { map[f.id] = 0; });
    purchases.forEach((p) => { map[p.farmerId] = (map[p.farmerId] || 0) + p.netPayment; });
    payments.forEach((pay) => { map[pay.farmerId] = (map[pay.farmerId] || 0) - pay.amount; });
    return map;
  }, [farmers, purchases, payments]);
  const totalDebt = Object.values(balances).reduce((s, v) => s + Math.max(v, 0), 0);

  const chartData = useMemo(() => {
    const days = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const key = d.toISOString().slice(0, 10);
      const kg = purchases.filter((p) => p.date === key).reduce((s, p) => s + p.netKg, 0);
      days.push({ date: fmtDateShort(key), kg: Math.round(kg * 10) / 10 });
    }
    return days;
  }, [purchases]);

  const recent = [...purchases].sort((a, b) => b.createdAt - a.createdAt).slice(0, 6);

  return (
    <div>
      <div className="zk-h1">Pano</div>
      <div className="zk-h1-sub">{fmtDate(today)} · genel durum</div>

      <div className="zk-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', marginBottom: 16 }}>
        <StatCard label="Bugün alınan" value={fmtKg(totalKgToday)} />
        <StatCard label="Bugünkü net ödeme" value={fmtTL(totalAmountToday)} />
        <StatCard label="Mevcut stok" value={fmtKg(currentStock)} tone={COLORS.blue} icon={Warehouse} />
        <StatCard label="Ödenecek bakiye" value={fmtTL(totalDebt)} tone={totalDebt > 0 ? COLORS.red : COLORS.olive} />
      </div>

      <div className="zk-card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Son 14 gün alım (kg)</div>
        <div style={{ width: '100%', height: 180 }}>
          <ResponsiveContainer>
            <LineChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EFEBDD" />
              <XAxis dataKey="date" tick={{ fontSize: 10, fill: COLORS.inkSoft }} />
              <YAxis tick={{ fontSize: 10, fill: COLORS.inkSoft }} width={40} />
              <Tooltip formatter={(v) => [v + ' kg', 'Alım']} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
              <Line type="monotone" dataKey="kg" stroke={COLORS.olive} strokeWidth={2} dot={{ r: 2 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="zk-card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div style={{ fontSize: 13.5, fontWeight: 700 }}>Son alımlar</div>
          <button className="zk-btn zk-btn-primary" onClick={() => setTab('purchase')}><Plus size={14} /> Yeni alım</button>
        </div>
        {recent.length === 0 ? (
          <div className="zk-empty"><Package size={26} className="zk-empty-icon" /><br/>Henüz alım kaydı yok. "Yeni alım" ile başlayın.</div>
        ) : (
          <table className="zk-table">
            <thead>
              <tr><th>Tarih</th><th>Çiftçi</th><th>Net kg</th><th>Fiyat</th><th>Net ödeme</th></tr>
            </thead>
            <tbody>
              {recent.map((p) => {
                const f = farmers.find((x) => x.id === p.farmerId);
                return (
                  <tr key={p.id}>
                    <td>{fmtDate(p.date)}</td>
                    <td>{f ? f.name : '—'}</td>
                    <td>{fmtKg(p.netKg)}</td>
                    <td>{fmtTL(p.pricePerKg)}/kg</td>
                    <td>{fmtTL(p.netPayment)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function AddFarmerModal({ onClose, onSave }) {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [tcNo, setTcNo] = useState('');
  const [address, setAddress] = useState('');
  const [bagkurStatus, setBagkurStatus] = useState(false);

  const submit = () => {
    if (!name.trim()) return;
    onSave({ name: name.trim(), phone: phone.trim(), tcNo: tcNo.trim(), address: address.trim(), bagkurStatus });
  };

  return (
    <Modal title="Yeni çiftçi ekle" onClose={onClose}>
      <div style={{ marginBottom: 12 }}>
        <label className="zk-label">Ad soyad</label>
        <input className="zk-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="örn. Mehmet Yılmaz" autoFocus />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 12 }}>
        <div>
          <label className="zk-label">Telefon</label>
          <input className="zk-input" value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="0532 xxx xx xx" />
        </div>
        <div>
          <label className="zk-label">TC Kimlik No</label>
          <input className="zk-input" value={tcNo} onChange={(e) => setTcNo(e.target.value)} placeholder="11 haneli" maxLength={11} />
        </div>
      </div>
      <div style={{ marginBottom: 12 }}>
        <label className="zk-label">Adres (müstahsil makbuzu için)</label>
        <input className="zk-input" value={address} onChange={(e) => setAddress(e.target.value)} placeholder="Köy / ilçe / il" />
      </div>
      <div style={{ marginBottom: 18 }}>
        <label className="zk-checkbox-row">
          <input type="checkbox" checked={bagkurStatus} onChange={(e) => setBagkurStatus(e.target.checked)} />
          Tarım BAĞ-KUR'lu (SGK kesintisi uygulanır)
        </label>
      </div>
      <button className="zk-btn zk-btn-primary" style={{ width: '100%', justifyContent: 'center' }} onClick={submit}>Kaydet</button>
    </Modal>
  );
}

function FarmersTab({ farmers, setFarmers, purchases, payments, setTab, setSelectedFarmerId }) {
  const [showAdd, setShowAdd] = useState(false);
  const [query, setQuery] = useState('');

  const balances = useMemo(() => {
    const map = {};
    farmers.forEach((f) => { map[f.id] = 0; });
    purchases.forEach((p) => { map[p.farmerId] = (map[p.farmerId] || 0) + p.netPayment; });
    payments.forEach((pay) => { map[pay.farmerId] = (map[pay.farmerId] || 0) - pay.amount; });
    return map;
  }, [farmers, purchases, payments]);

  const filtered = farmers.filter((f) => f.name.toLowerCase().includes(query.toLowerCase()));

  const addFarmer = async (data) => {
    const newFarmer = { id: uid(), ...data, createdAt: Date.now() };
    const next = [...farmers, newFarmer];
    setFarmers(next);
    await storageSet('zk:farmers', next);
    setShowAdd(false);
  };

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
        <div>
          <div className="zk-h1">Çiftçiler</div>
          <div className="zk-h1-sub">{farmers.length} kayıtlı çiftçi</div>
        </div>
        <button className="zk-btn zk-btn-primary" onClick={() => setShowAdd(true)}><Plus size={14} /> Çiftçi ekle</button>
      </div>

      <div style={{ position: 'relative', marginBottom: 14, maxWidth: 320 }}>
        <Search size={15} style={{ position: 'absolute', left: 10, top: 9, color: COLORS.inkSoft }} />
        <input className="zk-input" style={{ paddingLeft: 32 }} placeholder="Çiftçi ara..." value={query} onChange={(e) => setQuery(e.target.value)} />
      </div>

      <div className="zk-card">
        {filtered.length === 0 ? (
          <div className="zk-empty"><Users size={26} className="zk-empty-icon" /><br/>{farmers.length === 0 ? 'Henüz çiftçi eklenmedi.' : 'Sonuç bulunamadı.'}</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            {filtered.map((f) => {
              const bal = balances[f.id] || 0;
              return (
                <div key={f.id} className="zk-farmer-row" onClick={() => { setSelectedFarmerId(f.id); setTab('ledger'); }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                    <div className="zk-avatar">{f.name.split(' ').map((w) => w[0]).slice(0, 2).join('').toUpperCase()}</div>
                    <div>
                      <div style={{ fontWeight: 600, fontSize: 13.5, display: 'flex', alignItems: 'center', gap: 7 }}>
                        {f.name}
                        {f.bagkurStatus && <span className="zk-tag">BAĞ-KUR</span>}
                      </div>
                      <div style={{ fontSize: 11.5, color: COLORS.inkSoft, display: 'flex', alignItems: 'center', gap: 10, marginTop: 2 }}>
                        {f.phone && <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><Phone size={11} />{f.phone}</span>}
                        {f.tcNo && <span>TC: {f.tcNo}</span>}
                      </div>
                    </div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <span className={`zk-badge ${bal > 0 ? 'zk-badge-red' : 'zk-badge-olive'}`}>
                      {bal > 0 ? `${fmtTL(bal)} ödenecek` : 'Bakiye kapalı'}
                    </span>
                    <ChevronRight size={16} color={COLORS.inkSoft} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {showAdd && <AddFarmerModal onClose={() => setShowAdd(false)} onSave={addFarmer} />}
    </div>
  );
}

function PurchaseTab({ farmers, setFarmers, purchases, setPurchases, onPrintReceipt, settings, priceList }) {
  const [farmerId, setFarmerId] = useState('');
  const [showAddFarmer, setShowAddFarmer] = useState(false);
  const [date, setDate] = useState(todayStr());
  const [commissionRate, setCommissionRate] = useState('3');
  const [borsaTescilli, setBorsaTescilli] = useState(false);
  const [noDeduction, setNoDeduction] = useState(false);
  const [note, setNote] = useState('');
  const [applyBagkur, setApplyBagkur] = useState(false);
  const [bagkurRate, setBagkurRate] = useState('1');
  const [lastSaved, setLastSaved] = useState(null);

  const [items, setItems] = useState([]);
  const [lineVariety, setLineVariety] = useState(priceList[0]?.name || '');
  const [lineGradeName, setLineGradeName] = useState('');
  const [lineKg, setLineKg] = useState('');
  const [linePrice, setLinePrice] = useState('');

  const farmer = farmers.find((f) => f.id === farmerId);
  const selectedVariety = priceList.find((v) => v.name === lineVariety);

  useEffect(() => {
    setApplyBagkur(farmer ? !!farmer.bagkurStatus : false);
  }, [farmerId]);

  useEffect(() => {
    if (!selectedVariety) return;
    if (selectedVariety.hasGrades) {
      const firstGrade = selectedVariety.grades[0];
      setLineGradeName(firstGrade ? firstGrade.name : '');
      setLinePrice(firstGrade ? firstGrade.price.toString() : '');
    } else {
      setLineGradeName('');
      setLinePrice((selectedVariety.singlePrice || 0).toString());
    }
  }, [lineVariety, priceList]);

  useEffect(() => {
    if (!selectedVariety || !selectedVariety.hasGrades) return;
    const g = selectedVariety.grades.find((x) => x.name === lineGradeName);
    if (g) setLinePrice(g.price.toString());
  }, [lineGradeName]);

  const handleFarmerSelect = (value) => {
    if (value === '__add_new__') { setShowAddFarmer(true); return; }
    setFarmerId(value);
  };

  const saveNewFarmer = async (data) => {
    const newFarmer = { id: uid(), ...data, createdAt: Date.now() };
    const next = [...farmers, newFarmer];
    setFarmers(next);
    await storageSet('zk:farmers', next);
    setShowAddFarmer(false);
    setFarmerId(newFarmer.id);
  };

  const addLine = () => {
    const kgVal = parseFloat(lineKg);
    const priceVal = parseFloat(linePrice);
    if (!lineVariety || !kgVal || kgVal <= 0 || !priceVal || priceVal <= 0) return;
    const label = lineGradeName ? `${lineVariety} · ${lineGradeName}` : lineVariety;
    setItems((prev) => [...prev, { id: uid(), grade: label, kg: kgVal, pricePerKg: priceVal, amount: kgVal * priceVal }]);
    setLineKg('');
  };

  const removeLine = (id) => setItems((prev) => prev.filter((x) => x.id !== id));

  const netKg = items.reduce((s, i) => s + i.kg, 0);
  const amount = items.reduce((s, i) => s + i.amount, 0);
  const commissionAmount = noDeduction ? 0 : amount * ((parseFloat(commissionRate) || 0) / 100);
  const stopajOrani = stopajOraniHesapla(borsaTescilli);
  const stopajTutari = noDeduction ? 0 : amount * (stopajOrani / 100);
  const bagkurTutari = (!noDeduction && applyBagkur) ? amount * ((parseFloat(bagkurRate) || 0) / 100) : 0;
  const netPayment = amount - commissionAmount - stopajTutari - bagkurTutari;

  const canSave = farmerId && items.length > 0;

  const save = async () => {
    if (!canSave) return;
    const record = {
      id: uid(),
      makbuzNo: purchases.length + 1,
      farmerId,
      date,
      items,
      netKg,
      noDeduction,
      commissionRate: noDeduction ? 0 : (parseFloat(commissionRate) || 0),
      commissionAmount,
      borsaTescilli,
      stopajOrani: noDeduction ? 0 : stopajOrani,
      stopajTutari,
      applyBagkur: noDeduction ? false : applyBagkur,
      bagkurRate: parseFloat(bagkurRate) || 0,
      bagkurTutari,
      amount,
      netPayment,
      note,
      createdAt: Date.now(),
    };
    const next = [...purchases, record];
    setPurchases(next);
    await storageSet('zk:purchases', next);
    setLastSaved(record);
    setItems([]); setNote(''); setLineKg('');
  };

  return (
    <div>
      <div className="zk-h1">Yeni alım</div>
      <div className="zk-h1-sub">Elekten çıkan her sınıfı ayrı tartıp ekleyin, sonunda toplam üzerinden hesaplansın</div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.3fr 1fr', gap: 18, alignItems: 'start' }}>
        <div className="zk-card">
          <div className="zk-grid" style={{ gridTemplateColumns: '1fr 1fr', marginBottom: 14 }}>
            <div>
              <label className="zk-label">Çiftçi</label>
              <select className="zk-select" value={farmerId} onChange={(e) => handleFarmerSelect(e.target.value)}>
                <option value="">Seçin...</option>
                {farmers.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
                <option value="__add_new__">+ Yeni çiftçi ekle</option>
              </select>
            </div>
            <div>
              <label className="zk-label">Tarih</label>
              <input className="zk-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
            </div>
          </div>

          <div style={{ background: COLORS.paper, borderRadius: 10, padding: 12, marginBottom: 14 }}>
            <div className="zk-label" style={{ marginBottom: 8 }}>Tartım satırı ekle</div>
            <div className="zk-grid" style={{ gridTemplateColumns: selectedVariety?.hasGrades ? '1fr 0.9fr 0.7fr 0.7fr auto' : '1.4fr 0.7fr 0.7fr auto', gap: 8, alignItems: 'end' }}>
              <div>
                <label className="zk-label">Tür</label>
                <select className="zk-select" value={lineVariety} onChange={(e) => setLineVariety(e.target.value)}>
                  {priceList.length === 0 && <option value="">Fiyat listesi boş</option>}
                  {priceList.map((v) => <option key={v.id} value={v.name}>{v.name}</option>)}
                </select>
              </div>
              {selectedVariety?.hasGrades && (
                <div>
                  <label className="zk-label">Numara</label>
                  <select className="zk-select" value={lineGradeName} onChange={(e) => setLineGradeName(e.target.value)}>
                    {selectedVariety.grades.length === 0 && <option value="">Numara yok</option>}
                    {selectedVariety.grades.map((g) => <option key={g.id} value={g.name}>{g.name}</option>)}
                  </select>
                </div>
              )}
              <div>
                <label className="zk-label">Kg</label>
                <input className="zk-input" type="number" value={lineKg} onChange={(e) => setLineKg(e.target.value)} placeholder="0" />
              </div>
              <div>
                <label className="zk-label">Fiyat/kg</label>
                <input className="zk-input" type="number" value={linePrice} onChange={(e) => setLinePrice(e.target.value)} placeholder="0.00" />
              </div>
              <button className="zk-btn zk-btn-gold" onClick={addLine} style={{ height: 36 }}><Plus size={14} /> Ekle</button>
            </div>
          </div>

          {items.length > 0 && (
            <table className="zk-table" style={{ marginBottom: 14 }}>
              <thead><tr><th>Sınıf</th><th>Kg</th><th>Fiyat</th><th>Tutar</th><th></th></tr></thead>
              <tbody>
                {items.map((it) => (
                  <tr key={it.id}>
                    <td><span className="zk-badge zk-badge-blue">{it.grade}</span></td>
                    <td>{fmtKg(it.kg)}</td>
                    <td>{fmtTL(it.pricePerKg)}/kg</td>
                    <td>{fmtTL(it.amount)}</td>
                    <td><button className="zk-btn zk-btn-secondary" style={{ padding: '4px 8px' }} onClick={() => removeLine(it.id)}><X size={12} /></button></td>
                  </tr>
                ))}
                <tr>
                  <td style={{ fontWeight: 700 }}>Toplam</td>
                  <td style={{ fontWeight: 700 }}>{fmtKg(netKg)}</td>
                  <td></td>
                  <td style={{ fontWeight: 700 }}>{fmtTL(amount)}</td>
                  <td></td>
                </tr>
              </tbody>
            </table>
          )}

          <label className="zk-checkbox-row" style={{ background: COLORS.oliveSoft, padding: '9px 12px', borderRadius: 8, marginBottom: 12 }}>
            <input type="checkbox" checked={noDeduction} onChange={(e) => setNoDeduction(e.target.checked)} />
            Kesintisiz hesapla (komisyon / stopaj / BAĞ-KUR uygulanmasın, tutarın tamamı ödensin)
          </label>

          {!noDeduction && (
            <div className="zk-grid" style={{ gridTemplateColumns: '1fr 1fr', marginBottom: 12 }}>
              <div>
                <label className="zk-label">Komisyon oranı (%)</label>
                <input className="zk-input" type="number" value={commissionRate} onChange={(e) => setCommissionRate(e.target.value)} placeholder="3" />
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'flex-end', gap: 6 }}>
                <label className="zk-checkbox-row">
                  <input type="checkbox" checked={borsaTescilli} onChange={(e) => setBorsaTescilli(e.target.checked)} />
                  Ticaret borsasına tescilli
                </label>
                <label className="zk-checkbox-row">
                  <input type="checkbox" checked={applyBagkur} onChange={(e) => setApplyBagkur(e.target.checked)} />
                  BAĞ-KUR kesintisi uygula (%)
                  <input
                    className="zk-input"
                    type="number"
                    value={bagkurRate}
                    onChange={(e) => setBagkurRate(e.target.value)}
                    style={{ width: 55, padding: '4px 6px' }}
                    disabled={!applyBagkur}
                  />
                </label>
              </div>
            </div>
          )}

          <div style={{ marginBottom: 16 }}>
            <label className="zk-label">Not (opsiyonel)</label>
            <input className="zk-input" value={note} onChange={(e) => setNote(e.target.value)} placeholder="serbest not..." />
          </div>

          <div style={{ background: COLORS.paper, borderRadius: 10, padding: 14, marginBottom: 16, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, fontSize: 12.5 }}>
            <div>Ürün tutarı ({fmtKg(netKg)})</div><div style={{ textAlign: 'right', fontWeight: 600 }}>{fmtTL(amount)}</div>
            {!noDeduction && (
              <>
                <div style={{ color: COLORS.gold }}>Komisyon (%{commissionRate || 0})</div><div style={{ textAlign: 'right', fontWeight: 600, color: COLORS.gold }}>− {fmtTL(commissionAmount)}</div>
                <div style={{ color: COLORS.blue }}>Stopaj (%{stopajOrani})</div><div style={{ textAlign: 'right', fontWeight: 600, color: COLORS.blue }}>− {fmtTL(stopajTutari)}</div>
                {applyBagkur && (<><div style={{ color: COLORS.red }}>BAĞ-KUR (%{bagkurRate || 0})</div><div style={{ textAlign: 'right', fontWeight: 600, color: COLORS.red }}>− {fmtTL(bagkurTutari)}</div></>)}
              </>
            )}
            <div style={{ fontWeight: 700, borderTop: `1px solid ${COLORS.border}`, paddingTop: 8 }}>Çiftçiye ödenecek</div>
            <div style={{ textAlign: 'right', fontWeight: 700, borderTop: `1px solid ${COLORS.border}`, paddingTop: 8, color: COLORS.olive }}>{fmtTL(netPayment)}</div>
          </div>

          <button className="zk-btn zk-btn-primary" style={{ width: '100%', justifyContent: 'center' }} disabled={!canSave} onClick={save}>
            Alımı kaydet
          </button>

          {lastSaved && (
            <div style={{ marginTop: 12, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: COLORS.oliveSoft, padding: '10px 12px', borderRadius: 8 }}>
              <span style={{ fontSize: 12, color: COLORS.olive }}>Alım #{lastSaved.makbuzNo} kaydedildi.</span>
              <button className="zk-btn zk-btn-secondary" onClick={() => onPrintReceipt(lastSaved)}><Printer size={13} /> Müstahsil makbuzu yazdır</button>
            </div>
          )}
        </div>

        <ScaleWidget onWeightCapture={(v) => setLineKg(v.toFixed(1))} />
      </div>

      {showAddFarmer && <AddFarmerModal onClose={() => setShowAddFarmer(false)} onSave={saveNewFarmer} />}
    </div>
  );
}

function WarehouseTab({ purchases, buyers, setBuyers, sales, setSales }) {
  const [showAddBuyer, setShowAddBuyer] = useState(false);
  const [buyerName, setBuyerName] = useState('');
  const [buyerPhone, setBuyerPhone] = useState('');

  const [buyerId, setBuyerId] = useState('');
  const [date, setDate] = useState(todayStr());
  const [grade, setGrade] = useState('');
  const [kg, setKg] = useState('');
  const [pricePerKg, setPricePerKg] = useState('');
  const [note, setNote] = useState('');

  const totalPurchasedKg = purchases.reduce((s, p) => s + p.netKg, 0);
  const totalSoldKg = sales.reduce((s, s2) => s + s2.kg, 0);
  const currentStock = totalPurchasedKg - totalSoldKg;

  const purchasedByGrade = useMemo(() => {
    const map = {};
    purchases.forEach((p) => {
      (p.items || []).forEach((it) => { map[it.grade] = (map[it.grade] || 0) + it.kg; });
    });
    return map;
  }, [purchases]);

  const soldByGrade = useMemo(() => {
    const map = {};
    sales.forEach((s) => { map[s.grade || 'Etiketsiz'] = (map[s.grade || 'Etiketsiz'] || 0) + s.kg; });
    return map;
  }, [sales]);

  const stockByGrade = useMemo(() => {
    const grades = new Set([...Object.keys(purchasedByGrade), ...Object.keys(soldByGrade)]);
    return Array.from(grades).map((g) => ({ grade: g, stock: (purchasedByGrade[g] || 0) - (soldByGrade[g] || 0) })).sort((a, b) => b.stock - a.stock);
  }, [purchasedByGrade, soldByGrade]);

  const availableForGrade = grade ? (purchasedByGrade[grade] || 0) - (soldByGrade[grade] || 0) : 0;
  const amount = (parseFloat(kg) || 0) * (parseFloat(pricePerKg) || 0);
  const canSave = buyerId && grade && parseFloat(kg) > 0 && parseFloat(pricePerKg) > 0 && parseFloat(kg) <= availableForGrade + 0.001;

  const addBuyer = async () => {
    if (!buyerName.trim()) return;
    const b = { id: uid(), name: buyerName.trim(), phone: buyerPhone.trim(), createdAt: Date.now() };
    const next = [...buyers, b];
    setBuyers(next);
    await storageSet('zk:buyers', next);
    setBuyerName(''); setBuyerPhone(''); setShowAddBuyer(false);
    setBuyerId(b.id);
  };

  const saveSale = async () => {
    if (!canSave) return;
    const record = { id: uid(), buyerId, date, grade, kg: parseFloat(kg), pricePerKg: parseFloat(pricePerKg), amount, note, createdAt: Date.now() };
    const next = [...sales, record];
    setSales(next);
    await storageSet('zk:sales', next);
    setKg(''); setPricePerKg(''); setNote('');
  };

  const recentSales = [...sales].sort((a, b) => b.createdAt - a.createdAt).slice(0, 8);

  return (
    <div>
      <div className="zk-h1">Depo ve satış</div>
      <div className="zk-h1-sub">Sınıf bazında stok durumu ve alıcıya satış kayıtları</div>

      <div className="zk-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(150px,1fr))', marginBottom: 18 }}>
        <StatCard label="Toplam alınan" value={fmtKg(totalPurchasedKg)} icon={Package} />
        <StatCard label="Toplam satılan" value={fmtKg(totalSoldKg)} icon={ShoppingCart} />
        <StatCard label="Mevcut stok" value={fmtKg(currentStock)} tone={COLORS.blue} icon={Warehouse} />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: 16, alignItems: 'start' }}>
        <div className="zk-card">
          <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 12 }}>Yeni satış</div>
          <div className="zk-grid" style={{ gridTemplateColumns: '1fr 1fr', marginBottom: 10 }}>
            <div>
              <label className="zk-label">Alıcı</label>
              <div style={{ display: 'flex', gap: 6 }}>
                <select className="zk-select" value={buyerId} onChange={(e) => setBuyerId(e.target.value)}>
                  <option value="">Seçin...</option>
                  {buyers.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
                </select>
                <button className="zk-btn zk-btn-secondary" onClick={() => setShowAddBuyer(true)}><Plus size={13} /></button>
              </div>
            </div>
            <div>
              <label className="zk-label">Tarih</label>
              <input className="zk-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
            </div>
          </div>
          <div style={{ marginBottom: 10 }}>
            <label className="zk-label">Sınıf / numara</label>
            <select className="zk-select" value={grade} onChange={(e) => setGrade(e.target.value)}>
              <option value="">Seçin...</option>
              {stockByGrade.filter((g) => g.stock > 0.01).map((g) => <option key={g.grade} value={g.grade}>{g.grade} · stokta {fmtKg(g.stock)}</option>)}
            </select>
          </div>
          <div className="zk-grid" style={{ gridTemplateColumns: '1fr 1fr', marginBottom: 10 }}>
            <div>
              <label className="zk-label">Miktar (kg){grade ? ` · stokta ${fmtKg(availableForGrade)}` : ''}</label>
              <input className="zk-input" type="number" value={kg} onChange={(e) => setKg(e.target.value)} placeholder="0" />
            </div>
            <div>
              <label className="zk-label">Kg fiyatı (TL)</label>
              <input className="zk-input" type="number" value={pricePerKg} onChange={(e) => setPricePerKg(e.target.value)} placeholder="0.00" />
            </div>
          </div>
          <div style={{ marginBottom: 12 }}>
            <label className="zk-label">Not</label>
            <input className="zk-input" value={note} onChange={(e) => setNote(e.target.value)} placeholder="opsiyonel" />
          </div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12, fontSize: 13 }}>
            <span style={{ color: COLORS.inkSoft }}>Toplam tutar</span>
            <span style={{ fontWeight: 700 }}>{fmtTL(amount)}</span>
          </div>
          <button className="zk-btn zk-btn-primary" style={{ width: '100%', justifyContent: 'center' }} disabled={!canSave} onClick={saveSale}>
            Satışı kaydet
          </button>
          {grade && parseFloat(kg) > availableForGrade && kg && (
            <div style={{ fontSize: 11.5, color: COLORS.red, marginTop: 8 }}>Bu sınıfta stoktan fazla miktar girdiniz.</div>
          )}
        </div>

        <div className="zk-card">
          <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Sınıf bazında stok</div>
          {stockByGrade.length === 0 ? (
            <div className="zk-empty">Kayıt yok.</div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {stockByGrade.map(({ grade: g, stock }) => (
                <div key={g} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12.5 }}>
                  <span className="zk-badge zk-badge-blue">{g}</span>
                  <span style={{ fontWeight: 600, color: stock < 0 ? COLORS.red : COLORS.ink }}>{fmtKg(stock)}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="zk-card" style={{ marginTop: 16 }}>
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Son satışlar</div>
        {recentSales.length === 0 ? (
          <div className="zk-empty">Henüz satış kaydı yok.</div>
        ) : (
          <table className="zk-table">
            <thead><tr><th>Tarih</th><th>Alıcı</th><th>Sınıf</th><th>Kg</th><th>Fiyat</th><th>Tutar</th></tr></thead>
            <tbody>
              {recentSales.map((s) => {
                const b = buyers.find((x) => x.id === s.buyerId);
                return (
                  <tr key={s.id}>
                    <td>{fmtDate(s.date)}</td>
                    <td>{b ? b.name : '—'}</td>
                    <td><span className="zk-badge zk-badge-blue">{s.grade || '—'}</span></td>
                    <td>{fmtKg(s.kg)}</td>
                    <td>{fmtTL(s.pricePerKg)}/kg</td>
                    <td>{fmtTL(s.amount)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {showAddBuyer && (
        <Modal title="Yeni alıcı ekle" onClose={() => setShowAddBuyer(false)}>
          <div style={{ marginBottom: 12 }}>
            <label className="zk-label">Alıcı / firma adı</label>
            <input className="zk-input" value={buyerName} onChange={(e) => setBuyerName(e.target.value)} placeholder="örn. Ege Zeytinyağı A.Ş." autoFocus />
          </div>
          <div style={{ marginBottom: 18 }}>
            <label className="zk-label">Telefon (opsiyonel)</label>
            <input className="zk-input" value={buyerPhone} onChange={(e) => setBuyerPhone(e.target.value)} />
          </div>
          <button className="zk-btn zk-btn-primary" style={{ width: '100%', justifyContent: 'center' }} onClick={addBuyer}>Kaydet</button>
        </Modal>
      )}
    </div>
  );
}

function LedgerTab({ farmers, purchases, payments, setPayments, selectedFarmerId, setSelectedFarmerId, onPrintReceipt }) {
  const [payAmount, setPayAmount] = useState('');
  const [payNote, setPayNote] = useState('');

  const farmer = farmers.find((f) => f.id === selectedFarmerId);

  const entries = useMemo(() => {
    if (!farmer) return [];
    const p = purchases.filter((x) => x.farmerId === farmer.id).map((x) => ({ type: 'purchase', date: x.date, createdAt: x.createdAt, amount: x.netPayment, data: x }));
    const pay = payments.filter((x) => x.farmerId === farmer.id).map((x) => ({ type: 'payment', date: x.date, createdAt: x.createdAt, amount: -x.amount, data: x }));
    return [...p, ...pay].sort((a, b) => a.createdAt - b.createdAt);
  }, [farmer, purchases, payments]);

  let running = 0;
  const withRunning = entries.map((e) => { running += e.amount; return { ...e, running }; });
  const balance = running;

  const addPayment = async () => {
    const amt = parseFloat(payAmount);
    if (!farmer || !amt || amt <= 0) return;
    const record = { id: uid(), farmerId: farmer.id, date: todayStr(), amount: amt, note: payNote, createdAt: Date.now() };
    const next = [...payments, record];
    setPayments(next);
    await storageSet('zk:payments', next);
    setPayAmount(''); setPayNote('');
  };

  if (!farmer) {
    return (
      <div>
        <div className="zk-h1">Cari hesap</div>
        <div className="zk-h1-sub">Görüntülemek için bir çiftçi seçin</div>
        <div className="zk-card">
          <select className="zk-select" value="" onChange={(e) => setSelectedFarmerId(e.target.value)}>
            <option value="">Çiftçi seçin...</option>
            {farmers.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
          </select>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
        <div>
          <div className="zk-h1">{farmer.name}</div>
          <div className="zk-h1-sub">Cari hesap özeti</div>
        </div>
        <select className="zk-select" style={{ width: 200 }} value={selectedFarmerId} onChange={(e) => setSelectedFarmerId(e.target.value)}>
          {farmers.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
        </select>
      </div>

      <div className="zk-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(150px,1fr))', marginBottom: 18 }}>
        <StatCard label="Güncel bakiye" value={fmtTL(Math.abs(balance))} tone={balance > 0 ? COLORS.red : COLORS.olive} />
        <StatCard label="Durum" value={balance > 0 ? 'Ödenecek' : 'Kapalı'} />
        <StatCard label="Toplam işlem" value={entries.length} />
      </div>

      <div className="zk-card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Ödeme ekle</div>
        <div style={{ display: 'flex', gap: 10 }}>
          <input className="zk-input" type="number" placeholder="Tutar (TL)" value={payAmount} onChange={(e) => setPayAmount(e.target.value)} style={{ maxWidth: 160 }} />
          <input className="zk-input" placeholder="Not (opsiyonel)" value={payNote} onChange={(e) => setPayNote(e.target.value)} />
          <button className="zk-btn zk-btn-primary" onClick={addPayment}>Ekle</button>
        </div>
      </div>

      <div className="zk-card">
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Hareketler</div>
        {withRunning.length === 0 ? (
          <div className="zk-empty">Henüz hareket yok.</div>
        ) : (
          <table className="zk-table">
            <thead><tr><th>Tarih</th><th>İşlem</th><th>Tutar</th><th>Bakiye</th><th></th></tr></thead>
            <tbody>
              {withRunning.slice().reverse().map((e, i) => (
                <tr key={i}>
                  <td>{fmtDate(e.date)}</td>
                  <td>
                    {e.type === 'purchase'
                      ? <span className="zk-badge zk-badge-olive">Alım · {fmtKg(e.data.netKg)}</span>
                      : <span className="zk-badge zk-badge-gold">Ödeme{e.data.note ? ` · ${e.data.note}` : ''}</span>}
                  </td>
                  <td style={{ color: e.amount >= 0 ? COLORS.olive : COLORS.gold, fontWeight: 600 }}>
                    {e.amount >= 0 ? '+' : ''}{fmtTL(e.amount)}
                  </td>
                  <td style={{ fontWeight: 600 }}>{fmtTL(e.running)}</td>
                  <td>
                    {e.type === 'purchase' && (
                      <button className="zk-btn zk-btn-secondary" style={{ padding: '5px 9px' }} onClick={() => onPrintReceipt(e.data)}><Printer size={12} /></button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function ReportsTab({ farmers, purchases, sales, buyers }) {
  const [range, setRange] = useState('month');

  const filtered = useMemo(() => {
    const now = new Date();
    return purchases.filter((p) => {
      const d = new Date(p.date);
      if (range === 'today') return p.date === todayStr();
      if (range === 'month') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
      return true;
    });
  }, [purchases, range]);

  const filteredSales = useMemo(() => {
    const now = new Date();
    return sales.filter((s) => {
      const d = new Date(s.date);
      if (range === 'today') return s.date === todayStr();
      if (range === 'month') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
      return true;
    });
  }, [sales, range]);

  const totalKg = filtered.reduce((s, p) => s + p.netKg, 0);
  const totalAmount = filtered.reduce((s, p) => s + p.amount, 0);
  const totalCommission = filtered.reduce((s, p) => s + p.commissionAmount, 0);
  const totalStopaj = filtered.reduce((s, p) => s + (p.stopajTutari || 0), 0);
  const totalPayable = filtered.reduce((s, p) => s + p.netPayment, 0);
  const totalSalesAmount = filteredSales.reduce((s, s2) => s + s2.amount, 0);

  const byFarmer = useMemo(() => {
    const map = {};
    filtered.forEach((p) => {
      if (!map[p.farmerId]) map[p.farmerId] = { kg: 0, amount: 0, count: 0 };
      map[p.farmerId].kg += p.netKg;
      map[p.farmerId].amount += p.netPayment;
      map[p.farmerId].count += 1;
    });
    return Object.entries(map).map(([farmerId, v]) => ({ farmerId, ...v })).sort((a, b) => b.kg - a.kg);
  }, [filtered]);

  const chartData = byFarmer.slice(0, 8).map((row) => {
    const f = farmers.find((x) => x.id === row.farmerId);
    return { name: f ? f.name.split(' ')[0] : '—', kg: Math.round(row.kg * 10) / 10 };
  });

  const byGrade = useMemo(() => {
    const map = {};
    filtered.forEach((p) => {
      (p.items || []).forEach((it) => {
        if (!map[it.grade]) map[it.grade] = { kg: 0, amount: 0 };
        map[it.grade].kg += it.kg;
        map[it.grade].amount += it.amount;
      });
    });
    return Object.entries(map).map(([grade, v]) => ({ grade, ...v })).sort((a, b) => b.kg - a.kg);
  }, [filtered]);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    const purchaseRows = filtered.map((p) => {
      const f = farmers.find((x) => x.id === p.farmerId);
      return {
        'Makbuz No': p.makbuzNo,
        'Tarih': p.date,
        'Çiftçi': f ? f.name : '',
        'TC No': f ? f.tcNo : '',
        'Net kg': p.netKg,
        'Tutar': p.amount,
        'Komisyon %': p.commissionRate,
        'Komisyon tutarı': p.commissionAmount,
        'Stopaj %': p.stopajOrani,
        'Stopaj tutarı': p.stopajTutari,
        'BAĞ-KUR tutarı': p.bagkurTutari || 0,
        'Net ödenen': p.netPayment,
      };
    });
    const wsPurchases = XLSX.utils.json_to_sheet(purchaseRows);
    XLSX.utils.book_append_sheet(wb, wsPurchases, 'Alimlar');

    const detailRows = [];
    filtered.forEach((p) => {
      const f = farmers.find((x) => x.id === p.farmerId);
      (p.items || []).forEach((it) => {
        detailRows.push({
          'Makbuz No': p.makbuzNo,
          'Tarih': p.date,
          'Çiftçi': f ? f.name : '',
          'Sınıf': it.grade,
          'Kg': it.kg,
          'Kg fiyatı': it.pricePerKg,
          'Tutar': it.amount,
        });
      });
    });
    const wsDetail = XLSX.utils.json_to_sheet(detailRows);
    XLSX.utils.book_append_sheet(wb, wsDetail, 'Alim Detay (Sinif)');

    const salesRows = filteredSales.map((s) => {
      const b = buyers.find((x) => x.id === s.buyerId);
      return { 'Tarih': s.date, 'Alıcı': b ? b.name : '', 'Sınıf': s.grade || '', 'Kg': s.kg, 'Kg fiyatı': s.pricePerKg, 'Tutar': s.amount, 'Not': s.note || '' };
    });
    const wsSales = XLSX.utils.json_to_sheet(salesRows);
    XLSX.utils.book_append_sheet(wb, wsSales, 'Satislar');

    const summaryRows = byFarmer.map((row) => {
      const f = farmers.find((x) => x.id === row.farmerId);
      return { 'Çiftçi': f ? f.name : '', 'İşlem sayısı': row.count, 'Toplam kg': row.kg, 'Tutar': row.amount };
    });
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Ciftci Ozeti');

    XLSX.writeFile(wb, `zeytin-rapor-${todayStr()}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
        <div>
          <div className="zk-h1">Raporlar</div>
          <div className="zk-h1-sub">Toplam alım, satış ve kesinti özeti</div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <select className="zk-select" style={{ width: 130 }} value={range} onChange={(e) => setRange(e.target.value)}>
            <option value="today">Bugün</option>
            <option value="month">Bu ay</option>
            <option value="all">Tümü</option>
          </select>
          <button className="zk-btn zk-btn-blue" onClick={exportExcel}><Download size={13} /> Excel</button>
        </div>
      </div>

      <div className="zk-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(150px,1fr))', marginBottom: 16 }}>
        <StatCard label="Toplam kg" value={fmtKg(totalKg)} />
        <StatCard label="Ürün tutarı" value={fmtTL(totalAmount)} />
        <StatCard label="Komisyon" value={fmtTL(totalCommission)} tone={COLORS.gold} />
        <StatCard label="Stopaj" value={fmtTL(totalStopaj)} tone={COLORS.blue} />
        <StatCard label="Çiftçilere ödenen" value={fmtTL(totalPayable)} tone={COLORS.olive} />
        <StatCard label="Satış tutarı" value={fmtTL(totalSalesAmount)} />
      </div>

      <div className="zk-card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Çiftçi bazında kg dağılımı</div>
        {chartData.length === 0 ? (
          <div className="zk-empty">Bu aralıkta kayıt yok.</div>
        ) : (
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#EFEBDD" />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: COLORS.inkSoft }} />
                <YAxis tick={{ fontSize: 10, fill: COLORS.inkSoft }} width={40} />
                <Tooltip formatter={(v) => [v + ' kg', 'Alım']} contentStyle={{ fontSize: 12, borderRadius: 8 }} />
                <Bar dataKey="kg" fill={COLORS.olive} radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      <div className="zk-card" style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Sınıf / numara bazında toplam</div>
        {byGrade.length === 0 ? (
          <div className="zk-empty">Bu aralıkta kayıt yok.</div>
        ) : (
          <table className="zk-table">
            <thead><tr><th>Sınıf</th><th>Toplam kg</th><th>Tutar</th></tr></thead>
            <tbody>
              {byGrade.map((row) => (
                <tr key={row.grade}>
                  <td><span className="zk-badge zk-badge-blue">{row.grade}</span></td>
                  <td>{fmtKg(row.kg)}</td>
                  <td>{fmtTL(row.amount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div className="zk-card">
        <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 10 }}>Çiftçi bazında dağılım (tablo)</div>
        {byFarmer.length === 0 ? (
          <div className="zk-empty">Bu aralıkta kayıt yok.</div>
        ) : (
          <table className="zk-table">
            <thead><tr><th>Çiftçi</th><th>İşlem sayısı</th><th>Toplam kg</th><th>Tutar</th></tr></thead>
            <tbody>
              {byFarmer.map((row) => {
                const f = farmers.find((x) => x.id === row.farmerId);
                return (
                  <tr key={row.farmerId}>
                    <td>{f ? f.name : '—'}</td>
                    <td>{row.count}</td>
                    <td>{fmtKg(row.kg)}</td>
                    <td>{fmtTL(row.amount)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function VarietyEditor({ variety, onChange, onRemove }) {
  const [newGradeName, setNewGradeName] = useState('');
  const [newGradePrice, setNewGradePrice] = useState('');

  const toggleHasGrades = () => onChange({ ...variety, hasGrades: !variety.hasGrades });
  const setSinglePrice = (price) => onChange({ ...variety, singlePrice: parseFloat(price) || 0 });
  const updateGradePrice = (gradeId, price) => onChange({ ...variety, grades: variety.grades.map((g) => (g.id === gradeId ? { ...g, price: parseFloat(price) || 0 } : g)) });
  const removeGradeRow = (gradeId) => onChange({ ...variety, grades: variety.grades.filter((g) => g.id !== gradeId) });
  const addGradeRow = () => {
    if (!newGradeName.trim()) return;
    onChange({ ...variety, grades: [...variety.grades, { id: uid(), name: newGradeName.trim(), price: parseFloat(newGradePrice) || 0 }] });
    setNewGradeName(''); setNewGradePrice('');
  };

  return (
    <div style={{ border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 12, marginBottom: 10 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
        <span style={{ fontWeight: 700, fontSize: 13.5 }}>{variety.name}</span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <label className="zk-checkbox-row" style={{ fontSize: 11.5 }}>
            <input type="checkbox" checked={variety.hasGrades} onChange={toggleHasGrades} />
            Numaraya ayrılıyor
          </label>
          <button className="zk-btn zk-btn-secondary" style={{ padding: '4px 8px' }} onClick={onRemove}><X size={12} /></button>
        </div>
      </div>

      {!variety.hasGrades ? (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 12.5, color: COLORS.inkSoft }}>Kg fiyatı</span>
          <input className="zk-input" type="number" value={variety.singlePrice || 0} onChange={(e) => setSinglePrice(e.target.value)} style={{ width: 100 }} />
        </div>
      ) : (
        <>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 8 }}>
            {variety.grades.map((g) => (
              <div key={g.id} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ flex: 1, fontSize: 12.5 }}>{g.name}</span>
                <input className="zk-input" type="number" value={g.price} onChange={(e) => updateGradePrice(g.id, e.target.value)} style={{ width: 85 }} />
                <button className="zk-btn zk-btn-secondary" style={{ padding: '4px 7px' }} onClick={() => removeGradeRow(g.id)}><X size={11} /></button>
              </div>
            ))}
            {variety.grades.length === 0 && <div style={{ fontSize: 11.5, color: COLORS.inkSoft }}>Henüz numara eklenmedi.</div>}
          </div>
          <div style={{ display: 'flex', gap: 6 }}>
            <input className="zk-input" value={newGradeName} onChange={(e) => setNewGradeName(e.target.value)} placeholder="örn. 1 Numara" style={{ fontSize: 12.5 }} />
            <input className="zk-input" type="number" value={newGradePrice} onChange={(e) => setNewGradePrice(e.target.value)} placeholder="Fiyat" style={{ width: 80 }} />
            <button className="zk-btn zk-btn-secondary" style={{ padding: '6px 10px' }} onClick={addGradeRow}><Plus size={12} /></button>
          </div>
        </>
      )}
    </div>
  );
}

function SettingsTab({ settings, setSettings, priceList, setPriceList }) {
  const [businessName, setBusinessName] = useState(settings.businessName || '');
  const [taxNo, setTaxNo] = useState(settings.taxNo || '');
  const [address, setAddress] = useState(settings.address || '');
  const [newVarietyName, setNewVarietyName] = useState('');

  const save = async () => {
    const next = { businessName, taxNo, address };
    setSettings(next);
    await storageSet('zk:settings', next);
  };

  const updateVariety = async (updated) => {
    const next = priceList.map((v) => (v.id === updated.id ? updated : v));
    setPriceList(next);
    await storageSet('zk:priceList', next);
  };

  const removeVariety = async (id) => {
    const next = priceList.filter((v) => v.id !== id);
    setPriceList(next);
    await storageSet('zk:priceList', next);
  };

  const addVariety = async () => {
    if (!newVarietyName.trim()) return;
    const next = [...priceList, { id: uid(), name: newVarietyName.trim(), hasGrades: true, singlePrice: 0, grades: [] }];
    setPriceList(next);
    await storageSet('zk:priceList', next);
    setNewVarietyName('');
  };

  return (
    <div>
      <div className="zk-h1">Ayarlar</div>
      <div className="zk-h1-sub">İşletme bilgileri ve zeytin türü / fiyat listesi</div>

      <div style={{ display: 'grid', gridTemplateColumns: '0.8fr 1.2fr', gap: 16, alignItems: 'start', maxWidth: 920 }}>
        <div className="zk-card">
          <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 12 }}>İşletme bilgileri</div>
          <div style={{ marginBottom: 12 }}>
            <label className="zk-label">İşletme / komisyoncu adı</label>
            <input className="zk-input" value={businessName} onChange={(e) => setBusinessName(e.target.value)} placeholder="örn. Ahmet Yılmaz Zeytin Komisyonculuğu" />
          </div>
          <div style={{ marginBottom: 12 }}>
            <label className="zk-label">Vergi kimlik no</label>
            <input className="zk-input" value={taxNo} onChange={(e) => setTaxNo(e.target.value)} />
          </div>
          <div style={{ marginBottom: 18 }}>
            <label className="zk-label">Adres</label>
            <input className="zk-input" value={address} onChange={(e) => setAddress(e.target.value)} />
          </div>
          <button className="zk-btn zk-btn-primary" onClick={save}>Kaydet</button>
        </div>

        <div className="zk-card">
          <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 4 }}>Zeytin türleri ve fiyat listesi (bu hafta)</div>
          <div style={{ fontSize: 11.5, color: COLORS.inkSoft, marginBottom: 12 }}>Her tür için numaraya ayrılıp ayrılmadığını seçin. Yeni alım ekranında otomatik gelir.</div>

          {priceList.map((v) => (
            <VarietyEditor key={v.id} variety={v} onChange={updateVariety} onRemove={() => removeVariety(v.id)} />
          ))}
          {priceList.length === 0 && <div className="zk-empty">Henüz tür eklenmedi.</div>}

          <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
            <input className="zk-input" value={newVarietyName} onChange={(e) => setNewVarietyName(e.target.value)} placeholder="örn. Edremit" />
            <button className="zk-btn zk-btn-gold" onClick={addVariety}><Plus size={13} /> Tür ekle</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function PrintArea({ purchase, farmer, settings }) {
  if (!purchase || !farmer) return <div id="zk-print-area" />;
  return (
    <div id="zk-print-area">
      <div style={{ fontFamily: "'Courier New', monospace", width: '100%', fontSize: 11, lineHeight: 1.5 }}>
        <div style={{ textAlign: 'center', marginBottom: 6 }}>
          <div style={{ fontFamily: "'Fraunces', Georgia, serif", fontSize: 14, fontWeight: 600 }}>{settings.businessName || 'Zeytin Komisyonculuğu'}</div>
          {settings.address && <div style={{ fontSize: 10 }}>{settings.address}</div>}
          {settings.taxNo && <div style={{ fontSize: 10 }}>VKN: {settings.taxNo}</div>}
        </div>
        <div style={{ borderTop: '1px dashed #000', borderBottom: '1px dashed #000', padding: '4px 0', marginBottom: 6, textAlign: 'center', fontWeight: 700 }}>
          MÜSTAHSİL MAKBUZU No: {purchase.makbuzNo}
        </div>
        <div>Tarih: {fmtDate(purchase.date)}</div>
        <div>Satıcı: {farmer.name}</div>
        {farmer.tcNo && <div>TC No: {farmer.tcNo}</div>}
        {farmer.address && <div>Adres: {farmer.address}</div>}
        <div style={{ borderTop: '1px dashed #000', margin: '6px 0' }} />
        <div style={{ fontWeight: 700, marginBottom: 2 }}>Ürün: Zeytin</div>
        <table style={{ width: '100%', marginTop: 4 }}>
          <tbody>
            {purchase.items && purchase.items.map((it) => (
              <tr key={it.id}>
                <td>{it.grade}</td>
                <td style={{ textAlign: 'right' }}>{fmtKg(it.kg)}</td>
                <td style={{ textAlign: 'right' }}>×{fmtTL(it.pricePerKg)}</td>
                <td style={{ textAlign: 'right' }}>{fmtTL(it.amount)}</td>
              </tr>
            ))}
            <tr style={{ borderTop: '1px dashed #000' }}>
              <td style={{ fontWeight: 700, paddingTop: 3 }}>Toplam</td>
              <td style={{ textAlign: 'right', fontWeight: 700, paddingTop: 3 }}>{fmtKg(purchase.netKg)}</td>
              <td></td>
              <td style={{ textAlign: 'right', fontWeight: 700, paddingTop: 3 }}>{fmtTL(purchase.amount)}</td>
            </tr>
          </tbody>
        </table>
        <div style={{ borderTop: '1px dashed #000', margin: '6px 0' }} />
        <table style={{ width: '100%' }}>
          <tbody>
            {purchase.noDeduction ? (
              <tr><td>Kesintisiz</td><td style={{ textAlign: 'right' }}>—</td></tr>
            ) : (
              <>
                <tr><td>Komisyon (%{purchase.commissionRate})</td><td style={{ textAlign: 'right' }}>− {fmtTL(purchase.commissionAmount)}</td></tr>
                <tr><td>Stopaj (%{purchase.stopajOrani})</td><td style={{ textAlign: 'right' }}>− {fmtTL(purchase.stopajTutari)}</td></tr>
                {purchase.applyBagkur && <tr><td>BAĞ-KUR (%{purchase.bagkurRate})</td><td style={{ textAlign: 'right' }}>− {fmtTL(purchase.bagkurTutari)}</td></tr>}
              </>
            )}
            <tr style={{ borderTop: '1px solid #000' }}><td style={{ fontWeight: 700, paddingTop: 4 }}>ÖDENEN NET</td><td style={{ textAlign: 'right', fontWeight: 700, paddingTop: 4 }}>{fmtTL(purchase.netPayment)}</td></tr>
          </tbody>
        </table>
        {purchase.note && <div style={{ marginTop: 6 }}>Not: {purchase.note}</div>}
        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 26, fontSize: 10 }}>
          <div>Satıcı İmza<br/>........................</div>
          <div>Alıcı İmza<br/>........................</div>
        </div>
      </div>
    </div>
  );
}

export default function ZeytinDefteri() {
  const [tab, setTab] = useState('dashboard');
  const [farmers, setFarmers] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [payments, setPayments] = useState([]);
  const [buyers, setBuyers] = useState([]);
  const [sales, setSales] = useState([]);
  const [settings, setSettings] = useState({});
  const [priceList, setPriceList] = useState([]);
  const [selectedFarmerId, setSelectedFarmerId] = useState('');
  const [loaded, setLoaded] = useState(false);
  const [printTarget, setPrintTarget] = useState(null);

  useEffect(() => {
    (async () => {
      const [f, p, pay, b, s, set, pl] = await Promise.all([
        storageGet('zk:farmers'),
        storageGet('zk:purchases'),
        storageGet('zk:payments'),
        storageGet('zk:buyers'),
        storageGet('zk:sales'),
        storageGet('zk:settings'),
        storageGet('zk:priceList'),
      ]);
      setFarmers(f || []); setPurchases(p || []); setPayments(pay || []);
      setBuyers(b || []); setSales(s || []); setSettings(set || {});
      if (pl && pl.length > 0) {
        const normalized = pl.map((v) => ('grades' in v ? v : { id: v.id, name: v.name, hasGrades: false, singlePrice: v.price || 0, grades: [] }));
        setPriceList(normalized);
      } else {
        const defaults = [
          { id: uid(), name: 'Tirilye', hasGrades: true, singlePrice: 0, grades: [
            { id: uid(), name: '1 Numara', price: 100 },
            { id: uid(), name: '2 Numara', price: 90 },
            { id: uid(), name: '3 Numara', price: 80 },
            { id: uid(), name: '4 Numara', price: 70 },
          ] },
          { id: uid(), name: 'Edremit', hasGrades: true, singlePrice: 0, grades: [] },
          { id: uid(), name: 'Domat', hasGrades: true, singlePrice: 0, grades: [] },
          { id: uid(), name: 'Uslu', hasGrades: true, singlePrice: 0, grades: [] },
          { id: uid(), name: 'Aydın', hasGrades: true, singlePrice: 0, grades: [] },
          { id: uid(), name: 'Manzelin', hasGrades: true, singlePrice: 0, grades: [] },
          { id: uid(), name: 'Yağlık', hasGrades: false, singlePrice: 60, grades: [] },
        ];
        setPriceList(defaults);
        await storageSet('zk:priceList', defaults);
      }
      setLoaded(true);
    })();
  }, []);

  const handlePrintReceipt = (purchase) => {
    const farmer = farmers.find((f) => f.id === purchase.farmerId);
    setPrintTarget({ purchase, farmer });
    setTimeout(() => window.print(), 100);
  };

  const navItems = [
    { key: 'dashboard', label: 'Pano', icon: LayoutDashboard },
    { key: 'farmers', label: 'Çiftçiler', icon: Users },
    { key: 'purchase', label: 'Yeni alım', icon: ScaleIcon },
    { key: 'warehouse', label: 'Depo & satış', icon: Warehouse },
    { key: 'ledger', label: 'Cari hesap', icon: Wallet },
    { key: 'reports', label: 'Raporlar', icon: FileBarChart },
    { key: 'settings', label: 'Ayarlar', icon: SettingsIcon },
  ];

  if (!loaded) {
    return <div className="zk-app"><GlobalStyle /><div style={{ padding: 40, fontSize: 13, color: COLORS.inkSoft }}>Yükleniyor...</div></div>;
  }

  return (
    <div className="zk-app">
      <GlobalStyle />
      <div className="zk-shell">
        <div className="zk-sidebar">
          <div className="zk-brand-row">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
              <path d="M4 20c6-1 9-4 11-9 1.5-3.7 1-6.5-1-8.5-3 2-5 5-6 8-1.5 4-3 7-4 9.5Z" stroke="#D9C77E" strokeWidth="1.4" strokeLinejoin="round"/>
              <ellipse cx="9.3" cy="12.5" rx="1.5" ry="2.1" transform="rotate(-35 9.3 12.5)" fill="#D9C77E"/>
              <ellipse cx="12.6" cy="8.4" rx="1.3" ry="1.8" transform="rotate(-35 12.6 8.4)" fill="#D9C77E" opacity="0.85"/>
            </svg>
            <div className="zk-brand">Zeytin Defteri</div>
          </div>
          <div className="zk-brand-sub">Komisyon Yönetimi</div>
          {navItems.map((item) => (
            <button key={item.key} className={`zk-navbtn ${tab === item.key ? 'active' : ''}`} onClick={() => setTab(item.key)}>
              <item.icon size={16} /> {item.label}
            </button>
          ))}
        </div>
        <div className="zk-main">
          {tab === 'dashboard' && <DashboardTab farmers={farmers} purchases={purchases} payments={payments} sales={sales} setTab={setTab} />}
          {tab === 'farmers' && <FarmersTab farmers={farmers} setFarmers={setFarmers} purchases={purchases} payments={payments} setTab={setTab} setSelectedFarmerId={setSelectedFarmerId} />}
          {tab === 'purchase' && <PurchaseTab farmers={farmers} setFarmers={setFarmers} purchases={purchases} setPurchases={setPurchases} onPrintReceipt={handlePrintReceipt} settings={settings} priceList={priceList} />}
          {tab === 'warehouse' && <WarehouseTab purchases={purchases} buyers={buyers} setBuyers={setBuyers} sales={sales} setSales={setSales} />}
          {tab === 'ledger' && <LedgerTab farmers={farmers} purchases={purchases} payments={payments} setPayments={setPayments} selectedFarmerId={selectedFarmerId} setSelectedFarmerId={setSelectedFarmerId} onPrintReceipt={handlePrintReceipt} />}
          {tab === 'reports' && <ReportsTab farmers={farmers} purchases={purchases} sales={sales} buyers={buyers} />}
          {tab === 'settings' && <SettingsTab settings={settings} setSettings={setSettings} priceList={priceList} setPriceList={setPriceList} />}
        </div>
      </div>
      {printTarget && <PrintArea purchase={printTarget.purchase} farmer={printTarget.farmer} settings={settings} />}
    </div>
  );
}
