import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

// Firestore 및 Firebase 설정 (기존 구조 유지)
import { initializeApp } from 'firebase/app';
import { getFirestore, collection, getDocs, setDoc, doc, deleteDoc } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyDummyKeyForEnvironment",
  authDomain: "asset-management.firebaseapp.com",
  projectId: "asset-management",
  storageBucket: "asset-management.appspot.com",
  messagingSenderId: "1234567890",
  appId: "1:1234567890:web:abcdef"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// 구글 시트 단가현황 API URL
const GOOGLE_SHEET_API_URL = 'https://script.google.com/macros/s/AKfycbwXxM9sjxOAwHQDQ4kjX9tINeaD5aJg5eOVGglD8Z8IGs7WVdkTEw6nUKG-Tm2Kej1-/exec';

export default function AssetManagementApp() {
  // 탭 상태: 'payment' | 'stock' | 'transaction' | 'portfolio'
  const [activeTab, setActiveTab] = useState('payment');

  // 1. 납입금액 관리 State
  const [payments, setPayments] = useState([]);
  const [paymentMsg, setPaymentMsg] = useState('');

  // 2. 종목 관리 State
  const [stocks, setStocks] = useState([]);
  const [stockMsg, setStockMsg] = useState('');

  // 3. 거래현황 State
  const [transactions, setTransactions] = useState([]);
  const [txStartDate, setTxStartDate] = useState('');
  const [txEndDate, setTxEndDate] = useState('');
  const [txMsg, setTxMsg] = useState('');

  // 4. 포트폴리오 현황 State
  const [portfolioBaseDate, setPortfolioBaseDate] = useState('2026-09-05');
  const [portfolioBankFilter, setPortfolioBankFilter] = useState('전체 은행');
  const [portfolioPurposeFilter, setPortfolioPurposeFilter] = useState('전체 목적');
  const [portfolioList, setPortfolioList] = useState([]);
  const [portfolioMsg, setPortfolioMsg] = useState('');

  // 초기 데이터 로드
  useEffect(() => {
    loadAllData();
  }, []);

  const loadAllData = async () => {
    try {
      // 납입금액 로드
      const paySnap = await getDocs(collection(db, 'payments'));
      const payData = paySnap.docs.map(d => ({ id: d.id, ...d.data() }));
      setPayments(payData);

      // 종목관리 로드
      const stockSnap = await getDocs(collection(db, 'stocks'));
      const stockData = stockSnap.docs.map(d => ({ id: d.id, ...d.data() }));
      // 은행별, 목적별, 종목코드별 정렬
      stockData.sort((a, b) => {
        if ((a.bank || '') !== (b.bank || '')) return (a.bank || '').localeCompare(b.bank || '');
        if ((a.purpose || '') !== (b.purpose || '')) return (a.purpose || '').localeCompare(b.purpose || '');
        return (a.code || '').localeCompare(b.code || '');
      });
      setStocks(stockData);

      // 거래현황 로드
      const txSnap = await getDocs(collection(db, 'transactions'));
      const txData = txSnap.docs.map(d => ({ id: d.id, ...d.data() }));
      // 거래일자, 은행, 목적, 종목코드 순 정렬
      txData.sort((a, b) => {
        if ((a.date || '') !== (b.date || '')) return (a.date || '').localeCompare(b.date || '');
        if ((a.bank || '') !== (b.bank || '')) return (a.bank || '').localeCompare(b.bank || '');
        if ((a.purpose || '') !== (b.purpose || '')) return (a.purpose || '').localeCompare(b.purpose || '');
        return (a.code || '').localeCompare(b.code || '');
      });
      setTransactions(txData);

      // 포트폴리오 로드
      const pfSnap = await getDocs(collection(db, 'portfolios'));
      const pfData = pfSnap.docs.map(d => ({ id: d.id, ...d.data() }));
      setPortfolioList(pfData);

    } catch (e) {
      console.error("데이터 로딩 실패:", e);
    }
  };

  // ==================== [1. 납입금액 관리 로직] ====================
  const handlePaymentAdd = () => {
    const newRow = { id: 'temp_' + Date.now(), year: '2026', bank: '', amount: 0, isNew: true, checked: false };
    setPayments([...payments, newRow]);
  };

  const handlePaymentDelete = () => {
    setPayments(payments.filter(p => !p.checked));
  };

  const handlePaymentExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const formatted = data.map((item, idx) => ({
        id: 'excel_' + Date.now() + '_' + idx,
        year: String(item['납입연도'] || '2026'),
        bank: String(item['은행'] || ''),
        amount: Number(item['납입금액'] || 0),
        isNew: true
      }));
      setPayments([...payments, ...formatted]);
    };
    reader.readAsBinaryString(file);
  };

  const handlePaymentSave = async () => {
    try {
      for (let p of payments) {
        const docId = p.id.startsWith('temp_') || p.id.startsWith('excel_') ? 'pay_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5) : p.id;
        await setDoc(doc(db, 'payments', docId), {
          year: p.year,
          bank: p.bank,
          amount: Number(p.amount)
        });
      }
      setPaymentMsg('납입금액이 성공적으로 저장되었습니다.');
      setTimeout(() => setPaymentMsg(''), 3000);
      loadAllData();
    } catch (e) {
      alert("저장 실패: " + e.message);
    }
  };


  // ==================== [2. 종목 관리 로직] ====================
  const handleStockAdd = () => {
    const newRow = { id: 'temp_' + Date.now(), bank: '미래에셋', purpose: '연금', code: '', name: '', type: '', ratio: 0, currency: 'KRW', isNew: true, checked: false };
    setStocks([...stocks, newRow]);
  };

  const handleStockDelete = () => {
    setStocks(stocks.filter(s => !s.checked));
  };

  const handleStockExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const formatted = data.map((item, idx) => ({
        id: 'excel_' + Date.now() + '_' + idx,
        bank: String(item['은행'] || '미래에셋'),
        purpose: String(item['목적'] || '연금'),
        code: String(item['종목코드'] || ''),
        name: String(item['종목명'] || ''),
        type: String(item['종목 유형'] || ''),
        ratio: Number(item['유형 비율'] || 0),
        currency: String(item['통화'] || 'KRW'),
        isNew: true
      }));
      setStocks([...stocks, ...formatted]);
    };
    reader.readAsBinaryString(file);
  };

  const handleStockSave = async () => {
    try {
      for (let s of stocks) {
        const docId = s.id.startsWith('temp_') || s.id.startsWith('excel_') ? 'stock_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5) : s.id;
        await setDoc(doc(db, 'stocks', docId), {
          bank: s.bank,
          purpose: s.purpose,
          code: s.code,
          name: s.name,
          type: s.type,
          ratio: Number(s.ratio),
          currency: s.currency || 'KRW'
        });
      }
      setStockMsg('종목 정보가 성공적으로 저장되었습니다.');
      setTimeout(() => setStockMsg(''), 3000);
      loadAllData();
    } catch (e) {
      alert("저장 실패: " + e.message);
    }
  };


  // ==================== [3. 거래현황 로직] ====================
  const handleTxAdd = () => {
    const today = new Date().toISOString().split('T')[0];
    const defaultBank = stocks.length > 0 ? stocks[0].bank : '';
    const defaultPurpose = stocks.length > 0 ? stocks[0].purpose : '연금';
    const matchedStock = stocks.find(s => s.bank === defaultBank && s.purpose === defaultPurpose);
    
    const newRow = {
      id: 'temp_' + Date.now(),
      date: today,
      bank: defaultBank,
      purpose: defaultPurpose,
      name: matchedStock ? matchedStock.name : '',
      code: matchedStock ? matchedStock.code : '',
      price: 0,
        buyQty: 0,
      sellQty: 0,
      currency: matchedStock ? (matchedStock.currency || 'KRW') : 'KRW',
      isNew: true,
      checked: false
    };
    setTransactions([...transactions, newRow]);
  };

  const handleTxDelete = () => {
    setTransactions(transactions.filter(t => !t.checked));
  };

  const handleTxExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const formatted = data.map((item, idx) => ({
        id: 'excel_' + Date.now() + '_' + idx,
        date: String(item['거래일자'] || new Date().toISOString().split('T')[0]),
        bank: String(item['은행'] || ''),
        purpose: String(item['목적'] || '연금'),
        name: String(item['종목명'] || ''),
        code: String(item['종목코드'] || ''),
        price: Number(item['단가'] || 0),
        buyQty: Number(item['매수수량'] || 0),
        sellQty: Number(item['매도수량'] || 0),
        currency: String(item['통화'] || 'KRW'),
        isNew: true
      }));
      setTransactions([...transactions, ...formatted]);
    };
    reader.readAsBinaryString(file);
  };

  const handleTxSave = async () => {
    try {
      for (let t of transactions) {
        const docId = t.id.startsWith('temp_') || t.id.startsWith('excel_') ? 'tx_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5) : t.id;
        await setDoc(doc(db, 'transactions', docId), {
          date: t.date,
          bank: t.bank,
          purpose: t.purpose,
          name: t.name,
          code: t.code,
          price: Number(t.price),
          buyQty: Number(t.buyQty),
          sellQty: Number(t.sellQty),
          currency: t.currency || 'KRW'
        });
      }
      setTxMsg('거래현황이 성공적으로 저장되었습니다.');
      setTimeout(() => setTxMsg(''), 3000);
      loadAllData();
    } catch (e) {
      alert("저장 실패: " + e.message);
    }
  };


  // ==================== [4. 포트폴리오 현황 로직] ====================
  const handlePortfolioCalculate = async () => {
    try {
      // 1. 구글 시트 웹 앱 API에서 단가현황 가져오기
      let sheetPrices = {};
      try {
        const res = await fetch(GOOGLE_SHEET_API_URL);
        const sheetData = await res.json();
        // 데이터 구조가 배열 혹은 객체 형태일 때 종목코드 기준 맵핑
        if (Array.isArray(sheetData)) {
          sheetData.forEach(item => {
            const code = String(item['종목코드'] || item['code'] || '');
            const price = Number(item['단가'] || item['price'] || 0);
            if (code) sheetPrices[code] = price;
          });
        }
      } catch (err) {
        console.warn("구글 시트 API 호출 실패, 수동 입력 및 기본값 사용:", err);
      }

      // 2. 포트폴리오 계산 대상 그룹화 (은행, 목적, 종목명, 통화)
      const groupMap = {};

      // 거래현황을 기준으로 자산 집계
      transactions.forEach(tx => {
        if (tx.date && tx.date <= portfolioBaseDate) {
          const key = `${tx.bank}_${tx.purpose}_${tx.name}_${tx.code}_${tx.currency || 'KRW'}`;
          if (!groupMap[key]) {
            groupMap[key] = {
              bank: tx.bank,
              purpose: tx.purpose,
              name: tx.name,
              code: tx.code,
              currency: tx.currency || 'KRW',
              totalQty: 0,
              totalBuyCost: 0,
              totalSellQty: 0
            };
          }
          const buy = Number(tx.buyQty || 0);
          const sell = Number(tx.sellQty || 0);
          const price = Number(tx.price || 0);

          groupMap[key].totalQty += (buy - sell);
          if (buy > 0) {
            groupMap[key].totalBuyCost += (buy * price);
          }
        }
      });

      const newPortfolios = [];
      Object.keys(groupMap).forEach(key => {
        const item = groupMap[key];
        if (item.totalQty > 0) {
          const avgPrice = item.totalQty > 0 ? (item.totalBuyCost / item.totalQty) : 0;
          
          // 구글 시트 단가현황에서 가져오기, 없으면 기존 화면이나 0 (수동 입력 가능)
          const currentPrice = sheetPrices[item.code] !== undefined ? sheetPrices[item.code] : 0;
          
          const buyAmount = item.totalQty * avgPrice;
          const currentAmount = item.totalQty * currentPrice;
          const evalProfit = currentAmount - buyAmount;
          const sellProfit = 0; // 매매손익 기본값
          const profitRate = buyAmount > 0 ? (evalProfit / buyAmount) : 0;

          newPortfolios.push({
            id: 'pf_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
            bank: item.bank,
            purpose: item.purpose,
            name: item.name,
            code: item.code,
            avgPrice: Math.round(avgPrice * 100) / 100,
            currentPrice: currentPrice,
            qty: item.totalQty,
            buyAmount: Math.round(buyAmount),
            currentAmount: Math.round(currentAmount),
            evalProfit: Math.round(evalProfit),
            sellProfit: Math.round(sellProfit),
            profitRate: profitRate,
            currency: item.currency,
            checked: false
          });
        }
      });

      setPortfolioList(newPortfolios);
      setPortfolioMsg('포트폴리오 계산이 완료되었습니다. (구글 시트 단가 연동됨)');
      setTimeout(() => setPortfolioMsg(''), 3000);
    } catch (e) {
      alert("포트폴리오 계산 중 오류 발생: " + e.message);
    }
  };

  const handlePortfolioDelete = () => {
    setPortfolioList(portfolioList.filter(p => !p.checked));
  };

  const handlePortfolioSave = async () => {
    try {
      for (let p of portfolioList) {
        await setDoc(doc(db, 'portfolios', p.id), {
          bank: p.bank,
          purpose: p.purpose,
          name: p.name,
          code: p.code,
          avgPrice: Number(p.avgPrice),
          currentPrice: Number(p.currentPrice),
          qty: Number(p.qty),
          buyAmount: Number(p.buyAmount),
          currentAmount: Number(p.currentAmount),
          evalProfit: Number(p.evalProfit),
          sellProfit: Number(p.sellProfit),
          profitRate: Number(p.profitRate),
          currency: p.currency
        });
      }
      setPortfolioMsg('포트폴리오 현황이 성공적으로 저장되었습니다.');
      setTimeout(() => setPortfolioMsg(''), 3000);
      loadAllData();
    } catch (e) {
      alert("저장 실패: " + e.message);
    }
  };

  // 포트폴리오 필터링 적용된 목록
  const filteredPortfolios = portfolioList.filter(p => {
    if (portfolioBankFilter !== '전체 은행' && p.bank !== portfolioBankFilter) return false;
    if (portfolioPurposeFilter !== '전체 목적' && p.purpose !== portfolioPurposeFilter) return false;
    return true;
  });

  const totalCurrentAmount = filteredPortfolios.reduce((acc, cur) => acc + (cur.currentAmount || 0), 0);
  const totalEvalProfit = filteredPortfolios.reduce((acc, cur) => acc + (cur.evalProfit || 0), 0);
  const totalSellProfit = filteredPortfolios.reduce((acc, cur) => acc + (cur.sellProfit || 0), 0);


  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif', backgroundColor: '#f9f9f9', minHeight: '100vh' }}>
      <h2 style={{ marginBottom: '20px', color: '#333' }}>자산 관리 웹 애플리케이션</h2>

      {/* 네비게이션 탭 */}
      <div style={{ display: 'flex', gap: '10px', marginBottom: '20px', borderBottom: '2px solid #ddd', paddingBottom: '10px' }}>
        <button 
          onClick={() => setActiveTab('payment')}
          style={{ padding: '10px 20px', backgroundColor: activeTab === 'payment' ? '#007bff' : '#f1f1f1', color: activeTab === 'payment' ? '#fff' : '#333', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
          납입금액 관리
        </button>
        <button 
          onClick={() => setActiveTab('stock')}
          style={{ padding: '10px 20px', backgroundColor: activeTab === 'stock' ? '#007bff' : '#f1f1f1', color: activeTab === 'stock' ? '#fff' : '#333', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
          종목 관리
        </button>
        <button 
          onClick={() => setActiveTab('transaction')}
          style={{ padding: '10px 20px', backgroundColor: activeTab === 'transaction' ? '#007bff' : '#f1f1f1', color: activeTab === 'transaction' ? '#fff' : '#333', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
          거래현황
        </button>
        <button 
          onClick={() => setActiveTab('portfolio')}
          style={{ padding: '10px 20px', backgroundColor: activeTab === 'portfolio' ? '#007bff' : '#f1f1f1', color: activeTab === 'portfolio' ? '#fff' : '#333', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
          포트폴리오 현황
        </button>
      </div>

      {/* ==================== 1. 납입금액 관리 화면 ==================== */}
      {activeTab === 'payment' && (
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <div style={{ display: 'flex', gap: '10px' }}>
              <button onClick={loadAllData} style={{ padding: '8px 15px', cursor: 'pointer' }}>조회</button>
              <button onClick={handlePaymentAdd} style={{ padding: '8px 15px', cursor: 'pointer' }}>추가</button>
              <button onClick={handlePaymentDelete} style={{ padding: '8px 15px', cursor: 'pointer', backgroundColor: '#ff4d4d', color: '#fff', border: 'none' }}>삭제</button>
              <label style={{ padding: '8px 15px', backgroundColor: '#28a745', color: '#fff', borderRadius: '4px', cursor: 'pointer' }}>
                엑셀업로드
                <input type="file" accept=".xlsx, .xls" onChange={handlePaymentExcel} style={{ display: 'none' }} />
              </label>
              <button onClick={handlePaymentSave} style={{ padding: '8px 15px', backgroundColor: '#007bff', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>저장</button>
            </div>
            <div style={{ fontWeight: 'bold' }}>총 건수: {payments.length}건</div>
          </div>
          {paymentMsg && <div style={{ color: 'green', marginBottom: '10px', fontWeight: 'bold' }}>{paymentMsg}</div>}
          <table style={{ width: '100%', borderCollapse: 'collapse', backgroundColor: '#fff' }}>
            <thead>
              <tr style={{ backgroundColor: '#f1f1f1', borderBottom: '1px solid #ddd' }}>
                <th style={{ padding: '10px', width: '40px' }}><input type="checkbox" /></th>
                <th style={{ padding: '10px', textAlign: 'left' }}>납입연도</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>은행</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>납입금액</th>
              </tr>
            </thead>
            <tbody>
              {payments.map((p, idx) => (
                <tr key={p.id || idx} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '10px', textAlign: 'center' }}>
                    <input type="checkbox" checked={p.checked || false} onChange={(e) => {
                      const updated = [...payments];
                      updated[idx].checked = e.target.checked;
                      setPayments(updated);
                    }} />
                  </td>
                  <td style={{ padding: '10px' }}>
                    <input type="text" value={p.year || ''} onChange={(e) => {
                      const updated = [...payments];
                      updated[idx].year = e.target.value;
                      setPayments(updated);
                    }} style={{ width: '100px', padding: '4px' }} />
                  </td>
                  <td style={{ padding: '10px' }}>
                    <input type="text" value={p.bank || ''} onChange={(e) => {
                      const updated = [...payments];
                      updated[idx].bank = e.target.value;
                      setPayments(updated);
                    }} style={{ width: '150px', padding: '4px' }} />
                  </td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>
                    <input type="number" value={p.amount || 0} onChange={(e) => {
                      const updated = [...payments];
                      updated[idx].amount = e.target.value;
                      setPayments(updated);
                    }} style={{ width: '150px', padding: '4px', textAlign: 'right' }} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ==================== 2. 종목 관리 화면 ==================== */}
      {activeTab === 'stock' && (
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <div style={{ display: 'flex', gap: '10px' }}>
              <button onClick={loadAllData} style={{ padding: '8px 15px', cursor: 'pointer' }}>조회</button>
              <button onClick={handleStockAdd} style={{ padding: '8px 15px', cursor: 'pointer' }}>추가</button>
              <button onClick={handleStockDelete} style={{ padding: '8px 15px', cursor: 'pointer', backgroundColor: '#ff4d4d', color: '#fff', border: 'none' }}>삭제</button>
              <label style={{ padding: '8px 15px', backgroundColor: '#28a745', color: '#fff', borderRadius: '4px', cursor: 'pointer' }}>
                엑셀업로드
                <input type="file" accept=".xlsx, .xls" onChange={handleStockExcel} style={{ display: 'none' }} />
              </label>
              <button onClick={handleStockSave} style={{ padding: '8px 15px', backgroundColor: '#007bff', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>저장</button>
            </div>
            <div style={{ fontWeight: 'bold' }}>총 건수: {stocks.length}건</div>
          </div>
          {stockMsg && <div style={{ color: 'green', marginBottom: '10px', fontWeight: 'bold' }}>{stockMsg}</div>}
          <table style={{ width: '100%', borderCollapse: 'collapse', backgroundColor: '#fff' }}>
            <thead>
              <tr style={{ backgroundColor: '#f1f1f1', borderBottom: '1px solid #ddd' }}>
                <th style={{ padding: '10px', width: '40px' }}><input type="checkbox" /></th>
                <th style={{ padding: '10px', textAlign: 'left' }}>은행</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>목적</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>종목코드</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>종목명</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>종목 유형</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>유형 비율</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>통화</th>
              </tr>
            </thead>
            <tbody>
              {stocks.map((s, idx) => (
                <tr key={s.id || idx} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '10px', textAlign: 'center' }}>
                    <input type="checkbox" checked={s.checked || false} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].checked = e.target.checked;
                      setStocks(updated);
                    }} />
                  </td>
                  <td style={{ padding: '10px' }}>
                    <select value={s.bank || '미래에셋'} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].bank = e.target.value;
                      setStocks(updated);
                    }} style={{ padding: '4px' }}>
                      <option value="미래에셋">미래에셋</option>
                      <option value="KB증권">KB증권</option>
                      <option value="삼성증권">삼성증권</option>
                    </select>
                  </td>
                  <td style={{ padding: '10px' }}>
                    <select value={s.purpose || '연금'} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].purpose = e.target.value;
                      setStocks(updated);
                    }} style={{ padding: '4px' }}>
                      <option value="연금">연금</option>
                      <option value="IRP">IRP</option>
                      <option value="DC">DC</option>
                      <option value="기타">기타</option>
                    </select>
                  </td>
                  <td style={{ padding: '10px' }}>
                    <input type="text" value={s.code || ''} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].code = e.target.value;
                      setStocks(updated);
                    }} style={{ width: '100px', padding: '4px' }} />
                  </td>
                  <td style={{ padding: '10px' }}>
                    <input type="text" value={s.name || ''} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].name = e.target.value;
                      setStocks(updated);
                    }} style={{ width: '150px', padding: '4px' }} />
                  </td>
                  <td style={{ padding: '10px' }}>
                    <input type="text" value={s.type || ''} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].type = e.target.value;
                      setStocks(updated);
                    }} style={{ width: '100px', padding: '4px' }} />
                  </td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>
                    <input type="number" value={s.ratio || 0} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].ratio = e.target.value;
                      setStocks(updated);
                    }} style={{ width: '80px', padding: '4px', textAlign: 'right' }} />
                  </td>
                  <td style={{ padding: '10px' }}>
                    <select value={s.currency || 'KRW'} onChange={(e) => {
                      const updated = [...stocks];
                      updated[idx].currency = e.target.value;
                      setStocks(updated);
                    }} style={{ padding: '4px' }}>
                      <option value="KRW">KRW</option>
                      <option value="USD">USD</option>
                    </select>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ==================== 3. 거래현황 화면 ==================== */}
      {activeTab === 'transaction' && (
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <span>거래일자:</span>
              <input type="date" value={txStartDate} onChange={(e) => setTxStartDate(e.target.value)} style={{ padding: '6px' }} />
              ~
              <input type="date" value={txEndDate} onChange={(e) => setTxEndDate(e.target.value)} style={{ padding: '6px' }} />
              <button onClick={loadAllData} style={{ padding: '8px 15px', cursor: 'pointer' }}>조회</button>
              <button onClick={handleTxAdd} style={{ padding: '8px 15px', cursor: 'pointer' }}>추가</button>
              <button onClick={handleTxDelete} style={{ padding: '8px 15px', cursor: 'pointer', backgroundColor: '#ff4d4d', color: '#fff', border: 'none' }}>삭제</button>
              <label style={{ padding: '8px 15px', backgroundColor: '#28a745', color: '#fff', borderRadius: '4px', cursor: 'pointer' }}>
                엑셀업로드
                <input type="file" accept=".xlsx, .xls" onChange={handleTxExcel} style={{ display: 'none' }} />
              </label>
              <button onClick={handleTxSave} style={{ padding: '8px 15px', backgroundColor: '#007bff', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>저장</button>
            </div>
            <div style={{ fontWeight: 'bold' }}>
              총 건수: {transactions.filter(t => (!txStartDate || t.date >= txStartDate) && (!txEndDate || t.date <= txEndDate)).length}건
            </div>
          </div>
          {txMsg && <div style={{ color: 'green', marginBottom: '10px', fontWeight: 'bold' }}>{txMsg}</div>}
          <table style={{ width: '100%', borderCollapse: 'collapse', backgroundColor: '#fff' }}>
            <thead>
              <tr style={{ backgroundColor: '#f1f1f1', borderBottom: '1px solid #ddd' }}>
                <th style={{ padding: '10px', width: '40px' }}><input type="checkbox" /></th>
                <th style={{ padding: '10px', textAlign: 'left' }}>거래일자</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>은행</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>목적</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>종목명</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>종목코드</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>통화</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>단가</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>매수수량</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>매도수량</th>
              </tr>
            </thead>
            <tbody>
              {transactions
                .filter(t => (!txStartDate || t.date >= txStartDate) && (!txEndDate || t.date <= txEndDate))
                .map((t, idx) => {
                  // 종목관리에서 해당 은행의 고유 은행 목록 추출
                  const uniqueBanks = [...new Set(stocks.map(s => s.bank))];
                  // 선택된 은행의 목적 목록
                  const filteredPurposes = [...new Set(stocks.filter(s => s.bank === t.bank).map(s => s.purpose))];
                  // 선택된 은행 & 목적의 종목명 목록
                  const filteredNames = [...new Set(stocks.filter(s => s.bank === t.bank && s.purpose === t.purpose).map(s => s.name))];

                  return (
                    <tr key={t.id || idx} style={{ borderBottom: '1px solid #eee' }}>
                      <td style={{ padding: '10px', textAlign: 'center' }}>
                        <input type="checkbox" checked={t.checked || false} onChange={(e) => {
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].checked = e.target.checked;
                          setTransactions(updated);
                        }} />
                      </td>
                      <td style={{ padding: '10px' }}>
                        <input type="date" value={t.date || ''} onChange={(e) => {
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].date = e.target.value;
                          setTransactions(updated);
                        }} style={{ padding: '4px' }} />
                      </td>
                      <td style={{ padding: '10px' }}>
                        <select value={t.bank || ''} onChange={(e) => {
                          const val = e.target.value;
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].bank = val;
                          // 연관된 목적, 종목명, 종목코드, 통화 자동 갱신
                          const validPurposes = [...new Set(stocks.filter(s => s.bank === val).map(s => s.purpose))];
                          updated[realIdx].purpose = validPurposes[0] || '연금';
                          const validNames = [...new Set(stocks.filter(s => s.bank === val && s.purpose === updated[realIdx].purpose).map(s => s.name))];
                          updated[realIdx].name = validNames[0] || '';
                          const matchedStock = stocks.find(s => s.bank === val && s.purpose === updated[realIdx].purpose && s.name === updated[realIdx].name);
                          updated[realIdx].code = matchedStock ? matchedStock.code : '';
                          updated[realIdx].currency = matchedStock ? (matchedStock.currency || 'KRW') : 'KRW';
                          setTransactions(updated);
                        }} style={{ padding: '4px' }}>
                          {uniqueBanks.map(b => <option key={b} value={b}>{b}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: '10px' }}>
                        <select value={t.purpose || ''} onChange={(e) => {
                          const val = e.target.value;
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].purpose = val;
                          const validNames = [...new Set(stocks.filter(s => s.bank === updated[realIdx].bank && s.purpose === val).map(s => s.name))];
                          updated[realIdx].name = validNames[0] || '';
                          const matchedStock = stocks.find(s => s.bank === updated[realIdx].bank && s.purpose === val && s.name === updated[realIdx].name);
                          updated[realIdx].code = matchedStock ? matchedStock.code : '';
                          updated[realIdx].currency = matchedStock ? (matchedStock.currency || 'KRW') : 'KRW';
                          setTransactions(updated);
                        }} style={{ padding: '4px' }}>
                          {filteredPurposes.map(p => <option key={p} value={p}>{p}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: '10px' }}>
                        <select value={t.name || ''} onChange={(e) => {
                          const val = e.target.value;
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].name = val;
                          const matchedStock = stocks.find(s => s.bank === updated[realIdx].bank && s.purpose === updated[realIdx].purpose && s.name === val);
                          updated[realIdx].code = matchedStock ? matchedStock.code : '';
                          updated[realIdx].currency = matchedStock ? (matchedStock.currency || 'KRW') : 'KRW';
                          setTransactions(updated);
                        }} style={{ padding: '4px' }}>
                          {filteredNames.map(n => <option key={n} value={n}>{n}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: '10px' }}>
                        <input type="text" readOnly value={t.code || ''} style={{ width: '90px', padding: '4px', backgroundColor: '#e9ecef' }} />
                      </td>
                      <td style={{ padding: '10px' }}>
                        <input type="text" readOnly value={t.currency || 'KRW'} style={{ width: '60px', padding: '4px', backgroundColor: '#e9ecef', textAlign: 'center' }} />
                      </td>
                      <td style={{ padding: '10px', textAlign: 'right' }}>
                        <input type="text" value={t.price ? (t.currency === 'USD' ? Number(t.price).toLocaleString('en-US', { minimumFractionDigits: 2 }) : Number(t.price).toLocaleString()) : ''} onChange={(e) => {
                          const rawVal = e.target.value.replace(/,/g, '');
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].price = rawVal;
                          setTransactions(updated);
                        }} style={{ width: '100px', padding: '4px', textAlign: 'right' }} />
                      </td>
                      <td style={{ padding: '10px', textAlign: 'right' }}>
                        <input type="number" value={t.buyQty || 0} onChange={(e) => {
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].buyQty = e.target.value;
                          setTransactions(updated);
                        }} style={{ width: '80px', padding: '4px', textAlign: 'right' }} />
                      </td>
                      <td style={{ padding: '10px', textAlign: 'right' }}>
                        <input type="number" value={t.sellQty || 0} onChange={(e) => {
                          const updated = [...transactions];
                          const realIdx = transactions.findIndex(item => item.id === t.id);
                          updated[realIdx].sellQty = e.target.value;
                          setTransactions(updated);
                        }} style={{ width: '80px', padding: '4px', textAlign: 'right' }} />
                      </td>
                    </tr>
                  );
                })}
            </tbody>
          </table>
        </div>
      )}

      {/* ==================== 4. 포트폴리오 현황 화면 ==================== */}
      {activeTab === 'portfolio' && (
        <div>
          {/* 상단 컨트롤 영역 */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', backgroundColor: '#fff', padding: '15px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
            <div style={{ display: 'flex', gap: '15px', alignItems: 'center' }}>
              <div>
                <span style={{ marginRight: '5px', fontWeight: 'bold' }}>기준일자</span>
                <input type="date" value={portfolioBaseDate} onChange={(e) => setPortfolioBaseDate(e.target.value)} style={{ padding: '6px' }} />
              </div>
              <div>
                <span style={{ marginRight: '5px', fontWeight: 'bold' }}>은행</span>
                <select value={portfolioBankFilter} onChange={(e) => setPortfolioBankFilter(e.target.value)} style={{ padding: '6px' }}>
                  <option value="전체 은행">전체 은행</option>
                  {[...new Set(stocks.map(s => s.bank))].map(b => <option key={b} value={b}>{b}</option>)}
                </select>
              </div>
              <div>
                <span style={{ marginRight: '5px', fontWeight: 'bold' }}>목적</span>
                <select value={portfolioPurposeFilter} onChange={(e) => setPortfolioPurposeFilter(e.target.value)} style={{ padding: '6px' }}>
                  <option value="전체 목적">전체 목적</option>
                  <option value="연금">연금</option>
                  <option value="IRP">IRP</option>
                  <option value="DC">DC</option>
                  <option value="기타">기타</option>
                </select>
              </div>
              <button onClick={loadAllData} style={{ padding: '8px 15px', cursor: 'pointer' }}>조회</button>
              <button onClick={handlePortfolioCalculate} style={{ padding: '8px 15px', backgroundColor: '#28a745', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>계산</button>
              <button onClick={handlePortfolioDelete} style={{ padding: '8px 15px', backgroundColor: '#ff4d4d', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>삭제</button>
              <button onClick={handlePortfolioSave} style={{ padding: '8px 15px', backgroundColor: '#007bff', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>저장</button>
            </div>

            {/* 우측 금액 요약 정보 */}
            <div style={{ display: 'flex', gap: '20px', textAlign: 'right' }}>
              <div>
                <div style={{ fontSize: '12px', color: '#666' }}>현재금액 총액</div>
                <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#007bff' }}>{totalCurrentAmount.toLocaleString()}원</div>
              </div>
              <div>
                <div style={{ fontSize: '12px', color: '#666' }}>평가손익 총액</div>
                <div style={{ fontSize: '16px', fontWeight: 'bold', color: totalEvalProfit >= 0 ? '#28a745' : '#ff4d4d' }}>{totalEvalProfit.toLocaleString()}원</div>
              </div>
              <div>
                <div style={{ fontSize: '12px', color: '#666' }}>매매손익 총액</div>
                <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#333' }}>{totalSellProfit.toLocaleString()}원</div>
              </div>
            </div>
          </div>

          {portfolioMsg && <div style={{ color: 'green', marginBottom: '10px', fontWeight: 'bold' }}>{portfolioMsg}</div>}

          {/* 하단 데이터 그리드 (테이블) */}
          <table style={{ width: '100%', borderCollapse: 'collapse', backgroundColor: '#fff' }}>
            <thead>
              <tr style={{ backgroundColor: '#f1f1f1', borderBottom: '1px solid #ddd' }}>
                <th style={{ padding: '10px', width: '40px' }}><input type="checkbox" /></th>
                <th style={{ padding: '10px', textAlign: 'left' }}>은행</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>목적</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>종목명</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>평균단가</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>현재단가</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>수량</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>매입금액</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>현재금액</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>평가손익</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>매매손익</th>
                <th style={{ padding: '10px', textAlign: 'right' }}>이익율</th>
                <th style={{ padding: '10px', textAlign: 'center' }}>통화</th>
              </tr>
            </thead>
            <tbody>
              {filteredPortfolios.map((p, idx) => (
                <tr key={p.id || idx} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '10px', textAlign: 'center' }}>
                    <input type="checkbox" checked={p.checked || false} onChange={(e) => {
                      const updated = [...portfolioList];
                      const realIdx = portfolioList.findIndex(item => item.id === p.id);
                      updated[realIdx].checked = e.target.checked;
                      setPortfolioList(updated);
                    }} />
                  </td>
                  <td style={{ padding: '10px' }}>{p.bank}</td>
                  <td style={{ padding: '10px' }}>{p.purpose}</td>
                  <td style={{ padding: '10px' }}>{p.name} ({p.code})</td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>{Number(p.avgPrice).toLocaleString()}</td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>
                    <input 
                      type="number" 
                      value={p.currentPrice} 
                      onChange={(e) => {
                        const val = Number(e.target.value);
                        const updated = [...portfolioList];
                        const realIdx = portfolioList.findIndex(item => item.id === p.id);
                        updated[realIdx].currentPrice = val;
                        
                        // 파생 변수 재계산
                        const qty = updated[realIdx].qty;
                        const avgP = updated[realIdx].avgPrice;
                        const buyAmt = qty * avgP;
                        const curAmt = qty * val;
                        const evProfit = curAmt - buyAmt;
                        
                        updated[realIdx].buyAmount = Math.round(buyAmt);
                        updated[realIdx].currentAmount = Math.round(curAmt);
                        updated[realIdx].evalProfit = Math.round(evProfit);
                        updated[realIdx].profitRate = buyAmt > 0 ? (evProfit / buyAmount) : 0;
                        
                        setPortfolioList(updated);
                      }} 
                      style={{ width: '90px', padding: '4px', textAlign: 'right' }} 
                    />
                  </td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>{Number(p.qty).toLocaleString()}</td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>{Number(p.buyAmount).toLocaleString()}</td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>{Number(p.currentAmount).toLocaleString()}</td>
                  <td style={{ padding: '10px', textAlign: 'right', color: p.evalProfit >= 0 ? 'green' : 'red' }}>{Number(p.evalProfit).toLocaleString()}</td>
                  <td style={{ padding: '10px', textAlign: 'right' }}>{Number(p.sellProfit).toLocaleString()}</td>
                  <td style={{ padding: '10px', textAlign: 'right', color: p.profitRate >= 0 ? 'green' : 'red' }}>{(p.profitRate * 100).toFixed(2)}%</td>
                  <td style={{ padding: '10px', textAlign: 'center' }}>{p.currency}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
