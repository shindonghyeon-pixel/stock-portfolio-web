import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  Search, 
  Plus, 
  Trash2, 
  FileSpreadsheet, 
  Calendar, 
  Wallet,
  Landmark,
  PiggyBank,
  TrendingUp,
  CheckSquare,
  AlertCircle,
  Save,
  Layers,
  ArrowLeftRight,
  Calculator
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { db } from './firebase';
import { 
  collection, 
  getDocs, 
  doc, 
  writeBatch,
  serverTimestamp 
} from 'firebase/firestore';

const getTodayString = () => {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const getYearFirstDayString = () => {
  const today = new Date();
  const year = today.getFullYear();
  return `${year}-01-01`;
};

const formatCurrency = (amount, currency = 'KRW') => {
  if (amount === 0 || amount === '' || amount === undefined || isNaN(amount)) return '0';
  const num = Number(amount);
  const currUpper = (currency || 'KRW').toUpperCase();
  if (currUpper === 'USD') {
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return num.toLocaleString('ko-KR', { maximumFractionDigits: 0 });
};

const parseExcelDate = (val) => {
  if (val === undefined || val === null || val === '') return getTodayString();
  if (typeof val === 'number') {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const d = String(date.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'string') {
    let clean = val.trim().replace(/[\.\/]/g, '-');
    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(clean)) {
      const parts = clean.split('-');
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }
    const parsed = new Date(clean);
    if (!isNaN(parsed.getTime())) {
      const y = parsed.getFullYear();
      const m = String(parsed.getMonth() + 1).padStart(2, '0');
      const d = String(parsed.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return clean; 
  }
  return getTodayString();
};

export default function App() {
  const [isTailwindLoaded, setIsTailwindLoaded] = useState(false);
  const [activeTab, setActiveTab] = useState('payment');
  
  // 1. 납입금액 관리 상태
  const [payments, setPayments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [searchYearInput, setSearchYearInput] = useState('');
  const [appliedSearchYear, setAppliedSearchYear] = useState('');
  const [selectedIds, setSelectedIds] = useState([]);
  const [deletedIds, setDeletedIds] = useState([]);

  // 2. 종목 관리 상태
  const [stocks, setStocks] = useState([]);
  const [loadingStocks, setLoadingStocks] = useState(false);
  const [isSavingStocks, setIsSavingStocks] = useState(false);
  const [selectedStockIds, setSelectedStockIds] = useState([]);
  const [deletedStockIds, setDeletedStockIds] = useState([]);

  // 3. 거래현황 상태
  const [transactions, setTransactions] = useState([]);
  const [loadingTransactions, setLoadingTransactions] = useState(false);
  const [isSavingTransactions, setIsSavingTransactions] = useState(false);
  const [selectedTransactionIds, setSelectedTransactionIds] = useState([]);
  const [deletedTransactionIds, setDeletedTransactionIds] = useState([]);
  const [txStartDate, setTxStartDate] = useState('');
  const [txEndDate, setTxEndDate] = useState('');
  const [appliedTxStartDate, setAppliedTxStartDate] = useState('');
  const [appliedTxEndDate, setAppliedTxEndDate] = useState('');

  // 4. 포트폴리오 현황 상태
  const [portfolios, setPortfolios] = useState([]);
  const [loadingPortfolios, setLoadingPortfolios] = useState(false);
  const [isSavingPortfolios, setIsSavingPortfolios] = useState(false);
  const [selectedPortfolioIds, setSelectedPortfolioIds] = useState([]);
  const [deletedPortfolioIds, setDeletedPortfolioIds] = useState([]);
  const [pfBaseDate, setPfBaseDate] = useState(getTodayString());
  const [pfBankFilter, setPfBankFilter] = useState('');
  const [pfPurposeFilter, setPfPurposeFilter] = useState('');
  const [appliedPfBaseDate, setAppliedPfBaseDate] = useState(getTodayString());
  const [appliedPfBankFilter, setAppliedPfBankFilter] = useState('');
  const [appliedPfPurposeFilter, setAppliedPfPurposeFilter] = useState('');
  const [exchangeRate, setExchangeRate] = useState(1350);

  // 5. 총액 Trend 상태
  const [trendStartDate, setTrendStartDate] = useState(getYearFirstDayString());
  const [trendEndDate, setTrendEndDate] = useState(getTodayString());
  const [appliedTrendStartDate, setAppliedTrendStartDate] = useState(getYearFirstDayString());
  const [appliedTrendEndDate, setAppliedTrendEndDate] = useState(getTodayString());
  const [trendProfitRateInput, setTrendProfitRateInput] = useState(0.08); // 기본값 0.08
  const [trendRows, setTrendRows] = useState([]);
  const [selectedTrendIds, setSelectedTrendIds] = useState([]);
  const [deletedTrendIds, setDeletedTrendIds] = useState([]);
  const [isSavingTrend, setIsSavingTrend] = useState(false);

  const [errorModal, setErrorModal] = useState({ isOpen: false, message: '' });
  const [successMessage, setSuccessMessage] = useState('');
  
  const fileInputRef = useRef(null);
  const stockFileInputRef = useRef(null);
  const transactionFileInputRef = useRef(null);
  const trendFileInputRef = useRef(null);
  const [activeUploadType, setActiveUploadType] = useState('payment');

  useEffect(() => {
    if (document.getElementById('tailwind-cdn')) {
      setIsTailwindLoaded(true);
    } else {
      const script = document.createElement('script');
      script.id = 'tailwind-cdn';
      script.src = 'https://cdn.tailwindcss.com';
      script.onload = () => setIsTailwindLoaded(true);
      document.head.appendChild(script);
    }

    fetchPayments();
    fetchStocks();
    fetchTransactions();
    fetchPortfolios();
    fetchTrends();
    fetchExchangeRate();
  }, []);

  const fetchExchangeRate = async () => {
    try {
      const res = await fetch('https://open.er-api.com/v6/latest/USD');
      const data = await res.json();
      if (data && data.rates && data.rates.KRW) {
        setExchangeRate(data.rates.KRW);
      }
    } catch (e) {
      console.error("환율 정보 조회 실패:", e);
    }
  };

  const fetchPayments = async () => {
    try {
      setLoading(true);
      const querySnapshot = await getDocs(collection(db, "payments"));
      const dataList = querySnapshot.docs.map(docSnap => ({
        id: docSnap.id,
        ...docSnap.data()
      }));
      dataList.sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0));
      setPayments(dataList);
      setDeletedIds([]);
    } catch (error) {
      console.error("납입 데이터 불러오기 실패:", error);
    } finally {
      setLoading(false);
    }
  };

  const fetchStocks = async () => {
    try {
      setLoadingStocks(true);
      const querySnapshot = await getDocs(collection(db, "stocks"));
      let dataList = querySnapshot.docs.map(docSnap => ({
        id: docSnap.id,
        ...docSnap.data()
      }));
      dataList.sort((a, b) => {
        const bankA = (a.bank || '').trim();
        const bankB = (b.bank || '').trim();
        if (bankA !== bankB) return bankA.localeCompare(bankB, 'ko');
        const purposeA = (a.purpose || '').trim();
        const purposeB = (b.purpose || '').trim();
        if (purposeA !== purposeB) return purposeA.localeCompare(purposeB, 'ko');
        const codeA = (a.code || '').trim();
        const codeB = (b.code || '').trim();
        return codeA.localeCompare(codeB, 'ko');
      });
      setStocks(dataList);
      setDeletedStockIds([]);
    } catch (error) {
      console.error("종목 데이터 불러오기 실패:", error);
    } finally {
      setLoadingStocks(false);
    }
  };

  const fetchTransactions = async () => {
    try {
      setLoadingTransactions(true);
      const querySnapshot = await getDocs(collection(db, "transactions"));
      let dataList = querySnapshot.docs.map(docSnap => ({
        id: docSnap.id,
        ...docSnap.data()
      }));
      dataList.sort((a, b) => {
        const dateA = (a.date || '').trim();
        const dateB = (b.date || '').trim();
        if (dateA !== dateB) return dateB.localeCompare(dateA);
        const bankA = (a.bank || '').trim();
        const bankB = (b.bank || '').trim();
        if (bankA !== bankB) return bankA.localeCompare(bankB, 'ko');
        const purposeA = (a.purpose || '').trim();
        const purposeB = (b.purpose || '').trim();
        if (purposeA !== purposeB) return purposeA.localeCompare(purposeB, 'ko');
        const codeA = (a.code || '').trim();
        const codeB = (b.code || '').trim();
        return codeA.localeCompare(codeB, 'ko');
      });
      setTransactions(dataList);
      setDeletedTransactionIds([]);
    } catch (error) {
      console.error("거래현황 데이터 불러오기 실패:", error);
    } finally {
      setLoadingTransactions(false);
    }
  };

  const fetchPortfolios = async () => {
    try {
      setLoadingPortfolios(true);
      const querySnapshot = await getDocs(collection(db, "portfolios"));
      let dataList = querySnapshot.docs.map(docSnap => ({
        id: docSnap.id,
        ...docSnap.data()
      }));
      setPortfolios(dataList);
      setDeletedPortfolioIds([]);
    } catch (error) {
      console.error("포트폴리오 데이터 불러오기 실패:", error);
    } finally {
      setLoadingPortfolios(false);
    }
  };

  const fetchTrends = async () => {
    try {
      const querySnapshot = await getDocs(collection(db, "trends"));
      let dataList = querySnapshot.docs.map(docSnap => ({
        id: docSnap.id,
        ...docSnap.data()
      }));
      setTrendRows(dataList);
      setDeletedTrendIds([]);
    } catch (error) {
      console.error("총액 Trend 데이터 불러오기 실패:", error);
    }
  };

  const handleCalculatePortfolio = async () => {
    const baseDate = pfBaseDate || getTodayString();
    
    let externalPriceMap = new Map();
    try {
      const apiUrl = 'https://script.google.com/macros/s/AKfycbzsinoBtvlVMUQC69g2Aa6EmRM747h8ffB5_r1zM5hf1FReRQbLgJy-jHMgn6J7xFfC/exec';
      const res = await fetch(apiUrl);
      if (!res.ok) throw new Error(`HTTP 통신 에러: ${res.status}`);
      
      const textData = await res.text();
      const sheetData = JSON.parse(textData);
      
      let targetArray = Array.isArray(sheetData) ? sheetData : (sheetData.data || Object.values(sheetData).find(Array.isArray) || []);
      
      targetArray.forEach(item => {
        let matchedCode = '';
        let matchedPrice = 0;
        
        for (const [key, value] of Object.entries(item)) {
          const cleanKey = String(key).replace(/\s+/g, '').toLowerCase();
          if (cleanKey.includes('종목코드') || cleanKey === 'code') {
            matchedCode = String(value).trim();
          }
          if (cleanKey.includes('단가') || cleanKey.includes('현재가') || cleanKey === 'price') {
            matchedPrice = parseFloat(String(value).replace(/[^0-9.-]+/g, '')) || 0;
          }
        }
        
        if (matchedCode) {
          externalPriceMap.set(matchedCode, matchedPrice);
        }
      });
    } catch (err) {
      console.error("구글시트 API 연동 실패:", err);
    }

    const d = new Date(baseDate);
    d.setDate(d.getDate() - 1);
    const prevDateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

    const priorTxs = transactions.filter(t => (t.date || '') <= prevDateStr);
    const holdingMap = new Map();

    priorTxs.forEach(t => {
      const key = `${t.bank}|${t.purpose}|${t.name}|${t.code}|${t.currency}`;
      if (!holdingMap.has(key)) {
        holdingMap.set(key, { qty: 0, totalCost: 0, bank: t.bank, purpose: t.purpose, name: t.name, code: t.code, currency: t.currency });
      }
      const item = holdingMap.get(key);
      const buyQ = Number(t.buyQty || 0);
      const sellQ = Number(t.sellQty || 0);
      const price = Number(t.price || 0);
      
      const newQty = item.qty + buyQ - sellQ;
      if (buyQ > 0) {
        item.totalCost += (buyQ * price);
      }
      item.qty = newQty;
    });

    const todayTxs = transactions.filter(t => (t.date || '') === baseDate);
    const todayTxMap = new Map();
    todayTxs.forEach(t => {
      const key = `${t.bank}|${t.purpose}|${t.name}|${t.code}|${t.currency}`;
      if (!todayTxMap.has(key)) {
        todayTxMap.set(key, { buyQty: 0, sellQty: 0, buyAmount: 0 });
      }
      const item = todayTxMap.get(key);
      const bQ = Number(t.buyQty || 0);
      const sQ = Number(t.sellQty || 0);
      item.buyQty += bQ;
      item.sellQty += sQ;
      if (bQ > 0) {
        item.buyAmount += bQ * Number(t.price || 0);
      }
    });

    const allKeys = new Set([...holdingMap.keys(), ...todayTxMap.keys()]);
    const newPfList = [];

    for (const key of allKeys) {
      const prevItem = holdingMap.get(key) || { qty: 0, totalCost: 0 };
      const todayItem = todayTxMap.get(key) || { buyQty: 0, sellQty: 0, buyAmount: 0 };
      
      const [bank, purpose, name, code, currency] = key.split('|');
      
      const prevQty = prevItem.qty;
      const prevAvgPrice = prevQty > 0 ? (prevItem.totalCost / prevQty) : 0;
      
      const buyQty = todayItem.buyQty;
      const sellQty = todayItem.sellQty;
      const buyAmount = todayItem.buyAmount;

      const qty = prevQty + buyQty - sellQty;
      if (qty <= 0) continue;

      const denom = prevQty + buyQty;
      const avgPrice = denom > 0 ? ((prevAvgPrice * prevQty) + buyAmount) / denom : prevAvgPrice;

      let currentPrice = 0;
      const cleanCode = (code || '').trim();
      if (externalPriceMap.has(cleanCode)) {
        currentPrice = externalPriceMap.get(cleanCode);
      } else {
        const existingPf = portfolios.find(p => !p.isManual && p.bank === bank && p.purpose === purpose && p.code === code && p.baseDate === baseDate);
        if (existingPf && existingPf.currentPrice > 0) {
          currentPrice = existingPf.currentPrice;
        }
      }

      const purchaseAmount = qty * avgPrice;
      const currentAmount = qty * currentPrice;
      const evalProfitLoss = currentAmount - purchaseAmount;
      
      const todaySellProfitLoss = sellQty > 0 ? (currentPrice - avgPrice) * sellQty : 0;
      const prevPfItem = portfolios.find(p => !p.isManual && p.baseDate === prevDateStr && p.bank === bank && p.purpose === purpose && p.code === code);
      const prevSellProfitLoss = prevPfItem ? Number(prevPfItem.sellProfitLoss || 0) : 0;
      const sellProfitLoss = prevSellProfitLoss + todaySellProfitLoss;

      const profitRate = purchaseAmount > 0 ? evalProfitLoss / purchaseAmount : 0;

      newPfList.push({
        id: 'pf_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
        baseDate: baseDate,
        bank: bank,
        purpose: purpose,
        name: name,
        code: code,
        currency: currency,
        avgPrice: avgPrice,
        currentPrice: currentPrice,
        qty: qty,
        purchaseAmount: purchaseAmount,
        currentAmount: currentAmount,
        evalProfitLoss: evalProfitLoss,
        sellProfitLoss: sellProfitLoss,
        profitRate: profitRate,
        isManual: false
      });
    }

    const existingManualItems = portfolios.filter(p => p.isManual);
    existingManualItems.forEach(manual => {
      const currentAmount = Number(manual.currentAmount || 0);
      const purchaseAmount = Number(manual.purchaseAmount || 0);
      const evalProfitLoss = currentAmount - purchaseAmount;

      newPfList.push({
        ...manual,
        id: manual.id || ('manual_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9)),
        baseDate: baseDate,
        evalProfitLoss: evalProfitLoss
      });
    });

    setPortfolios(newPfList);
    setSuccessMessage('포트폴리오 계산이 완료되었습니다.');
    setTimeout(() => setSuccessMessage(''), 3000);
  };

  const handleCalculateTrend = () => {
    if (!appliedTrendStartDate || !appliedTrendEndDate) {
      showError('시작일자와 종료일자를 반드시 입력해주세요.');
      return;
    }

    const start = new Date(appliedTrendStartDate);
    const end = new Date(appliedTrendEndDate);
    if (start > end) {
      showError('시작일자가 종료일자보다 클 수 없습니다.');
      return;
    }

    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();

    // 포트폴리오 데이터에서 존재하는 모든 기준일자 집합 추출 (날짜 오름차순 정렬)
    const availablePortfolioDates = Array.from(new Set(portfolios.map(p => p.baseDate).filter(Boolean))).sort();

    // 월별로 그룹화하여 데이터가 존재하는 날짜들을 관리
    const monthlyDataMap = new Map(); // "YYYY-MM" -> [dateStr1, dateStr2, ...]
    availablePortfolioDates.forEach(dt => {
      const ym = dt.substring(0, 7);
      if (!monthlyDataMap.has(ym)) {
        monthlyDataMap.set(ym, []);
      }
      monthlyDataMap.get(ym).push(dt);
    });

    const targetDatesSet = new Set();

    // 시작일부터 종료일까지의 월들을 순회하며 조건에 맞는 날짜 추출
    let currIter = new Date(start.getFullYear(), start.getMonth(), 1);
    const endIter = new Date(end.getFullYear(), end.getMonth(), 1);

    while (currIter <= endIter) {
      const y = currIter.getFullYear();
      const m = currIter.getMonth();
      const ym = `${y}-${String(m + 1).padStart(2, '0')}`;
      const isCurrentMonth = (y === currentYear && m === currentMonth);

      const datesInMonth = monthlyDataMap.get(ym) || [];

      if (!isCurrentMonth) {
        // 3. 당월이 아닌 경우: 해당 월의 가장 빠른 일자의 데이터를 보여줌
        if (datesInMonth.length > 0) {
          // 해당 월에 속하고 [시작일, 종료일] 범위 내에 포함되는 가장 빠른 날짜 또는 월의 첫 데이터
          const validDates = datesInMonth.filter(dt => dt >= appliedTrendStartDate && dt <= appliedTrendEndDate);
          if (validDates.length > 0) {
            targetDatesSet.add(validDates[0]);
          }
        }
      } else {
        // 4. 당월인 경우: 
        // - 당월의 첫 번째 데이터가 있는 일자
        // - 그 이후는 오늘 이전일자는 매주 월요일 데이터만
        // - 마지막에 오늘 날짜
        const monthDates = datesInMonth.filter(dt => dt >= appliedTrendStartDate && dt <= appliedTrendEndDate);
        if (monthDates.length > 0) {
          // 첫 번째 데이터
          targetDatesSet.add(monthDates[0]);

          // 그 이후 날짜 중 오늘 이전일자이면서 월요일(getDay() === 1)인 날짜들
          monthDates.forEach(dt => {
            const dtObj = new Date(dt);
            const todayStr = getTodayString();
            if (dt > monthDates[0] && dt < todayStr && dtObj.getDay() === 1) {
              targetDatesSet.add(dt);
            }
          });
        }
        // 마지막에 오늘 날짜 포함 (종료일이 오늘이거나 포함될 때)
        const todayStr = getTodayString();
        if (todayStr >= appliedTrendStartDate && todayStr <= appliedTrendEndDate) {
          targetDatesSet.add(todayStr);
        }
      }

      currIter.setMonth(currIter.getMonth() + 1);
    }

    // 만약 portfolio에 직접 계산된 날짜가 없거나 기간 내 일자가 직접 선택되지 않은 경우를 대비해 
    // 사용자가 조회 기간 내 지정한 규칙대로 일자 배열 생성 후 필터링
    if (targetDatesSet.size === 0) {
      // 기본적으로 시작일부터 종료일까지 조건에 맞는 날짜들 추가
      let c = new Date(start);
      while (c <= end) {
        const y = c.getFullYear();
        const m = c.getMonth();
        const d = c.getDate();
        const dateStr = `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        const isCurrentMonth = (y === currentYear && m === currentMonth);

        if (!isCurrentMonth) {
          // 당월이 아닌 경우 각 월의 1일 또는 해당월 첫 날짜
          if (d === 1) targetDatesSet.add(dateStr);
        } else {
          // 당월인 경우 첫날, 월요일, 오늘
          const todayStr = getTodayString();
          if (c.getDay() === 1 || dateStr === todayStr || d === 1) {
            targetDatesSet.add(dateStr);
          }
        }
        c.setDate(c.getDate() + 1);
      }
    }

    const finalDates = Array.from(targetDatesSet).sort();

    // 포트폴리오 데이터에서 날짜별 현재금액 총액 가져오기 함수
    const getPortfolioTotalForDate = (dateStr) => {
      const itemsForDate = portfolios.filter(p => p.baseDate === dateStr);
      if (itemsForDate.length === 0) return 0;
      return itemsForDate.reduce((sum, p) => {
        const isUSD = (p.currency || 'KRW').toUpperCase() === 'USD';
        const mult = isUSD ? exchangeRate : 1;
        return sum + (Number(p.currentAmount || 0) * mult);
      }, 0);
    };

    const newRows = finalDates.map((dt, idx) => {
      const amount = getPortfolioTotalForDate(dt);
      return {
        id: 'trend_' + Date.now() + '_' + idx,
        date: dt,
        amount: amount,
        isTemp: true
      };
    });

    setTrendRows(newRows);
    setSuccessMessage('총액 Trend 조회가 완료되었습니다.');
    setTimeout(() => setSuccessMessage(''), 3000);
  };

  const handleAddTrendRow = () => {
    setTrendRows(prev => [{
      id: 'temp_trend_' + Date.now(),
      date: getTodayString(),
      amount: 0,
      isTemp: true
    }, ...prev]);
  };

  const handleAddPortfolioRow = () => {
    const defaultBank = availableBanks[0] || '';
    const defaultPurpose = availablePurposes[0] || '연금';
    setPortfolios(prev => [{
      id: 'temp_manual_' + Date.now(),
      baseDate: pfBaseDate || getTodayString(),
      bank: defaultBank,
      purpose: defaultPurpose,
      name: '',
      code: '',
      currency: 'KRW',
      avgPrice: 0,
      currentPrice: 0,
      qty: 0,
      purchaseAmount: 0,
      currentAmount: 0,
      evalProfitLoss: 0,
      sellProfitLoss: 0,
      profitRate: 0,
      isManual: true
    }, ...prev]);
  };

  const handleAddRow = () => {
    setPayments(prev => [{
      id: 'temp_' + Date.now(),
      date: getTodayString(),
      bank: '',
      purpose: '연금',
      amount: 0,
    }, ...prev]);
  };

  const handleAddStockRow = () => {
    setStocks(prev => [...prev, {
      id: 'temp_stock_' + Date.now(),
      bank: '미래에셋',
      purpose: '연금',
      code: '',
      name: '',
      category: '',
      ratio: 0,
      currency: 'KRW',
    }]);
  };

  const handleAddTransactionRow = () => {
    setTransactions(prev => [{
      id: 'temp_tx_' + Date.now(),
      date: getTodayString(),
      bank: '',
      purpose: '연금',
      name: '',
      code: '',
      currency: 'KRW',
      price: 0,
      buyQty: 0,
      sellQty: 0,
    }, ...prev]);
  };

  const handleDeleteRows = () => {
    if (selectedIds.length === 0) return;
    setDeletedIds(prev => [...prev, ...selectedIds.filter(id => !String(id).startsWith('temp_') && !String(id).startsWith('excel_'))]);
    setPayments(prev => prev.filter(p => !selectedIds.includes(p.id)));
    setSelectedIds([]);
  };

  const handleDeleteStockRows = () => {
    if (selectedStockIds.length === 0) return;
    setDeletedStockIds(prev => [...prev, ...selectedStockIds.filter(id => !String(id).startsWith('temp_stock_') && !String(id).startsWith('excel_stock_'))]);
    setStocks(prev => prev.filter(s => !selectedStockIds.includes(s.id)));
    setSelectedStockIds([]);
  };

  const handleDeleteTransactionRows = () => {
    if (selectedTransactionIds.length === 0) return;
    setDeletedTransactionIds(prev => [...prev, ...selectedTransactionIds.filter(id => !String(id).startsWith('temp_tx_') && !String(id).startsWith('excel_tx_'))]);
    setTransactions(prev => prev.filter(t => !selectedTransactionIds.includes(t.id)));
    setSelectedTransactionIds([]);
  };

  const handleDeletePortfolioRows = () => {
    if (selectedPortfolioIds.length === 0) return;
    setDeletedPortfolioIds(prev => [...prev, ...selectedPortfolioIds.filter(id => !String(id).startsWith('temp_pf_') && !String(id).startsWith('pf_') && !String(id).startsWith('temp_manual_') && !String(id).startsWith('manual_'))]);
    setPortfolios(prev => prev.filter(p => !selectedPortfolioIds.includes(p.id)));
    setSelectedPortfolioIds([]);
  };

  const handleDeleteTrendRows = () => {
    if (selectedTrendIds.length === 0) return;
    setDeletedTrendIds(prev => [...prev, ...selectedTrendIds.filter(id => !String(id).startsWith('temp_trend_') && !String(id).startsWith('trend_'))]);
    setTrendRows(prev => prev.filter(t => !selectedTrendIds.includes(t.id)));
    setSelectedTrendIds([]);
  };

  const triggerExcelUpload = (type) => {
    setActiveUploadType(type);
    const refMap = { 
      payment: fileInputRef, 
      stock: stockFileInputRef, 
      transaction: transactionFileInputRef,
      trend: trendFileInputRef
    };
    if (refMap[type].current) {
      refMap[type].current.value = "";
      refMap[type].current.click();
    }
  };

  const showError = (message) => {
    setErrorModal({ isOpen: true, message });
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

        if (jsonData.length < 2) {
          showError('엑셀 시트에 데이터가 부족합니다.');
          return;
        }

        const headers = jsonData[0];
        const dataRows = jsonData.slice(1).filter(row => row && row.length > 0);

        if (activeUploadType === 'payment') {
          const excelRows = dataRows.map((row, index) => ({
            id: 'excel_' + Date.now() + '_' + index,
            date: parseExcelDate(row[headers.indexOf('납입일자')]),
            bank: String(row[headers.indexOf('은행')] || '').trim(),
            purpose: String(row[headers.indexOf('목적')] || '연금').trim(),
            amount: parseInt(String(row[headers.indexOf('금액')] || '0').replace(/[^0-9]/g, ''), 10) || 0,
          }));
          setPayments(prev => [...excelRows, ...prev]);
        } else if (activeUploadType === 'stock') {
          const excelRows = dataRows.map((row, index) => ({
            id: 'excel_stock_' + Date.now() + '_' + index,
            bank: String(row[headers.indexOf('은행')] || '').trim(),
            purpose: String(row[headers.indexOf('목적')] || '연금').trim(),
            code: String(row[headers.indexOf('종목코드')] || '').trim(),
            name: String(row[headers.indexOf('종목명')] || '').trim(),
            category: String(row[headers.indexOf('종목 유형')] || '').trim(),
            ratio: parseFloat(String(row[headers.indexOf('유형 비율')] || '0').replace(/[^0-9.]/g, '')) || 0,
            currency: String(row[headers.indexOf('통화')] || 'KRW').trim(),
          }));
          setStocks(prev => [...excelRows, ...prev]);
        } else if (activeUploadType === 'trend') {
          const uploadedRows = dataRows.map((row, index) => {
            const dt = parseExcelDate(row[headers.indexOf('일자')]);
            const amt = parseFloat(String(row[headers.indexOf('금액')] || '0').replace(/[^0-9.]/g, '')) || 0;
            return {
              id: 'excel_trend_' + Date.now() + '_' + index,
              date: dt,
              amount: amt
            };
          });

          setTrendRows(prev => {
            const map = new Map();
            prev.forEach(r => map.set(r.date, r));
            uploadedRows.forEach(r => {
              map.set(r.date, { ...r, id: map.has(r.date) ? map.get(r.date).id : r.id });
            });
            return Array.from(map.values()).sort((a, b) => b.date.localeCompare(a.date));
          });
        } else {
          const excelRows = dataRows.map((row, index) => {
            const bankVal = String(row[headers.indexOf('은행')] || '').trim();
            const purposeVal = String(row[headers.indexOf('목적')] || '연금').trim();
            const nameVal = String(row[headers.indexOf('종목명')] || '').trim();
            const matchedStock = stocks.find(s => (s.bank || '').trim() === bankVal && (s.purpose || '').trim() === purposeVal && (s.name || '').trim() === nameVal);
            return {
              id: 'excel_tx_' + Date.now() + '_' + index,
              date: parseExcelDate(row[headers.indexOf('거래일자')]),
              bank: bankVal,
              purpose: purposeVal,
              name: nameVal,
              code: matchedStock ? matchedStock.code : '',
              currency: matchedStock ? (matchedStock.currency || 'KRW') : 'KRW',
              price: parseFloat(String(row[headers.indexOf('단가')] || '0').replace(/[^0-9.]/g, '')) || 0,
              buyQty: parseFloat(String(row[headers.indexOf('매수수량')] || '0').replace(/[^0-9.]/g, '')) || 0,
              sellQty: parseFloat(String(row[headers.indexOf('매도수량')] || '0').replace(/[^0-9.]/g, '')) || 0,
            };
          });
          setTransactions(prev => [...excelRows, ...prev]);
        }
        setSuccessMessage('엑셀이 로드되었습니다. [저장] 버튼을 눌러 DB에 반영하세요.');
        setTimeout(() => setSuccessMessage(''), 3000);
      } catch (err) {
        showError('엑셀 파일을 읽는 중 오류가 발생했습니다.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSaveToDatabase = async () => {
    if (isSaving) return;
    setIsSaving(true);
    try {
      const batch = writeBatch(db);
      deletedIds.forEach(id => batch.delete(doc(db, "payments", id)));
      payments.forEach(p => {
        const ref = String(p.id).startsWith('temp_') || String(p.id).startsWith('excel_') ? doc(collection(db, "payments")) : doc(db, "payments", p.id);
        const data = { date: p.date, bank: p.bank, purpose: p.purpose, amount: Number(p.amount || 0) };
        if (String(p.id).startsWith('temp_') || String(p.id).startsWith('excel_')) data.createdAt = serverTimestamp();
        batch.set(ref, data, { merge: true });
      });
      await batch.commit();
      await fetchPayments();
      setSuccessMessage('납입금액 저장 완료!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (err) {
      showError(`저장 실패: ${err.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSaveStocksToDatabase = async () => {
    if (isSavingStocks) return;
    setIsSavingStocks(true);
    try {
      const batch = writeBatch(db);
      deletedStockIds.forEach(id => batch.delete(doc(db, "stocks", id)));
      stocks.forEach(s => {
        const ref = String(s.id).startsWith('temp_stock_') || String(s.id).startsWith('excel_stock_') ? doc(collection(db, "stocks")) : doc(db, "stocks", s.id);
        const data = { bank: s.bank, purpose: s.purpose, code: s.code, name: s.name, category: s.category, ratio: Number(s.ratio || 0), currency: s.currency || 'KRW' };
        if (String(s.id).startsWith('temp_stock_') || String(s.id).startsWith('excel_stock_')) data.createdAt = serverTimestamp();
        batch.set(ref, data, { merge: true });
      });
      await batch.commit();
      await fetchStocks();
      setSuccessMessage('종목 관리 저장 완료!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (err) {
      showError(`저장 실패: ${err.message}`);
    } finally {
      setIsSavingStocks(false);
    }
  };

  const handleSaveTransactionsToDatabase = async () => {
    if (isSavingTransactions) return;
    setIsSavingTransactions(true);
    try {
      const batch = writeBatch(db);
      deletedTransactionIds.forEach(id => batch.delete(doc(db, "transactions", id)));
      transactions.forEach(t => {
        const ref = String(t.id).startsWith('temp_tx_') || String(t.id).startsWith('excel_tx_') ? doc(collection(db, "transactions")) : doc(db, "transactions", t.id);
        const data = { date: t.date, bank: t.bank, purpose: t.purpose, name: t.name, code: t.code, currency: t.currency || 'KRW', price: Number(t.price || 0), buyQty: Number(t.buyQty || 0), sellQty: Number(t.sellQty || 0) };
        if (String(t.id).startsWith('temp_tx_') || String(t.id).startsWith('excel_tx_')) data.createdAt = serverTimestamp();
        batch.set(ref, data, { merge: true });
      });
      await batch.commit();
      await fetchTransactions();
      setSuccessMessage('거래현황 저장 완료!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (err) {
      showError(`저장 실패: ${err.message}`);
    } finally {
      setIsSavingTransactions(false);
    }
  };

  const handleSavePortfoliosToDatabase = async () => {
    if (isSavingPortfolios) return;
    setIsSavingPortfolios(true);
    try {
      const batch = writeBatch(db);
      deletedPortfolioIds.forEach(id => batch.delete(doc(db, "portfolios", id)));
      portfolios.forEach(pf => {
        const ref = String(pf.id).startsWith('temp_pf_') || String(pf.id).startsWith('pf_') || String(pf.id).startsWith('temp_manual_') || String(pf.id).startsWith('manual_') ? doc(collection(db, "portfolios")) : doc(db, "portfolios", pf.id);
        const data = { 
          baseDate: pf.baseDate, 
          bank: pf.bank, 
          purpose: pf.purpose, 
          name: pf.name, 
          code: pf.code || '', 
          currency: pf.currency || 'KRW', 
          avgPrice: pf.avgPrice || 0, 
          currentPrice: pf.currentPrice || 0, 
          qty: pf.qty || 0, 
          purchaseAmount: pf.purchaseAmount || 0, 
          currentAmount: pf.currentAmount || 0, 
          evalProfitLoss: pf.evalProfitLoss || 0, 
          sellProfitLoss: pf.sellProfitLoss || 0, 
          profitRate: pf.profitRate || 0,
          isManual: !!pf.isManual
        };
        if (String(pf.id).startsWith('temp_pf_') || String(pf.id).startsWith('pf_') || String(pf.id).startsWith('temp_manual_') || String(pf.id).startsWith('manual_')) data.createdAt = serverTimestamp();
        batch.set(ref, data, { merge: true });
      });
      await batch.commit();
      await fetchPortfolios();
      setSuccessMessage('포트폴리오 저장 완료!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (err) {
      showError(`저장 실패: ${err.message}`);
    } finally {
      setIsSavingPortfolios(false);
    }
  };

  const handleSaveTrendToDatabase = async () => {
    if (isSavingTrend) return;
    setIsSavingTrend(true);
    try {
      const batch = writeBatch(db);
      deletedTrendIds.forEach(id => batch.delete(doc(db, "trends", id)));
      trendRows.forEach(tr => {
        const ref = String(tr.id).startsWith('temp_trend_') || String(tr.id).startsWith('trend_') || String(tr.id).startsWith('excel_trend_') ? doc(collection(db, "trends")) : doc(db, "trends", tr.id);
        const data = { date: tr.date, amount: Number(tr.amount || 0) };
        if (String(tr.id).startsWith('temp_trend_') || String(tr.id).startsWith('trend_') || String(tr.id).startsWith('excel_trend_')) data.createdAt = serverTimestamp();
        batch.set(ref, data, { merge: true });
      });
      await batch.commit();
      await fetchTrends();
      setSuccessMessage('총액 Trend 저장 완료!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (err) {
      showError(`저장 실패: ${err.message}`);
    } finally {
      setIsSavingTrend(false);
    }
  };

  const handleRowChange = (id, field, val) => setPayments(prev => prev.map(p => p.id === id ? { ...p, [field]: val } : p));
  const handleStockRowChange = (id, field, val) => {
    setStocks(prev => prev.map(s => {
      if (s.id === id) {
        const updated = { ...s, [field]: val };
        if (field === 'code') {
          const cUpper = (val || '').trim().toUpperCase();
          if (/^[A-Z]+$/.test(cUpper) && cUpper.length <= 5) {
            updated.currency = 'USD';
          } else if (/^\d{6}$/.test(cUpper)) {
            updated.currency = 'KRW';
          }
        }
        return updated;
      }
      return s;
    }));
  };
  
  const handleTransactionRowChange = (id, field, val) => {
    setTransactions(prev => prev.map(t => {
      if (t.id === id) {
        const updated = { ...t, [field]: val };
        if (field === 'bank' || field === 'purpose' || field === 'name' || field === 'code') {
          const matched = stocks.find(s => 
            (s.bank || '').trim() === (field === 'bank' ? val : updated.bank).trim() &&
            (s.purpose || '').trim() === (field === 'purpose' ? val : updated.purpose).trim() &&
            ((field === 'name' ? val : updated.name) ? (s.name || '').trim() === (field === 'name' ? val : updated.name).trim() : true) &&
            ((field === 'code' ? val : updated.code) ? (s.code || '').trim() === (field === 'code' ? val : updated.code).trim() : true)
          );
          if (matched) {
            updated.code = matched.code || updated.code;
            updated.name = matched.name || updated.name;
            updated.currency = matched.currency || 'KRW';
          } else if (field === 'code') {
            const cUpper = (val || '').trim().toUpperCase();
            if (/^[A-Z]+$/.test(cUpper) && cUpper.length <= 5) {
              updated.currency = 'USD';
            } else if (/^\d{6}$/.test(cUpper)) {
              updated.currency = 'KRW';
            }
          }
        }
        return updated;
      }
      return t;
    }));
  };

  const handlePortfolioRowChange = (id, field, val) => {
    setPortfolios(prev => prev.map(pf => {
      if (pf.id === id) {
        const updated = { ...pf, [field]: val };
        if (pf.isManual) {
          if (field === 'bank' || field === 'purpose') {
            updated.name = '';
          }
          if (field === 'currentAmount' || field === 'purchaseAmount') {
            const curAmt = field === 'currentAmount' ? Number(val || 0) : Number(pf.currentAmount || 0);
            const purAmt = field === 'purchaseAmount' ? Number(val || 0) : Number(pf.purchaseAmount || 0);
            updated.evalProfitLoss = curAmt - purAmt;
          }
        } else {
          if (field === 'currentPrice') {
            const newCurrentPrice = Number(val || 0);
            updated.currentAmount = updated.qty * newCurrentPrice;
            updated.evalProfitLoss = updated.currentAmount - updated.purchaseAmount;
            updated.profitRate = updated.purchaseAmount > 0 ? updated.evalProfitLoss / updated.purchaseAmount : 0;
          }
        }
        return updated;
      }
      return pf;
    }));
  };

  const handleTrendRowChange = (id, field, val) => {
    setTrendRows(prev => prev.map(tr => tr.id === id ? { ...tr, [field]: val } : tr));
  };

  const filteredPayments = useMemo(() => appliedSearchYear ? payments.filter(p => p.date && p.date.startsWith(appliedSearchYear)) : payments, [payments, appliedSearchYear]);
  
  const filteredStocks = useMemo(() => {
    return [...stocks].sort((a, b) => {
      const bankA = (a.bank || '').trim();
      const bankB = (b.bank || '').trim();
      if (bankA !== bankB) return bankA.localeCompare(bankB, 'ko');
      const purposeA = (a.purpose || '').trim();
      const purposeB = (b.purpose || '').trim();
      if (purposeA !== purposeB) return purposeA.localeCompare(purposeB, 'ko');
      const codeA = (a.code || '').trim();
      const codeB = (b.code || '').trim();
      return codeA.localeCompare(codeB, 'ko');
    });
  }, [stocks]);

  const filteredTransactions = useMemo(() => {
    return transactions.filter(t => {
      if (!t.date) return true;
      if (appliedTxStartDate && t.date < appliedTxStartDate) return false;
      if (appliedTxEndDate && t.date > appliedTxEndDate) return false;
      return true;
    }).sort((a, b) => {
      if ((a.date || '') !== (b.date || '')) return (b.date || '').localeCompare(a.date || '');
      if ((a.bank || '') !== (b.bank || '')) return (a.bank || '').localeCompare(b.bank || '', 'ko');
      if ((a.purpose || '') !== (b.purpose || '')) return (a.purpose || '').localeCompare(b.purpose || '', 'ko');
      return (a.code || '').localeCompare(b.code || '', 'ko');
    });
  }, [transactions, appliedTxStartDate, appliedTxEndDate]);

  const filteredPortfolios = useMemo(() => portfolios.filter(pf => {
    if (appliedPfBaseDate && pf.baseDate !== appliedPfBaseDate) return false;
    if (appliedPfBankFilter && pf.bank !== appliedPfBankFilter) return false;
    if (appliedPfPurposeFilter && pf.purpose !== appliedPfPurposeFilter) return false;
    return true;
  }), [portfolios, appliedPfBaseDate, appliedPfBankFilter, appliedPfPurposeFilter]);

  // 총액 Trend 관련 요약 및 계산 (요청사항 2 반영: 기존 납입총액 + 시작일자보다 적은 날짜 중에 가장 가까운 일자의 납입 누적액)
  const trendTotalPayment = useMemo(() => {
    if (!appliedTrendStartDate || !appliedTrendEndDate) return 0;
    
    // 1. 기간 내 납입 총액
    const inRangeSum = payments
      .filter(p => p.date && p.date >= appliedTrendStartDate && p.date <= appliedTrendEndDate)
      .reduce((sum, p) => sum + Number(p.amount || 0), 0);

    // 2. 시작일자보다 적은 날짜 중에 가장 가까운 일자의 납입 누적액(또는 해당일까지의 총합)
    const priorPayments = payments.filter(p => p.date && p.date < appliedTrendStartDate);
    let priorSum = 0;
    if (priorPayments.length > 0) {
      // 날짜 기준 내림차순 정렬하여 가장 가까운(최신) 과거 일자 찾기
      priorPayments.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      const nearestDate = priorPayments[0].date;
      // 해당 날짜 이하(또는 해당 날짜)의 모든 누적 납입액
      priorSum = payments
        .filter(p => p.date && p.date <= nearestDate)
        .reduce((sum, p) => sum + Number(p.amount || 0), 0);
    }

    return inRangeSum + priorSum;
  }, [payments, appliedTrendStartDate, appliedTrendEndDate]);

  const trendTargetAmount = useMemo(() => {
    const rate = Number(trendProfitRateInput || 0);
    return trendTotalPayment * (1 + rate);
  }, [trendTotalPayment, trendProfitRateInput]);

  const sortedTrendRows = useMemo(() => {
    return [...trendRows].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  }, [trendRows]);

  const availableBanks = useMemo(() => Array.from(new Set(stocks.map(s => s.bank?.trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'ko')), [stocks]);
  const availablePurposes = useMemo(() => Array.from(new Set(stocks.map(s => s.purpose?.trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'ko')), [stocks]);

  if (!isTailwindLoaded) return <div className="flex justify-center items-center h-screen text-slate-500 font-sans">준비 중...</div>;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans antialiased">
      <input type="file" accept=".xlsx, .xls" className="hidden" ref={fileInputRef} onChange={handleFileChange} />
      <input type="file" accept=".xlsx, .xls" className="hidden" ref={stockFileInputRef} onChange={handleFileChange} />
      <input type="file" accept=".xlsx, .xls" className="hidden" ref={transactionFileInputRef} onChange={handleFileChange} />
      <input type="file" accept=".xlsx, .xls" className="hidden" ref={trendFileInputRef} onChange={handleFileChange} />

      {errorModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-6">
            <div className="flex items-center gap-3 text-rose-600 mb-4"><AlertCircle size={28} /><h3 className="text-xl font-bold">오류 안내</h3></div>
            <div className="text-slate-600 mb-6 whitespace-pre-line leading-relaxed">{errorModal.message}</div>
            <div className="flex justify-end"><button onClick={() => setErrorModal({ isOpen: false, message: '' })} className="px-6 py-2 bg-slate-800 text-white rounded-lg font-medium hover:bg-slate-700">확인</button></div>
          </div>
        </div>
      )}

      {successMessage && <div className="fixed top-5 right-5 z-50 bg-emerald-600 text-white px-6 py-3 rounded-xl shadow-lg font-medium">{successMessage}</div>}

      <header className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-3 text-indigo-600"><Landmark size={28} /><h1 className="text-xl font-bold text-slate-900">주식 포트폴리오 관리 시스템</h1></div>
        <div className="text-sm text-emerald-600 font-medium bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200 flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>Firebase 클라우드 연동 완료</div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 sm:px-6 lg:px-8">
        <div className="flex space-x-1 border-b border-slate-200 mb-6 overflow-x-auto">
          <button onClick={() => setActiveTab('payment')} className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 whitespace-nowrap ${activeTab === 'payment' ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}><Wallet size={18} />납입금액 관리</button>
          <button onClick={() => setActiveTab('stock')} className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 whitespace-nowrap ${activeTab === 'stock' ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}><Layers size={18} />종목 관리</button>
          <button onClick={() => setActiveTab('transaction')} className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 whitespace-nowrap ${activeTab === 'transaction' ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}><ArrowLeftRight size={18} />거래현황</button>
          <button onClick={() => setActiveTab('portfolio')} className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 whitespace-nowrap ${activeTab === 'portfolio' ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}><PiggyBank size={18} />포트폴리오 현황</button>
          <button onClick={() => setActiveTab('trend')} className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 whitespace-nowrap ${activeTab === 'trend' ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}><TrendingUp size={18} />총액 Trend</button>
        </div>

        {/* 1. 납입금액 관리 탭 */}
        {activeTab === 'payment' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-start xl:items-center">
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="relative"><Calendar size={16} className="absolute left-3 top-3 text-slate-400" /><input type="text" placeholder="납입연도 (예: 2026)" className="pl-10 pr-4 py-2 w-44 border border-slate-300 rounded-lg text-sm bg-white" value={searchYearInput} onChange={e => setSearchYearInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && setAppliedSearchYear(searchYearInput)} /></div>
                <button onClick={() => setAppliedSearchYear(searchYearInput)} className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700"><Search size={16} />조회</button>
                <button onClick={handleAddRow} className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700"><Plus size={16} />추가</button>
                <button onClick={handleDeleteRows} disabled={selectedIds.length === 0} className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium ${selectedIds.length > 0 ? 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}><Trash2 size={16} />삭제 {selectedIds.length > 0 && `(${selectedIds.length})`}</button>
                <button onClick={() => triggerExcelUpload('payment')} className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50"><FileSpreadsheet size={16} className="text-green-600" />엑셀 업로드</button>
                <button onClick={handleSaveToDatabase} disabled={isSaving} className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700"><Save size={16} />{isSaving ? '저장 중...' : '저장'}</button>
              </div>
              <div className="flex items-center gap-6 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm w-full xl:w-auto">
                <div className="flex flex-col"><span className="text-xs text-slate-500 font-medium">납입 총액 (전체)</span><span className="text-lg font-bold text-slate-800">{formatCurrency(payments.reduce((s, c) => s + Number(c.amount || 0), 0))}원</span></div>
                <div className="w-px h-10 bg-slate-200"></div>
                <div className="flex flex-col"><span className="text-xs text-indigo-500 font-medium">조회 총액 (현재 화면)</span><span className="text-lg font-bold text-indigo-700">{formatCurrency(filteredPayments.reduce((s, c) => s + Number(c.amount || 0), 0))}원</span></div>
              </div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[800px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={filteredPayments.length > 0 && selectedIds.length === filteredPayments.length} onChange={e => setSelectedIds(e.target.checked ? filteredPayments.map(p => p.id) : [])} /></th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-48">납입일자</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-48">은행</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-48">목적</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm text-right">금액 (원)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredPayments.length > 0 ? filteredPayments.map(row => (
                    <tr key={row.id} className={`hover:bg-slate-50 ${selectedIds.includes(row.id) ? 'bg-indigo-50/30' : ''}`}>
                      <td className="py-3 px-4"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={selectedIds.includes(row.id)} onChange={() => setSelectedIds(prev => prev.includes(row.id) ? prev.filter(i => i !== row.id) : [...prev, row.id])} /></td>
                      <td className="py-2 px-4"><input type="date" value={row.date || ''} onChange={e => handleRowChange(row.id, 'date', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                      <td className="py-2 px-4"><input type="text" value={row.bank || ''} onChange={e => handleRowChange(row.id, 'bank', e.target.value)} placeholder="은행 입력" className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                      <td className="py-2 px-4"><select value={row.purpose || '연금'} onChange={e => handleRowChange(row.id, 'purpose', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white"><option value="연금">연금</option><option value="IRP">IRP</option><option value="DC">DC</option><option value="기타">기타</option></select></td>
                      <td className="py-2 px-4 text-right"><input type="text" value={row.amount === 0 ? '' : formatCurrency(row.amount)} onChange={e => handleRowChange(row.id, 'amount', parseInt(e.target.value.replace(/[^0-9]/g, ''), 10) || 0)} placeholder="0" className="w-full md:w-3/4 px-3 py-1.5 border rounded-md text-sm text-right font-medium bg-white" /></td>
                    </tr>
                  )) : <tr><td colSpan="5" className="py-16 text-center text-slate-500">조회된 납입 내역이 없습니다.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 2. 종목 관리 탭 */}
        {activeTab === 'stock' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center">
              <div className="flex flex-wrap items-center gap-2.5">
                <button onClick={fetchStocks} className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700"><Search size={16} />조회</button>
                <button onClick={handleAddStockRow} className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700"><Plus size={16} />추가</button>
                <button onClick={handleDeleteStockRows} disabled={selectedStockIds.length === 0} className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium ${selectedStockIds.length > 0 ? 'bg-rose-50 text-rose-600 border border-rose-200' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}><Trash2 size={16} />삭제 {selectedStockIds.length > 0 && `(${selectedStockIds.length})`}</button>
                <button onClick={() => triggerExcelUpload('stock')} className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50"><FileSpreadsheet size={16} className="text-green-600" />엑셀 업로드</button>
                <button onClick={handleSaveStocksToDatabase} disabled={isSavingStocks} className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700"><Save size={16} />{isSavingStocks ? '저장 중...' : '저장'}</button>
              </div>
              <div className="flex items-center gap-4 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm"><span className="text-xs text-slate-500 font-medium">등록된 종목 수:</span><span className="text-base font-bold text-slate-800">{stocks.length} 개</span></div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={stocks.length > 0 && selectedStockIds.length === stocks.length} onChange={e => setSelectedStockIds(e.target.checked ? stocks.map(s => s.id) : [])} /></th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-44">은행</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-44">목적</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36">종목코드</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm">종목명</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-40">종목 유형</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-28 text-right">유형 비율 (%)</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-28 text-center">통화</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredStocks.length > 0 ? filteredStocks.map(stock => (
                    <tr key={stock.id} className={`hover:bg-slate-50 ${selectedStockIds.includes(stock.id) ? 'bg-indigo-50/30' : ''}`}>
                      <td className="py-3 px-4"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={selectedStockIds.includes(stock.id)} onChange={() => setSelectedStockIds(prev => prev.includes(stock.id) ? prev.filter(i => i !== stock.id) : [...prev, stock.id])} /></td>
                      <td className="py-2 px-4"><select value={stock.bank || ''} onChange={e => handleStockRowChange(stock.id, 'bank', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white"><option value="">은행 선택</option><option value="미래에셋">미래에셋</option><option value="KB증권">KB증권</option><option value="삼성증권">삼성증권</option></select></td>
                      <td className="py-2 px-4"><select value={stock.purpose || '연금'} onChange={e => handleStockRowChange(stock.id, 'purpose', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white"><option value="연금">연금</option><option value="IRP">IRP</option><option value="DC">DC</option><option value="기타">기타</option></select></td>
                      <td className="py-2 px-4"><input type="text" value={stock.code || ''} onChange={e => handleStockRowChange(stock.id, 'code', e.target.value)} placeholder="종목코드" className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                      <td className="py-2 px-4"><input type="text" value={stock.name || ''} onChange={e => handleStockRowChange(stock.id, 'name', e.target.value)} placeholder="종목명" className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                      <td className="py-2 px-4"><input type="text" value={stock.category || ''} onChange={e => handleStockRowChange(stock.id, 'category', e.target.value)} placeholder="종목 유형" className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                      <td className="py-2 px-4 text-right"><input type="number" value={stock.ratio || 0} onChange={e => handleStockRowChange(stock.id, 'ratio', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm text-right bg-white" /></td>
                      <td className="py-2 px-4 text-center"><input type="text" value={stock.currency || 'KRW'} onChange={e => handleStockRowChange(stock.id, 'currency', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm text-center bg-white font-medium" /></td>
                    </tr>
                  )) : <tr><td colSpan="8" className="py-16 text-center text-slate-500">등록된 종목 내역이 없습니다.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 3. 거래현황 탭 */}
        {activeTab === 'transaction' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center">
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5"><Calendar size={16} className="text-slate-400" /><input type="date" value={txStartDate} onChange={e => setTxStartDate(e.target.value)} className="text-sm outline-none bg-transparent" /><span className="text-slate-400">~</span><input type="date" value={txEndDate} onChange={e => setTxEndDate(e.target.value)} className="text-sm outline-none bg-transparent" /></div>
                <button onClick={() => { setAppliedTxStartDate(txStartDate); setAppliedTxEndDate(txEndDate); }} className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700"><Search size={16} />조회</button>
                <button onClick={handleAddTransactionRow} className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700"><Plus size={16} />추가</button>
                <button onClick={handleDeleteTransactionRows} disabled={selectedTransactionIds.length === 0} className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium ${selectedTransactionIds.length > 0 ? 'bg-rose-50 text-rose-600 border border-rose-200' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}><Trash2 size={16} />삭제 {selectedTransactionIds.length > 0 && `(${selectedTransactionIds.length})`}</button>
                <button onClick={() => triggerExcelUpload('transaction')} className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50"><FileSpreadsheet size={16} className="text-green-600" />엑셀 업로드</button>
                <button onClick={handleSaveTransactionsToDatabase} disabled={isSavingTransactions} className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700"><Save size={16} />{isSavingTransactions ? '저장 중...' : '저장'}</button>
              </div>
              <div className="flex items-center gap-4 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm"><span className="text-xs text-slate-500 font-medium">조회 건수:</span><span className="text-base font-bold text-slate-800">{filteredTransactions.length} 건</span></div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[1100px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={filteredTransactions.length > 0 && selectedTransactionIds.length === filteredTransactions.length} onChange={e => setSelectedTransactionIds(e.target.checked ? filteredTransactions.map(t => t.id) : [])} /></th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-40">거래일자</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36">은행</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-32">목적</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm">종목명</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-32">종목코드</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-24 text-center">통화</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-32 text-right">단가</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-28 text-right">매수수량</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-28 text-right">매도수량</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredTransactions.length > 0 ? filteredTransactions.map(tx => {
                    const rowNames = Array.from(new Set(stocks.filter(s => (s.bank || '').trim() === (tx.bank || '').trim() && (s.purpose || '').trim() === (tx.purpose || '').trim()).map(s => s.name?.trim()).filter(Boolean)));
                    return (
                      <tr key={tx.id} className={`hover:bg-slate-50 ${selectedTransactionIds.includes(tx.id) ? 'bg-indigo-50/30' : ''}`}>
                        <td className="py-3 px-4"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={selectedTransactionIds.includes(tx.id)} onChange={() => setSelectedTransactionIds(prev => prev.includes(tx.id) ? prev.filter(i => i !== tx.id) : [...prev, tx.id])} /></td>
                        <td className="py-2 px-4"><input type="date" value={tx.date || ''} onChange={e => handleTransactionRowChange(tx.id, 'date', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                        <td className="py-2 px-4"><select value={tx.bank || ''} onChange={e => handleTransactionRowChange(tx.id, 'bank', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white"><option value="">은행 선택</option>{availableBanks.map(b => <option key={b} value={b}>{b}</option>)}</select></td>
                        <td className="py-2 px-4"><select value={tx.purpose || '연금'} onChange={e => handleTransactionRowChange(tx.id, 'purpose', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white"><option value="연금">연금</option><option value="IRP">IRP</option><option value="DC">DC</option><option value="기타">기타</option></select></td>
                        <td className="py-2 px-4"><select value={tx.name || ''} onChange={e => handleTransactionRowChange(tx.id, 'name', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white"><option value="">종목명 선택</option>{rowNames.map(n => <option key={n} value={n}>{n}</option>)}</select></td>
                        <td className="py-2 px-4"><input type="text" value={tx.code || ''} onChange={e => handleTransactionRowChange(tx.id, 'code', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm bg-white" /></td>
                        <td className="py-2 px-4 text-center"><input type="text" value={tx.currency || 'KRW'} onChange={e => handleTransactionRowChange(tx.id, 'currency', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm text-center bg-white font-medium" /></td>
                        <td className="py-2 px-4 text-right"><input type="text" value={tx.price === 0 || tx.price === '' ? '' : formatCurrency(tx.price, tx.currency)} onChange={e => handleTransactionRowChange(tx.id, 'price', parseFloat(e.target.value.replace(/[^0-9.]/g, '')) || 0)} className="w-full px-3 py-1.5 border rounded-md text-sm text-right bg-white font-medium" /></td>
                        <td className="py-2 px-4 text-right"><input type="number" step="any" value={tx.buyQty === 0 ? '' : tx.buyQty} onChange={e => handleTransactionRowChange(tx.id, 'buyQty', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm text-right bg-white" /></td>
                        <td className="py-2 px-4 text-right"><input type="number" step="any" value={tx.sellQty === 0 ? '' : tx.sellQty} onChange={e => handleTransactionRowChange(tx.id, 'sellQty', e.target.value)} className="w-full px-3 py-1.5 border rounded-md text-sm text-right bg-white" /></td>
                      </tr>
                    );
                  }) : <tr><td colSpan="10" className="py-16 text-center text-slate-500">조회된 거래현황 내역이 없습니다.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 4. 포트폴리오 현황 탭 */}
        {activeTab === 'portfolio' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-center">
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5"><span className="text-xs text-slate-500 font-medium">기준일자</span><input type="date" value={pfBaseDate} onChange={e => setPfBaseDate(e.target.value)} className="text-sm outline-none bg-transparent font-medium" /></div>
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5"><span className="text-xs text-slate-500 font-medium">환율(USD/KRW)</span><input type="number" step="any" value={exchangeRate} onChange={e => setExchangeRate(parseFloat(e.target.value) || 0)} className="w-24 text-sm outline-none bg-transparent font-medium text-right" /></div>
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5"><span className="text-xs text-slate-500 font-medium">은행</span><select value={pfBankFilter} onChange={e => setPfBankFilter(e.target.value)} className="text-sm outline-none bg-transparent font-medium"><option value="">전체 은행</option>{availableBanks.map(b => <option key={b} value={b}>{b}</option>)}</select></div>
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5"><span className="text-xs text-slate-500 font-medium">목적</span><select value={pfPurposeFilter} onChange={e => setPfPurposeFilter(e.target.value)} className="text-sm outline-none bg-transparent font-medium"><option value="">전체 목적</option>{availablePurposes.map(p => <option key={p} value={p}>{p}</option>)}</select></div>
                <button onClick={() => { setAppliedPfBaseDate(pfBaseDate); setAppliedPfBankFilter(pfBankFilter); setAppliedPfPurposeFilter(pfPurposeFilter); }} className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700"><Search size={16} />조회</button>
                <button onClick={handleCalculatePortfolio} className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700"><Calculator size={16} />계산</button>
                <button onClick={handleAddPortfolioRow} className="flex items-center gap-1.5 px-3.5 py-2 bg-teal-600 text-white rounded-lg text-sm font-medium hover:bg-teal-700"><Plus size={16} />추가</button>
                <button onClick={handleDeletePortfolioRows} disabled={selectedPortfolioIds.length === 0} className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium ${selectedPortfolioIds.length > 0 ? 'bg-rose-50 text-rose-600 border border-rose-200' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}><Trash2 size={16} />삭제</button>
                <button onClick={handleSavePortfoliosToDatabase} disabled={isSavingPortfolios} className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700"><Save size={16} />{isSavingPortfolios ? '저장 중...' : '저장'}</button>
              </div>
              <div className="flex items-center gap-4 bg-white px-5 py-2 rounded-lg border border-slate-200 shadow-sm">
                <div className="flex flex-col"><span className="text-[11px] text-slate-500 font-medium">현재금액 총액</span><span className="text-sm font-bold text-slate-800">{formatCurrency(filteredPortfolios.reduce((s, c) => s + (Number(c.currentAmount || 0) * ((c.currency || 'KRW').toUpperCase() === 'USD' ? exchangeRate : 1)), 0))}원</span></div>
                <div className="w-px h-8 bg-slate-200"></div>
                <div className="flex flex-col"><span className="text-[11px] text-indigo-500 font-medium">평가손익 총액</span><span className={`text-sm font-bold ${filteredPortfolios.reduce((s, c) => s + (Number(c.evalProfitLoss || 0) * ((c.currency || 'KRW').toUpperCase() === 'USD' ? exchangeRate : 1)), 0) >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{formatCurrency(filteredPortfolios.reduce((s, c) => s + (Number(c.evalProfitLoss || 0) * ((c.currency || 'KRW').toUpperCase() === 'USD' ? exchangeRate : 1)), 0))}원</span></div>
                <div className="w-px h-8 bg-slate-200"></div>
                <div className="flex flex-col"><span className="text-[11px] text-emerald-600 font-medium">매매손익 총액</span><span className={`text-sm font-bold ${filteredPortfolios.reduce((s, c) => s + (Number(c.sellProfitLoss || 0) * ((c.currency || 'KRW').toUpperCase() === 'USD' ? exchangeRate : 1)), 0) >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{formatCurrency(filteredPortfolios.reduce((s, c) => s + (Number(c.sellProfitLoss || 0) * ((c.currency || 'KRW').toUpperCase() === 'USD' ? exchangeRate : 1)), 0))}원</span></div>
              </div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[1250px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={filteredPortfolios.length > 0 && selectedPortfolioIds.length === filteredPortfolios.length} onChange={e => setSelectedPortfolioIds(e.target.checked ? filteredPortfolios.map(p => p.id) : [])} /></th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36">은행</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-32">목적</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm">종목명</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-32 text-right">평균단가</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">현재단가</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-28 text-right">수량</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">매입금액</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">현재금액</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">평가손익</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">매매손익</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-28 text-right">이익율</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-24 text-center">통화</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredPortfolios.length > 0 ? filteredPortfolios.map(pf => {
                    const isUSD = (pf.currency || 'KRW').toUpperCase() === 'USD';
                    const multiplier = isUSD ? exchangeRate : 1;

                    if (pf.isManual) {
                      const availableManualNames = Array.from(new Set(
                        stocks
                          .filter(s => (s.bank || '').trim() === (pf.bank || '').trim() && (s.purpose || '').trim() === (pf.purpose || '').trim() && !s.code?.trim())
                          .map(s => s.name?.trim())
                          .filter(Boolean)
                      ));

                      return (
                        <tr key={pf.id} className={`hover:bg-slate-50 ${selectedPortfolioIds.includes(pf.id) ? 'bg-indigo-50/30' : ''}`}>
                          <td className="py-3 px-4"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={selectedPortfolioIds.includes(pf.id)} onChange={() => setSelectedPortfolioIds(prev => prev.includes(pf.id) ? prev.filter(i => i !== pf.id) : [...prev, pf.id])} /></td>
                          <td className="py-2 px-4"><select value={pf.bank || ''} onChange={e => handlePortfolioRowChange(pf.id, 'bank', e.target.value)} className="w-full px-2 py-1 border rounded text-sm bg-white"><option value="">선택</option>{availableBanks.map(b => <option key={b} value={b}>{b}</option>)}</select></td>
                          <td className="py-2 px-4"><select value={pf.purpose || '연금'} onChange={e => handlePortfolioRowChange(pf.id, 'purpose', e.target.value)} className="w-full px-2 py-1 border rounded text-sm bg-white"><option value="연금">연금</option><option value="IRP">IRP</option><option value="DC">DC</option><option value="기타">기타</option></select></td>
                          <td className="py-2 px-4">
                            <select 
                              value={pf.name || ''} 
                              onChange={e => setPortfolios(prev => prev.map(p => p.id === pf.id ? { ...p, name: e.target.value } : p))} 
                              className="w-full px-2 py-1 border rounded text-sm bg-white"
                            >
                              <option value="">종목명 선택</option>
                              {availableManualNames.map(n => <option key={n} value={n}>{n}</option>)}
                            </select>
                          </td>
                          <td className="py-3 px-4 text-sm text-right text-slate-400">-</td>
                          <td className="py-3 px-4 text-sm text-right text-slate-400">-</td>
                          <td className="py-3 px-4 text-sm text-right text-slate-400">-</td>
                          <td className="py-2 px-4 text-right">
                            <input 
                              type="text" 
                              value={pf.purchaseAmount === 0 || pf.purchaseAmount === '' ? '' : formatCurrency(pf.purchaseAmount, pf.currency)}
                              onChange={e => handlePortfolioRowChange(pf.id, 'purchaseAmount', parseFloat(e.target.value.replace(/[^0-9.]/g, '')) || 0)}
                              placeholder="0"
                              className="w-full px-2 py-1 border border-teal-300 rounded text-sm text-right font-medium bg-teal-50/30 text-teal-900 focus:outline-none focus:ring-1 focus:ring-teal-500"
                            />
                          </td>
                          <td className="py-2 px-4 text-right">
                            <input 
                              type="text" 
                              value={pf.currentAmount === 0 || pf.currentAmount === '' ? '' : formatCurrency(pf.currentAmount, pf.currency)}
                              onChange={e => handlePortfolioRowChange(pf.id, 'currentAmount', parseFloat(e.target.value.replace(/[^0-9.]/g, '')) || 0)}
                              placeholder="0"
                              className="w-full px-2 py-1 border border-teal-300 rounded text-sm text-right font-medium bg-teal-50/30 text-teal-900 focus:outline-none focus:ring-1 focus:ring-teal-500"
                            />
                          </td>
                          <td className={`py-3 px-4 text-sm text-right font-bold ${Number(pf.evalProfitLoss || 0) >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{formatCurrency(pf.evalProfitLoss, pf.currency)}</td>
                          <td className="py-3 px-4 text-sm text-right text-slate-400">-</td>
                          <td className="py-3 px-4 text-sm text-right text-slate-400">-</td>
                          <td className="py-3 px-4 text-sm text-center text-slate-600 font-medium">{pf.currency || 'KRW'}</td>
                        </tr>
                      );
                    }
                    return (
                      <tr key={pf.id} className={`hover:bg-slate-50 ${selectedPortfolioIds.includes(pf.id) ? 'bg-indigo-50/30' : ''}`}>
                        <td className="py-3 px-4"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={selectedPortfolioIds.includes(pf.id)} onChange={() => setSelectedPortfolioIds(prev => prev.includes(pf.id) ? prev.filter(i => i !== pf.id) : [...prev, pf.id])} /></td>
                        <td className="py-3 px-4 text-sm font-medium text-slate-800">{pf.bank}</td>
                        <td className="py-3 px-4 text-sm text-slate-600">{pf.purpose}</td>
                        <td className="py-3 px-4 text-sm font-medium text-slate-900">{pf.name} <span className="text-xs text-slate-400 font-normal">({pf.code})</span></td>
                        <td className="py-3 px-4 text-sm text-right text-slate-700 font-medium">{formatCurrency(pf.avgPrice, pf.currency)}</td>
                        <td className="py-2 px-4 text-right">
                          <input 
                            type="text" 
                            value={pf.currentPrice === 0 || pf.currentPrice === '' ? '' : formatCurrency(pf.currentPrice, pf.currency)}
                            onChange={e => handlePortfolioRowChange(pf.id, 'currentPrice', parseFloat(e.target.value.replace(/[^0-9.]/g, '')) || 0)}
                            placeholder="0"
                            className="w-full px-2 py-1 border border-indigo-200 rounded text-sm text-right font-medium bg-indigo-50/30 text-indigo-900 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                          />
                        </td>
                        <td className="py-3 px-4 text-sm text-right text-slate-700">{Number(pf.qty || 0).toLocaleString()}</td>
                        <td className="py-3 px-4 text-sm text-right text-slate-800 font-medium">{formatCurrency(pf.purchaseAmount * multiplier, isUSD ? 'KRW' : pf.currency)}{isUSD && <span className="block text-[11px] text-slate-400 font-normal">($ {formatCurrency(pf.purchaseAmount, 'USD')})</span>}</td>
                        <td className="py-3 px-4 text-sm text-right text-slate-900 font-semibold">{formatCurrency(pf.currentAmount * multiplier, isUSD ? 'KRW' : pf.currency)}{isUSD && <span className="block text-[11px] text-slate-400 font-normal">($ {formatCurrency(pf.currentAmount, 'USD')})</span>}</td>
                        <td className={`py-3 px-4 text-sm text-right font-bold ${Number(pf.evalProfitLoss || 0) >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{formatCurrency(pf.evalProfitLoss * multiplier, isUSD ? 'KRW' : pf.currency)}{isUSD && <span className="block text-[11px] font-normal opacity-75">($ {formatCurrency(pf.evalProfitLoss, 'USD')})</span>}</td>
                        <td className={`py-3 px-4 text-sm text-right font-bold ${Number(pf.sellProfitLoss || 0) >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{formatCurrency(pf.sellProfitLoss * multiplier, isUSD ? 'KRW' : pf.currency)}{isUSD && <span className="block text-[11px] font-normal opacity-75">($ {formatCurrency(pf.sellProfitLoss, 'USD')})</span>}</td>
                        <td className={`py-3 px-4 text-sm text-right font-bold ${Number(pf.profitRate || 0) >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{(Number(pf.profitRate || 0) * 100).toFixed(2)}%</td>
                        <td className="py-3 px-4 text-sm text-center text-slate-600 font-medium">{pf.currency}</td>
                      </tr>
                    );
                  }) : <tr><td colSpan="13" className="py-16 text-center text-slate-500">조회된 포트폴리오 현황이 없습니다. 상단의 [계산] 버튼을 눌러보세요.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 5. 총액 Trend 탭 */}
        {activeTab === 'trend' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-center">
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5">
                  <span className="text-xs text-slate-500 font-medium">기간</span>
                  <input type="date" value={trendStartDate} onChange={e => setTrendStartDate(e.target.value)} className="text-sm outline-none bg-transparent font-medium" />
                  <span className="text-slate-400">~</span>
                  <input type="date" value={trendEndDate} onChange={e => setTrendEndDate(e.target.value)} className="text-sm outline-none bg-transparent font-medium" />
                </div>
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5">
                  <span className="text-xs text-slate-500 font-medium">이익율</span>
                  <input type="number" step="any" value={trendProfitRateInput} onChange={e => setTrendProfitRateInput(parseFloat(e.target.value) || 0)} className="w-20 text-sm outline-none bg-transparent font-medium text-right" placeholder="0.08" />
                </div>
                <button onClick={() => { setAppliedTrendStartDate(trendStartDate); setAppliedTrendEndDate(trendEndDate); handleCalculateTrend(); }} className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700"><Search size={16} />조회</button>
                <button onClick={handleAddTrendRow} className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700"><Plus size={16} />추가</button>
                <button onClick={handleDeleteTrendRows} disabled={selectedTrendIds.length === 0} className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium ${selectedTrendIds.length > 0 ? 'bg-rose-50 text-rose-600 border border-rose-200' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}><Trash2 size={16} />삭제</button>
                <button onClick={() => triggerExcelUpload('trend')} className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50"><FileSpreadsheet size={16} className="text-green-600" />엑셀 업로드</button>
                <button onClick={handleSaveTrendToDatabase} disabled={isSavingTrend} className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700"><Save size={16} />{isSavingTrend ? '저장 중...' : '저장'}</button>
              </div>
              <div className="flex items-center gap-4 bg-white px-5 py-2 rounded-lg border border-slate-200 shadow-sm">
                <div className="flex flex-col"><span className="text-[11px] text-slate-500 font-medium">납입 총액</span><span className="text-sm font-bold text-slate-800">{formatCurrency(trendTotalPayment)}원</span></div>
                <div className="w-px h-8 bg-slate-200"></div>
                <div className="flex flex-col"><span className="text-[11px] text-indigo-500 font-medium">목표 금액</span><span className="text-sm font-bold text-indigo-600">{formatCurrency(trendTargetAmount)}원</span></div>
              </div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={trendRows.length > 0 && selectedTrendIds.length === trendRows.length} onChange={e => setSelectedTrendIds(e.target.checked ? trendRows.map(t => t.id) : [])} /></th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-48">일자</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-44 text-right">금액</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">비율</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-36 text-right">이익율</th>
                    <th className="py-3 px-4 border-b font-semibold text-slate-600 text-sm w-44 text-right">이익금액</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {trendRows.length > 0 ? trendRows.map((row, index) => {
                    const lastRow = sortedTrendRows.length > 0 ? sortedTrendRows[sortedTrendRows.length - 1] : null;
                    const baseAmountForRatio = lastRow ? Number(lastRow.amount || 0) : 0;

                    const ratio = baseAmountForRatio !== 0 ? (Number(row.amount || 0) - baseAmountForRatio) / baseAmountForRatio : 0;
                    const profitRate = trendTotalPayment !== 0 ? (Number(row.amount || 0) - trendTotalPayment) / trendTotalPayment : 0;
                    const profitAmount = Number(row.amount || 0) - trendTotalPayment;

                    return (
                      <tr key={row.id} className={`hover:bg-slate-50 ${selectedTrendIds.includes(row.id) ? 'bg-indigo-50/30' : ''}`}>
                        <td className="py-3 px-4"><input type="checkbox" className="w-4 h-4 rounded border-slate-300 text-indigo-600" checked={selectedTrendIds.includes(row.id)} onChange={() => setSelectedTrendIds(prev => prev.includes(row.id) ? prev.filter(i => i !== row.id) : [...prev, row.id])} /></td>
                        <td className="py-2 px-4"><input type="date" value={row.date || ''} onChange={e => handleTrendRowChange(row.id, 'date', e.target.value)} className="w-full px-2 py-1 border rounded text-sm bg-white" /></td>
                        <td className="py-2 px-4 text-right"><input type="text" value={row.amount === 0 || row.amount === '' ? '' : formatCurrency(row.amount)} onChange={e => handleTrendRowChange(row.id, 'amount', parseFloat(e.target.value.replace(/[^0-9.]/g, '')) || 0)} placeholder="0" className="w-full px-2 py-1 border rounded text-sm text-right font-medium bg-white" /></td>
                        <td className={`py-3 px-4 text-sm text-right font-medium ${ratio >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{(ratio * 100).toFixed(2)}%</td>
                        <td className={`py-3 px-4 text-sm text-right font-medium ${profitRate >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{(profitRate * 100).toFixed(2)}%</td>
                        <td className={`py-3 px-4 text-sm text-right font-bold ${profitAmount >= 0 ? 'text-rose-600' : 'text-blue-600'}`}>{formatCurrency(profitAmount)}원</td>
                      </tr>
                    );
                  }) : <tr><td colSpan="6" className="py-16 text-center text-slate-500">조회된 총액 Trend 내역이 없습니다. 기간 설정 후 [조회]를 누르거나 추가/엑셀업로드를 해주세요.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
