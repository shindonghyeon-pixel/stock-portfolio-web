import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  Search, 
  Plus, 
  Trash2, 
  FileSpreadsheet, 
  Calendar, 
  Wallet,
  Landmark,
  PiggyBank,
  CheckSquare,
  AlertCircle,
  Save,
  Layers,
  ArrowLeftRight
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

const formatCurrency = (amount) => {
  if (amount === 0 || !amount) return '0';
  return Number(amount).toLocaleString('ko-KR');
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

  // 2. 종목 관리 상태 ([은행], [목적], [종목코드], [종목명], [종목 유형], [유형 비율], [통화])
  const [stocks, setStocks] = useState([]);
  const [loadingStocks, setLoadingStocks] = useState(false);
  const [isSavingStocks, setIsSavingStocks] = useState(false);
  const [selectedStockIds, setSelectedStockIds] = useState([]);
  const [deletedStockIds, setDeletedStockIds] = useState([]);

  // 3. 거래현황 상태 (거래시작일자, 거래종료일자, 거래목록)
  const [transactions, setTransactions] = useState([]);
  const [loadingTransactions, setLoadingTransactions] = useState(false);
  const [isSavingTransactions, setIsSavingTransactions] = useState(false);
  const [selectedTransactionIds, setSelectedTransactionIds] = useState([]);
  const [deletedTransactionIds, setDeletedTransactionIds] = useState([]);
  const [txStartDate, setTxStartDate] = useState('');
  const [txEndDate, setTxEndDate] = useState('');
  const [appliedTxStartDate, setAppliedTxStartDate] = useState('');
  const [appliedTxEndDate, setAppliedTxEndDate] = useState('');

  const [errorModal, setErrorModal] = useState({ isOpen: false, message: '' });
  const [successMessage, setSuccessMessage] = useState('');
  
  const fileInputRef = useRef(null);
  const stockFileInputRef = useRef(null);
  const transactionFileInputRef = useRef(null);
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
  }, []);

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
      showError(`데이터를 불러오지 못했습니다.\n(${error.message})`);
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
      // 은행별, 목적별, 종목코드별 오름차순 정렬
      dataList.sort((a, b) => {
        const bankA = (a.bank || '').trim();
        const bankB = (b.bank || '').trim();
        if (bankA !== bankB) {
          return bankA.localeCompare(bankB, 'ko');
        }
        const purposeA = (a.purpose || '').trim();
        const purposeB = (b.purpose || '').trim();
        if (purposeA !== purposeB) {
          return purposeA.localeCompare(purposeB, 'ko');
        }
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
      dataList.sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0));
      setTransactions(dataList);
      setDeletedTransactionIds([]);
    } catch (error) {
      console.error("거래현황 데이터 불러오기 실패:", error);
    } finally {
      setLoadingTransactions(false);
    }
  };

  const handleSearch = () => {
    setAppliedSearchYear(searchYearInput);
    setSelectedIds([]); 
  };

  const handleTxSearch = () => {
    setAppliedTxStartDate(txStartDate);
    setAppliedTxEndDate(txEndDate);
    setSelectedTransactionIds([]);
  };

  const handleAddRow = () => {
    const newRow = {
      id: 'temp_' + Date.now(),
      date: getTodayString(),
      bank: '',
      purpose: '연금',
      amount: 0,
    };
    setPayments(prev => [newRow, ...prev]);
  };

  const handleAddStockRow = () => {
    const newStockRow = {
      id: 'temp_stock_' + Date.now(),
      bank: '',
      purpose: '연금',
      code: '',
      name: '',
      category: '',
      ratio: 0,
      currency: 'KRW',
    };
    setStocks(prev => [...prev, newStockRow]);
  };

  const handleAddTransactionRow = () => {
    const newTxRow = {
      id: 'temp_tx_' + Date.now(),
      date: getTodayString(),
      bank: '',
      purpose: '연금',
      name: '',
      code: '',
      price: 0,
      buyQty: 0,
      sellQty: 0,
    };
    setTransactions(prev => [newTxRow, ...prev]);
  };

  const handleDeleteRows = () => {
    if (selectedIds.length === 0) return;
    const targetRealIds = selectedIds.filter(id => !String(id).startsWith('temp_') && !String(id).startsWith('excel_'));
    setDeletedIds(prev => [...prev, ...targetRealIds]);
    setPayments(prev => prev.filter(p => !selectedIds.includes(p.id)));
    setSelectedIds([]);
  };

  const handleDeleteStockRows = () => {
    if (selectedStockIds.length === 0) return;
    const targetRealIds = selectedStockIds.filter(id => !String(id).startsWith('temp_stock_') && !String(id).startsWith('excel_stock_'));
    setDeletedStockIds(prev => [...prev, ...targetRealIds]);
    setStocks(prev => prev.filter(s => !selectedStockIds.includes(s.id)));
    setSelectedStockIds([]);
  };

  const handleDeleteTransactionRows = () => {
    if (selectedTransactionIds.length === 0) return;
    const targetRealIds = selectedTransactionIds.filter(id => !String(id).startsWith('temp_tx_') && !String(id).startsWith('excel_tx_'));
    setDeletedTransactionIds(prev => [...prev, ...targetRealIds]);
    setTransactions(prev => prev.filter(t => !selectedTransactionIds.includes(t.id)));
    setSelectedTransactionIds([]);
  };

  const triggerExcelUpload = (type) => {
    setActiveUploadType(type);
    if (type === 'payment' && fileInputRef.current) {
      fileInputRef.current.value = "";
      fileInputRef.current.click();
    } else if (type === 'stock' && stockFileInputRef.current) {
      stockFileInputRef.current.value = "";
      stockFileInputRef.current.click();
    } else if (type === 'transaction' && transactionFileInputRef.current) {
      transactionFileInputRef.current.value = "";
      transactionFileInputRef.current.click();
    }
  };

  const showError = (message) => {
    setErrorModal({ isOpen: true, message });
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const fileExt = file.name.split('.').pop().toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      showError('유효한 엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (workbook.SheetNames.length === 0) {
          showError('엑셀 파일에 시트가 존재하지 않습니다.');
          return;
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

        if (jsonData.length < 2) {
          showError('엑셀 시트에 데이터가 부족합니다.');
          return;
        }

        const headers = jsonData[0];

        if (activeUploadType === 'payment') {
          const requiredColumns = ['납입일자', '은행', '목적', '금액'];
          const missingColumns = requiredColumns.filter(col => !headers.includes(col));
          
          if (missingColumns.length > 0) {
            showError(`필수 열 누락: [ ${missingColumns.join(', ')} ]\n(1행에 '납입일자', '은행', '목적', '금액' 열이 있어야 합니다.)`);
            return;
          }

          const colIndices = {
            date: headers.indexOf('납입일자'),
            bank: headers.indexOf('은행'),
            purpose: headers.indexOf('목적'),
            amount: headers.indexOf('금액'),
          };

          const dataRows = jsonData.slice(1).filter(row => row && row.length > 0);
          
          const excelRows = dataRows.map((row, index) => {
            const rawDate = row[colIndices.date];
            const parsedDate = parseExcelDate(rawDate);
            const rawAmount = String(row[colIndices.amount] || '0').replace(/[^0-9]/g, '');
            const amountNum = parseInt(rawAmount, 10) || 0;

            return {
              id: 'excel_' + Date.now() + '_' + index,
              date: parsedDate,
              bank: String(row[colIndices.bank] || '').trim(),
              purpose: String(row[colIndices.purpose] || '연금').trim(),
              amount: amountNum,
            };
          });

          setPayments(prev => [...excelRows, ...prev]);
          setSuccessMessage('납입금액 엑셀이 로드되었습니다. [저장] 버튼을 눌러 DB에 반영하세요.');
        } else if (activeUploadType === 'stock') {
          const requiredColumns = ['은행', '목적', '종목코드', '종목명', '종목 유형', '유형 비율', '통화'];
          const missingColumns = requiredColumns.filter(col => !headers.includes(col));
          
          if (missingColumns.length > 0) {
            showError(`필수 열 누락: [ ${missingColumns.join(', ')} ]\n(1행에 '은행', '목적', '종목코드', '종목명', '종목 유형', '유형 비율', '통화' 열이 있어야 합니다.)`);
            return;
          }

          const colIndices = {
            bank: headers.indexOf('은행'),
            purpose: headers.indexOf('목적'),
            code: headers.indexOf('종목코드'),
            name: headers.indexOf('종목명'),
            category: headers.indexOf('종목 유형'),
            ratio: headers.indexOf('유형 비율'),
            currency: headers.indexOf('통화'),
          };

          const dataRows = jsonData.slice(1).filter(row => row && row.length > 0);
          
          const excelRows = dataRows.map((row, index) => {
            const rawRatio = String(row[colIndices.ratio] || '0').replace(/[^0-9.]/g, '');
            const ratioNum = parseFloat(rawRatio) || 0;
            const bankVal = String(row[colIndices.bank] || '').trim();
            const purposeVal = String(row[colIndices.purpose] || '연금').trim();
            const validPurpose = ['연금', 'IRP', 'DC', '기타'].includes(purposeVal) ? purposeVal : '연금';

            return {
              id: 'excel_stock_' + Date.now() + '_' + index,
              bank: bankVal,
              purpose: validPurpose,
              code: String(row[colIndices.code] || '').trim(),
              name: String(row[colIndices.name] || '').trim(),
              category: String(row[colIndices.category] || '').trim(),
              ratio: ratioNum,
              currency: String(row[colIndices.currency] || 'KRW').trim(),
            };
          });

          setStocks(prev => [...excelRows, ...prev]);
          setSuccessMessage('종목 엑셀이 로드되었습니다. [저장] 버튼을 눌러 DB에 반영하세요.');
        } else {
          // 거래현황 엑셀 업로드 [거래일자, 은행, 목적, 종목명, 단가, 매수수량, 매도수량]
          const requiredColumns = ['거래일자', '은행', '목적', '종목명', '단가', '매수수량', '매도수량'];
          const missingColumns = requiredColumns.filter(col => !headers.includes(col));
          
          if (missingColumns.length > 0) {
            showError(`필수 열 누락: [ ${missingColumns.join(', ')} ]\n(1행에 '거래일자', '은행', '목적', '종목명', '단가', '매수수량', '매도수량' 열이 있어야 합니다.)`);
            return;
          }

          const colIndices = {
            date: headers.indexOf('거래일자'),
            bank: headers.indexOf('은행'),
            purpose: headers.indexOf('목적'),
            name: headers.indexOf('종목명'),
            price: headers.indexOf('단가'),
            buyQty: headers.indexOf('매수수량'),
            sellQty: headers.indexOf('매도수량'),
          };

          const dataRows = jsonData.slice(1).filter(row => row && row.length > 0);
          
          const excelRows = dataRows.map((row, index) => {
            const rawDate = row[colIndices.date];
            const parsedDate = parseExcelDate(rawDate);
            const bankVal = String(row[colIndices.bank] || '').trim();
            const purposeVal = String(row[colIndices.purpose] || '연금').trim();
            const nameVal = String(row[colIndices.name] || '').trim();
            const rawPrice = String(row[colIndices.price] || '0').replace(/[^0-9]/g, '');
            const priceNum = parseInt(rawPrice, 10) || 0;
            const buyQtyNum = parseFloat(String(row[colIndices.buyQty] || '0').replace(/[^0-9.]/g, '')) || 0;
            const sellQtyNum = parseFloat(String(row[colIndices.sellQty] || '0').replace(/[^0-9.]/g, '')) || 0;

            // 종목 관리에서 일치하는 종목 찾아 코드 자동 바인딩
            const matchedStock = stocks.find(s => 
              (s.bank || '').trim() === bankVal &&
              (s.purpose || '').trim() === purposeVal &&
              (s.name || '').trim() === nameVal
            );

            return {
              id: 'excel_tx_' + Date.now() + '_' + index,
              date: parsedDate,
              bank: bankVal,
              purpose: ['연금', 'IRP', 'DC', '기타'].includes(purposeVal) ? purposeVal : '연금',
              name: nameVal,
              code: matchedStock ? matchedStock.code : '',
              price: priceNum,
              buyQty: buyQtyNum,
              sellQty: sellQtyNum,
            };
          });

          setTransactions(prev => [...excelRows, ...prev]);
          setSuccessMessage('거래현황 엑셀이 로드되었습니다. [저장] 버튼을 눌러 DB에 반영하세요.');
        }

        setTimeout(() => setSuccessMessage(''), 4000);
      } catch (err) {
        console.error(err);
        showError('엑셀 파일을 읽는 중 오류가 발생했습니다.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSaveToDatabase = async () => {
    if (isSaving) return;
    
    setIsSaving(true);
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error("서버 응답 시간이 초과되었습니다.")), 30000)
    );

    const saveExecution = async () => {
      const batch = writeBatch(db);

      for (const id of deletedIds) {
        const docRef = doc(db, "payments", id);
        batch.delete(docRef);
      }

      for (const payment of payments) {
        if (String(payment.id).startsWith('temp_') || String(payment.id).startsWith('excel_')) {
          const newDocRef = doc(collection(db, "payments"));
          batch.set(newDocRef, {
            date: payment.date || getTodayString(),
            bank: payment.bank || '',
            purpose: payment.purpose || '연금',
            amount: Number(payment.amount || 0),
            createdAt: serverTimestamp()
          });
        } else {
          const docRef = doc(db, "payments", payment.id);
          batch.update(docRef, {
            date: payment.date || getTodayString(),
            bank: payment.bank || '',
            purpose: payment.purpose || '연금',
            amount: Number(payment.amount || 0)
          });
        }
      }

      await batch.commit();
    };

    try {
      await Promise.race([saveExecution(), timeoutPromise]);
      await fetchPayments();
      setSuccessMessage('납입금액 데이터가 성공적으로 저장되었습니다!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      console.error("저장 중 에러 발생:", error);
      showError(`저장 실패: ${error.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleSaveStocksToDatabase = async () => {
    if (isSavingStocks) return;
    
    setIsSavingStocks(true);
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error("서버 응답 시간이 초과되었습니다.")), 30000)
    );

    const saveExecution = async () => {
      const batch = writeBatch(db);

      for (const id of deletedStockIds) {
        const docRef = doc(db, "stocks", id);
        batch.delete(docRef);
      }

      for (const stock of stocks) {
        if (String(stock.id).startsWith('temp_stock_') || String(stock.id).startsWith('excel_stock_')) {
          const newDocRef = doc(collection(db, "stocks"));
          batch.set(newDocRef, {
            bank: stock.bank || '',
            purpose: stock.purpose || '연금',
            code: stock.code || '',
            name: stock.name || '',
            category: stock.category || '',
            ratio: Number(stock.ratio || 0),
            currency: stock.currency || 'KRW',
            createdAt: serverTimestamp()
          });
        } else {
          const docRef = doc(db, "stocks", stock.id);
          batch.update(docRef, {
            bank: stock.bank || '',
            purpose: stock.purpose || '연금',
            code: stock.code || '',
            name: stock.name || '',
            category: stock.category || '',
            ratio: Number(stock.ratio || 0),
            currency: stock.currency || 'KRW'
          });
        }
      }

      await batch.commit();
    };

    try {
      await Promise.race([saveExecution(), timeoutPromise]);
      await fetchStocks();
      setSuccessMessage('종목 관리 데이터가 성공적으로 저장되었습니다!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      console.error("종목 저장 중 에러 발생:", error);
      showError(`저장 실패: ${error.message}`);
    } finally {
      setIsSavingStocks(false);
    }
  };

  const handleSaveTransactionsToDatabase = async () => {
    if (isSavingTransactions) return;
    
    setIsSavingTransactions(true);
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error("서버 응답 시간이 초과되었습니다.")), 30000)
    );

    const saveExecution = async () => {
      const batch = writeBatch(db);

      for (const id of deletedTransactionIds) {
        const docRef = doc(db, "transactions", id);
        batch.delete(docRef);
      }

      for (const tx of transactions) {
        if (String(tx.id).startsWith('temp_tx_') || String(tx.id).startsWith('excel_tx_')) {
          const newDocRef = doc(collection(db, "transactions"));
          batch.set(newDocRef, {
            date: tx.date || getTodayString(),
            bank: tx.bank || '',
            purpose: tx.purpose || '연금',
            name: tx.name || '',
            code: tx.code || '',
            price: Number(tx.price || 0),
            buyQty: Number(tx.buyQty || 0),
            sellQty: Number(tx.sellQty || 0),
            createdAt: serverTimestamp()
          });
        } else {
          const docRef = doc(db, "transactions", tx.id);
          batch.update(docRef, {
            date: tx.date || getTodayString(),
            bank: tx.bank || '',
            purpose: tx.purpose || '연금',
            name: tx.name || '',
            code: tx.code || '',
            price: Number(tx.price || 0),
            buyQty: Number(tx.buyQty || 0),
            sellQty: Number(tx.sellQty || 0)
          });
        }
      }

      await batch.commit();
    };

    try {
      await Promise.race([saveExecution(), timeoutPromise]);
      await fetchTransactions();
      setSuccessMessage('거래현황 데이터가 성공적으로 저장되었습니다!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      console.error("거래현황 저장 중 에러 발생:", error);
      showError(`저장 실패: ${error.message}`);
    } finally {
      setIsSavingTransactions(false);
    }
  };

  const handleRowChange = (id, field, value) => {
    setPayments(prev => prev.map(payment => {
      if (payment.id === id) {
        return { ...payment, [field]: value };
      }
      return payment;
    }));
  };

  const handleStockRowChange = (id, field, value) => {
    setStocks(prev => prev.map(stock => {
      if (stock.id === id) {
        return { ...stock, [field]: value };
      }
      return stock;
    }));
  };

  const handleTransactionRowChange = (id, field, value) => {
    setTransactions(prev => prev.map(tx => {
      if (tx.id === id) {
        const updated = { ...tx, [field]: value };
        // 만약 은행, 목적, 혹은 종목명이 변경되었다면 종목코드 자동 바인딩
        if (field === 'bank' || field === 'purpose' || field === 'name') {
          const targetBank = field === 'bank' ? value : updated.bank;
          const targetPurpose = field === 'purpose' ? value : updated.purpose;
          const targetName = field === 'name' ? value : updated.name;

          const matchedStock = stocks.find(s => 
            (s.bank || '').trim() === (targetBank || '').trim() &&
            (s.purpose || '').trim() === (targetPurpose || '').trim() &&
            (s.name || '').trim() === (targetName || '').trim()
          );

          if (matchedStock) {
            updated.code = matchedStock.code || '';
          } else if (field === 'name') {
            updated.code = '';
          }
        }
        return updated;
      }
      return tx;
    }));
  };

  const toggleSelection = (id) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
    );
  };

  const toggleStockSelection = (id) => {
    setSelectedStockIds(prev => 
      prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
    );
  };

  const toggleTransactionSelection = (id) => {
    setSelectedTransactionIds(prev => 
      prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
    );
  };

  const filteredPayments = useMemo(() => {
    if (!appliedSearchYear) return payments;
    return payments.filter(payment => payment.date && payment.date.startsWith(appliedSearchYear));
  }, [payments, appliedSearchYear]);

  const filteredTransactions = useMemo(() => {
    return transactions.filter(tx => {
      if (!tx.date) return true;
      if (appliedTxStartDate && tx.date < appliedTxStartDate) return false;
      if (appliedTxEndDate && tx.date > appliedTxEndDate) return false;
      return true;
    });
  }, [transactions, appliedTxStartDate, appliedTxEndDate]);

  const handleSelectAll = (e) => {
    if (e.target.checked) {
      setSelectedIds(filteredPayments.map(p => p.id));
    } else {
      setSelectedIds([]);
    }
  };

  const handleSelectAllStocks = (e) => {
    if (e.target.checked) {
      setSelectedStockIds(stocks.map(s => s.id));
    } else {
      setSelectedStockIds([]);
    }
  };

  const handleSelectAllTransactions = (e) => {
    if (e.target.checked) {
      setSelectedTransactionIds(filteredTransactions.map(t => t.id));
    } else {
      setSelectedTransactionIds([]);
    }
  };

  const totalAmount = useMemo(() => {
    return payments.reduce((sum, current) => sum + Number(current.amount || 0), 0);
  }, [payments]);

  const filteredTotalAmount = useMemo(() => {
    return filteredPayments.reduce((sum, current) => sum + Number(current.amount || 0), 0);
  }, [filteredPayments]);

  // 종목관리 기반 고유 은행 목록 (DISTINCT)
  const availableBanks = useMemo(() => {
    const set = new Set();
    stocks.forEach(s => {
      if (s.bank && s.bank.trim()) set.add(s.bank.trim());
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ko'));
  }, [stocks]);

  // 특정 행의 선택된 은행, 목적에 따른 고유 종목명 목록 (DISTINCT)
  const getAvailableStockNames = (bank, purpose) => {
    const set = new Set();
    stocks.forEach(s => {
      if (
        (s.bank || '').trim() === (bank || '').trim() &&
        (s.purpose || '').trim() === (purpose || '').trim() &&
        s.name && s.name.trim()
      ) {
        set.add(s.name.trim());
      }
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ko'));
  };

  if (!isTailwindLoaded) {
    return (
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh', fontFamily: 'sans-serif', color: '#64748b' }}>
        화면을 준비 중입니다...
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans antialiased">
      
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        className="hidden" 
        ref={fileInputRef} 
        onChange={handleFileChange} 
      />

      <input 
        type="file" 
        accept=".xlsx, .xls" 
        className="hidden" 
        ref={stockFileInputRef} 
        onChange={handleFileChange} 
      />

      <input 
        type="file" 
        accept=".xlsx, .xls" 
        className="hidden" 
        ref={transactionFileInputRef} 
        onChange={handleFileChange} 
      />

      {errorModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-6">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <AlertCircle size={28} />
              <h3 className="text-xl font-bold">오류 안내</h3>
            </div>
            <div className="text-slate-600 mb-6 whitespace-pre-line leading-relaxed">
              {errorModal.message}
            </div>
            <div className="flex justify-end">
              <button 
                onClick={() => setErrorModal({ isOpen: false, message: '' })}
                className="px-6 py-2 bg-slate-800 text-white rounded-lg font-medium hover:bg-slate-700 transition-colors"
              >
                확인
              </button>
            </div>
          </div>
        </div>
      )}

      {successMessage && (
        <div className="fixed top-5 right-5 z-50 bg-emerald-600 text-white px-6 py-3 rounded-xl shadow-lg flex items-center gap-2 font-medium transition-all">
          <span>{successMessage}</span>
        </div>
      )}

      <header className="bg-white border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row items-start sm:items-center justify-between shadow-sm gap-4">
        <div className="flex items-center gap-3 text-indigo-600">
          <Landmark size={28} />
          <h1 className="text-xl font-bold tracking-tight text-slate-900">주식 포트폴리오 관리 시스템</h1>
        </div>
        <div className="text-sm text-emerald-600 font-medium bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200 flex items-center gap-1.5 shadow-xs">
          <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
          Firebase 클라우드 연동 완료
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 sm:px-6 lg:px-8">
        
        {/* Tab Navigation */}
        <div className="flex space-x-1 border-b border-slate-200 mb-6 overflow-x-auto">
          <button
            onClick={() => setActiveTab('payment')}
            className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 transition-colors duration-200 whitespace-nowrap
              ${activeTab === 'payment' 
                ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' 
                : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
              }`}
          >
            <Wallet size={18} />
            납입금액 관리
          </button>
          <button
            onClick={() => setActiveTab('stock')}
            className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 transition-colors duration-200 whitespace-nowrap
              ${activeTab === 'stock' 
                ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' 
                : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
              }`}
          >
            <Layers size={18} />
            종목 관리
          </button>
          <button
            onClick={() => setActiveTab('transaction')}
            className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 transition-colors duration-200 whitespace-nowrap
              ${activeTab === 'transaction' 
                ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' 
                : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
              }`}
          >
            <ArrowLeftRight size={18} />
            거래현황
          </button>
          <button
            onClick={() => setActiveTab('portfolio')}
            className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 transition-colors duration-200 whitespace-nowrap
              ${activeTab === 'portfolio' 
                ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' 
                : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
              }`}
          >
            <PiggyBank size={18} />
            포트폴리오 현황
          </button>
        </div>

        {/* 1. 납입금액 관리 탭 */}
        {activeTab === 'payment' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-start xl:items-center">
              
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Calendar size={16} className="text-slate-400" />
                  </div>
                  <input
                    type="text"
                    placeholder="납입연도 (예: 2026)"
                    className="pl-10 pr-4 py-2 w-44 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-slate-900"
                    value={searchYearInput}
                    onChange={(e) => setSearchYearInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleSearch()}
                  />
                </div>
                
                <button 
                  onClick={handleSearch}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 transition-colors shadow-sm"
                >
                  <Search size={16} />
                  조회
                </button>

                <button 
                  onClick={handleAddRow}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 transition-colors shadow-sm"
                >
                  <Plus size={16} />
                  추가
                </button>

                <button 
                  onClick={handleDeleteRows}
                  disabled={selectedIds.length === 0}
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium transition-colors shadow-sm
                    ${selectedIds.length > 0 
                      ? 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200 text-rose-600' 
                      : 'bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed'
                    }`}
                >
                  <Trash2 size={16} />
                  삭제 {selectedIds.length > 0 && `(${selectedIds.length})`}
                </button>

                <button 
                  onClick={() => triggerExcelUpload('payment')}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 transition-colors shadow-sm"
                >
                  <FileSpreadsheet size={16} className="text-green-600" />
                  엑셀 업로드
                </button>

                <button 
                  onClick={handleSaveToDatabase}
                  disabled={isSaving}
                  className={`flex items-center gap-1.5 px-4 py-2 text-white rounded-lg text-sm font-medium transition-colors shadow-sm ml-1
                    ${isSaving ? 'bg-blue-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}`}
                >
                  <Save size={16} />
                  {isSaving ? '저장 중...' : '저장'}
                </button>
              </div>

              <div className="flex items-center gap-6 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm min-w-max w-full xl:w-auto">
                <div className="flex flex-col">
                  <span className="text-xs text-slate-500 font-medium">납입 총액 (전체)</span>
                  <span className="text-lg font-bold text-slate-800 tracking-tight">
                    {formatCurrency(totalAmount)}<span className="text-sm font-normal text-slate-500 ml-1">원</span>
                  </span>
                </div>
                <div className="w-px h-10 bg-slate-200"></div>
                <div className="flex flex-col">
                  <span className="text-xs text-indigo-500 font-medium">조회 총액 (현재 화면)</span>
                  <span className="text-lg font-bold text-indigo-700 tracking-tight">
                    {formatCurrency(filteredTotalAmount)}<span className="text-sm font-normal text-indigo-500 ml-1">원</span>
                  </span>
                </div>
              </div>
            </div>

            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[800px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b border-slate-200 bg-slate-50">
                      <input 
                        type="checkbox" 
                        className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600 cursor-pointer"
                        checked={filteredPayments.length > 0 && selectedIds.length === filteredPayments.length}
                        onChange={handleSelectAll}
                      />
                    </th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-48">납입일자</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-48">은행</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-48">목적</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm text-right">금액 (원)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {loading && payments.length === 0 ? (
                    <tr>
                      <td colSpan="5" className="py-16 text-center text-slate-400">데이터를 불러오는 중입니다...</td>
                    </tr>
                  ) : filteredPayments.length > 0 ? (
                    filteredPayments.map((row) => (
                      <tr 
                        key={row.id} 
                        className={`hover:bg-slate-50 transition-colors ${selectedIds.includes(row.id) ? 'bg-indigo-50/30' : ''}`}
                      >
                        <td className="py-3 px-4 w-12">
                          <input 
                            type="checkbox" 
                            className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600 cursor-pointer"
                            checked={selectedIds.includes(row.id)}
                            onChange={() => toggleSelection(row.id)}
                          />
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="date" 
                            value={row.date || ''}
                            onChange={(e) => handleRowChange(row.id, 'date', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="text" 
                            value={row.bank || ''}
                            onChange={(e) => handleRowChange(row.id, 'bank', e.target.value)}
                            placeholder="은행 입력"
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <select 
                            value={row.purpose || '연금'}
                            onChange={(e) => handleRowChange(row.id, 'purpose', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          >
                            <option value="연금">연금</option>
                            <option value="IRP">IRP</option>
                            <option value="DC">DC</option>
                            <option value="기타">기타</option>
                          </select>
                        </td>
                        <td className="py-2 px-4">
                          <div className="relative flex items-center justify-end">
                            <input 
                              type="text" 
                              value={row.amount === 0 ? '' : formatCurrency(row.amount)}
                              onChange={(e) => {
                                const rawValue = e.target.value.replace(/[^0-9]/g, '');
                                const numValue = parseInt(rawValue, 10) || 0;
                                handleRowChange(row.id, 'amount', numValue);
                              }}
                              placeholder="0"
                              className="w-full md:w-3/4 px-3 py-1.5 border border-slate-200 rounded-md text-sm text-right font-medium focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            />
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="5" className="py-16 text-center text-slate-500">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <CheckSquare size={32} className="text-slate-300" />
                          <p>조회된 납입 내역이 없습니다.</p>
                          <p className="text-sm text-slate-400">추가 버튼을 누르거나 엑셀 파일을 업로드해보세요.</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 2. 종목 관리 탭 ([은행], [목적], [종목코드], [종목명], [종목 유형], [유형 비율], [통화]) */}
        {activeTab === 'stock' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-start xl:items-center">
              
              <div className="flex flex-wrap items-center gap-2.5">
                <button 
                  onClick={fetchStocks}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 transition-colors shadow-sm"
                >
                  <Search size={16} />
                  조회
                </button>

                <button 
                  onClick={handleAddStockRow}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 transition-colors shadow-sm"
                >
                  <Plus size={16} />
                  추가
                </button>

                <button 
                  onClick={handleDeleteStockRows}
                  disabled={selectedStockIds.length === 0}
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium transition-colors shadow-sm
                    ${selectedStockIds.length > 0 
                      ? 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200 text-rose-600' 
                      : 'bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed'
                    }`}
                >
                  <Trash2 size={16} />
                  삭제 {selectedStockIds.length > 0 && `(${selectedStockIds.length})`}
                </button>

                <button 
                  onClick={() => triggerExcelUpload('stock')}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 transition-colors shadow-sm"
                >
                  <FileSpreadsheet size={16} className="text-green-600" />
                  엑셀 업로드
                </button>

                <button 
                  onClick={handleSaveStocksToDatabase}
                  disabled={isSavingStocks}
                  className={`flex items-center gap-1.5 px-4 py-2 text-white rounded-lg text-sm font-medium transition-colors shadow-sm ml-1
                    ${isSavingStocks ? 'bg-blue-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}`}
                >
                  <Save size={16} />
                  {isSavingStocks ? '저장 중...' : '저장'}
                </button>
              </div>

              <div className="flex items-center gap-4 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm">
                <span className="text-xs text-slate-500 font-medium">등록된 종목 수:</span>
                <span className="text-base font-bold text-slate-800">{stocks.length} 개</span>
              </div>
            </div>

            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b border-slate-200 bg-slate-50">
                      <input 
                        type="checkbox" 
                        className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600 cursor-pointer"
                        checked={stocks.length > 0 && selectedStockIds.length === stocks.length}
                        onChange={handleSelectAllStocks}
                      />
                    </th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-44">은행</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-44">목적</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-36">종목코드</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm">종목명</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-40">종목 유형</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-28 text-right">유형 비율 (%)</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-28 text-center">통화</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {loadingStocks && stocks.length === 0 ? (
                    <tr>
                      <td colSpan="8" className="py-16 text-center text-slate-400">종목 데이터를 불러오는 중입니다...</td>
                    </tr>
                  ) : stocks.length > 0 ? (
                    stocks.map((stock) => (
                      <tr 
                        key={stock.id} 
                        className={`hover:bg-slate-50 transition-colors ${selectedStockIds.includes(stock.id) ? 'bg-indigo-50/30' : ''}`}
                      >
                        <td className="py-3 px-4 w-12">
                          <input 
                            type="checkbox" 
                            className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600 cursor-pointer"
                            checked={selectedStockIds.includes(stock.id)}
                            onChange={() => toggleStockSelection(stock.id)}
                          />
                        </td>
                        <td className="py-2 px-4">
                          <select 
                            value={stock.bank || ''}
                            onChange={(e) => handleStockRowChange(stock.id, 'bank', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          >
                            <option value="">은행 선택</option>
                            <option value="미래에셋">미래에셋</option>
                            <option value="KB증권">KB증권</option>
                            <option value="삼성증권">삼성증권</option>
                          </select>
                        </td>
                        <td className="py-2 px-4">
                          <select 
                            value={stock.purpose || '연금'}
                            onChange={(e) => handleStockRowChange(stock.id, 'purpose', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          >
                            <option value="연금">연금</option>
                            <option value="IRP">IRP</option>
                            <option value="DC">DC</option>
                            <option value="기타">기타</option>
                          </select>
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="text" 
                            value={stock.code || ''}
                            onChange={(e) => handleStockRowChange(stock.id, 'code', e.target.value)}
                            placeholder="종목코드"
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="text" 
                            value={stock.name || ''}
                            onChange={(e) => handleStockRowChange(stock.id, 'name', e.target.value)}
                            placeholder="종목명"
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="text" 
                            value={stock.category || ''}
                            onChange={(e) => handleStockRowChange(stock.id, 'category', e.target.value)}
                            placeholder="종목 유형"
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="number" 
                            value={stock.ratio || 0}
                            onChange={(e) => handleStockRowChange(stock.id, 'ratio', e.target.value)}
                            placeholder="0"
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm text-right focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <input 
                            type="text" 
                            value={stock.currency || 'KRW'}
                            onChange={(e) => handleStockRowChange(stock.id, 'currency', e.target.value)}
                            placeholder="KRW"
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm text-center focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          />
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="8" className="py-16 text-center text-slate-500">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <CheckSquare size={32} className="text-slate-300" />
                          <p>등록된 종목 내역이 없습니다.</p>
                          <p className="text-sm text-slate-400">추가 버튼을 누르거나 엑셀 파일을 업로드해보세요.</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 3. 거래현황 탭 */}
        {activeTab === 'transaction' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[75vh]">
            
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-start xl:items-center">
              
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-1.5 shadow-xs">
                  <Calendar size={16} className="text-slate-400" />
                  <input
                    type="date"
                    value={txStartDate}
                    onChange={(e) => setTxStartDate(e.target.value)}
                    className="text-sm outline-none bg-transparent text-slate-900"
                  />
                  <span className="text-slate-400">~</span>
                  <input
                    type="date"
                    value={txEndDate}
                    onChange={(e) => setTxEndDate(e.target.value)}
                    className="text-sm outline-none bg-transparent text-slate-900"
                  />
                </div>
                
                <button 
                  onClick={handleTxSearch}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 transition-colors shadow-sm"
                >
                  <Search size={16} />
                  조회
                </button>

                <button 
                  onClick={handleAddTransactionRow}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 transition-colors shadow-sm"
                >
                  <Plus size={16} />
                  추가
                </button>

                <button 
                  onClick={handleDeleteTransactionRows}
                  disabled={selectedTransactionIds.length === 0}
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium transition-colors shadow-sm
                    ${selectedTransactionIds.length > 0 
                      ? 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200 text-rose-600' 
                      : 'bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed'
                    }`}
                >
                  <Trash2 size={16} />
                  삭제 {selectedTransactionIds.length > 0 && `(${selectedTransactionIds.length})`}
                </button>

                <button 
                  onClick={() => triggerExcelUpload('transaction')}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 transition-colors shadow-sm"
                >
                  <FileSpreadsheet size={16} className="text-green-600" />
                  엑셀 업로드
                </button>

                <button 
                  onClick={handleSaveTransactionsToDatabase}
                  disabled={isSavingTransactions}
                  className={`flex items-center gap-1.5 px-4 py-2 text-white rounded-lg text-sm font-medium transition-colors shadow-sm ml-1
                    ${isSavingTransactions ? 'bg-blue-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}`}
                >
                  <Save size={16} />
                  {isSavingTransactions ? '저장 중...' : '저장'}
                </button>
              </div>

              <div className="flex items-center gap-4 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm">
                <span className="text-xs text-slate-500 font-medium">조회 건수:</span>
                <span className="text-base font-bold text-slate-800">{filteredTransactions.length} 건</span>
              </div>
            </div>

            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="py-3 px-4 w-12 border-b border-slate-200 bg-slate-50">
                      <input 
                        type="checkbox" 
                        className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600 cursor-pointer"
                        checked={filteredTransactions.length > 0 && selectedTransactionIds.length === filteredTransactions.length}
                        onChange={handleSelectAllTransactions}
                      />
                    </th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-44">거래일자</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-40">은행</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-36">목적</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm">종목명</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-36">종목코드</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-32 text-right">단가</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-28 text-right">매수수량</th>
                    <th className="py-3 px-4 border-b border-slate-200 font-semibold text-slate-600 text-sm w-28 text-right">매도수량</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {loadingTransactions && transactions.length === 0 ? (
                    <tr>
                      <td colSpan="9" className="py-16 text-center text-slate-400">거래현황 데이터를 불러오는 중입니다...</td>
                    </tr>
                  ) : filteredTransactions.length > 0 ? (
                    filteredTransactions.map((tx) => {
                      const rowAvailableNames = getAvailableStockNames(tx.bank, tx.purpose);
                      return (
                        <tr 
                          key={tx.id} 
                          className={`hover:bg-slate-50 transition-colors ${selectedTransactionIds.includes(tx.id) ? 'bg-indigo-50/30' : ''}`}
                        >
                          <td className="py-3 px-4 w-12">
                            <input 
                              type="checkbox" 
                              className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600 cursor-pointer"
                              checked={selectedTransactionIds.includes(tx.id)}
                              onChange={() => toggleTransactionSelection(tx.id)}
                            />
                          </td>
                          <td className="py-2 px-4">
                            <input 
                              type="date" 
                              value={tx.date || getTodayString()}
                              onChange={(e) => handleTransactionRowChange(tx.id, 'date', e.target.value)}
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            />
                          </td>
                          <td className="py-2 px-4">
                            <select 
                              value={tx.bank || ''}
                              onChange={(e) => handleTransactionRowChange(tx.id, 'bank', e.target.value)}
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            >
                              <option value="">은행 선택</option>
                              {availableBanks.map(b => (
                                <option key={b} value={b}>{b}</option>
                              ))}
                            </select>
                          </td>
                          <td className="py-2 px-4">
                            <select 
                              value={tx.purpose || '연금'}
                              onChange={(e) => handleTransactionRowChange(tx.id, 'purpose', e.target.value)}
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            >
                              <option value="연금">연금</option>
                              <option value="IRP">IRP</option>
                              <option value="DC">DC</option>
                              <option value="기타">기타</option>
                            </select>
                          </td>
                          <td className="py-2 px-4">
                            <select 
                              value={tx.name || ''}
                              onChange={(e) => handleTransactionRowChange(tx.id, 'name', e.target.value)}
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            >
                              <option value="">종목명 선택</option>
                              {rowAvailableNames.map(name => (
                                <option key={name} value={name}>{name}</option>
                              ))}
                            </select>
                          </td>
                          <td className="py-2 px-4">
                            <input 
                              type="text" 
                              value={tx.code || ''}
                              readOnly
                              placeholder="자동입력"
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm bg-slate-100 text-slate-600 outline-none cursor-not-allowed"
                            />
                          </td>
                          <td className="py-2 px-4">
                            <input 
                              type="text" 
                              value={tx.price === 0 ? '' : formatCurrency(tx.price)}
                              onChange={(e) => {
                                const raw = e.target.value.replace(/[^0-9]/g, '');
                                handleTransactionRowChange(tx.id, 'price', parseInt(raw, 10) || 0);
                              }}
                              placeholder="0"
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm text-right focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900 font-medium"
                            />
                          </td>
                          <td className="py-2 px-4">
                            <input 
                              type="number" 
                              value={tx.buyQty === 0 ? '' : tx.buyQty}
                              onChange={(e) => handleTransactionRowChange(tx.id, 'buyQty', e.target.value)}
                              placeholder="0"
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm text-right focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            />
                          </td>
                          <td className="py-2 px-4">
                            <input 
                              type="number" 
                              value={tx.sellQty === 0 ? '' : tx.sellQty}
                              onChange={(e) => handleTransactionRowChange(tx.id, 'sellQty', e.target.value)}
                              placeholder="0"
                              className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm text-right focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                            />
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan="9" className="py-16 text-center text-slate-500">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <CheckSquare size={32} className="text-slate-300" />
                          <p>조회된 거래현황 내역이 없습니다.</p>
                          <p className="text-sm text-slate-400">추가 버튼을 누르거나 엑셀 파일을 업로드해보세요.</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 4. 포트폴리오 현황 탭 */}
        {activeTab === 'portfolio' && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 text-center flex flex-col items-center justify-center min-h-[50vh]">
             <PiggyBank size={48} className="text-slate-300 mb-4" />
             <h2 className="text-xl font-bold text-slate-700">포트폴리오 현황 화면</h2>
             <p className="text-slate-500 mt-2">이 탭의 내용은 다음 단계에서 구성할 예정입니다.</p>
          </div>
        )}
      </main>
    </div>
  );
}
