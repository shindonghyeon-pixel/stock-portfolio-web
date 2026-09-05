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
  Save
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { db, auth } from './firebase';
import { signInAnonymously } from 'firebase/auth';
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
  const [activeTab, setActiveTab] = useState('payment');
  const [payments, setPayments] = useState([]);
  
  const [loading, setLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  
  const [searchYearInput, setSearchYearInput] = useState('');
  const [appliedSearchYear, setAppliedSearchYear] = useState('');
  const [selectedIds, setSelectedIds] = useState([]);
  const [deletedIds, setDeletedIds] = useState([]);
  
  const [errorModal, setErrorModal] = useState({ isOpen: false, message: '' });
  const [successMessage, setSuccessMessage] = useState('');
  const fileInputRef = useRef(null);

  // 1. 앱 실행 시 인증이 끝날 때까지 확실히 기다린 후 데이터 로드
  useEffect(() => {
    const initApp = async () => {
      try {
        if (auth) {
          await signInAnonymously(auth);
          console.log("🔥 Firebase 익명 인증 완료");
        }
      } catch (authErr) {
        console.warn("⚠️ 인증 경고:", authErr);
      }
      await fetchPayments();
    };

    initApp();
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
      console.error("데이터 불러오기 실패:", error);
      showError(`데이터를 불러오지 못했습니다.\n(${error.message})`);
    } finally {
      setLoading(false);
    }
  };

  const handleSearch = () => {
    setAppliedSearchYear(searchYearInput);
    setSelectedIds([]); 
  };

  const handleAddRow = () => {
    const newRow = {
      id: 'temp_' + Date.now(),
      date: getTodayString(),
      bank: '미래에셋',
      purpose: '연금',
      amount: 0,
    };
    setPayments(prev => [newRow, ...prev]);
  };

  const handleDeleteRows = () => {
    if (selectedIds.length === 0) return;
    const targetRealIds = selectedIds.filter(id => !String(id).startsWith('temp_') && !String(id).startsWith('excel_'));
    setDeletedIds(prev => [...prev, ...targetRealIds]);
    setPayments(prev => prev.filter(p => !selectedIds.includes(p.id)));
    setSelectedIds([]);
  };

  const triggerExcelUpload = () => {
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
      fileInputRef.current.click();
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
            bank: String(row[colIndices.bank] || '기타').trim(),
            purpose: String(row[colIndices.purpose] || '기타').trim(),
            amount: amountNum,
          };
        });

        setPayments(prev => [...excelRows, ...prev]);
        setSuccessMessage('엑셀이 로드되었습니다. [저장] 버튼을 눌러 DB에 반영하세요.');
        setTimeout(() => setSuccessMessage(''), 4000);
      } catch (err) {
        console.error(err);
        showError('엑셀 파일을 읽는 중 오류가 발생했습니다.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 2. 저장 버튼: 인증 보장 상태에서 Batch 일괄 처리 실행
  const handleSaveToDatabase = async () => {
    if (isSaving) return;
    
    setIsSaving(true);
    console.log("=== 저장 작업 시작 ===");

    try {
      // 저장 전 인증 상태 재확인
      if (auth && !auth.currentUser) {
        await signInAnonymously(auth);
      }

      const batch = writeBatch(db);

      // 삭제 대기열 처리
      for (const id of deletedIds) {
        const docRef = doc(db, "payments", id);
        batch.delete(docRef);
      }

      // 추가 및 수정 처리
      for (const payment of payments) {
        if (String(payment.id).startsWith('temp_') || String(payment.id).startsWith('excel_')) {
          const newDocRef = doc(collection(db, "payments"));
          batch.set(newDocRef, {
            date: payment.date || getTodayString(),
            bank: payment.bank || '미래에셋',
            purpose: payment.purpose || '연금',
            amount: Number(payment.amount || 0),
            createdAt: serverTimestamp()
          });
        } else {
          const docRef = doc(db, "payments", payment.id);
          batch.update(docRef, {
            date: payment.date || getTodayString(),
            bank: payment.bank || '미래에셋',
            purpose: payment.purpose || '연금',
            amount: Number(payment.amount || 0)
          });
        }
      }

      await batch.commit();
      await fetchPayments();
      setSuccessMessage('성공적으로 데이터베이스에 저장되었습니다!');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      console.error("저장 중 에러 발생:", error);
      showError(`저장 실패: ${error.message}`);
    } finally {
      setIsSaving(false);
      console.log("=== 저장 작업 종료 ===");
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

  const toggleSelection = (id) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(item !== id) : [...prev, id]
    );
  };

  const filteredPayments = useMemo(() => {
    if (!appliedSearchYear) return payments;
    return payments.filter(payment => payment.date && payment.date.startsWith(appliedSearchYear));
  }, [payments, appliedSearchYear]);

  const handleSelectAll = (e) => {
    if (e.target.checked) {
      setSelectedIds(filteredPayments.map(p => p.id));
    } else {
      setSelectedIds([]);
    }
  };

  const totalAmount = useMemo(() => {
    return payments.reduce((sum, current) => sum + Number(current.amount || 0), 0);
  }, [payments]);

  const filteredTotalAmount = useMemo(() => {
    return filteredPayments.reduce((sum, current) => sum + Number(current.amount || 0), 0);
  }, [filteredPayments]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans antialiased">
      
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        className="hidden" 
        id="excel-upload-input"
        ref={fileInputRef} 
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
        
        <div className="flex space-x-1 border-b border-slate-200 mb-6">
          <button
            onClick={() => setActiveTab('payment')}
            className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 transition-colors duration-200
              ${activeTab === 'payment' 
                ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' 
                : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
              }`}
          >
            <Wallet size={18} />
            납입금액 관리
          </button>
          <button
            onClick={() => setActiveTab('portfolio')}
            className={`flex items-center gap-2 px-6 py-3 text-sm font-medium border-b-2 transition-colors duration-200
              ${activeTab === 'portfolio' 
                ? 'border-indigo-600 text-indigo-600 bg-indigo-50/50' 
                : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
              }`}
          >
            <PiggyBank size={18} />
            포트폴리오 현황
          </button>
        </div>

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

                <label 
                  htmlFor="excel-upload-input"
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 transition-colors shadow-sm cursor-pointer"
                >
                  <FileSpreadsheet size={16} className="text-green-600" />
                  엑셀 업로드
                </label>

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
                          <select 
                            value={row.bank || '미래에셋'}
                            onChange={(e) => handleRowChange(row.id, 'bank', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none bg-white text-slate-900"
                          >
                            <option value="미래에셋">미래에셋</option>
                            <option value="KB증권">KB증권</option>
                            <option value="삼성증권">삼성증권</option>
                            <option value="기타">기타</option>
                          </select>
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
