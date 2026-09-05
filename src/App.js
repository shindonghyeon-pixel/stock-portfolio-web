import React, { useState, useMemo, useEffect } from 'react';
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
import { db } from './firebase';
import { 
  collection, 
  getDocs, 
  addDoc, 
  updateDoc,
  deleteDoc, 
  doc, 
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

export default function App() {
  const [activeTab, setActiveTab] = useState('payment');
  const [payments, setPayments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [searchYearInput, setSearchYearInput] = useState('');
  const [appliedSearchYear, setAppliedSearchYear] = useState('');
  const [selectedIds, setSelectedIds] = useState([]);
  
  // 삭제할 ID들을 모아두는 상태
  const [deletedIds, setDeletedIds] = useState([]);
  const [errorModal, setErrorModal] = useState({ isOpen: false, message: '' });
  const [successMessage, setSuccessMessage] = useState('');

  // 1. Firebase 데이터 불러오기
  const fetchPayments = async () => {
    try {
      setLoading(true);
      const querySnapshot = await getDocs(collection(db, "payments"));
      const dataList = querySnapshot.docs.map(docSnap => ({
        id: docSnap.id,
        isNew: false, // 기존 DB 데이터
        ...docSnap.data()
      }));
      dataList.sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0));
      setPayments(dataList);
      setDeletedIds([]);
    } catch (error) {
      console.error("데이터 불러오기 실패:", error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchPayments();
  }, []);

  const handleSearch = () => {
    setAppliedSearchYear(searchYearInput);
    setSelectedIds([]); 
  };

  // 2. 추가 버튼: 화면에 임시 행 추가 (저장 버튼을 눌러야 DB 반영)
  const handleAddRow = () => {
    const newRow = {
      id: 'temp_' + Date.now(), // 임시 ID
      date: getTodayString(),
      bank: '미래에셋',
      purpose: '연금',
      amount: 0,
      isNew: true // 새로 추가된 항목 표시
    };
    setPayments(prev => [newRow, ...prev]);
  };

  // 3. 삭제 버튼: 화면상에서 체크된 항목을 목록에서 제외하고 삭제 대기열에 등록
  const handleDeleteRows = () => {
    if (selectedIds.length === 0) return;
    
    // DB에 이미 존재하는 데이터(임시 ID가 아닌 것)는 삭제 대기열에 추가
    const targetRealIds = selectedIds.filter(id => !String(id).startsWith('temp_'));
    setDeletedIds(prev => [...prev, ...targetRealIds]);

    // 화면에서 즉시 제거
    setPayments(prev => prev.filter(p => !selectedIds.includes(p.id)));
    setSelectedIds([]);
  };

  const showError = (message) => {
    setErrorModal({ isOpen: true, message });
  };

  // 4. 엑셀 업로드: 데이터를 파싱하여 화면에 일괄 추가 (저장 버튼을 눌러야 DB 반영)
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

        if (workbook.SheetNames.length !== 1) {
          showError(`엑셀 파일에 시트가 ${workbook.SheetNames.length}개 있습니다. 반드시 1개의 시트만 존재해야 합니다.`);
          return;
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        if (jsonData.length === 0) {
          showError('엑셀 시트가 비어있습니다.');
          return;
        }

        const headers = jsonData[0];
        const requiredColumns = ['납입일자', '은행', '목적', '금액'];
        const missingColumns = requiredColumns.filter(col => !headers.includes(col));
        if (missingColumns.length > 0) {
          showError(`다음 필수 열이 누락되었습니다: [ ${missingColumns.join(', ')} ]\n(1행에 '납입일자', '은행', '목적', '금액' 열이 모두 있어야 합니다.)`);
          return;
        }

        const dataRows = jsonData.slice(1).filter(row => row.length > 0);
        if (dataRows.length === 0) {
          showError('열 제목 아래에 실제 데이터가 존재하지 않습니다.');
          return;
        }

        const colIndices = {
          date: headers.indexOf('납입일자'),
          bank: headers.indexOf('은행'),
          purpose: headers.indexOf('목적'),
          amount: headers.indexOf('금액'),
        };

        const excelRows = dataRows.map((row, index) => {
          let dateStr = String(row[colIndices.date] || getTodayString()).trim();
          dateStr = dateStr.replace(/[\.\/]/g, '-');
          
          const amountStr = String(row[colIndices.amount] || '0').replace(/[^0-9]/g, '');
          const amountNum = parseInt(amountStr, 10) || 0;

          return {
            id: 'excel_' + Date.now() + '_' + index,
            date: dateStr,
            bank: row[colIndices.bank] || '기타',
            purpose: row[colIndices.purpose] || '기타',
            amount: amountNum,
            isNew: true // 새로 추가된 항목으로 처리
          };
        });

        setPayments(prev => [...excelRows, ...prev]);
        setSuccessMessage('엑셀 데이터가 성공적으로 로드되었습니다. [저장] 버튼을 누르면 DB에 반영됩니다.');
        setTimeout(() => setSuccessMessage(''), 4000);
      } catch (err) {
        console.error(err);
        showError('파일을 읽는 중 오류가 발생했습니다. 올바른 엑셀 파일인지 확인해 주세요.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 5. [저장] 버튼: 추가된 항목은 DB에 추가, 삭제된 항목은 DB에서 삭제, 수정된 내용은 반영
  const handleSaveToDatabase = async () => {
    try {
      setLoading(true);

      // A. 삭제 대기열에 있는 항목 DB에서 삭제
      for (const id of deletedIds) {
        await deleteDoc(doc(db, "payments", id));
      }

      // B. 현재 화면의 데이터 처리 (신규 추가분 DB 반영 및 기존 데이터 업데이트)
      for (const payment of payments) {
        if (String(payment.id).startsWith('temp_') || String(payment.id).startsWith('excel_')) {
          // 새로 추가된 항목인 경우 DB에 추가
          await addDoc(collection(db, "payments"), {
            date: payment.date,
            bank: payment.bank,
            purpose: payment.purpose,
            amount: Number(payment.amount || 0),
            createdAt: serverTimestamp()
          });
        } else {
          // 기존 항목인 경우 내용 변경사항 업데이트
          const docRef = doc(db, "payments", payment.id);
          await updateDoc(docRef, {
            date: payment.date,
            bank: payment.bank,
            purpose: payment.purpose,
            amount: Number(payment.amount || 0)
          });
        }
      }

      await fetchPayments();
      setSuccessMessage('변경사항이 데이터베이스에 성공적으로 저장되었습니다!');
      setTimeout(() => setSuccessMessage(''), 4000);
    } catch (error) {
      console.error("저장 실패:", error);
      showError('데이터 저장 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
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
      prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
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
        onChange={handleFileChange} 
      />

      {errorModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-6">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <AlertCircle size={28} />
              <h3 className="text-xl font-bold">오류</h3>
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

      {/* 상단 알림 메시지 */}
      {successMessage && (
        <div className="fixed top-5 right-5 z-50 bg-emerald-600 text-white px-6 py-3 rounded-xl shadow-lg flex items-center gap-2 font-medium">
          <span>{successMessage}</span>
        </div>
      )}

      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-3 text-indigo-600">
          <Landmark size={28} />
          <h1 className="text-xl font-bold tracking-tight text-slate-900">주식 포트폴리오 관리 시스템</h1>
        </div>
        <div className="text-sm text-emerald-600 font-medium bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200 flex items-center gap-1.5 shadow-xs">
          <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
          Firebase 클라우드 DB 연동됨
        </div>
      </header>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 py-8 sm:px-6 lg:px-8">
        
        {/* Tab Navigation */}
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
            
            {/* Top Section: Controls & Summary (버튼들 한 줄에 배치) */}
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col xl:flex-row gap-4 justify-between items-start xl:items-center">
              
              {/* Left Controls: 한 줄에 모두 배치되도록 flex-wrap 설정 및 간격 조정 */}
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Calendar size={16} className="text-slate-400" />
                  </div>
                  <input
                    type="text"
                    placeholder="납입연도 (예: 2026)"
                    className="pl-10 pr-4 py-2 w-44 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
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
                      ? 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200' 
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

                {/* 저장 버튼 신설 */}
                <button 
                  onClick={handleSaveToDatabase}
                  disabled={loading}
                  className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 transition-colors shadow-sm ml-2"
                >
                  <Save size={16} />
                  {loading ? '처리중...' : '저장'}
                </button>
              </div>

              {/* Right Summary */}
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

            {/* Bottom Section: Data Table */}
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

        {/* Portfolio Tab Content */}
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
