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
  CheckSquare,
  AlertCircle
} from 'lucide-react';

// 엑셀 처리를 위한 SheetJS 라이브러리를 CDN에서 동적 로드하도록 수정 (esbuild 에러 해결)

const initialData = [
  { id: 1, date: '2026-01-15', bank: '미래에셋', purpose: '연금', amount: 5000000 },
  { id: 2, date: '2026-03-10', bank: 'KB증권', purpose: 'IRP', amount: 3000000 },
  { id: 3, date: '2025-12-05', bank: '삼성증권', purpose: 'DC', amount: 4500000 },
];

const getTodayString = () => {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const formatCurrency = (amount) => {
  if (amount === 0 || !amount) return '0';
  return amount.toLocaleString('ko-KR');
};

export default function App() {
  const [activeTab, setActiveTab] = useState('payment');
  const [payments, setPayments] = useState(initialData);
  const [searchYearInput, setSearchYearInput] = useState('');
  const [appliedSearchYear, setAppliedSearchYear] = useState('');
  const [selectedIds, setSelectedIds] = useState([]);
  
  // 엑셀 업로드 관련 상태 및 Ref
  const fileInputRef = useRef(null);
  const [errorModal, setErrorModal] = useState({ isOpen: false, message: '' });

  useEffect(() => {
    // 환경에 XLSX 라이브러리가 없으면 스크립트 태그를 동적으로 생성하여 추가
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.async = true;
      document.body.appendChild(script);
    }
  }, []);

  const handleSearch = () => {
    setAppliedSearchYear(searchYearInput);
    setSelectedIds([]); 
  };

  const handleAddRow = () => {
    const newRow = {
      id: Date.now(),
      date: getTodayString(),
      bank: '미래에셋',
      purpose: '연금',
      amount: 0,
    };
    setPayments([newRow, ...payments]);
  };

  const handleDeleteRows = () => {
    if (selectedIds.length === 0) return;
    setPayments(payments.filter(p => !selectedIds.includes(p.id)));
    setSelectedIds([]);
  };

  const triggerExcelUpload = () => {
    if (fileInputRef.current) {
      fileInputRef.current.value = ""; // 동일한 파일 연속 업로드 가능하도록 초기화
      fileInputRef.current.click();
    }
  };

  const showError = (message) => {
    setErrorModal({ isOpen: true, message });
  };

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    // CDN을 통해 라이브러리가 완전히 로드되었는지 확인
    if (!window.XLSX) {
      showError('엑셀 처리 라이브러리를 불러오는 중입니다. 잠시 후 다시 시도해주세요.');
      return;
    }

    // 1. 파일 확장자 기본 검증
    const fileExt = file.name.split('.').pop().toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      showError('유효한 엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // 2. 단일 시트 검증
        if (workbook.SheetNames.length !== 1) {
          showError(`엑셀 파일에 시트가 ${workbook.SheetNames.length}개 있습니다. 반드시 1개의 시트만 존재해야 합니다.`);
          return;
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // 2차원 배열 형태로 데이터 추출 (헤더 포함)
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        if (jsonData.length === 0) {
          showError('엑셀 시트가 비어있습니다.');
          return;
        }

        const headers = jsonData[0];
        const requiredColumns = ['납입일자', '은행', '목적', '금액'];

        // 3. 필수 컬럼(열) 검증
        const missingColumns = requiredColumns.filter(col => !headers.includes(col));
        if (missingColumns.length > 0) {
          showError(`다음 필수 열이 누락되었습니다: [ ${missingColumns.join(', ')} ]\n(1행에 '납입일자', '은행', '목적', '금액' 열이 모두 있어야 합니다.)`);
          return;
        }

        // 4. 실제 데이터 존재 여부 검증 (2행 이상부터 데이터가 있어야 함)
        const dataRows = jsonData.slice(1).filter(row => row.length > 0);
        if (dataRows.length === 0) {
          showError('열 제목 아래에 실제 데이터가 존재하지 않습니다.');
          return;
        }

        // 컬럼 인덱스 매핑
        const colIndices = {
          date: headers.indexOf('납입일자'),
          bank: headers.indexOf('은행'),
          purpose: headers.indexOf('목적'),
          amount: headers.indexOf('금액'),
        };

        // 데이터 파싱 및 가공
        const newPayments = dataRows.map((row, index) => {
          // 날짜 서식 정리 (2026.01.15 또는 2026/01/15 -> 2026-01-15)
          let dateStr = String(row[colIndices.date] || getTodayString()).trim();
          dateStr = dateStr.replace(/[\.\/]/g, '-');
          
          // 금액 숫자만 추출
          const amountStr = String(row[colIndices.amount] || '0').replace(/[^0-9]/g, '');
          const amountNum = parseInt(amountStr, 10) || 0;

          return {
            id: Date.now() + index, // 고유 ID 부여
            date: dateStr,
            bank: row[colIndices.bank] || '기타',
            purpose: row[colIndices.purpose] || '기타',
            amount: amountNum,
          };
        });

        // 기존 데이터 맨 위에 새 데이터 추가
        setPayments(prev => [...newPayments, ...prev]);

      } catch (err) {
        console.error(err);
        showError('파일을 읽는 중 오류가 발생했습니다. 올바른 엑셀 파일인지 확인해 주세요.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleRowChange = (id, field, value) => {
    setPayments(payments.map(payment => {
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
    return payments.filter(payment => payment.date.startsWith(appliedSearchYear));
  }, [payments, appliedSearchYear]);

  const handleSelectAll = (e) => {
    if (e.target.checked) {
      setSelectedIds(filteredPayments.map(p => p.id));
    } else {
      setSelectedIds([]);
    }
  };

  const totalAmount = useMemo(() => {
    return payments.reduce((sum, current) => sum + current.amount, 0);
  }, [payments]);

  const filteredTotalAmount = useMemo(() => {
    return filteredPayments.reduce((sum, current) => sum + current.amount, 0);
  }, [filteredPayments]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      
      {/* 엑셀 파일 업로드를 위한 숨겨진 인풋 */}
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        className="hidden" 
        ref={fileInputRef} 
        onChange={handleFileChange} 
      />

      {/* 에러 모달 (alert 대체) */}
      {errorModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-6 transform transition-all">
            <div className="flex items-center gap-3 text-rose-600 mb-4">
              <AlertCircle size={28} />
              <h3 className="text-xl font-bold">엑셀 업로드 오류</h3>
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

      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 flex items-center shadow-sm">
        <div className="flex items-center gap-2 text-indigo-600">
          <Landmark size={28} />
          <h1 className="text-xl font-bold tracking-tight">주식 포트폴리오 관리</h1>
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
            
            {/* Top Section: Controls & Summary */}
            <div className="p-5 border-b border-slate-200 bg-slate-50/50 flex flex-col lg:flex-row gap-4 justify-between items-start lg:items-center">
              
              {/* Left Controls */}
              <div className="flex flex-wrap items-center gap-3">
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Calendar size={16} className="text-slate-400" />
                  </div>
                  <input
                    type="text"
                    placeholder="납입연도 (예: 2026)"
                    className="pl-10 pr-4 py-2 w-48 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 bg-white"
                    value={searchYearInput}
                    onChange={(e) => setSearchYearInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleSearch()}
                  />
                </div>
                
                <button 
                  onClick={handleSearch}
                  className="flex items-center gap-1.5 px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 transition-colors shadow-sm"
                >
                  <Search size={16} />
                  조회
                </button>
                <div className="w-px h-6 bg-slate-300 mx-1 hidden sm:block"></div>
                <button 
                  onClick={handleAddRow}
                  className="flex items-center gap-1.5 px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 transition-colors shadow-sm"
                >
                  <Plus size={16} />
                  추가
                </button>
                <button 
                  onClick={handleDeleteRows}
                  disabled={selectedIds.length === 0}
                  className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-medium transition-colors shadow-sm
                    ${selectedIds.length > 0 
                      ? 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200' 
                      : 'bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed'
                    }`}
                >
                  <Trash2 size={16} />
                  삭제 {selectedIds.length > 0 && `(${selectedIds.length})`}
                </button>
                <button 
                  onClick={triggerExcelUpload}
                  className="flex items-center gap-1.5 px-4 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 transition-colors shadow-sm"
                >
                  <FileSpreadsheet size={16} className="text-green-600" />
                  엑셀 업로드
                </button>
              </div>

              {/* Right Summary */}
              <div className="flex items-center gap-6 bg-white px-5 py-2.5 rounded-lg border border-slate-200 shadow-sm min-w-max w-full lg:w-auto">
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
                  {filteredPayments.length > 0 ? (
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
                            value={row.date}
                            onChange={(e) => handleRowChange(row.id, 'date', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none"
                          />
                        </td>
                        <td className="py-2 px-4">
                          <select 
                            value={row.bank}
                            onChange={(e) => handleRowChange(row.id, 'bank', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none"
                          >
                            <option value="미래에셋">미래에셋</option>
                            <option value="KB증권">KB증권</option>
                            <option value="삼성증권">삼성증권</option>
                            <option value="기타">기타</option>
                          </select>
                        </td>
                        <td className="py-2 px-4">
                          <select 
                            value={row.purpose}
                            onChange={(e) => handleRowChange(row.id, 'purpose', e.target.value)}
                            className="w-full px-3 py-1.5 border border-slate-200 rounded-md text-sm focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none"
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
                              className="w-full md:w-3/4 px-3 py-1.5 border border-slate-200 rounded-md text-sm text-right font-medium focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none"
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
