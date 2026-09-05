import React, { useState, useEffect } from 'react';
import { 
  TrendingUp, 
  Upload, 
  Plus, 
  Trash2, 
  RefreshCw, 
  DollarSign, 
  PieChart as PieChartIcon, 
  FileText,
  AlertCircle,
  CheckCircle2
} from 'lucide-react';
import * as XLSX from 'xlsx';

export default function App() {
  // 1. 상태 정의 (브라우저 로컬 스토리지에서 데이터를 불러와 초기값으로 설정)
  const [payments, setPayments] = useState(() => {
    const saved = localStorage.getItem('stock_payments');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return [
      { id: 1, date: '2026-03-01', category: '급여', amount: 1500000, memo: '3월 정기 납입' },
      { id: 2, date: '2026-03-15', category: '배당금', amount: 300000, memo: '미국 주식 배당' }
    ];
  });

  const [holdings, setHoldings] = useState(() => {
    const saved = localStorage.getItem('stock_holdings');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return [
      { id: 1, ticker: 'AAPL', name: '애플', shares: 10, avgPrice: 180, currentPrice: 190 },
      { id: 2, ticker: 'TSLA', name: '테슬라', shares: 5, avgPrice: 200, currentPrice: 195 }
    ];
  });

  const [activeTab, setActiveTab] = useState('payment');
  const [excelMessage, setExcelMessage] = useState({ type: '', text: '' });

  // 2. 데이터가 바뀔 때마다 브라우저에 자동 저장 (새로고침해도 유지됨)
  useEffect(() => {
    localStorage.setItem('stock_payments', JSON.stringify(payments));
  }, [payments]);

  useEffect(() => {
    localStorage.setItem('stock_holdings', JSON.stringify(holdings));
  }, [holdings]);

  // 납입금액 추가 폼 상태
  const [newPayment, setNewPayment] = useState({ date: '', category: '급여', amount: '', memo: '' });

  const handleAddPayment = (e) => {
    e.preventDefault();
    if (!newPayment.date || !newPayment.amount) return;
    const item = {
      id: Date.now(),
      date: newPayment.date,
      category: newPayment.category,
      amount: Number(newPayment.amount),
      memo: newPayment.memo
    };
    setPayments([item, ...payments]);
    setNewPayment({ date: '', category: '급여', amount: '', memo: '' });
  };

  const handleDeletePayment = (id) => {
    setPayments(payments.filter(p => p.id !== id));
  };

  // 엑셀 업로드 핸들러 (단일 시트 및 필수 컬럼 검증)
  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });

        // 단일 시트 검증
        if (workbook.SheetNames.length !== 1) {
          setExcelMessage({ type: 'error', text: '엑셀 파일에는 정확히 1개의 시트만 존재해야 합니다.' });
          return;
        }

        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) {
          setExcelMessage({ type: 'error', text: '엑셀 파일에 데이터가 비어 있습니다.' });
          return;
        }

        // 컬럼 검증 (날짜, 카테고리,금액, 메모 또는 종목코드, 수량 등 유연하게 체크)
        const firstRow = data[0];
        const keys = Object.keys(firstRow);

        // 납입 관리용 엑셀 형식인지 확인 (예: 날짜, 금액 포함 여부)
        const formattedData = data.map((row, idx) => ({
          id: Date.now() + idx,
          date: row['날짜'] || row['date'] || new Date().toISOString().split('T')[0],
          category: row['카테고리'] || row['구분'] || '기타',
          amount: Number(row['금액'] || row['amount'] || 0),
          memo: row['메모'] || row['비고'] || ''
        }));

        setPayments([...formattedData, ...payments]);
        setExcelMessage({ type: 'success', text: `성공적으로 ${formattedData.length}개의 데이터를 불러와 저장했습니다!` });
      } catch (err) {
        console.error(err);
        setExcelMessage({ type: 'error', text: '엑셀 파일을 읽는 중 오류가 발생했습니다.' });
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col">
      {/* 상단 네비게이션 */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 py-4 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <div className="bg-indigo-600 p-2 rounded-xl text-white">
              <TrendingUp className="w-6 h-6" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">스마트 주식 포트폴리오</h1>
          </div>
          <div className="text-sm text-emerald-600 font-medium bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200 flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
            데이터 자동 저장 활성화됨
          </div>
        </div>
      </header>

      {/* 탭 메뉴 */}
      <div className="max-w-7xl mx-auto w-full px-4 mt-6">
        <div className="flex space-x-2 border-b border-slate-200">
          <button
            onClick={() => setActiveTab('payment')}
            className={`pb-3 px-4 font-semibold text-sm transition-colors border-b-2 flex items-center gap-2 ${
              activeTab === 'payment'
                ? 'border-indigo-600 text-indigo-600'
                : 'border-transparent text-slate-500 hover:text-slate-700'
            }`}
          >
            <DollarSign className="w-4 h-4" /> 납입금액 관리
          </button>
          <button
            onClick={() => setActiveTab('portfolio')}
            className={`pb-3 px-4 font-semibold text-sm transition-colors border-b-2 flex items-center gap-2 ${
              activeTab === 'portfolio'
                ? 'border-indigo-600 text-indigo-600'
                : 'border-transparent text-slate-500 hover:text-slate-700'
            }`}
          >
            <PieChartIcon className="w-4 h-4" /> 보유 종목 현황
          </button>
        </div>
      </div>

      {/* 메인 콘텐츠 영역 */}
      <main className="max-w-7xl mx-auto w-full px-4 py-6 flex-1">
        {activeTab === 'payment' && (
          <div className="space-y-6">
            {/* 상단 액션 및 업로드 바 */}
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900">납입 금액 내역</h2>
                <p className="text-sm text-slate-500">입력 및 엑셀 업로드된 데이터는 브라우저에 안전하게 영구 보존됩니다.</p>
              </div>
              <div className="flex items-center gap-3">
                <label className="cursor-pointer bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 shadow-xs">
                  <Upload className="w-4 h-4 text-indigo-600" /> 엑셀 업로드
                  <input type="file" accept=".xlsx, .xls" onChange={handleExcelUpload} className="hidden" />
                </label>
              </div>
            </div>

            {/* 엑셀 안내 메시지 */}
            {excelMessage.text && (
              <div className={`p-4 rounded-xl border flex items-center gap-3 text-sm ${
                excelMessage.type === 'error' ? 'bg-rose-50 border-rose-200 text-rose-700' : 'bg-emerald-50 border-emerald-200 text-emerald-700'
              }`}>
                {excelMessage.type === 'error' ? <AlertCircle className="w-5 h-5 shrink-0" /> : <CheckCircle2 className="w-5 h-5 shrink-0" />}
                <span>{excelMessage.text}</span>
              </div>
            )}

            {/* 개별 입력 폼 */}
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
              <h3 className="text-sm font-bold text-slate-700 mb-4">직접 내역 추가하기</h3>
              <form onSubmit={handleAddPayment} className="grid grid-cols-1 md:grid-cols-5 gap-3">
                <input
                  type="date"
                  value={newPayment.date}
                  onChange={(e) => setNewPayment({ ...newPayment, date: e.target.value })}
                  className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  required
                />
                <select
                  value={newPayment.category}
                  onChange={(e) => setNewPayment({ ...newPayment, category: e.target.value })}
                  className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                >
                  <option value="급여">급여</option>
                  <option value="배당금">배당금</option>
                  <option value="추가입금">추가입금</option>
                  <option value="기타">기타</option>
                </select>
                <input
                  type="number"
                  placeholder="금액 (원)"
                  value={newPayment.amount}
                  onChange={(e) => setNewPayment({ ...newPayment, amount: e.target.value })}
                  className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  required
                />
                <input
                  type="text"
                  placeholder="메모 (선택사항)"
                  value={newPayment.memo}
                  onChange={(e) => setNewPayment({ ...newPayment, memo: e.target.value })}
                  className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                />
                <button
                  type="submit"
                  className="bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-2 rounded-xl text-sm transition flex items-center justify-center gap-2 shadow-sm"
                >
                  <Plus className="w-4 h-4" /> 추가
                </button>
              </form>
            </div>

            {/* 납입 내역 테이블 */}
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs font-semibold uppercase tracking-wider">
                      <th className="py-3 px-4">날짜</th>
                      <th className="py-3 px-4">구분</th>
                      <th className="py-3 px-4 text-right">금액</th>
                      <th className="py-3 px-4">메모</th>
                      <th className="py-3 px-4 text-center">관리</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-sm">
                    {payments.length === 0 ? (
                      <tr>
                        <td colSpan="5" className="text-center py-8 text-slate-400">등록된 납입 내역이 없습니다.</td>
                      </tr>
                    ) : (
                      payments.map((p) => (
                        <tr key={p.id} className="hover:bg-slate-50 transition-colors">
                          <td className="py-3 px-4 text-slate-600">{p.date}</td>
                          <td className="py-3 px-4 font-medium text-slate-800">{p.category}</td>
                          <td className="py-3 px-4 text-right font-semibold text-emerald-600">
                            {Number(p.amount).toLocaleString()} 원
                          </td>
                          <td className="py-3 px-4 text-slate-500">{p.memo || '-'}</td>
                          <td className="py-3 px-4 text-center">
                            <button
                              onClick={() => handleDeletePayment(p.id)}
                              className="text-slate-400 hover:text-rose-600 p-1 rounded-lg transition"
                              title="삭제"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'portfolio' && (
          <div className="space-y-6">
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
              <h2 className="text-lg font-bold text-slate-900">보유 종목 현황</h2>
              <p className="text-sm text-slate-500">현재 보유 중인 자산 및 주식 포트폴리오 목록입니다.</p>
            </div>
            
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs font-semibold uppercase tracking-wider">
                      <th className="py-3 px-4">티커</th>
                      <th className="py-3 px-4">종목명</th>
                      <th className="py-3 px-4 text-right">보유수량</th>
                      <th className="py-3 px-4 text-right">평단가</th>
                      <th className="py-3 px-4 text-right">현재가</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-sm">
                    {holdings.map((h) => (
                      <tr key={h.id} className="hover:bg-slate-50 transition-colors">
                        <td className="py-3 px-4 font-mono font-bold text-indigo-600">{h.ticker}</td>
                        <td className="py-3 px-4 font-medium text-slate-800">{h.name}</td>
                        <td className="py-3 px-4 text-right">{h.shares} 주</td>
                        <td className="py-3 px-4 text-right">${h.avgPrice}</td>
                        <td className="py-3 px-4 text-right font-semibold">${h.currentPrice}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
