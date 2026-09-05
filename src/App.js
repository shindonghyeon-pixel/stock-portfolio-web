import React, { useState, useEffect } from 'react';
import { 
  TrendingUp, 
  Upload, 
  Plus, 
  Trash2, 
  Search, 
  DollarSign, 
  PieChart as PieChartIcon, 
  AlertCircle,
  CheckCircle2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { db } from './firebase';
import { 
  collection, 
  getDocs, 
  addDoc, 
  deleteDoc, 
  doc, 
  serverTimestamp 
} from 'firebase/firestore';

export default function App() {
  const [activeTab, setActiveTab] = useState('payment');
  
  const [payments, setPayments] = useState([]);
  const [loading, setLoading] = useState(true);

  const [searchYear, setSearchYear] = useState('');
  const [filteredPayments, setFilteredPayments] = useState([]);
  const [selectedIds, setSelectedIds] = useState([]);
  const [excelMessage, setExcelMessage] = useState({ type: '', text: '' });

  const [newRow, setNewRow] = useState({
    date: new Date().toISOString().split('T')[0],
    bank: '미래에셋',
    purpose: '연금',
    amount: ''
  });

  const fetchPayments = async () => {
    try {
      setLoading(true);
      const querySnapshot = await getDocs(collection(db, "payments"));
      const dataList = querySnapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      dataList.sort((a, b) => new Date(b.date) - new Date(a.date));
      setPayments(dataList);
      setFilteredPayments(dataList);
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
    if (!searchYear.trim()) {
      setFilteredPayments(payments);
    } else {
      const filtered = payments.filter(p => p.date && p.date.startsWith(searchYear));
      setFilteredPayments(filtered);
    }
    setSelectedIds([]);
  };

  const handleAddRow = async () => {
    if (!newRow.date || !newRow.amount) {
      alert('납입일자와 금액은 필수 입력입니다.');
      return;
    }

    try {
      const itemToSave = {
        date: newRow.date,
        bank: newRow.bank,
        purpose: newRow.purpose,
        amount: Number(newRow.amount),
        createdAt: serverTimestamp()
      };

      await addDoc(collection(db, "payments"), itemToSave);
      await fetchPayments();
      
      setNewRow({
        date: new Date().toISOString().split('T')[0],
        bank: '미래에셋',
        purpose: '연금',
        amount: ''
      });
      setExcelMessage({ type: 'success', text: '새 내역이 데이터베이스에 안전하게 저장되었습니다.' });
    } catch (error) {
      console.error("추가 실패:", error);
      alert('데이터 저장 중 오류가 발생했습니다.');
    }
  };

  const handleSelectOne = (id) => {
    if (selectedIds.includes(id)) {
      setSelectedIds(selectedIds.filter(item => item !== id));
    } else {
      setSelectedIds([...selectedIds, id]);
    }
  };

  const handleSelectAll = (e) => {
    if (e.target.checked) {
      setSelectedIds(filteredPayments.map(p => p.id));
    } else {
      setSelectedIds([]);
    }
  };

  const handleDeleteSelected = async () => {
    if (selectedIds.length === 0) {
      alert('삭제할 항목을 선택해주세요.');
      return;
    }

    if (!window.confirm('선택한 항목들을 정말 삭제하시겠습니까?')) return;

    try {
      for (const id of selectedIds) {
        await deleteDoc(doc(db, "payments", id));
      }
      await fetchPayments();
      setSelectedIds([]);
      setExcelMessage({ type: 'success', text: '선택한 항목이 삭제되었습니다.' });
    } catch (error) {
      console.error("삭제 실패:", error);
      alert('삭제 중 오류가 발생했습니다.');
    }
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) {
          setExcelMessage({ type: 'error', text: '엑셀 파일에 데이터가 비어 있습니다.' });
          return;
        }

        for (const row of data) {
          const itemToSave = {
            date: row['납입일자'] || row['date'] || new Date().toISOString().split('T')[0],
            bank: row['은행'] || row['bank'] || '미래에셋',
            purpose: row['목적'] || row['purpose'] || '연금',
            amount: Number(row['금액'] || row['amount'] || 0),
            createdAt: serverTimestamp()
          };
          await addDoc(collection(db, "payments"), itemToSave);
        }

        await fetchPayments();
        setExcelMessage({ type: 'success', text: '엑셀 데이터를 성공적으로 업로드했습니다!' });
      } catch (err) {
        console.error(err);
        setExcelMessage({ type: 'error', text: '엑셀 파일을 읽는 중 오류가 발생했습니다.' });
      }
    };
    reader.readAsBinaryString(file);
  };

  const totalAmountAll = payments.reduce((sum, p) => sum + Number(p.amount || 0), 0);
  const totalAmountFiltered = filteredPayments.reduce((sum, p) => sum + Number(p.amount || 0), 0);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 py-4 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <div className="bg-indigo-600 p-2 rounded-xl text-white">
              <TrendingUp className="w-6 h-6" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">주식 포트폴리오 관리 시스템</h1>
          </div>
          <div className="text-sm text-emerald-600 font-medium bg-emerald-50 px-3 py-1.5 rounded-full border border-emerald-200 flex items-center gap-1.5 shadow-xs">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
            클라우드 DB 실시간 동기화 중
          </div>
        </div>
      </header>

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

      <main className="max-w-7xl mx-auto w-full px-4 py-6 flex-1 space-y-6">
        {activeTab === 'payment' && (
          <div className="space-y-6">
            <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
              <div className="flex flex-wrap items-center gap-3">
                <div className="flex items-center gap-2 bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5">
                  <span className="text-sm font-medium text-slate-600">납입연도</span>
                  <input
                    type="text"
                    placeholder="예: 2026"
                    value={searchYear}
                    onChange={(e) => setSearchYear(e.target.value)}
                    className="w-24 bg-transparent text-sm focus:outline-none font-semibold text-slate-800"
                  />
                </div>

                <button
                  onClick={handleSearch}
                  className="bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-1.5 shadow-xs"
                >
                  <Search className="w-4 h-4" /> 조회
                </button>

                <button
                  onClick={handleDeleteSelected}
                  className="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-1.5 shadow-xs"
                >
                  <Trash2 className="w-4 h-4" /> 삭제
                </button>

                <label className="cursor-pointer bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 shadow-xs">
                  <Upload className="w-4 h-4 text-indigo-600" /> 엑셀 업로드
                  <input type="file" accept=".xlsx, .xls" onChange={handleExcelUpload} className="hidden" />
                </label>
              </div>

              <div className="flex items-center gap-6 bg-slate-50 border border-slate-200 px-5 py-3 rounded-xl w-full lg:w-auto justify-around lg:justify-end">
                <div className="text-right">
                  <p className="text-xs text-slate-500 font-medium">납입 총액 (전체)</p>
                  <p className="text-base font-bold text-slate-800">{totalAmountAll.toLocaleString()} 원</p>
                </div>
                <div className="h-8 w-px bg-slate-200"></div>
                <div className="text-right">
                  <p className="text-xs text-indigo-600 font-medium">조회 총액 (필터)</p>
                  <p className="text-base font-bold text-indigo-600">{totalAmountFiltered.toLocaleString()} 원</p>
                </div>
              </div>
            </div>

            {excelMessage.text && (
              <div className={`p-4 rounded-xl border flex items-center gap-3 text-sm ${
                excelMessage.type === 'error' ? 'bg-rose-50 border-rose-200 text-rose-700' : 'bg-emerald-50 border-emerald-200 text-emerald-700'
              }`}>
                {excelMessage.type === 'error' ? <AlertCircle className="w-5 h-5 shrink-0" /> : <CheckCircle2 className="w-5 h-5 shrink-0" />}
                <span>{excelMessage.text}</span>
              </div>
            )}

            <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-wrap items-center gap-3">
              <span className="text-sm font-bold text-slate-700 ml-1">직접 추가:</span>
              <input
                type="date"
                value={newRow.date}
                onChange={(e) => setNewRow({ ...newRow, date: e.target.value })}
                className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              />
              <select
                value={newRow.bank}
                onChange={(e) => setNewRow({ ...newRow, bank: e.target.value })}
                className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              >
                <option value="미래에셋">미래에셋</option>
                <option value="KB증권">KB증권</option>
                <option value="삼성증권">삼성증권</option>
              </select>
              <select
                value={newRow.purpose}
                onChange={(e) => setNewRow({ ...newRow, purpose: e.target.value })}
                className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              >
                <option value="연금">연금</option>
                <option value="IRP">IRP</option>
                <option value="DC">DC</option>
                <option value="기타">기타</option>
              </select>
              <input
                type="number"
                placeholder="금액 (원 입력)"
                value={newRow.amount}
                onChange={(e) => setNewRow({ ...newRow, amount: e.target.value })}
                className="px-3 py-2 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 w-40"
              />
              <button
                onClick={handleAddRow}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded-xl text-sm font-medium transition flex items-center gap-1.5 shadow-sm ml-auto"
              >
                <Plus className="w-4 h-4" /> 추가
              </button>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs font-semibold uppercase tracking-wider">
                      <th className="py-3 px-4 w-12 text-center">
                        <input
                          type="checkbox"
                          onChange={handleSelectAll}
                          checked={filteredPayments.length > 0 && selectedIds.length === filteredPayments.length}
                          className="rounded text-indigo-600 focus:ring-indigo-500 w-4 h-4 cursor-pointer"
                        />
                      </th>
                      <th className="py-3 px-4">납입일자</th>
                      <th className="py-3 px-4">은행</th>
                      <th className="py-3 px-4">목적 (연금유형)</th>
                      <th className="py-3 px-4 text-right">금액</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-sm">
                    {loading ? (
                      <tr>
                        <td colSpan="5" className="text-center py-12 text-slate-400">데이터를 불러오는 중입니다...</td>
                      </tr>
                    ) : filteredPayments.length === 0 ? (
                      <tr>
                        <td colSpan="5" className="text-center py-12 text-slate-400">조회된 납입 내역이 없습니다.</td>
                      </tr>
                    ) : (
                      filteredPayments.map((p) => (
                        <tr key={p.id} className="hover:bg-slate-50 transition-colors">
                          <td className="py-3 px-4 text-center">
                            <input
                              type="checkbox"
                              checked={selectedIds.includes(p.id)}
                              onChange={() => handleSelectOne(p.id)}
                              className="rounded text-indigo-600 focus:ring-indigo-500 w-4 h-4 cursor-pointer"
                            />
                          </td>
                          <td className="py-3 px-4 text-slate-700 font-medium">{p.date}</td>
                          <td className="py-3 px-4 text-slate-800">{p.bank}</td>
                          <td className="py-3 px-4">
                            <span className="inline-block bg-indigo-50 text-indigo-700 font-medium px-2.5 py-1 rounded-lg text-xs border border-indigo-100">
                              {p.purpose}
                            </span>
                          </td>
                          <td className="py-3 px-4 text-right font-bold text-emerald-600">
                            {Number(p.amount || 0).toLocaleString()} 원
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
              <p className="text-sm text-slate-500">현재 포트폴리오에 담긴 자산 목록입니다.</p>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
