import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import { FileDown, RefreshCw, AlertTriangle, FileSpreadsheet, Building2, Users, Calendar } from 'lucide-react';
import FileUpload from './components/FileUpload';
import { processExcelFiles } from './utils/excelProcessor';

function App() {
  const [arrearsFile, setArrearsFile] = useState(null);
  const [residentsFile, setResidentsFile] = useState(null);
  const [rocYear, setRocYear] = useState('');
  const [rocMonth, setRocMonth] = useState('');
  const [status, setStatus] = useState('idle');
  const [errorMessage, setErrorMessage] = useState('');

  const handleProcess = async () => {
    if (!arrearsFile || !residentsFile) {
      setErrorMessage('請先上傳兩個檔案');
      return;
    }
    const year = parseInt(rocYear, 10);
    const month = parseInt(rocMonth, 10);
    if (!year || !month || month < 1 || month > 12) {
      setErrorMessage('請輸入正確的民國年份及月份');
      return;
    }

    setStatus('processing');
    setErrorMessage('');

    try {
      const blob = await processExcelFiles(arrearsFile, residentsFile, year, month);
      const fileName = `${year}年${month}月管理費催收.xlsx`;
      saveAs(blob, fileName);
      setStatus('success');
    } catch (error) {
      console.error(error);
      setStatus('error');
      setErrorMessage(error.message || '處理檔案時發生錯誤');
    }
  };

  const reset = () => {
    setArrearsFile(null);
    setResidentsFile(null);
    setRocYear('');
    setRocMonth('');
    setStatus('idle');
    setErrorMessage('');
  };

  const canProcess = arrearsFile && residentsFile && rocYear && rocMonth;

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col items-center justify-center p-6 font-sans text-gray-800">
      <div className="w-full max-w-4xl grid grid-cols-1 md:grid-cols-2 gap-8">

        {/* Header Section */}
        <div className="col-span-1 md:col-span-2 text-center mb-4">
          <div className="inline-flex items-center justify-center w-20 h-20 rounded-2xl bg-gradient-to-br from-blue-500 to-indigo-600 text-white mb-6 shadow-lg shadow-blue-200">
            <FileSpreadsheet size={40} />
          </div>
          <h1 className="text-4xl font-extrabold text-gray-900 tracking-tight">美河市房屋管理費欠繳名單整理</h1>
        </div>

        {/* Date Input Section */}
        <div className="col-span-1 md:col-span-2">
          <div className="bg-white rounded-2xl shadow-xl shadow-gray-100 border border-gray-100 overflow-hidden">
            <div className="bg-amber-500 p-4 flex items-center justify-between">
              <div className="flex items-center space-x-3 text-white">
                <Calendar className="w-6 h-6" />
                <h2 className="text-xl font-bold">指定年月</h2>
              </div>
              <span className="bg-amber-400 text-amber-50 text-xs font-bold px-2 py-1 rounded-full border border-amber-300">必填</span>
            </div>
            <div className="p-6 flex items-center gap-4 justify-center bg-amber-50/30">
              <label className="text-gray-700 font-medium">民國</label>
              <input
                type="number"
                value={rocYear}
                onChange={(e) => setRocYear(e.target.value)}
                placeholder="115"
                className="w-24 px-3 py-2 border border-gray-300 rounded-lg text-center text-lg font-semibold focus:outline-none focus:ring-2 focus:ring-amber-400 focus:border-transparent"
              />
              <span className="text-gray-700 font-medium">年</span>
              <input
                type="number"
                value={rocMonth}
                onChange={(e) => setRocMonth(e.target.value)}
                placeholder="2"
                min="1"
                max="12"
                className="w-20 px-3 py-2 border border-gray-300 rounded-lg text-center text-lg font-semibold focus:outline-none focus:ring-2 focus:ring-amber-400 focus:border-transparent"
              />
              <span className="text-gray-700 font-medium">月</span>
            </div>
          </div>
        </div>

        {/* Upload Card 1: Arrears List */}
        <div className="bg-white rounded-2xl shadow-xl shadow-gray-100 border border-gray-100 overflow-hidden flex flex-col h-full transform transition-all hover:-translate-y-1 duration-300">
          <div className="bg-blue-600 p-6 flex items-center justify-between">
            <div className="flex items-center space-x-3 text-white">
              <Building2 className="w-6 h-6" />
              <h2 className="text-xl font-bold">房屋管理費欠繳名單</h2>
            </div>
            <span className="bg-blue-500 text-blue-50 text-xs font-bold px-2 py-1 rounded-full border border-blue-400">步驟 1</span>
          </div>
          <div className="p-8 flex-1 flex flex-col justify-center bg-blue-50/30">
            <FileUpload
              label="請上傳管理費欠繳名單 (.xlsx)"
              file={arrearsFile}
              onFileSelect={setArrearsFile}
              variant="blue"
            />
            <div className="mt-4 p-4 bg-blue-50 rounded-lg text-sm text-blue-700 space-y-1">
              <p className="font-semibold mb-1">💡 檔案需求：</p>
              <ul className="list-disc list-inside space-y-1 opacity-80 pl-1">
                <li>C 欄：地址代碼 (如 530902)</li>
                <li>H 欄：欠費期間</li>
                <li>K 欄：欠費金額</li>
              </ul>
            </div>
          </div>
        </div>

        {/* Upload Card 2: Residents List */}
        <div className="bg-white rounded-2xl shadow-xl shadow-gray-100 border border-gray-100 overflow-hidden flex flex-col h-full transform transition-all hover:-translate-y-1 duration-300">
          <div className="bg-emerald-600 p-6 flex items-center justify-between">
            <div className="flex items-center space-x-3 text-white">
              <Users className="w-6 h-6" />
              <h2 className="text-xl font-bold">住戶名單</h2>
            </div>
            <span className="bg-emerald-500 text-emerald-50 text-xs font-bold px-2 py-1 rounded-full border border-emerald-400">步驟 2</span>
          </div>
          <div className="p-8 flex-1 flex flex-col justify-center bg-emerald-50/30">
            <FileUpload
              label="請上傳住戶名單 (.xlsx)"
              file={residentsFile}
              onFileSelect={setResidentsFile}
              variant="green"
            />
            <div className="mt-4 p-4 bg-emerald-50 rounded-lg text-sm text-emerald-700 space-y-1">
              <p className="font-semibold mb-1">💡 檔案需求：</p>
              <ul className="list-disc list-inside space-y-1 opacity-80 pl-1">
                <li>工作表名稱：<span className="font-mono bg-emerald-100 px-1 rounded">新店機廠捷17.18.19</span></li>
                <li>C 欄：地址</li>
                <li>H/I 欄：姓名/電話</li>
              </ul>
            </div>
          </div>
        </div>

        {/* Action Section */}
        <div className="col-span-1 md:col-span-2 mt-4 space-y-6">
          {errorMessage && (
            <div className="p-4 bg-red-50 border border-red-200 rounded-xl flex items-center text-red-700 animate-fade-in">
              <AlertTriangle className="w-6 h-6 mr-3 flex-shrink-0" />
              <span className="font-medium">{errorMessage}</span>
            </div>
          )}

          {status === 'success' && (
            <div className="p-4 bg-green-50 border border-green-200 rounded-xl flex items-center text-green-700 animate-fade-in">
              <FileDown className="w-6 h-6 mr-3" />
              <span className="font-medium">檔案已成功處理並開始下載！</span>
            </div>
          )}

          <div className="flex justify-center gap-4">
            <button
              onClick={handleProcess}
              disabled={status === 'processing' || !canProcess}
              className="w-full max-w-sm bg-gray-900 hover:bg-black text-white font-bold py-4 px-8 rounded-xl shadow-lg hover:shadow-xl transition-all duration-200 flex items-center justify-center disabled:opacity-50 disabled:cursor-not-allowed disabled:shadow-none transform active:scale-95"
            >
              {status === 'processing' ? (
                <>
                  <RefreshCw className="w-5 h-5 mr-2 animate-spin" />
                  處理中...
                </>
              ) : (
                <>
                  <RefreshCw className="w-5 h-5 mr-2" />
                  開始處理並下載
                </>
              )}
            </button>

            {(status === 'success' || status === 'error') && (
              <button
                onClick={reset}
                className="bg-white hover:bg-gray-50 text-gray-700 font-bold py-4 px-8 rounded-xl border border-gray-200 shadow-sm hover:shadow-md transition-all duration-200"
              >
                重置
              </button>
            )}
          </div>
        </div>

      </div>
      <footer className="mt-8 text-center text-gray-400 text-sm">
        <p>Last Updated: {new Date().toLocaleString()}</p>
      </footer>
    </div>
  );
}

export default App;
