import React, { useState, useMemo } from 'react';
import { createRoot } from 'react-dom/client';

/**
 * Excel Processing Tool v4
 * Features:
 * 1. Upload XLSX/CSV file
 * 2. Optional date range filtering (Leave blank for all data)
 * 3. Group by email (unique key)
 * 4. Priority Logic:
 *    - Earliest date wins
 *    - If same date, "社媒端" source wins
 *    - If same date and same source, combine all owners found
 * 5. Dynamic Owners: Uses all names found in the file within the selected (or total) scope.
 * 6. Export: Email, Owners, Date, and Order Source.
 */

const PRIORITY_SOURCE = "社媒端";

interface RecordInfo {
  date: Date;
  source: string;
  owners: Set<string>;
}

function App() {
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");

  // Config states
  const [emailCol, setEmailCol] = useState("");
  const [ownerCol, setOwnerCol] = useState("");
  const [dateCol, setDateCol] = useState("");
  const [sourceCol, setSourceCol] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = (window as any).XLSX.read(bstr, { type: 'binary', cellDates: true });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = (window as any).XLSX.utils.sheet_to_json(ws);

        if (jsonData.length > 0) {
          setData(jsonData);
          const firstRow = jsonData[0];
          const cols = Object.keys(firstRow);
          setColumns(cols);
          
          // Auto-detect columns based on common names
          setEmailCol(cols.find(c => c.includes('邮箱') || c.toLowerCase().includes('email')) || "");
          setOwnerCol(cols.find(c => c.includes('负责人') || c.toLowerCase().includes('owner')) || "");
          setDateCol(cols.find(c => c.includes('时间') || c.includes('日期') || c.toLowerCase().includes('date')) || "");
          setSourceCol(cols.find(c => c.includes('来源') || c.includes('端口') || c.toLowerCase().includes('source')) || "");
        }
      } catch (err) {
        alert("文件读取失败，请确保格式正确。");
        console.error(err);
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  const { processedData, detectedOwners } = useMemo(() => {
    if (!data.length || !emailCol || !ownerCol) return { processedData: [], detectedOwners: new Set<string>() };

    const emailMap = new Map<string, RecordInfo>();
    const allOwnersInScope = new Set<string>();

    data.forEach(row => {
      const email = String(row[emailCol] || "").trim();
      if (!email) return;

      const rawDate = row[dateCol];
      // Convert rawDate to Date object or use a placeholder if empty/invalid
      let rowDate = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(rawDate) : null);
      const isInvalidDate = !rowDate || isNaN(rowDate.getTime());

      // Filter logic: Only filter if user provided start/end dates
      const s = startDate ? new Date(startDate) : null;
      const e = endDate ? new Date(endDate) : null;

      if (s || e) {
          // If a filter range is active, rows without valid dates are skipped
          if (isInvalidDate) return;
          if (s && rowDate! < s) return;
          if (e) {
              const adjustedEnd = new Date(e);
              adjustedEnd.setHours(23, 59, 59, 999);
              if (rowDate! > adjustedEnd) return;
          }
      }

      // If no date range specified by user, all data is "In Scope"
      const effectiveDate = isInvalidDate ? new Date(8640000000000000) : rowDate!;
      const source = String(row[sourceCol] || "").trim();
      const rawOwner = String(row[ownerCol] || "").trim();
      
      const rowOwners = new Set<string>();
      const splitOwners = rawOwner.split(/[,，;；\s\/\\]+/).map(o => o.trim()).filter(o => o.length > 0);
      
      splitOwners.forEach(o => {
          rowOwners.add(o);
          allOwnersInScope.add(o);
      });

      if (rowOwners.size === 0) return;

      if (!emailMap.has(email)) {
        emailMap.set(email, {
          date: effectiveDate,
          source: source,
          owners: rowOwners
        });
      } else {
        const best = emailMap.get(email)!;
        
        const bestDateStr = best.date === null || best.date.getTime() === 8640000000000000 ? "N/A" : best.date.toISOString().split('T')[0];
        const rowDateStr = isInvalidDate ? "N/A" : rowDate!.toISOString().split('T')[0];

        if (effectiveDate < best.date && rowDateStr !== bestDateStr) {
          // Rule 1: Earliest date wins
          emailMap.set(email, { date: effectiveDate, source: source, owners: rowOwners });
        } else if (rowDateStr === bestDateStr) {
          // Rule 2: Same day tie-breaker
          const rowIsPriority = source.includes(PRIORITY_SOURCE);
          const bestIsPriority = best.source.includes(PRIORITY_SOURCE);

          if (rowIsPriority && !bestIsPriority) {
            emailMap.set(email, { date: effectiveDate, source: source, owners: rowOwners });
          } else if (rowIsPriority === bestIsPriority) {
            // Rule 3: Merge owners
            rowOwners.forEach(o => best.owners.add(o));
          }
        }
      }
    });

    const result = Array.from(emailMap.entries()).map(([email, info]) => ({
      "邮箱": email,
      "负责人": Array.from(info.owners).join("、"),
      "日期": info.date.getTime() === 8640000000000000 ? "未知日期" : info.date.toISOString().split('T')[0],
      "订单来源": info.source
    }));

    return { processedData: result, detectedOwners: allOwnersInScope };
  }, [data, emailCol, ownerCol, dateCol, sourceCol, startDate, endDate]);

  const exportToExcel = () => {
    if (!processedData.length) {
      alert("没有符合条件的数据可供导出。");
      return;
    }

    const ws = (window as any).XLSX.utils.json_to_sheet(processedData);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "统计结果");
    (window as any).XLSX.writeFile(wb, `统计结果_${new Date().getTime()}.xlsx`);
  };

  return (
    <div className="min-h-screen pb-12">
      {/* Header */}
      <header className="gradient-bg text-white py-12 px-4 shadow-lg mb-8">
        <div className="max-w-4xl mx-auto">
          <h1 className="text-3xl font-bold mb-2 flex items-center">
            <i className="fas fa-file-excel mr-3"></i>
            Excel 数据统计工具 v4
          </h1>
          <p className="opacity-90">
            全量模式：日期未指定时，将识别并统计上传的所有数据。
          </p>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 space-y-6">
        {/* Step 1: Upload */}
        <section className="glass-card p-6 rounded-2xl shadow-sm">
          <h2 className="text-xl font-semibold mb-4 text-gray-800 flex items-center">
            <span className="bg-indigo-100 text-indigo-600 w-8 h-8 rounded-full flex items-center justify-center mr-3 text-sm">1</span>
            上传目标表格
          </h2>
          <div className="border-2 border-dashed border-gray-200 rounded-xl p-8 transition-colors hover:border-indigo-300 group relative">
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            />
            <div className="text-center">
              <i className={`fas ${fileName ? 'fa-check-circle text-green-500' : 'fa-cloud-upload-alt text-gray-400'} text-4xl mb-3`}></i>
              <p className="text-gray-600 font-medium">
                {fileName ? `已选择: ${fileName}` : "点击或拖拽文件到此处上传"}
              </p>
              <p className="text-xs text-gray-400 mt-2">支持 .xlsx, .xls, .csv 格式</p>
            </div>
          </div>
        </section>

        {data.length > 0 && (
          <>
            {/* Step 2: Configuration */}
            <section className="glass-card p-6 rounded-2xl shadow-sm">
              <h2 className="text-xl font-semibold mb-4 text-gray-800 flex items-center">
                <span className="bg-indigo-100 text-indigo-600 w-8 h-8 rounded-full flex items-center justify-center mr-3 text-sm">2</span>
                配置筛选与映射
              </h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-4">
                  <h3 className="text-sm font-medium text-gray-500 uppercase tracking-wider">列名映射</h3>
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">邮箱字段 (Key)</label>
                    <select 
                      value={emailCol} 
                      onChange={(e) => setEmailCol(e.target.value)}
                      className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                    >
                      <option value="">请选择列...</option>
                      {columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">负责人字段</label>
                    <select 
                      value={ownerCol} 
                      onChange={(e) => setOwnerCol(e.target.value)}
                      className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                    >
                      <option value="">请选择列...</option>
                      {columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">订单来源/端口</label>
                    <select 
                      value={sourceCol} 
                      onChange={(e) => setSourceCol(e.target.value)}
                      className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                    >
                      <option value="">请选择列...</option>
                      {columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-gray-600 mb-1">日期字段</label>
                    <select 
                      value={dateCol} 
                      onChange={(e) => setDateCol(e.target.value)}
                      className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                    >
                      <option value="">请选择列...</option>
                      {columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                </div>

                <div className="space-y-4">
                  <h3 className="text-sm font-medium text-gray-500 uppercase tracking-wider">时间区间筛选</h3>
                  <div className="grid grid-cols-1 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-gray-600 mb-1">开始日期 (可选，留空则全选)</label>
                      <input 
                        type="date" 
                        value={startDate}
                        onChange={(e) => setStartDate(e.target.value)}
                        className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-gray-600 mb-1">结束日期 (可选，留空则全选)</label>
                      <input 
                        type="date" 
                        value={endDate}
                        onChange={(e) => setEndDate(e.target.value)}
                        className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                      />
                    </div>
                  </div>
                  
                  {detectedOwners.size > 0 && (
                    <div className="mt-4">
                      <h4 className="text-xs font-semibold text-gray-500 mb-2 uppercase">作用域内识别到的负责人:</h4>
                      <div className="flex flex-wrap gap-1 max-h-32 overflow-y-auto p-2 bg-gray-50 rounded-lg border border-gray-100">
                        {Array.from(detectedOwners).map(name => (
                          <span key={name} className="px-2 py-0.5 bg-white border border-gray-200 rounded text-[10px] text-gray-600">
                            {name}
                          </span>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </div>
              
              <div className="mt-6 p-4 bg-green-50 rounded-xl border border-green-100">
                <h4 className="text-sm font-bold text-green-800 mb-1 flex items-center">
                  <i className="fas fa-check-double mr-2"></i> 筛选逻辑说明
                </h4>
                <p className="text-xs text-green-700">
                  当前处于<strong>全量识别模式</strong>：如果您未选择开始日期和结束日期，系统将自动识别并统计上传文件中<strong>所有</strong>有效的记录，并依然遵循日期最早优先、同日期端口优先的规则进行归属判定。
                </p>
              </div>
            </section>

            {/* Step 3: Result Preview & Export */}
            <section className="glass-card p-6 rounded-2xl shadow-sm">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
                <h2 className="text-xl font-semibold text-gray-800 flex items-center">
                  <span className="bg-indigo-100 text-indigo-600 w-8 h-8 rounded-full flex items-center justify-center mr-3 text-sm">3</span>
                  处理结果预览 ({processedData.length} 条)
                </h2>
                <button
                  onClick={exportToExcel}
                  disabled={processedData.length === 0}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-2.5 rounded-xl font-semibold transition-all shadow-lg active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed flex items-center"
                >
                  <i className="fas fa-file-download mr-2"></i>
                  导出统计结果
                </button>
              </div>

              {processedData.length > 0 ? (
                <div className="overflow-hidden border border-gray-100 rounded-xl">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-sm">
                      <thead>
                        <tr className="bg-gray-50 text-gray-500 border-b border-gray-100">
                          <th className="px-6 py-4 font-bold uppercase text-[10px] tracking-widest">邮箱</th>
                          <th className="px-6 py-4 font-bold uppercase text-[10px] tracking-widest">最终负责人</th>
                          <th className="px-6 py-4 font-bold uppercase text-[10px] tracking-widest text-center">归属日期</th>
                          <th className="px-6 py-4 font-bold uppercase text-[10px] tracking-widest">订单来源</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100 bg-white">
                        {processedData.slice(0, 15).map((row, idx) => (
                          <tr key={idx} className="hover:bg-indigo-50/30 transition-colors">
                            <td className="px-6 py-3.5 text-gray-800 font-medium">{row["邮箱"]}</td>
                            <td className="px-6 py-3.5">
                              <div className="flex flex-wrap gap-1">
                                {row["负责人"].split('、').map((name, nIdx) => (
                                    <span key={nIdx} className="inline-block bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded text-[10px] font-bold">
                                        {name}
                                    </span>
                                ))}
                              </div>
                            </td>
                            <td className="px-6 py-3.5 text-gray-500 text-center font-mono text-xs">{row["日期"]}</td>
                            <td className="px-6 py-3.5 text-gray-400 italic text-xs truncate max-w-[120px]">{row["订单来源"]}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {processedData.length > 15 && (
                    <div className="px-6 py-4 text-center bg-gray-50 text-gray-400 text-xs italic border-t border-gray-100">
                      预览仅显示前 15 条，请导出 Excel 以获取完整 {processedData.length} 条数据。
                    </div>
                  )}
                </div>
              ) : (
                <div className="text-center py-16 border-2 border-dashed border-gray-100 rounded-2xl bg-gray-50/50">
                  <i className="fas fa-search text-gray-200 text-5xl mb-4"></i>
                  <p className="text-gray-400 font-medium">配置字段后，所有匹配数据将即刻呈现</p>
                </div>
              )}
            </section>
          </>
        )}
      </main>

      {loading && (
        <div className="fixed inset-0 bg-indigo-900/20 backdrop-blur-md flex items-center justify-center z-50">
          <div className="bg-white p-10 rounded-3xl shadow-2xl flex flex-col items-center">
            <div className="relative">
              <div className="animate-spin rounded-full h-16 w-16 border-[3px] border-indigo-600 border-t-transparent"></div>
              <i className="fas fa-sync-alt absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-indigo-600"></i>
            </div>
            <p className="font-bold text-gray-800 mt-6 tracking-wide">全量数据分析中...</p>
            <p className="text-gray-400 text-xs mt-1">请稍候，正在为您生成统计报表</p>
          </div>
        </div>
      )}
    </div>
  );
}

const root = createRoot(document.getElementById('root')!);
root.render(<App />);