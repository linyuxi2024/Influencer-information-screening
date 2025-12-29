import React, { useState, useMemo } from 'react';
import { createRoot } from 'react-dom/client';

/**
 * Excel Processing Tool v5
 * 新增：负责人维度统计。
 * 逻辑：负责人为 Key -> 邮箱去重 (日期早优先, 社媒优先) -> 分人导出文件。
 */

const PRIORITY_SOURCE = "社媒端";

interface RecordInfo {
  email: string;
  owner: string;
  date: Date;
  source: string;
}

function App() {
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [activeTab, setActiveTab] = useState<'email' | 'owner'>('email');

  // 配置状态
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
          
          setEmailCol(cols.find(c => c.includes('邮箱') || c.toLowerCase().includes('email')) || "");
          setOwnerCol(cols.find(c => c.includes('负责人') || c.toLowerCase().includes('owner')) || "");
          setDateCol(cols.find(c => c.includes('时间') || c.includes('日期') || c.toLowerCase().includes('date')) || "");
          setSourceCol(cols.find(c => c.includes('来源') || c.includes('端口') || c.toLowerCase().includes('source')) || "");
        }
      } catch (err) {
        alert("文件读取失败。");
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  // --- 逻辑1: 邮箱维度 (原有功能) ---
  const emailCentricData = useMemo(() => {
    if (activeTab !== 'email' || !data.length || !emailCol || !ownerCol) return [];
    const emailMap = new Map<string, any>();
    data.forEach(row => {
      const email = String(row[emailCol] || "").trim();
      if (!email) return;
      const rawDate = row[dateCol];
      let rowDate = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(rawDate) : null);
      const isInvalidDate = !rowDate || isNaN(rowDate.getTime());
      const s = startDate ? new Date(startDate) : null;
      const e = endDate ? new Date(endDate) : null;
      if (s || e) {
        if (isInvalidDate) return;
        if (s && rowDate! < s) return;
        if (e) { const adjE = new Date(e); adjE.setHours(23, 59, 59, 999); if (rowDate! > adjE) return; }
      }
      const effectiveDate = isInvalidDate ? new Date(8640000000000000) : rowDate!;
      const source = String(row[sourceCol] || "").trim();
      const ownersList = String(row[ownerCol] || "").split(/[,，;；\s\/\\]+/).map(o => o.trim()).filter(o => o.length > 0);
      
      if (!emailMap.has(email)) {
        emailMap.set(email, { date: effectiveDate, source, owners: new Set(ownersList) });
      } else {
        const best = emailMap.get(email)!;
        const bestDateStr = best.date.getTime() === 8640000000000000 ? "N/A" : best.date.toISOString().split('T')[0];
        const rowDateStr = isInvalidDate ? "N/A" : rowDate!.toISOString().split('T')[0];
        if (effectiveDate < best.date && rowDateStr !== bestDateStr) {
          emailMap.set(email, { date: effectiveDate, source, owners: new Set(ownersList) });
        } else if (rowDateStr === bestDateStr) {
          if (source.includes(PRIORITY_SOURCE) && !best.source.includes(PRIORITY_SOURCE)) {
            emailMap.set(email, { date: effectiveDate, source, owners: new Set(ownersList) });
          } else if (source.includes(PRIORITY_SOURCE) === best.source.includes(PRIORITY_SOURCE)) {
            ownersList.forEach(o => best.owners.add(o));
          }
        }
      }
    });
    return Array.from(emailMap.entries()).map(([email, info]) => ({
      "邮箱": email,
      "负责人": Array.from(info.owners).join("、"),
      "日期": info.date.getTime() === 8640000000000000 ? "未知日期" : info.date.toISOString().split('T')[0],
      "订单来源": info.source
    }));
  }, [data, emailCol, ownerCol, dateCol, sourceCol, startDate, endDate, activeTab]);

  // --- 逻辑2: 负责人维度 (新增功能) ---
  const ownerCentricData = useMemo(() => {
    if (activeTab !== 'owner' || !data.length || !emailCol || !ownerCol) return new Map<string, RecordInfo[]>();
    const masterMap = new Map<string, Map<string, RecordInfo>>();
    data.forEach(row => {
      const rawOwners = String(row[ownerCol] || "").trim();
      if (!rawOwners) return;
      const rawDate = row[dateCol];
      let rowDate = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(rawDate) : null);
      const isInvalidDate = !rowDate || isNaN(rowDate.getTime());
      const s = startDate ? new Date(startDate) : null;
      const e = endDate ? new Date(endDate) : null;
      if (s || e) {
        if (isInvalidDate) return;
        if (s && rowDate! < s) return;
        if (e) { const adjE = new Date(e); adjE.setHours(23, 59, 59, 999); if (rowDate! > adjE) return; }
      }
      const email = String(row[emailCol] || "").trim();
      if (!email) return;
      const source = String(row[sourceCol] || "").trim();
      const effectiveDate = isInvalidDate ? new Date(8640000000000000) : rowDate!;
      const ownersList = rawOwners.split(/[,，;；\s\/\\]+/).map(o => o.trim()).filter(o => o.length > 0);

      ownersList.forEach(owner => {
        if (!masterMap.has(owner)) masterMap.set(owner, new Map());
        const ownerEmails = masterMap.get(owner)!;
        const current: RecordInfo = { email, owner, date: effectiveDate, source };
        if (!ownerEmails.has(email)) {
          ownerEmails.set(email, current);
        } else {
          const best = ownerEmails.get(email)!;
          const bestDStr = best.date.getTime() === 8640000000000000 ? "N/A" : best.date.toISOString().split('T')[0];
          const currDStr = effectiveDate.getTime() === 8640000000000000 ? "N/A" : effectiveDate.toISOString().split('T')[0];
          if (effectiveDate < best.date && currDStr !== bestDStr) {
            ownerEmails.set(email, current);
          } else if (currDStr === bestDStr) {
            if (source.includes(PRIORITY_SOURCE) && !best.source.includes(PRIORITY_SOURCE)) {
              ownerEmails.set(email, current);
            }
          }
        }
      });
    });
    const result = new Map<string, RecordInfo[]>();
    masterMap.forEach((m, o) => result.set(o, Array.from(m.values())));
    return result;
  }, [data, emailCol, ownerCol, dateCol, sourceCol, startDate, endDate, activeTab]);

  // 导出处理
  const handleExportEmailCentric = () => {
    const ws = (window as any).XLSX.utils.json_to_sheet(emailCentricData);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "邮箱统计");
    (window as any).XLSX.writeFile(wb, `邮箱维度统计_${new Date().getTime()}.xlsx`);
  };

  const exportSingleOwner = (owner: string, records: RecordInfo[]) => {
    const out = records.map(r => ({ "负责人": r.owner, "邮箱": r.email, "日期": r.date.getTime() === 8640000000000000 ? "未知日期" : r.date.toISOString().split('T')[0], "订单来源": r.source }));
    const ws = (window as any).XLSX.utils.json_to_sheet(out);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "数据");
    (window as any).XLSX.writeFile(wb, `${owner}_统计结果.xlsx`);
  };

  const handleExportAllZip = async () => {
    if (ownerCentricData.size === 0) return;
    setLoading(true);
    const zip = new (window as any).JSZip();
    ownerCentricData.forEach((recs, owner) => {
      const out = recs.map(r => ({ "负责人": r.owner, "邮箱": r.email, "日期": r.date.getTime() === 8640000000000000 ? "未知日期" : r.date.toISOString().split('T')[0], "订单来源": r.source }));
      const ws = (window as any).XLSX.utils.json_to_sheet(out);
      const wb = (window as any).XLSX.utils.book_new();
      (window as any).XLSX.utils.book_append_sheet(wb, ws, "数据");
      const wbout = (window as any).XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      zip.file(`${owner}_数据统计.xlsx`, wbout);
    });
    const blob = await zip.generateAsync({ type: "blob" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `按负责人分包统计_${new Date().getTime()}.zip`;
    link.click();
    setLoading(false);
  };

  return (
    <div className="min-h-screen pb-16">
      <header className="gradient-bg text-white py-10 px-6 shadow-lg mb-8">
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold flex items-center">
              <i className="fas fa-file-invoice mr-3"></i> Excel 高级统计工具
            </h1>
            <p className="opacity-80 text-sm mt-1">支持邮箱维度聚合与负责人维度去重统计</p>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 space-y-6">
        {/* Step 1: Upload */}
        <section className="glass-card p-6 rounded-2xl shadow-sm">
          <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
            <span className="bg-indigo-600 text-white w-6 h-6 rounded-full flex items-center justify-center mr-2 text-[10px]">1</span>
            上传目标文件
          </h2>
          <div className="border-2 border-dashed border-gray-200 rounded-xl p-8 hover:border-indigo-400 transition-all relative bg-gray-50/50 group text-center">
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
            <i className={`fas ${fileName ? 'fa-check-circle text-emerald-500' : 'fa-cloud-upload-alt text-indigo-300'} text-4xl mb-3 group-hover:scale-110 transition-transform`}></i>
            <p className="text-gray-600 font-medium">{fileName || "点击此处或拖拽文件上传"}</p>
          </div>
        </section>

        {data.length > 0 && (
          <>
            {/* Step 2: Mapping */}
            <section className="glass-card p-6 rounded-2xl shadow-sm">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                  <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                    <span className="bg-indigo-600 text-white w-6 h-6 rounded-full flex items-center justify-center mr-2 text-[10px]">2</span>
                    字段映射
                  </h2>
                  <div className="space-y-3">
                    {[
                      { label: "邮箱字段", state: emailCol, setter: setEmailCol },
                      { label: "负责人字段", state: ownerCol, setter: setOwnerCol },
                      { label: "日期字段", state: dateCol, setter: setDateCol },
                      { label: "来源/端口字段", state: sourceCol, setter: setSourceCol },
                    ].map((item, idx) => (
                      <div key={idx}>
                        <label className="text-xs font-bold text-gray-400 mb-1 block uppercase">{item.label}</label>
                        <select 
                          value={item.state} 
                          onChange={e => item.setter(e.target.value)} 
                          className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm text-gray-400 focus:ring-2 focus:ring-indigo-500 outline-none bg-white"
                        >
                          <option value="" className="text-gray-400">未选择</option>
                          {columns.map(c => (
                            <option key={c} value={c} className="text-gray-400">
                              {c}
                            </option>
                          ))}
                        </select>
                      </div>
                    ))}
                  </div>
                </div>
                <div>
                  <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                    <span className="bg-indigo-600 text-white w-6 h-6 rounded-full flex items-center justify-center mr-2 text-[10px]">3</span>
                    筛选条件
                  </h2>
                  <div className="space-y-4">
                    <div>
                      <label className="text-xs font-bold text-gray-400 mb-1 block uppercase">开始日期 (可选)</label>
                      <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none" />
                    </div>
                    <div>
                      <label className="text-xs font-bold text-gray-400 mb-1 block uppercase">结束日期 (可选)</label>
                      <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none" />
                    </div>
                    <div className="p-3 bg-indigo-50 rounded-lg border border-indigo-100 text-[10px] text-indigo-700 leading-relaxed">
                      <p className="font-bold mb-1 italic">去重规则说明：</p>
                      1. 日期越早越优先。<br/>
                      2. 同日期，来源包含“{PRIORITY_SOURCE}”优先。
                    </div>
                  </div>
                </div>
              </div>
            </section>

            {/* Step 3: Analysis Tabs */}
            <section className="glass-card overflow-hidden rounded-2xl shadow-sm border-0">
              <div className="flex border-b border-gray-100">
                <button 
                  onClick={() => setActiveTab('email')}
                  className={`flex-1 py-4 text-sm font-bold transition-all ${activeTab === 'email' ? 'bg-white text-indigo-600 border-b-2 border-indigo-600' : 'bg-gray-50/50 text-gray-400 hover:text-gray-600'}`}
                >
                  <i className="fas fa-at mr-2"></i> 邮箱维度 (汇总)
                </button>
                <button 
                  onClick={() => setActiveTab('owner')}
                  className={`flex-1 py-4 text-sm font-bold transition-all ${activeTab === 'owner' ? 'bg-white text-indigo-600 border-b-2 border-indigo-600' : 'bg-gray-50/50 text-gray-400 hover:text-gray-600'}`}
                >
                  <i className="fas fa-user-friends mr-2"></i> 负责人维度 (分表)
                </button>
              </div>

              <div className="p-6">
                {activeTab === 'email' ? (
                  <div className="space-y-6">
                    <div className="flex justify-between items-center">
                      <h3 className="text-gray-800 font-bold">邮箱聚合结果 ({emailCentricData.length} 条)</h3>
                      <button onClick={handleExportEmailCentric} className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded-xl text-xs font-bold shadow-md transition-all">
                        <i className="fas fa-file-export mr-2"></i> 导出汇总表
                      </button>
                    </div>
                    <div className="overflow-x-auto border border-gray-100 rounded-xl">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-gray-50 text-gray-400">
                          <tr>
                            <th className="px-4 py-3 font-bold uppercase tracking-wider">邮箱</th>
                            <th className="px-4 py-3 font-bold uppercase tracking-wider">负责人</th>
                            <th className="px-4 py-3 font-bold uppercase tracking-wider text-center">日期</th>
                            <th className="px-4 py-3 font-bold uppercase tracking-wider">来源</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {emailCentricData.slice(0, 10).map((r, i) => (
                            <tr key={i} className="hover:bg-indigo-50/20">
                              <td className="px-4 py-3 font-medium text-gray-700">{r["邮箱"]}</td>
                              <td className="px-4 py-3"><span className="bg-indigo-100 text-indigo-700 px-1.5 py-0.5 rounded font-bold text-[10px]">{r["负责人"]}</span></td>
                              <td className="px-4 py-3 text-center text-gray-400 font-mono">{r["日期"]}</td>
                              <td className="px-4 py-3 text-gray-400 italic">{r["订单来源"]}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {emailCentricData.length > 10 && <div className="text-center py-3 bg-gray-50/50 text-[10px] text-gray-400 italic">预览仅显示前10条...</div>}
                    </div>
                  </div>
                ) : (
                  <div className="space-y-6">
                    <div className="flex justify-between items-center">
                      <h3 className="text-gray-800 font-bold">按负责人分组 ({ownerCentricData.size} 个负责人)</h3>
                      <button onClick={handleExportAllZip} className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-xl text-xs font-bold shadow-md transition-all">
                        <i className="fas fa-file-archive mr-2"></i> 打包导出 (ZIP)
                      </button>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                      {Array.from(ownerCentricData.keys()).map(owner => {
                        const recs = ownerCentricData.get(owner)!;
                        return (
                          <div key={owner} className="p-4 bg-white border border-gray-100 rounded-2xl hover:border-indigo-300 transition-all shadow-sm group">
                            <div className="flex justify-between items-start mb-4">
                              <div>
                                <p className="text-[10px] text-gray-400 font-bold uppercase">负责人</p>
                                <h4 className="text-lg font-black text-gray-800">{owner}</h4>
                              </div>
                              <span className="bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded-full text-[10px] font-bold">{recs.length} 记录</span>
                            </div>
                            <button onClick={() => exportSingleOwner(owner, recs)} className="w-full bg-gray-50 group-hover:bg-indigo-600 group-hover:text-white text-gray-500 py-2 rounded-xl text-xs font-bold transition-all">
                              单独导出 Excel
                            </button>
                          </div>
                        );
                      })}
                    </div>
                    {ownerCentricData.size === 0 && <div className="text-center py-10 text-gray-300 italic">未发现匹配负责人</div>}
                  </div>
                )}
              </div>
            </section>
          </>
        )}
      </main>

      {loading && (
        <div className="fixed inset-0 bg-indigo-950/40 backdrop-blur-md flex items-center justify-center z-50">
          <div className="bg-white p-10 rounded-3xl shadow-2xl text-center">
            <div className="animate-spin rounded-full h-12 w-12 border-4 border-indigo-600 border-t-transparent mx-auto mb-4"></div>
            <p className="font-bold text-gray-800">正在生成统计报表...</p>
          </div>
        </div>
      )}
    </div>
  );
}

const root = createRoot(document.getElementById('root')!);
root.render(<App />);