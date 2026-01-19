
import React, { useState, useMemo } from 'react';
import { createRoot } from 'react-dom/client';

/**
 * Excel Processing Tool v6.5
 * 专项优化：
 * 1. 新增“红人姓名”字段映射与全链路展示。
 * 2. 延续 v6.4 的时区修复方案，确保 2025-12-03 等日期在任何时区下均能正确显示。
 */

const PRIORITY_SOURCE = "社媒端";

// 助手函数：安全地获取本地日期字符串 (YYYY-MM-DD)，避免时区偏移导致的日期偏差
const formatLocalDate = (date: Date) => {
  if (!date || isNaN(date.getTime())) return "N/A";
  if (date.getTime() === 8640000000000000) return "N/A";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

interface RecordInfo {
  email: string;
  owner: string;
  date: Date;
  source: string;
  partTimeName: string;
  influencerName: string; // 新增：红人姓名
  earliestOwner?: string; // 最早负责人
  earliestDate?: string;  // 最早日期
}

function App() {
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [activeTab, setActiveTab] = useState<'email' | 'owner' | 'earliest'>('email');

  // 配置状态
  const [emailCol, setEmailCol] = useState("");
  const [ownerCol, setOwnerCol] = useState("");
  const [dateCol, setDateCol] = useState("");
  const [sourceCol, setSourceCol] = useState("");
  const [partTimeCol, setPartTimeCol] = useState("");
  const [influencerCol, setInfluencerCol] = useState(""); // 新增：红人姓名列
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (evt: ProgressEvent<FileReader>) => {
      try {
        const bstr = evt.target?.result;
        if (typeof bstr !== 'string') {
          throw new Error("Invalid file content format.");
        }
        const wb = (window as any).XLSX.read(bstr, { type: 'binary', cellDates: true });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];

        const rawRows = (window as any).XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (!rawRows || rawRows.length === 0) {
          throw new Error("表格似乎是空的。");
        }

        const headerRow = rawRows[0] as any[];
        const cols = headerRow
          .map(c => String(c || "").trim())
          .filter(c => c !== ""); 
        
        setColumns(cols);

        const jsonData = (window as any).XLSX.utils.sheet_to_json(ws);
        setData(jsonData as any[]);

        if (cols.length > 0) {
          setEmailCol(cols.find(c => c.includes('邮箱') || c.toLowerCase().includes('email')) || "");
          setOwnerCol(cols.find(c => c.includes('负责人') || c.toLowerCase().includes('owner')) || "");
          setDateCol(cols.find(c => c.includes('时间') || c.includes('日期') || c.toLowerCase().includes('date')) || "");
          setSourceCol(cols.find(c => c.includes('来源') || c.includes('端口') || c.toLowerCase().includes('source')) || "");
          setPartTimeCol(cols.find(c => c.includes('兼职')) || "");
          setInfluencerCol(cols.find(c => c.includes('红人') || c.includes('姓名') || c.toLowerCase().includes('influencer')) || "");
        }
      } catch (err: any) {
        const errorMessage = err instanceof Error ? err.message : String(err);
        console.error("File processing error:", errorMessage);
        alert(`文件读取失败: ${errorMessage}`);
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  // --- 核心工具：计算全局最早归属字典 (Map: Email -> { owner, date, influencerName }) ---
  const globalEarliestInfoMap = useMemo(() => {
    if (!data.length || !emailCol || !ownerCol) return new Map<string, { owner: string, date: string, influencerName: string }>();
    const emailToEarliestMap = new Map<string, { owner: string, date: Date, source: string, influencerName: string }>();

    data.forEach(row => {
      const email = String(row[emailCol] || "").trim();
      if (!email) return;

      const rawDate = row[dateCol];
      let rowDate = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(rawDate) : null);
      const isInvalidDate = !rowDate || isNaN(rowDate.getTime());
      const effectiveDate = isInvalidDate ? new Date(8640000000000000) : rowDate!;
      const source = String(row[sourceCol] || "").trim();
      const owner = String(row[ownerCol] || "").trim();
      const influencerName = String(row[influencerCol] || "").trim();

      if (!emailToEarliestMap.has(email)) {
        emailToEarliestMap.set(email, { owner, date: effectiveDate, source, influencerName });
      } else {
        const best = emailToEarliestMap.get(email)!;
        const bestDStr = formatLocalDate(best.date);
        const currDStr = formatLocalDate(effectiveDate);
        
        if (effectiveDate < best.date && currDStr !== bestDStr) {
          emailToEarliestMap.set(email, { owner, date: effectiveDate, source, influencerName });
        } else if (currDStr === bestDStr) {
          if (source.includes(PRIORITY_SOURCE) && !best.source.includes(PRIORITY_SOURCE)) {
            emailToEarliestMap.set(email, { owner, date: effectiveDate, source, influencerName });
          }
        }
      }
    });

    const result = new Map<string, { owner: string, date: string, influencerName: string }>();
    emailToEarliestMap.forEach((val, key) => {
      result.set(key, { 
        owner: val.owner, 
        date: val.date.getTime() === 8640000000000000 ? "未知日期" : formatLocalDate(val.date),
        influencerName: val.influencerName
      });
    });
    return result;
  }, [data, emailCol, ownerCol, dateCol, sourceCol, influencerCol]);

  // --- 逻辑1: 邮箱维度 (汇总) ---
  const emailCentricData = useMemo(() => {
    if (!data.length || !emailCol || !ownerCol) return [];
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
      const ptName = String(row[partTimeCol] || "").trim();
      const infName = String(row[influencerCol] || "").trim();
      const ownersList = String(row[ownerCol] || "").split(/[,，;；\s\/\\]+/).map(o => o.trim()).filter(o => o.length > 0);
      
      if (!emailMap.has(email)) {
        emailMap.set(email, { date: effectiveDate, source, partTimeName: ptName, influencerName: infName, owners: new Set(ownersList) });
      } else {
        const best = emailMap.get(email)!;
        const bestDateStr = formatLocalDate(best.date);
        const rowDateStr = isInvalidDate ? "N/A" : formatLocalDate(rowDate!);
        
        if (effectiveDate < best.date && rowDateStr !== bestDateStr) {
          emailMap.set(email, { date: effectiveDate, source, partTimeName: ptName, influencerName: infName, owners: new Set(ownersList) });
        } else if (rowDateStr === bestDateStr) {
          if (source.includes(PRIORITY_SOURCE) && !best.source.includes(PRIORITY_SOURCE)) {
            emailMap.set(email, { date: effectiveDate, source, partTimeName: ptName, influencerName: infName, owners: new Set(ownersList) });
          } else if (source.includes(PRIORITY_SOURCE) === best.source.includes(PRIORITY_SOURCE)) {
            ownersList.forEach(o => best.owners.add(o));
          }
        }
      }
    });
    return Array.from(emailMap.entries()).map(([email, info]) => {
      const earliestInfo = globalEarliestInfoMap.get(email);
      return {
        "邮箱": email,
        "负责人": Array.from(info.owners).join("、"),
        "红人姓名": info.influencerName,
        "最早负责人": earliestInfo?.owner || "未识别",
        "最早日期": earliestInfo?.date || "未识别",
        "日期": info.date.getTime() === 8640000000000000 ? "未知日期" : formatLocalDate(info.date),
        "订单来源": info.source,
        "兼职名字": info.partTimeName
      };
    });
  }, [data, emailCol, ownerCol, dateCol, sourceCol, partTimeCol, influencerCol, startDate, endDate, globalEarliestInfoMap]);

  // --- 逻辑2: 负责人维度 (分表) ---
  const ownerCentricData = useMemo(() => {
    if (!data.length || !emailCol || !ownerCol) return new Map<string, RecordInfo[]>();
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
      const ptName = String(row[partTimeCol] || "").trim();
      const infName = String(row[influencerCol] || "").trim();
      const effectiveDate = isInvalidDate ? new Date(8640000000000000) : rowDate!;
      const ownersList = rawOwners.split(/[,，;；\s\/\\]+/).map(o => o.trim()).filter(o => o.length > 0);

      ownersList.forEach(owner => {
        if (!masterMap.has(owner)) masterMap.set(owner, new Map());
        const ownerEmails = masterMap.get(owner)!;
        const earliestInfo = globalEarliestInfoMap.get(email);
        const current: RecordInfo = { 
          email, 
          owner, 
          date: effectiveDate, 
          source, 
          partTimeName: ptName,
          influencerName: infName,
          earliestOwner: earliestInfo?.owner || "未识别",
          earliestDate: earliestInfo?.date || "未识别"
        };
        if (!ownerEmails.has(email)) {
          ownerEmails.set(email, current);
        } else {
          const best = ownerEmails.get(email)!;
          const bestDStr = formatLocalDate(best.date);
          const currDStr = formatLocalDate(effectiveDate);
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
  }, [data, emailCol, ownerCol, dateCol, sourceCol, partTimeCol, influencerCol, startDate, endDate, globalEarliestInfoMap]);

  // --- 逻辑3: 全局最早归属 (展示用) ---
  const earliestCentricData = useMemo(() => {
    if (!data.length || !emailCol || !ownerCol) return [];
    const globalEmailMap = new Map<string, RecordInfo>();
    
    data.forEach(row => {
      const email = String(row[emailCol] || "").trim();
      if (!email) return;
      
      const rawDate = row[dateCol];
      let rowDate = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(rawDate) : null);
      const isInvalidDate = !rowDate || isNaN(rowDate.getTime());
      const effectiveDate = isInvalidDate ? new Date(8640000000000000) : rowDate!;
      
      const source = String(row[sourceCol] || "").trim();
      const ptName = String(row[partTimeCol] || "").trim();
      const infName = String(row[influencerCol] || "").trim();
      const owner = String(row[ownerCol] || "").trim();
      
      const current: RecordInfo = { email, owner, date: effectiveDate, source, partTimeName: ptName, influencerName: infName };

      if (!globalEmailMap.has(email)) {
        globalEmailMap.set(email, current);
      } else {
        const best = globalEmailMap.get(email)!;
        const bestDStr = formatLocalDate(best.date);
        const currDStr = formatLocalDate(effectiveDate);
        
        if (effectiveDate < best.date && currDStr !== bestDStr) {
          globalEmailMap.set(email, current);
        } else if (currDStr === bestDStr) {
          if (source.includes(PRIORITY_SOURCE) && !best.source.includes(PRIORITY_SOURCE)) {
            globalEmailMap.set(email, current);
          }
        }
      }
    });

    return Array.from(globalEmailMap.values()).map(r => ({
      "邮箱": r.email,
      "负责人": r.owner,
      "红人姓名": r.influencerName,
      "日期": r.date.getTime() === 8640000000000000 ? "未知日期" : formatLocalDate(r.date),
      "来源": r.source,
      "兼职名字": r.partTimeName
    }));
  }, [data, emailCol, ownerCol, dateCol, sourceCol, partTimeCol, influencerCol]);

  // 导出处理
  const handleExportEmailCentric = () => {
    const ws = (window as any).XLSX.utils.json_to_sheet(emailCentricData);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "汇总统计");
    (window as any).XLSX.writeFile(wb, `邮箱汇总表_${new Date().getTime()}.xlsx`);
  };

  const handleExportEarliest = () => {
    const ws = (window as any).XLSX.utils.json_to_sheet(earliestCentricData);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "全局最早归属");
    (window as any).XLSX.writeFile(wb, `全局最早归属表_${new Date().getTime()}.xlsx`);
  };

  const exportSingleOwner = (owner: string, records: RecordInfo[]) => {
    const out = records.map(r => ({
      "负责人": r.owner,
      "邮箱": r.email,
      "红人姓名": r.influencerName,
      "最早负责人": r.earliestOwner,
      "最早日期": r.earliestDate,
      "日期": r.date.getTime() === 8640000000000000 ? "未知日期" : formatLocalDate(r.date),
      "订单来源": r.source,
      "兼职名字": r.partTimeName
    }));
    const ws = (window as any).XLSX.utils.json_to_sheet(out);
    const wb = (window as any).XLSX.utils.book_new();
    (window as any).XLSX.utils.book_append_sheet(wb, ws, "数据");
    (window as any).XLSX.writeFile(wb, `${owner}_统计结果.xlsx`);
  };

  const handleExportAllZip = async () => {
    if (ownerCentricData.size === 0) return;
    setLoading(true);
    try {
      const zip = new (window as any).JSZip();
      ownerCentricData.forEach((recs, owner) => {
        const out = recs.map(r => ({
          "负责人": r.owner,
          "邮箱": r.email,
          "红人姓名": r.influencerName,
          "最早负责人": r.earliestOwner,
          "最早日期": r.earliestDate,
          "日期": r.date.getTime() === 8640000000000000 ? "未知日期" : formatLocalDate(r.date),
          "订单来源": r.source,
          "兼职名字": r.partTimeName
        }));
        const ws = (window as any).XLSX.utils.json_to_sheet(out);
        const wb = (window as any).XLSX.utils.book_new();
        (window as any).XLSX.utils.book_append_sheet(wb, ws, "数据");
        const wbout = (window as any).XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        zip.file(`${owner}_数据统计.xlsx`, wbout);
      });
      // Explicitly type blob from generateAsync to handle potentially ambiguous Promise result
      const blob = await zip.generateAsync({ type: "blob" });
      const link = document.createElement("a");
      // Use explicit cast to any or Blob to satisfy strict URL.createObjectURL type requirements if needed
      link.href = URL.createObjectURL(blob as any);
      link.download = `负责人分表打包_${new Date().getTime()}.zip`;
      link.click();
    } catch (err: any) {
      const errorMessage = err instanceof Error ? err.message : String(err);
      alert(`打包导出失败: ${errorMessage}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen pb-16">
      <header className="gradient-bg text-white py-10 px-6 shadow-lg mb-8">
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold flex items-center">
              <i className="fas fa-file-invoice mr-3"></i> Excel 高级统计工具 v6.5
            </h1>
            <p className="opacity-80 text-sm mt-1">
              <span className="bg-white/20 px-2 py-0.5 rounded-md mr-2">核心优化</span>
              已修复日期显示偏差，支持红人姓名、最早负责人/日期索引分析
            </p>
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
                  <div className="grid grid-cols-2 gap-3">
                    {[
                      { label: "邮箱字段", state: emailCol, setter: setEmailCol },
                      { label: "负责人字段", state: ownerCol, setter: setOwnerCol },
                      { label: "日期字段", state: dateCol, setter: setDateCol },
                      { label: "来源字段", state: sourceCol, setter: setSourceCol },
                      { label: "兼职名字", state: partTimeCol, setter: setPartTimeCol },
                      { label: "红人姓名", state: influencerCol, setter: setInfluencerCol }, // 新增映射
                    ].map((item, idx) => (
                      <div key={idx} className={(item.label === "兼职名字" || item.label === "红人姓名") ? "col-span-1" : ""}>
                        <label className="text-[10px] font-bold text-gray-400 mb-1 block uppercase">{item.label}</label>
                        <select 
                          value={item.state} 
                          onChange={e => item.setter(e.target.value)} 
                          className="w-full border border-gray-200 rounded-lg px-3 py-1.5 text-xs text-gray-400 focus:ring-2 focus:ring-indigo-500 outline-none bg-white"
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
                    筛选条件 (仅影响前两个标签)
                  </h2>
                  <div className="space-y-3">
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="text-[10px] font-bold text-gray-400 mb-1 block uppercase">开始日期</label>
                        <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="w-full border border-gray-200 rounded-lg px-3 py-1.5 text-xs focus:ring-2 focus:ring-indigo-500 outline-none" />
                      </div>
                      <div>
                        <label className="text-[10px] font-bold text-gray-400 mb-1 block uppercase">结束日期</label>
                        <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="w-full border border-gray-200 rounded-lg px-3 py-1.5 text-xs focus:ring-2 focus:ring-indigo-500 outline-none" />
                      </div>
                    </div>
                    <div className="p-3 bg-indigo-50 rounded-lg border border-indigo-100 text-[10px] text-indigo-700 leading-relaxed">
                      <p className="font-bold mb-1 italic">逻辑提示：</p>
                      • 重复邮箱取<strong>日期最早</strong>的记录。<br/>
                      • 同日期则<strong>“{PRIORITY_SOURCE}”</strong>优先。<br/>
                      • “最早负责人/日期”通过匹配<strong>全表原始数据</strong>得出。
                    </div>
                  </div>
                </div>
              </div>
            </section>

            {/* Step 3: Analysis Tabs */}
            <section className="glass-card overflow-hidden rounded-2xl shadow-sm border-0">
              <div className="flex border-b border-gray-100 bg-gray-50/30">
                <button 
                  onClick={() => setActiveTab('email')}
                  className={`flex-1 py-4 text-xs font-bold transition-all ${activeTab === 'email' ? 'bg-white text-indigo-600 border-b-2 border-indigo-600' : 'text-gray-400 hover:text-gray-600'}`}
                >
                  <i className="fas fa-at mr-2"></i> 邮箱维度 (汇总)
                </button>
                <button 
                  onClick={() => setActiveTab('owner')}
                  className={`flex-1 py-4 text-xs font-bold transition-all ${activeTab === 'owner' ? 'bg-white text-indigo-600 border-b-2 border-indigo-600' : 'text-gray-400 hover:text-gray-600'}`}
                >
                  <i className="fas fa-user-friends mr-2"></i> 负责人维度 (分表)
                </button>
                <button 
                  onClick={() => setActiveTab('earliest')}
                  className={`flex-1 py-4 text-xs font-bold transition-all ${activeTab === 'earliest' ? 'bg-white text-indigo-600 border-b-2 border-indigo-600' : 'text-gray-400 hover:text-gray-600'}`}
                >
                  <i className="fas fa-history mr-2"></i> 全局最早归属
                </button>
              </div>

              <div className="p-6">
                {activeTab === 'email' && (
                  <div className="space-y-6">
                    <div className="flex justify-between items-center">
                      <h3 className="text-gray-800 font-bold text-sm">邮箱聚合结果 ({emailCentricData.length} 条)</h3>
                      <button onClick={handleExportEmailCentric} className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded-xl text-xs font-bold shadow-md transition-all">
                        <i className="fas fa-file-export mr-2"></i> 导出汇总表
                      </button>
                    </div>
                    <div className="overflow-x-auto border border-gray-100 rounded-xl">
                      <table className="w-full text-left text-[11px]">
                        <thead className="bg-gray-50 text-gray-400">
                          <tr>
                            <th className="px-4 py-3 font-bold uppercase">邮箱</th>
                            <th className="px-4 py-3 font-bold uppercase">红人姓名</th>
                            <th className="px-4 py-3 font-bold uppercase">当前负责人</th>
                            <th className="px-4 py-3 font-bold uppercase text-indigo-600">最早负责人</th>
                            <th className="px-4 py-3 font-bold uppercase text-indigo-600">最早日期</th>
                            <th className="px-4 py-3 font-bold uppercase">兼职名字</th>
                            <th className="px-4 py-3 font-bold uppercase">当前日期</th>
                            <th className="px-4 py-3 font-bold uppercase">来源</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {emailCentricData.slice(0, 10).map((r: any, i) => (
                            <tr key={i} className="hover:bg-indigo-50/20">
                              <td className="px-4 py-3 font-medium text-gray-700">{r["邮箱"]}</td>
                              <td className="px-4 py-3 font-bold text-emerald-600">{r["红人姓名"] || "-"}</td>
                              <td className="px-4 py-3"><span className="bg-indigo-50 text-indigo-600 px-1.5 py-0.5 rounded font-bold text-[10px]">{r["负责人"]}</span></td>
                              <td className="px-4 py-3 font-black text-gray-800">{r["最早负责人"]}</td>
                              <td className="px-4 py-3 font-black text-gray-800">{r["最早日期"]}</td>
                              <td className="px-4 py-3 text-gray-500">{r["兼职名字"] || "-"}</td>
                              <td className="px-4 py-3 text-gray-400 font-mono">{r["日期"]}</td>
                              <td className="px-4 py-3 text-gray-400 italic">{r["订单来源"]}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {activeTab === 'owner' && (
                  <div className="space-y-6">
                    <div className="flex justify-between items-center">
                      <h3 className="text-gray-800 font-bold text-sm">按负责人分组 ({ownerCentricData.size} 个负责人)</h3>
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
                                <h4 className="text-base font-black text-gray-800">{owner}</h4>
                              </div>
                              <span className="bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded-full text-[10px] font-bold">{recs.length} 记录</span>
                            </div>
                            <div className="mb-3">
                                <p className="text-[9px] text-gray-400 mb-1 uppercase">导出内容预览</p>
                                <div className="text-[10px] text-gray-500 space-y-0.5">
                                    <p>• 包含“红人姓名”字段</p>
                                    <p>• 包含“最早负责人”字段</p>
                                    <p>• 包含“最早日期”字段</p>
                                    <p>• 包含“兼职名字”字段</p>
                                </div>
                            </div>
                            <button onClick={() => exportSingleOwner(owner, recs)} className="w-full bg-gray-50 group-hover:bg-indigo-600 group-hover:text-white text-gray-500 py-2 rounded-xl text-[11px] font-bold transition-all">
                              单独导出 Excel
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {activeTab === 'earliest' && (
                  <div className="space-y-6">
                    <div className="flex justify-between items-center">
                      <h3 className="text-gray-800 font-bold text-sm">全局最早归属名单 (全量匹配库)</h3>
                      <button onClick={handleExportEarliest} className="bg-purple-600 hover:bg-purple-700 text-white px-5 py-2 rounded-xl text-xs font-bold shadow-md transition-all">
                        <i className="fas fa-history mr-2"></i> 导出全局最早表
                      </button>
                    </div>
                    <div className="overflow-x-auto border border-gray-100 rounded-xl">
                      <table className="w-full text-left text-[11px]">
                        <thead className="bg-purple-50 text-purple-400">
                          <tr>
                            <th className="px-4 py-3 font-bold uppercase">邮箱</th>
                            <th className="px-4 py-3 font-bold uppercase">红人姓名</th>
                            <th className="px-4 py-3 font-bold uppercase">最早负责人</th>
                            <th className="px-4 py-3 font-bold uppercase text-indigo-600">最早日期</th>
                            <th className="px-4 py-3 font-bold uppercase">来源</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {/* Add explicit any type to 'r' to avoid unknown type errors on line 573 */}
                          {earliestCentricData.slice(0, 10).map((r: any, i) => (
                            <tr key={i} className="hover:bg-purple-50/20">
                              <td className="px-4 py-3 font-medium text-gray-700">{r["邮箱"]}</td>
                              <td className="px-4 py-3 font-bold text-emerald-600">{r["红人姓名"] || "-"}</td>
                              <td className="px-4 py-3 font-bold text-purple-700">{r["负责人"]}</td>
                              <td className="px-4 py-3 text-gray-400 font-mono">{r["日期"]}</td>
                              <td className="px-4 py-3 text-gray-400 italic">{r["来源"]}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
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
            <p className="font-bold text-gray-800">正在进行全局归属透视分析与时区校准...</p>
          </div>
        </div>
      )}
    </div>
  );
}

const rootElement = document.getElementById('root');
if (!rootElement) throw new Error('Root element not found');
const root = createRoot(rootElement);
root.render(<App />);
