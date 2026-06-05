import React, { useState, useEffect, useCallback } from "react";
import {
  LOGO_B64,
  KHO_LABEL as FALLBACK_KHO_LABEL,
  META as FALLBACK_META,
  VUNG_INFO as FALLBACK_VUNG_INFO,
  RAW as FALLBACK_RAW,
  DAILY as FALLBACK_DAILY,
} from "./data.js";
import { EXCEL_POLL_INTERVAL_MS, loadExcelDashboardData, getExcelFileName } from "./excelData.js";

// ── helpers ────────────────────────────────────────────────────────────────
const sumML      = a => a.reduce((s,r)=>s+r[5],0);
const sumSPK     = a => a.reduce((s,r)=>s+r[6],0);
const avgML      = a => a.length ? Math.round(sumML(a)/a.length) : 0;
const colorByLoai= l => l==="Nhân Viên"?"#00d4ff":l==="Cộng Tác Viên"?"#ff6b35":"#a78bfa";
const getLevel   = ml => {
  if(ml===0) return {label:"ML=0",cls:"b-bad"};
  if(ml<30)  return {label:"Thấp",cls:"b-bad"};
  if(ml<100) return {label:"TB",  cls:"b-warn"};
  if(ml<300) return {label:"Khá", cls:"b-ok"};
  return          {label:"Cao", cls:"b-ok"};
};
const pad2 = value => String(value).padStart(2, "0");
const formatLoadedAt = value => value
  ? new Intl.DateTimeFormat("vi-VN", { hour: "2-digit", minute: "2-digit", second: "2-digit" }).format(new Date(value))
  : "";

const FALLBACK_DASHBOARD_DATA = {
  raw: FALLBACK_RAW,
  daily: FALLBACK_DAILY,
  khoLabel: FALLBACK_KHO_LABEL,
  vungInfo: FALLBACK_VUNG_INFO,
  meta: FALLBACK_META,
  source: "data.js",
  loadedAt: null,
};

// ── CSS-in-JS (paste toàn bộ CSS từ HTML vào đây) ─────────────────────────
const STYLES = `
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
:root{--bg:#0a0e1a;--surface:#111827;--surface2:#1a2234;--border:#1e2d45;--accent:#00d4ff;--accent2:#ff6b35;--accent3:#22c55e;--accent4:#f59e0b;--danger:#ef4444;--text:#e2e8f0;--muted:#64748b;}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Be Vietnam Pro',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden;}
body::before{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(0,212,255,.03)1px,transparent 1px),linear-gradient(90deg,rgba(0,212,255,.03)1px,transparent 1px);background-size:40px 40px;pointer-events:none;z-index:0;}
.wrap{position:relative;z-index:1;max-width:1440px;margin:0 auto;padding:22px;}
header{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;padding-bottom:18px;border-bottom:1px solid var(--border);}
.logo-area h1{font-size:1.4rem;font-weight:800;letter-spacing:-.5px;background:linear-gradient(135deg,var(--accent),#7c3aed);-webkit-background-clip:text;-webkit-text-fill-color:transparent;display:flex;align-items:center;gap:10px;}
.logo-area p{font-size:.75rem;color:var(--muted);margin-top:3px;}
.badge-live{display:flex;align-items:center;gap:8px;background:rgba(0,212,255,.08);border:1px solid rgba(0,212,255,.2);padding:6px 14px;border-radius:20px;font-size:.75rem;}
.dot-live{width:8px;height:8px;border-radius:50%;background:var(--accent3);animation:pulse 1.5s infinite;}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
.vung-bar{display:flex;gap:8px;margin-bottom:20px;flex-wrap:wrap;align-items:center;}
.vung-label{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;margin-right:4px;white-space:nowrap;}
.vung-btn{padding:7px 16px;border-radius:8px;font-size:.75rem;cursor:pointer;font-family:'Be Vietnam Pro';border:1px solid var(--border);background:var(--surface2);color:var(--muted);transition:all .18s;white-space:nowrap;}
.vung-btn:hover{border-color:rgba(0,212,255,.4);color:var(--text);}
.vung-btn.active{background:rgba(0,212,255,.12);border-color:rgba(0,212,255,.45);color:var(--accent);font-weight:700;}
.vung-btn.v-dtb.active{background:rgba(167,139,250,.12);border-color:rgba(167,139,250,.45);color:#a78bfa;}
.vung-info{margin-left:auto;font-size:.7rem;color:var(--muted);background:var(--surface2);padding:5px 12px;border-radius:6px;border:1px solid var(--border);}
.kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px;}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px;position:relative;overflow:hidden;transition:transform .2s;}
.kpi:hover{transform:translateY(-2px);}
.kpi::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;}
.kpi.c1::after{background:linear-gradient(90deg,var(--accent),transparent);}
.kpi.c2::after{background:linear-gradient(90deg,var(--accent2),transparent);}
.kpi.c3::after{background:linear-gradient(90deg,var(--accent3),transparent);}
.kpi.c4::after{background:linear-gradient(90deg,var(--accent4),transparent);}
.kpi.c5::after{background:linear-gradient(90deg,#a78bfa,transparent);}
.kpi-label{font-size:.68rem;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;}
.kpi-val{font-size:1.9rem;font-weight:800;font-family:'JetBrains Mono',monospace;}
.kpi.c1 .kpi-val{color:var(--accent);}
.kpi.c2 .kpi-val{color:var(--accent2);}
.kpi.c3 .kpi-val{color:var(--accent3);}
.kpi.c4 .kpi-val{color:var(--accent4);}
.kpi.c5 .kpi-val{color:#a78bfa;}
.kpi-sub{font-size:.68rem;color:var(--muted);margin-top:4px;}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:18px;}
.row3{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:18px;}
.panel{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px;overflow:hidden;}
.panel-title{font-size:.76rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:14px;display:flex;align-items:center;gap:8px;flex-shrink:0;}
.panel-title span{color:var(--accent);}
.bar-list{display:flex;flex-direction:column;gap:7px;}
.bar-item{display:flex;align-items:center;gap:10px;}
.bar-name{font-size:.7rem;color:var(--text);min-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.bar-track{flex:1;height:7px;background:var(--surface2);border-radius:4px;overflow:hidden;}
.bar-fill{height:100%;border-radius:4px;transition:width .8s cubic-bezier(.4,0,.2,1);}
.bar-val{font-size:.68rem;font-family:'JetBrains Mono';color:var(--muted);min-width:34px;text-align:right;}
.tbl-wrap{overflow-x:auto;height:528px;overflow-y:auto;}
.tbl-wrap::-webkit-scrollbar{width:4px;height:4px;}
.tbl-wrap::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px;}
table{width:100%;border-collapse:collapse;font-size:.7rem;}
thead th{position:sticky;top:0;z-index:2;background:var(--surface2);padding:9px 9px;text-align:left;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--border);}
tbody tr{border-bottom:1px solid rgba(30,45,69,.5);transition:background .15s;}
tbody tr:hover{background:var(--surface2);}
tbody td{padding:8px 9px;}
.badge{display:inline-flex;align-items:center;padding:2px 7px;border-radius:4px;font-size:.63rem;font-weight:600;white-space:nowrap;}
.b-nv{background:rgba(0,212,255,.1);color:var(--accent);}
.b-ctv{background:rgba(255,107,53,.1);color:var(--accent2);}
.b-dt{background:rgba(167,139,250,.1);color:#a78bfa;}
.b-ok{background:rgba(34,197,94,.1);color:var(--accent3);}
.b-warn{background:rgba(245,158,11,.1);color:var(--accent4);}
.b-bad{background:rgba(239,68,68,.1);color:var(--danger);}
.b-vung1{background:rgba(0,212,255,.08);color:var(--accent);font-size:.6rem;}
.b-vung2{background:rgba(167,139,250,.08);color:#a78bfa;font-size:.6rem;}
.mini-bar{display:inline-block;width:56px;height:9px;background:var(--surface2);border-radius:3px;overflow:hidden;vertical-align:middle;}
.mini-fill{height:100%;border-radius:3px;}
.donut-area{display:flex;align-items:center;gap:20px;}
.donut-svg{flex-shrink:0;}
.donut-legend{display:flex;flex-direction:column;gap:9px;flex:1;}
.dl-item{display:flex;align-items:center;gap:8px;}
.dl-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0;}
.dl-label{font-size:.7rem;color:var(--muted);}
.dl-val{font-size:.78rem;font-weight:700;margin-left:auto;padding-left:10px;font-family:'JetBrains Mono';}
.alert-list{display:flex;flex-direction:column;gap:7px;}
.alert-item{display:flex;align-items:flex-start;gap:9px;background:var(--surface2);border-radius:7px;padding:9px 11px;border-left:3px solid;}
.alert-item.red{border-color:var(--danger);}
.alert-item.amber{border-color:var(--accent4);}
.alert-item.green{border-color:var(--accent3);}
.alert-icon{font-size:.82rem;margin-top:1px;}
.alert-text p{font-size:.71rem;color:var(--text);}
.alert-text small{font-size:.63rem;color:var(--muted);}
.filter-row{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center;}
.filter-row select,.filter-row input{background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:6px 10px;border-radius:7px;font-size:.72rem;font-family:'Be Vietnam Pro';outline:none;cursor:pointer;}
.filter-row select:focus,.filter-row input:focus{border-color:var(--accent);}
.filter-btn{padding:6px 14px;border-radius:7px;background:rgba(0,212,255,.1);border:1px solid rgba(0,212,255,.25);color:var(--accent);font-size:.72rem;cursor:pointer;font-family:'Be Vietnam Pro';transition:background .2s;}
.filter-btn:hover{background:rgba(0,212,255,.2);}
.count-badge{font-size:.68rem;color:var(--muted);}
.tag-ml{width:9px;height:9px;border-radius:2px;background:var(--accent);display:inline-block;}
.tag-spk{width:9px;height:9px;border-radius:2px;background:var(--accent2);display:inline-block;}
.footer{text-align:center;font-size:.63rem;color:var(--muted);margin-top:22px;padding-top:14px;border-top:1px solid var(--border);}
.month-bar{display:flex;gap:6px;margin-bottom:16px;flex-wrap:wrap;align-items:center;}
.month-btn{padding:5px 11px;border-radius:7px;font-size:.72rem;cursor:pointer;font-family:'Be Vietnam Pro';border:1px solid var(--border);background:var(--surface2);color:var(--muted);transition:all .18s;white-space:nowrap;min-width:38px;text-align:center;}
.month-btn:hover{border-color:rgba(0,212,255,.4);color:var(--text);}
.month-btn.active{background:rgba(0,212,255,.12);border-color:rgba(0,212,255,.45);color:var(--accent);font-weight:700;}
.month-btn.no-data{opacity:.4;cursor:default;}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(3,1fr);}.row2,.row3{grid-template-columns:1fr;}}
`;

// ══════════════════════════════════════════════════════════════════════════
// MAIN COMPONENT
// ══════════════════════════════════════════════════════════════════════════
export default function Dashboard() {
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
  const [activeVung, setActiveVung] = useState("ALL");
  const [dailyType, setDailyType]   = useState("ML");
  const [dailyKho,  setDailyKho]    = useState("");
  const [selLoai,   setSelLoai]     = useState("");
  const [selMuc,    setSelMuc]      = useState("");
  const [selKho,    setSelKho]      = useState("");
  const [searchUID, setSearchUID]   = useState("");
  const [dashboardData, setDashboardData] = useState(FALLBACK_DASHBOARD_DATA);
  const [dataStatus, setDataStatus] = useState({ loading: true, error: "" });

  // inject CSS once
  useEffect(() => {
    const id = "dmx-styles";
    if (!document.getElementById(id)) {
      const s = document.createElement("style");
      s.id = id;
      s.textContent = STYLES;
      document.head.appendChild(s);
    }
  }, []);

  useEffect(() => {
    let cancelled = false;

    const refreshExcel = async ({ silent = false } = {}) => {
      if (!silent) setDataStatus(prev => ({ ...prev, loading: true }));

      try {
        const nextData = await loadExcelDashboardData(selectedMonth);
        if (cancelled) return;

        setDashboardData(nextData);
        setDataStatus({ loading: false, error: "" });
      } catch (error) {
        if (cancelled) return;

        setDataStatus({
          loading: false,
          error: error?.message || "Không đọc được file Excel",
        });
      }
    };

    refreshExcel();
    const intervalId = window.setInterval(() => refreshExcel({ silent: true }), EXCEL_POLL_INTERVAL_MS);

    return () => {
      cancelled = true;
      window.clearInterval(intervalId);
    };
  }, [selectedMonth]);

  const RAW = dashboardData.raw;
  const DAILY = dashboardData.daily;
  const KHO_LABEL = dashboardData.khoLabel;
  const VUNG_INFO = dashboardData.vungInfo;
  const META = dashboardData.meta;
  const dailyDays = META.days?.length
    ? META.days
    : Array.from({length:META.nDays || 31},(_,i)=>i+1);

  useEffect(() => {
    if (!VUNG_INFO[activeVung]) setActiveVung("ALL");
  }, [VUNG_INFO, activeVung]);

  // Reset kho filter khi đổi tháng
  useEffect(() => {
    setDailyKho("");
    setSelKho("");
  }, [selectedMonth]);

  // ── derived data ─────────────────────────────────────────────────────
  const getData = useCallback(() =>
    activeVung === "ALL" ? RAW : RAW.filter(r => r[3] === activeVung),
  [RAW, activeVung]);

  const d    = getData();
  const nv   = d.filter(r=>r[2]==="Nhân Viên");
  const ctv  = d.filter(r=>r[2]==="Cộng Tác Viên");
  const dt   = d.filter(r=>r[2]==="8 - ĐỐI TÁC GHLĐ");
  const totalML  = sumML(d);
  const totalSPK = sumSPK(d);
  const zeroML   = d.filter(r=>r[5]===0).length;
  const lowML    = d.filter(r=>r[5]>0&&r[5]<50).length;

  const khoList = [...new Set(d.map(r=>r[4]))].sort();

  // ── filtered table ───────────────────────────────────────────────────
  const tableRows = d.filter(r => {
    if (selLoai && r[2]!==selLoai) return false;
    if (selKho  && r[4]!==selKho)  return false;
    if (searchUID && !String(r[0]).includes(searchUID)) return false;
    if (selMuc==="cao"  && r[5]<200) return false;
    if (selMuc==="tb"   && (r[5]<50||r[5]>=200)) return false;
    if (selMuc==="thap" && (r[5]<=0||r[5]>49))   return false;
    if (selMuc==="zero" && r[5]!==0) return false;
    return true;
  }).sort((a,b)=>b[5]-a[5]);

  const maxTableML = tableRows.length ? Math.max(...tableRows.map(r=>r[5])) : 1;

  // ── top10 / low bars ─────────────────────────────────────────────────
  const sorted  = [...d].filter(r=>r[5]>0).sort((a,b)=>b[5]-a[5]);
  const maxML2  = sorted.length ? sorted[0][5] : 1;
  const top10   = sorted.slice(0,10);
  const lowBars = [...d].filter(r=>r[5]<50&&r[2]!=="8 - ĐỐI TÁC GHLĐ")
                        .sort((a,b)=>a[5]-b[5]).slice(0,10);

  // ── daily rows ───────────────────────────────────────────────────────
  const idx = dailyType==="ML" ? 2 : 3;
  const normalizeDayValues = values => dailyDays.map((_,i)=>Number(values?.[i] || 0));
  const dailyRows = dailyKho ? (() => {
    const rawInKho = RAW.filter(r=>r[4]===dailyKho);
    return rawInKho.map(r => {
      const daily   = DAILY.find(d=>d[0]===r[0]&&d[1]===dailyKho);
      const dayVals = daily ? normalizeDayValues(daily[idx]) : Array(dailyDays.length).fill(0);
      const rawTotal= dailyType==="ML" ? r[5] : r[6];
      const total   = daily ? dayVals.reduce((s,v)=>s+v,0) : rawTotal;
      return {uid:r[0],name:r[1],loai:r[2],dayVals,total,hasDailyData:!!daily};
    }).sort((a,b)=>b.total-a.total);
  })() : [];

  const realVals = dailyRows.filter(r=>r.hasDailyData).flatMap(r=>r.dayVals).filter(v=>v>0);
  const maxDayVal= realVals.length ? Math.max(...realVals) : 1;

  // ── donut data ───────────────────────────────────────────────────────
  const donutData = [
    {label:"Nhân Viên",    val:nv.length,  color:"#00d4ff"},
    {label:"Cộng Tác Viên",val:ctv.length, color:"#ff6b35"},
    {label:"Đối Tác",       val:dt.length,  color:"#a78bfa"},
  ];
  const donutTotal = donutData.reduce((s,x)=>s+x.val,0);

  const donutPaths = (() => {
    let angle = -90;
    return donutData.filter(x=>x.val>0).map(x => {
      const sweep = x.val/donutTotal*360;
      const [s,e] = [angle, angle+sweep];
      const R=48,r=27,cx=62,cy=62;
      const x1=cx+R*Math.cos(s*Math.PI/180), y1=cy+R*Math.sin(s*Math.PI/180);
      const x2=cx+R*Math.cos(e*Math.PI/180), y2=cy+R*Math.sin(e*Math.PI/180);
      const ix1=cx+r*Math.cos(s*Math.PI/180), iy1=cy+r*Math.sin(s*Math.PI/180);
      const ix2=cx+r*Math.cos(e*Math.PI/180), iy2=cy+r*Math.sin(e*Math.PI/180);
      const d = `M${x1},${y1} A${R},${R} 0 ${sweep>180?1:0},1 ${x2},${y2} L${ix2},${iy2} A${r},${r} 0 ${sweep>180?1:0},0 ${ix1},${iy1} Z`;
      angle += sweep;
      return {d, color:x.color, label:x.label, val:x.val};
    });
  })();

  // ── comp chart ───────────────────────────────────────────────────────
  const compGroups = [
    {label:"NV",  ml:avgML(nv),  spk:nv.length?Math.round(sumSPK(nv)/nv.length):0,  c:"#00d4ff"},
    {label:"CTV", ml:avgML(ctv), spk:ctv.length?Math.round(sumSPK(ctv)/ctv.length):0, c:"#ff6b35"},
    {label:"ĐT",  ml:avgML(dt),  spk:dt.length?Math.round(sumSPK(dt)/dt.length):0,   c:"#a78bfa"},
  ];
  const maxComp = Math.max(...compGroups.flatMap(g=>[g.ml,g.spk]),1);

  // ── alerts ───────────────────────────────────────────────────────────
  const alerts = [];
  if(zeroML>0) alerts.push({t:"red",  i:"🚫",p:`${zeroML} nhân sự ML = 0`,    s:"Kiểm tra lý do không có đơn lắp máy lạnh"});
  if(lowML>0)  alerts.push({t:"red",  i:"⚠️",p:`${lowML} NV/CTV có ML < 50`, s:"Năng suất thấp, cần hỗ trợ hoặc rà soát"});
  const tk = d.filter(r=>r[8]==="Tạm khóa");
  if(tk.length) alerts.push({t:"amber",i:"🔒",p:`${tk.length} nhân sự Tạm Khóa`,s:tk.map(r=>r[1]).join(", ")});
  const dtML = sumML(dt);
  if(totalML>0) alerts.push({t:"amber",i:"💡",p:`Đối Tác: ${Math.round(dtML/totalML*100)}% tổng ML (${dt.length} người)`,s:`TB ML/ĐT=${avgML(dt)} | TB ML/NV=${avgML(nv)}`});
  if(!alerts.length) alerts.push({t:"green",i:"✅",p:"Không có cảnh báo đặc biệt",s:"Tất cả chỉ số trong ngưỡng bình thường"});

  // ── low table ────────────────────────────────────────────────────────
  const lowTableRows = d.filter(r=>r[5]<=30&&r[2]!=="8 - ĐỐI TÁC GHLĐ").sort((a,b)=>a[5]-b[5]);

  // ── vs chart ─────────────────────────────────────────────────────────
  const vsGroups = [
    {label:"Đối Tác (ĐT)", ml:sumML(dt),  cnt:dt.length,  color:"#a78bfa"},
    {label:"Nhân Viên",    ml:sumML(nv),  cnt:nv.length,  color:"#00d4ff"},
    {label:"Cộng Tác Viên",ml:sumML(ctv), cnt:ctv.length, color:"#ff6b35"},
  ];
  const maxVS = Math.max(...vsGroups.map(g=>g.ml),1);

  // ── dist chart ───────────────────────────────────────────────────────
  const distBands = [
    {label:"= 0",     fn:r=>r[5]===0,                color:"#ef4444"},
    {label:"1–30",    fn:r=>r[5]>0&&r[5]<=30,        color:"#f59e0b"},
    {label:"31–99",   fn:r=>r[5]>=31&&r[5]<=99,      color:"#eab308"},
    {label:"100–199", fn:r=>r[5]>=100&&r[5]<=199,    color:"#00d4ff"},
    {label:"200–499", fn:r=>r[5]>=200&&r[5]<=499,    color:"#22c55e"},
    {label:"≥ 500",   fn:r=>r[5]>=500,               color:"#a78bfa"},
  ].map(b=>({...b, cnt:d.filter(b.fn).length}));
  const maxDist = Math.max(...distBands.map(b=>b.cnt),1);

  // ── vung info text ────────────────────────────────────────────────────
  const vi = VUNG_INFO[activeVung] || VUNG_INFO.ALL;
  const vungOptions = ["ALL", ...Object.keys(VUNG_INFO).filter(v=>v!=="ALL")];
  const vungLabel = v => v==="ALL" ? "🌐 Tất cả"
    : v==="DBSH" ? "🔵 Đồng Bằng Sông Hồng"
    : v==="DTB" ? "🟣 Đông Tây Bắc"
    : VUNG_INFO[v]?.name || v;
  const loadedAtText = formatLoadedAt(dashboardData.loadedAt);
  const dataStatusText = dataStatus.loading
    ? "Đang đọc Excel..."
    : dataStatus.error
      ? `${META.nDays} ngày · Dữ liệu dự phòng`
      : `${META.nDays} ngày · Excel ${loadedAtText}`;
  const dataStatusTitle = dataStatus.error || `Nguồn: ${dashboardData.source}`;
  const firstDay = dailyDays[0] || 1;
  const lastDay = dailyDays[dailyDays.length - 1] || META.nDays || 31;

  // ══════════════════════════════════════════════════════════════════════
  // RENDER
  // ══════════════════════════════════════════════════════════════════════
  return (
    <div className="wrap">

      {/* HEADER */}
      <header>
        <div className="logo-area">
          <h1>
            <img src={LOGO_B64} alt="Thợ ĐMX" style={{height:38,width:"auto",borderRadius:6,padding:"2px 4px",objectFit:"contain"}} />
            Dashboard Năng Suất Lắp Đặt
          </h1>
          <p>Điện Máy Xanh · {META.label} · {vi.name}</p>
        </div>
        <div className="badge-live" title={dataStatusTitle}><span className="dot-live"></span> {dataStatusText}</div>
      </header>

      {/* MONTH SELECTOR */}
      <div className="month-bar">
        <span className="vung-label">📅 Chọn tháng:</span>
        {Array.from({length:12},(_,i)=>i+1).map(m=>(
          <button key={m}
            className={`month-btn${selectedMonth===m?" active":""}`}
            onClick={()=>setSelectedMonth(m)}>
            T{m}
          </button>
        ))}
        <span style={{fontSize:".68rem",color:"var(--muted)",marginLeft:"auto",background:"var(--surface2)",padding:"4px 10px",borderRadius:6,border:"1px solid var(--border)"}}>
          File: {getExcelFileName(selectedMonth)}
        </span>
      </div>

      {/* VUNG SELECTOR */}
      <div className="vung-bar">
        <span className="vung-label">📍 Chọn vùng:</span>
        {vungOptions.map(v=>(
          <button key={v}
            className={`vung-btn${v==="DTB"?" v-dtb":""}${activeVung===v?" active":""}`}
            onClick={()=>setActiveVung(v)}>
            {vungLabel(v)}
          </button>
        ))}
        <span className="vung-info">{vi.name} · {d.length} nhân sự</span>
      </div>

      {/* KPI */}
      <div className="kpi-row">
        {[
          {c:"c1",label:"Tổng Nhân Sự",  val:d.length,            sub:`NV ${nv.length} · CTV ${ctv.length} · ĐT ${dt.length}`},
          {c:"c2",label:"Tổng ML",       val:totalML.toLocaleString(), sub:"Máy lạnh lắp đặt trong tháng"},
          {c:"c3",label:"Tổng SPK",      val:totalSPK.toLocaleString(),sub:"Sản phẩm khác"},
          {c:"c4",label:"TB ML/Người",   val:avgML(d),              sub:`NV:${avgML(nv)} · CTV:${avgML(ctv)} · ĐT:${avgML(dt)}`},
          {c:"c5",label:"Cảnh Báo Thấp", val:zeroML+lowML,          sub:`ML=0: ${zeroML} · ML<50: ${lowML}`},
        ].map(k=>(
          <div key={k.label} className={`kpi ${k.c}`}>
            <div className="kpi-label">{k.label}</div>
            <div className="kpi-val">{k.val}</div>
            <div className="kpi-sub">{k.sub}</div>
          </div>
        ))}
      </div>

      {/* TOP 10 + LOW BAR */}
      <div className="row2">
        <div className="panel">
          <div className="panel-title">🏆 <span>Top 10</span> Năng Suất Cao Nhất — ML</div>
          <div className="bar-list">
            {top10.map(r=>{
              const pct=Math.round(r[5]/maxML2*100);
              const lb=r[2]==="Nhân Viên"?"b-nv":r[2]==="Cộng Tác Viên"?"b-ctv":"b-dt";
              const ll=r[2]==="Nhân Viên"?"NV":r[2]==="Cộng Tác Viên"?"CTV":"ĐT";
              return (
                <div key={r[0]} className="bar-item">
                  <div className="bar-name" style={{minWidth:200}}>
                    <span style={{fontFamily:"JetBrains Mono",fontSize:".6rem",color:"var(--accent)",marginRight:5}}>{r[0]}</span>
                    {r[1]} <span style={{fontSize:".6rem",color:"var(--muted)"}}>({ll}·{r[4]})</span>
                  </div>
                  <div className="bar-track"><div className="bar-fill" style={{width:`${pct}%`,background:colorByLoai(r[2])}}></div></div>
                  <div className="bar-val">{r[5]}</div>
                </div>
              );
            })}
          </div>
        </div>
        <div className="panel">
          <div className="panel-title">⚠️ <span>Cảnh Báo</span> NV/CTV Năng Suất ML ≤ 30</div>
          <div className="bar-list">
            {lowBars.length===0
              ? <p style={{color:"var(--accent3)",fontSize:".72rem"}}>✅ Không có NV/CTV nào dưới ngưỡng</p>
              : lowBars.map(r=>{
                  const pct=r[5]===0?0:Math.round(r[5]/50*100);
                  const color=r[5]===0?"#ef4444":"#f59e0b";
                  return (
                    <div key={r[0]} className="bar-item">
                      <div className="bar-name" style={{minWidth:200}}>
                        <span style={{fontFamily:"JetBrains Mono",fontSize:".6rem",color:"var(--accent)",marginRight:5}}>{r[0]}</span>
                        {r[1]} <span style={{fontSize:".6rem",color:"var(--muted)"}}>({r[2]==="Nhân Viên"?"NV":"CTV"}·{r[4]})</span>
                      </div>
                      <div className="bar-track"><div className="bar-fill" style={{width:`${Math.max(pct,3)}%`,background:color}}></div></div>
                      <div className="bar-val" style={{color}}>{r[5]}</div>
                    </div>
                  );
                })
            }
          </div>
        </div>
      </div>

      {/* TABLE + DONUT */}
      <div className="row3">
        <div className="panel">
          <div className="panel-title">📋 <span>Chi Tiết</span> Nhân Sự</div>
          <div className="filter-row">
            <select value={selLoai} onChange={e=>setSelLoai(e.target.value)}>
              <option value="">Tất cả loại</option>
              <option value="Nhân Viên">Nhân Viên</option>
              <option value="Cộng Tác Viên">Cộng Tác Viên</option>
              <option value="8 - ĐỐI TÁC GHLĐ">Đối Tác</option>
            </select>
            <select value={selMuc} onChange={e=>setSelMuc(e.target.value)}>
              <option value="">Tất cả mức</option>
              <option value="cao">ML Cao (≥200)</option>
              <option value="tb">ML Trung Bình (50–199)</option>
              <option value="thap">ML Thấp (1–49)</option>
              <option value="zero">ML = 0</option>
            </select>
            <select value={selKho} onChange={e=>setSelKho(e.target.value)}>
              <option value="">Tất cả kho</option>
              {khoList.map(k=><option key={k} value={k}>{KHO_LABEL[k]||k} ({k})</option>)}
            </select>
            <input value={searchUID} onChange={e=>setSearchUID(e.target.value)} placeholder="🔍 Tìm theo User ID..." style={{minWidth:160}} />
            <button className="filter-btn" onClick={()=>{setSelLoai("");setSelMuc("");setSelKho("");setSearchUID("");}}>Reset</button>
            <span className="count-badge">{tableRows.length} nhân sự</span>
          </div>
          <div className="tbl-wrap">
            <table>
              <thead>
                <tr>
                  {["#","User","Tên Nhân Viên","Loại","Vùng","Kho","ML","SPK","Tổng QĐ","Mức","%"].map(h=><th key={h}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {tableRows.map((r,i)=>{
                  const lv=getLevel(r[5]);
                  const pct=maxTableML>0?Math.round(r[5]/maxTableML*100):0;
                  const color=r[5]===0?"#ef4444":r[5]<50?"#f59e0b":r[5]<200?"#00d4ff":"#22c55e";
                  const lb=r[2]==="Nhân Viên"?"b-nv":r[2]==="Cộng Tác Viên"?"b-ctv":"b-dt";
                  const ll=r[2]==="Nhân Viên"?"NV":r[2]==="Cộng Tác Viên"?"CTV":"ĐT";
                  const vb=r[3]==="DBSH"?"b-vung1":"b-vung2";
                  const vl=r[3]==="DBSH"?"ĐBSH":"ĐTB";
                  const hl=searchUID&&String(r[0]).includes(searchUID);
                  return (
                    <tr key={r[0]}>
                      <td style={{color:"var(--muted)"}}>{i+1}</td>
                      <td style={{fontFamily:"JetBrains Mono",fontSize:".68rem"}}>
                        <span style={{background:hl?"rgba(0,212,255,.18)":undefined,color:"var(--accent)",borderRadius:hl?3:undefined,padding:hl?"1px 3px":undefined}}>{r[0]}</span>
                      </td>
                      <td><strong>{r[1]}</strong></td>
                      <td><span className={`badge ${lb}`}>{ll}</span></td>
                      <td><span className={`badge ${vb}`}>{vl}</span></td>
                      <td style={{color:"var(--muted)",fontSize:".66rem"}}>{KHO_LABEL[r[4]]||r[4]}</td>
                      <td><strong style={{color,fontFamily:"JetBrains Mono"}}>{r[5]}</strong></td>
                      <td style={{fontFamily:"JetBrains Mono",color:"var(--accent2)"}}>{r[6]}</td>
                      <td style={{fontFamily:"JetBrains Mono"}}>{r[7].toLocaleString()}</td>
                      <td><span className={`badge ${lv.cls}`}>{lv.label}</span></td>
                      <td><span className="mini-bar"><span className="mini-fill" style={{width:`${pct}%`,background:color,display:"block"}}></span></span></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* DONUT + COMP + ALERTS */}
        <div className="panel">
          <div className="panel-title">🥧 <span>Cơ Cấu</span> Nhân Sự</div>
          <div className="donut-area">
            <svg className="donut-svg" width="124" height="124" viewBox="0 0 124 124">
              {donutPaths.map((p,i)=><path key={i} d={p.d} fill={p.color} opacity=".85"/>)}
              <text x="62" y="57" textAnchor="middle" fill="#e2e8f0" fontSize="13" fontWeight="800" fontFamily="JetBrains Mono">{donutTotal}</text>
              <text x="62" y="71" textAnchor="middle" fill="#64748b" fontSize="7.5">người</text>
            </svg>
            <div className="donut-legend">
              {donutData.map(x=>(
                <div key={x.label} className="dl-item">
                  <div className="dl-dot" style={{background:x.color}}></div>
                  <div className="dl-label">{x.label}</div>
                  <div className="dl-val" style={{color:x.color}}>{x.val} <small style={{fontSize:".58rem",color:"var(--muted)"}}>({donutTotal>0?Math.round(x.val/donutTotal*100):0}%)</small></div>
                </div>
              ))}
              <div className="dl-item" style={{marginTop:4,paddingTop:7,borderTop:"1px solid var(--border)"}}>
                <div className="dl-label" style={{fontSize:".63rem",width:"100%"}}>TB ML: NV={avgML(nv)} · CTV={avgML(ctv)} · ĐT={avgML(dt)}</div>
              </div>
            </div>
          </div>
          <div style={{marginTop:18}}>
            <div className="panel-title" style={{marginBottom:10}}>📊 <span>TB ML</span> Theo Loại</div>
            <div style={{display:"flex",gap:5,marginBottom:8}}>
              <span className="tag-ml"></span><span style={{fontSize:".62rem",color:"var(--muted)"}}>ML</span>
              <span className="tag-spk" style={{marginLeft:8}}></span><span style={{fontSize:".62rem",color:"var(--muted)"}}>SPK (TB)</span>
            </div>
            <div style={{display:"flex",gap:16,alignItems:"flex-end",height:110}}>
              {compGroups.map(g=>{
                const hml=Math.round(g.ml/maxComp*90), hspk=Math.round(g.spk/maxComp*90);
                return (
                  <div key={g.label} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3,flex:1}}>
                    <div style={{display:"flex",gap:3,alignItems:"flex-end",height:90}}>
                      <div style={{width:24,height:Math.max(hml,2),background:g.c,borderRadius:"3px 3px 0 0",position:"relative"}}>
                        <span style={{position:"absolute",top:-15,left:"50%",transform:"translateX(-50%)",fontSize:".58rem",color:g.c,fontFamily:"JetBrains Mono",whiteSpace:"nowrap"}}>{g.ml}</span>
                      </div>
                      <div style={{width:24,height:Math.max(hspk,2),background:"var(--accent2)",borderRadius:"3px 3px 0 0",position:"relative"}}>
                        <span style={{position:"absolute",top:-15,left:"50%",transform:"translateX(-50%)",fontSize:".58rem",color:"var(--accent2)",fontFamily:"JetBrains Mono",whiteSpace:"nowrap"}}>{g.spk}</span>
                      </div>
                    </div>
                    <div style={{fontSize:".7rem",fontWeight:700,color:g.c}}>{g.label}</div>
                  </div>
                );
              })}
            </div>
          </div>
          <div style={{marginTop:18}}>
            <div className="panel-title" style={{marginBottom:8}}>🚨 <span>Cảnh Báo</span></div>
            <div className="alert-list">
              {alerts.map((a,i)=>(
                <div key={i} className={`alert-item ${a.t}`}>
                  <div className="alert-icon">{a.i}</div>
                  <div className="alert-text"><p>{a.p}</p><small>{a.s}</small></div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* LOW TABLE */}
      <div className="panel" style={{marginBottom:18}}>
        <div className="panel-title">🔴 <span>Danh Sách</span> NV/CTV ML = 0 Hoặc Dưới 30 Cần Xem Xét</div>
        <div className="tbl-wrap" style={{height:260}}>
          <table>
            <thead>
              <tr>{["#","User","Tên","Loại","Vùng","Kho","ML","SPK","Trạng Thái","Nhận Xét"].map(h=><th key={h}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {lowTableRows.length===0
                ? <tr><td colSpan={10} style={{textAlign:"center",color:"var(--accent3)",padding:16}}>✅ Không có NV/CTV nào dưới ngưỡng trong vùng này</td></tr>
                : lowTableRows.map((r,i)=>{
                    const color=r[5]===0?"#ef4444":"#f59e0b";
                    const lb=r[2]==="Nhân Viên"?"b-nv":"b-ctv";
                    const vb=r[3]==="DBSH"?"b-vung1":"b-vung2";
                    const nx=r[5]===0?"Không có đơn ML trong tháng":r[5]<10?"Rất thấp, cần xác minh":"Thấp, cần hỗ trợ";
                    return (
                      <tr key={r[0]}>
                        <td style={{color:"var(--muted)"}}>{i+1}</td>
                        <td style={{fontFamily:"JetBrains Mono",fontSize:".63rem",color:"var(--accent)"}}>{r[0]}</td>
                        <td><strong>{r[1]}</strong></td>
                        <td><span className={`badge ${lb}`}>{r[2]==="Nhân Viên"?"NV":"CTV"}</span></td>
                        <td><span className={`badge ${vb}`}>{r[3]==="DBSH"?"ĐBSH":"ĐTB"}</span></td>
                        <td style={{color:"var(--muted)",fontSize:".65rem"}}>{KHO_LABEL[r[4]]||r[4]}</td>
                        <td><strong style={{color,fontFamily:"JetBrains Mono"}}>{r[5]}</strong></td>
                        <td style={{fontFamily:"JetBrains Mono"}}>{r[6]}</td>
                        <td>{r[8]==="Tạm khóa"?<span className="badge b-bad">Tạm khóa</span>:<span className="badge b-ok">HĐ</span>}</td>
                        <td style={{color,fontSize:".66rem"}}>{nx}</td>
                      </tr>
                    );
                  })
              }
            </tbody>
          </table>
        </div>
      </div>

      {/* DAILY PANEL */}
      <div className="panel" style={{marginBottom:18}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14,flexWrap:"wrap"}}>
          <div className="panel-title" style={{marginBottom:0}}>📅 <span>Chi Tiết</span> Đơn Hàng Theo Ngày — {META.label}</div>
          <div style={{display:"flex",gap:8,marginLeft:"auto",alignItems:"center",flexWrap:"wrap"}}>
            <select value={dailyKho} onChange={e=>setDailyKho(e.target.value)}
              style={{background:"var(--surface2)",border:"1px solid var(--border)",color:"var(--text)",padding:"6px 10px",borderRadius:7,fontSize:".72rem",fontFamily:"Be Vietnam Pro",outline:"none",cursor:"pointer"}}>
              <option value="">— Chọn kho để xem chi tiết theo ngày —</option>
              {khoList.map(k=><option key={k} value={k}>{KHO_LABEL[k]||k} ({k})</option>)}
            </select>
            <div style={{display:"flex",gap:4}}>
              {["ML","SPK"].map(t=>(
                <button key={t} onClick={()=>setDailyType(t)}
                  style={{padding:"5px 14px",borderRadius:6,fontSize:".72rem",cursor:"pointer",fontFamily:"Be Vietnam Pro",
                    border:dailyType===t?(t==="ML"?"1px solid rgba(0,212,255,.4)":"1px solid rgba(255,107,53,.4)"):"1px solid var(--border)",
                    background:dailyType===t?(t==="ML"?"rgba(0,212,255,.15)":"rgba(255,107,53,.15)"):"var(--surface2)",
                    color:dailyType===t?(t==="ML"?"var(--accent)":"var(--accent2)"):"var(--muted)",
                    fontWeight:dailyType===t?700:400}}>
                  {t==="ML"?"🧊 Máy Lạnh":"📦 Sản Phẩm Khác"}
                </button>
              ))}
            </div>
            {dailyKho&&<span style={{fontSize:".68rem",color:"var(--muted)",whiteSpace:"nowrap"}}>
              Kho: {KHO_LABEL[dailyKho]||dailyKho} · {dailyRows.length} nhân sự · {dailyType==="ML"?"Máy Lạnh":"Sản Phẩm Khác"}
            </span>}
          </div>
        </div>
        {!dailyKho
          ? <div style={{textAlign:"center",padding:"32px 20px",color:"var(--muted)"}}>
              <div style={{fontSize:"2rem",marginBottom:10}}>📅</div>
              <div style={{fontSize:".8rem",fontWeight:600,color:"var(--text)",marginBottom:6}}>Chọn kho để xem chi tiết đơn hàng theo ngày</div>
              <div style={{fontSize:".72rem"}}>Dữ liệu sẽ hiển thị từ ngày {pad2(firstDay)} đến ngày {pad2(lastDay)} · {META.label}</div>
            </div>
          : <div style={{overflowX:"auto",overflowY:"auto",maxHeight:420}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:".68rem",whiteSpace:"nowrap"}}>
                <thead>
                  <tr>
                    <th style={{position:"sticky",left:0,zIndex:3,background:"var(--surface2)",minWidth:56,padding:"8px",textAlign:"center",fontSize:".63rem",color:"var(--muted)",borderBottom:"1px solid var(--border)",borderRight:"1px solid var(--border)"}}>User</th>
                    <th style={{position:"sticky",left:56,zIndex:3,background:"var(--surface2)",minWidth:140,padding:"8px",fontSize:".63rem",color:"var(--muted)",borderBottom:"1px solid var(--border)",borderRight:"1px solid var(--border)"}}>Họ &amp; Tên</th>
                    <th style={{position:"sticky",left:196,zIndex:3,background:"var(--surface2)",minWidth:38,padding:"8px 6px",textAlign:"center",fontSize:".63rem",color:"var(--muted)",borderBottom:"1px solid var(--border)",borderRight:"1px solid var(--border)"}}>Loại</th>
                    {dailyDays.map(d=><th key={d} style={{background:"var(--surface2)",minWidth:30,padding:"8px 3px",textAlign:"center",fontSize:".63rem",color:"var(--muted)",borderBottom:"1px solid var(--border)"}}>{d}</th>)}
                    <th style={{position:"sticky",right:0,zIndex:3,background:"var(--surface2)",minWidth:48,padding:"8px 6px",textAlign:"center",fontSize:".63rem",color:dailyType==="ML"?"var(--accent)":"var(--accent2)",borderBottom:"1px solid var(--border)",borderLeft:"1px solid var(--border)"}}>∑{dailyType}</th>
                  </tr>
                </thead>
                <tbody>
                  {dailyRows.map((r,ri)=>{
                    const lb=r.loai==="Nhân Viên"?"b-nv":r.loai==="Cộng Tác Viên"?"b-ctv":"b-dt";
                    const ll=r.loai==="Nhân Viên"?"NV":r.loai==="Cộng Tác Viên"?"CTV":"ĐT";
                    const rowBg=ri%2===0?"var(--surface)":"var(--surface2)";
                    const totalColor=r.total===0?"var(--danger)":r.total<30?"var(--accent4)":r.total<100?"var(--accent)":"var(--accent3)";
                    return (
                      <tr key={r.uid} style={{background:rowBg}}>
                        <td style={{position:"sticky",left:0,zIndex:1,background:rowBg,padding:"5px 8px",textAlign:"center",fontFamily:"JetBrains Mono",fontSize:".65rem",color:"var(--accent)",borderRight:"1px solid var(--border)"}}>{r.uid}</td>
                        <td style={{position:"sticky",left:56,zIndex:1,background:rowBg,padding:"5px 8px",fontSize:".68rem",fontWeight:600,borderRight:"1px solid var(--border)",whiteSpace:"nowrap"}}>{r.name}</td>
                        <td style={{position:"sticky",left:196,zIndex:1,background:rowBg,padding:"5px 6px",textAlign:"center",borderRight:"1px solid var(--border)"}}><span className={`badge ${lb}`} style={{display:"inline-flex",alignItems:"center",padding:"2px 5px",borderRadius:4,fontSize:".6rem",fontWeight:600}}>{ll}</span></td>
                        {r.hasDailyData
                          ? r.dayVals.map((v,di)=>{
                              if(v===0) return <td key={di} style={{textAlign:"center",padding:"5px 3px",color:"var(--border)",fontSize:".62rem"}}>·</td>;
                              const heat=Math.min(Math.round(v/maxDayVal*100),100);
                              const bg=dailyType==="ML"?`rgba(0,212,255,${(heat/100*0.55+0.08).toFixed(2)})`:`rgba(255,107,53,${(heat/100*0.55+0.08).toFixed(2)})`;
                              const txtColor=heat>55?"#fff":dailyType==="ML"?"var(--accent)":"var(--accent2)";
                              return <td key={di} style={{textAlign:"center",padding:"5px 3px",background:bg,borderRadius:3,fontFamily:"JetBrains Mono",fontSize:".62rem",fontWeight:600,color:txtColor}}>{v}</td>;
                            })
                          : dailyDays.map((_,di)=><td key={di} style={{textAlign:"center",padding:"5px 3px",color:"var(--border)",fontSize:".58rem"}}>–</td>)
                        }
                        <td style={{position:"sticky",right:0,zIndex:1,background:rowBg,padding:"5px 6px",textAlign:"center",fontFamily:"JetBrains Mono",fontSize:".7rem",fontWeight:800,color:totalColor,borderLeft:"1px solid var(--border)"}}>{r.total}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
        }
      </div>

      {/* VS + DIST */}
      <div className="row2">
        <div className="panel">
          <div className="panel-title">⚖️ <span>Đối Tác vs NV+CTV</span> — Tổng ML theo Loại</div>
          <div className="bar-list" style={{marginTop:8}}>
            {vsGroups.map(g=>{
              const pct=Math.round(g.ml/maxVS*100);
              const share=totalML>0?Math.round(g.ml/totalML*100):0;
              return (
                <div key={g.label} className="bar-item">
                  <div className="bar-name" style={{minWidth:130}}>{g.label} <span style={{fontSize:".6rem",color:"var(--muted)"}}>({g.cnt} người)</span></div>
                  <div className="bar-track"><div className="bar-fill" style={{width:`${Math.max(pct,1)}%`,background:g.color}}></div></div>
                  <div className="bar-val" style={{color:g.color,minWidth:90}}>{g.ml.toLocaleString()} <span style={{color:"var(--muted)",fontSize:".6rem"}}>({share}%)</span></div>
                </div>
              );
            })}
          </div>
          <div style={{marginTop:14,padding:10,background:"var(--surface2)",borderRadius:8,fontSize:".7rem",color:"var(--muted)"}}>
            ⚡ <span style={{color:"var(--text)"}}>ĐT chiếm <strong style={{color:"#a78bfa"}}>{totalML>0?Math.round(sumML(dt)/totalML*100):0}%</strong> tổng ML với {dt.length} người. TB ML/ĐT=<strong style={{color:"#a78bfa"}}>{avgML(dt)}</strong> vs NV=<strong style={{color:"#00d4ff"}}>{avgML(nv)}</strong>.</span>
          </div>
        </div>
        <div className="panel">
          <div className="panel-title">📈 <span>Phân Phối</span> ML (Số người theo dải)</div>
          <div className="bar-list" style={{marginTop:8}}>
            {distBands.map(b=>{
              const pct=Math.round(b.cnt/maxDist*100);
              const share=d.length?Math.round(b.cnt/d.length*100):0;
              return (
                <div key={b.label} className="bar-item">
                  <div className="bar-name" style={{minWidth:64,fontFamily:"JetBrains Mono",fontSize:".66rem"}}>{b.label}</div>
                  <div className="bar-track"><div className="bar-fill" style={{width:`${Math.max(pct,2)}%`,background:b.color}}></div></div>
                  <div className="bar-val" style={{color:b.color,minWidth:72}}>{b.cnt} người <span style={{color:"var(--muted)",fontSize:".58rem"}}>({share}%)</span></div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div className="footer">Điện Máy Xanh · Bộ phận Vận Hành &amp; Quản Lý Dữ Liệu · {META.label}</div>
    </div>
  );
}
