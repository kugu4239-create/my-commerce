import React, { useState, useEffect, useCallback, useRef, useMemo } from "react";
import { createPortal } from "react-dom";
import dayjs from "dayjs";
// XLSX and papaparse are lazy-loaded on first use to keep initial bundle small
const getXLSX = () => import("xlsx").then(m => m);
const getPapa = () => import("papaparse").then(m => m.default);
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend,
  ScatterChart, Scatter, ZAxis, AreaChart, Area, ReferenceLine,
} from "recharts";

// ─────────────────────────────────────────────
// NOTE: Supabase 테이블/컬럼 추가 필요:
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS order_count integer DEFAULT 0;
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS refund_amount integer DEFAULT 0;
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS refund_count integer DEFAULT 0;
//   ALTER TABLE orders ADD COLUMN IF NOT EXISTS amount integer DEFAULT 0;
//
//   -- 인벤토리 상품코드 (Inventory Trend / Reorder Calculator):
//   ALTER TABLE inventory_snapshot ADD COLUMN IF NOT EXISTS product_code text DEFAULT '';
//   ALTER TABLE reorder_recommendations ADD COLUMN IF NOT EXISTS reorder_product_code text DEFAULT '';
//
//   -- 콘텐츠 임팩트 (인스타그램 포스트 캘린더):
//   CREATE TABLE IF NOT EXISTS instagram_posts (
//     id            serial PRIMARY KEY,
//     post_date     date    NOT NULL,
//     url           text    NOT NULL,
//     caption_memo  text,
//     created_at    timestamptz DEFAULT now()
//   );
//   CREATE INDEX IF NOT EXISTS instagram_posts_date_idx ON instagram_posts (post_date);
//   CREATE TABLE IF NOT EXISTS instagram_post_products (
//     post_id      int  NOT NULL REFERENCES instagram_posts(id) ON DELETE CASCADE,
//     product_name text NOT NULL,
//     PRIMARY KEY (post_id, product_name)
//   );
// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// DESIGN TOKENS
// ─────────────────────────────────────────────
const D = {
  bg:         "#f8f8f6",
  surface:    "#ffffff",
  surfaceAlt: "#f4f4f2",
  border:     "#e0e0da",
  borderMid:  "#ccccca",
  text:       "#111111",
  textSub:    "#222222",
  textMeta:   "#444444",
  black:      "#111111",
  green:      "#1a7a4f",
  red:        "#c0392b",
  amber:      "#b07d00",
  blue:       "#1a4fa5",
  SANKEY: [
    "#7EADD4","#7EB89E","#9E92C8","#D4A574","#82C4D8",
    "#9EC48C","#C4A8D4","#D4C07E","#82B8C4","#C49E82",
    "#A8C4D4","#C4B09E","#9EA8D4","#B8D49E","#D4A8B8",
    "#8CB4C4","#C4C49E","#B89EC4","#9EC4B4","#D4B88C",
  ],
};
const CH_COLOR={"자사몰":"#7EADD4","29CM":"#7EB89E","무신사":"#9E92C8"};
const chColor=ch=>CH_COLOR[ch]||"#A8B8C8";

// ─────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────
const nowStr = () => {
  const d = new Date();
  return `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")} ${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
};

function UpdatedAt({ ts }) {
  if (!ts) return null;
  return <span style={{ color:D.textMeta, fontSize:10, marginLeft:6 }}>업데이트 {ts}</span>;
}

const toNum = v => parseFloat(String(v||"0").replace(/[^0-9.-]/g,""))||0;

// 필터 버튼 햅틱 피드백 — data-hf 속성 버튼 클릭 시 짧은 더블 진동 (Android) / 클릭음 (iOS)
function _hapticTap(){
  if(navigator.vibrate){ navigator.vibrate([20,40,20]); return; }
  try{
    const ctx=new (window.AudioContext||window.webkitAudioContext)();
    const osc=ctx.createOscillator(); const gain=ctx.createGain();
    osc.connect(gain); gain.connect(ctx.destination);
    osc.frequency.value=800;
    gain.gain.setValueAtTime(0.06,ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+0.04);
    osc.start(); osc.stop(ctx.currentTime+0.04);
  }catch(_){}
}
if(typeof document!=="undefined"){
  document.addEventListener("click",e=>{
    if(e.target.closest("[data-hf]")) _hapticTap();
  },{passive:true});
}

// 스티키 닫기 버튼 상호 배제: 한 표만 스티키 표시
let _stickyActiveId = null;
const _stickyListeners = new Set();
const _stickyBroadcast = id => { _stickyActiveId = id; _stickyListeners.forEach(cb => cb(id)); };
function useStickyTable(myId, isVisible) {
  const { useState: _us, useEffect: _ue } = React;
  const [activeId, setActiveId] = _us(_stickyActiveId);
  _ue(() => { _stickyListeners.add(setActiveId); return () => _stickyListeners.delete(setActiveId); }, []);
  _ue(() => {
    if (isVisible) _stickyBroadcast(myId);
    else if (_stickyActiveId === myId) _stickyBroadcast(null);
  }, [isVisible, myId]);
  return activeId === myId && isVisible;
}

const toDate = raw => {
  if (!raw) return null;
  // Excel serial number (숫자로 저장된 날짜, 예: 46162 → 2026-05-11)
  const num = Number(raw);
  if (!isNaN(num) && num > 40000 && num < 80000 && String(raw).trim().length <= 5) {
    const d = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
    return d.toISOString().slice(0, 10);
  }
  // 시간/AM/PM/오전/오후 suffix 제거 → 날짜 부분만 추출 후 파싱
  // 예: "5/14/26 9:47" → "5/14/26", "2026-05-14T09:47:00" → "2026-05-14"
  const s = String(raw).trim()
    .replace(/T\d{2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$/, '')
    .replace(/\s+(오전|오후|AM|PM)?\s*\d{1,2}:\d{2}(:\d{2})?(\s*(AM|PM|오전|오후))?$/i, '')
    .trim();
  // YYYY년 M월 D일 (한국어 날짜 포맷)
  const mKr = s.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
  if (mKr) return `${mKr[1]}-${mKr[2].padStart(2,"0")}-${mKr[3].padStart(2,"0")}`;
  // YYYY. M. D (점+공백 포함, 예: "2025. 10. 16")
  const m0 = s.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (m0) return `${m0[1]}-${m0[2].padStart(2,"0")}-${m0[3].padStart(2,"0")}`;
  const m1 = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m1) return `${m1[1]}-${m1[2].padStart(2,"0")}-${m1[3].padStart(2,"0")}`;
  const m2 = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m2) return `${m2[3]}-${m2[2].padStart(2,"0")}-${m2[1].padStart(2,"0")}`;
  const m3 = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m3) return `${m3[1]}-${m3[2]}-${m3[3]}`;
  // YYMMDD 6자리 (예: "241223" → 2024-12-23)
  const m4 = s.match(/^(\d{2})(\d{2})(\d{2})$/);
  if (m4) { const yr=2000+parseInt(m4[1],10); return `${yr}-${m4[2]}-${m4[3]}`; }
  // M/D/YY (Excel 포맷, 예: "5/14/26" → 2026-05-14)
  const m5 = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2})$/);
  if (m5) { const yr=2000+parseInt(m5[3],10); return `${yr}-${m5[1].padStart(2,"0")}-${m5[2].padStart(2,"0")}`; }
  return null;
};

const fmtDays = days => {
  const d = Math.round(days || 0);
  if (d < 30) return `${d}일`;
  const months = Math.round(d / 30);
  if (months < 12) return `${months}개월`;
  const y = Math.floor(months / 12), m = months % 12;
  return m > 0 ? `${y}년 ${m}개월` : `${y}년`;
};

// 판매처 이름 정규화
const normChannel = raw => {
  if (!raw) return "미분류";
  const v = String(raw).trim();
  if (v === "MERRYON") return "자사몰";
  if (v === "예약거래") return "오프라인 스토어";
  return v;
};

// 이지어드민 CS 컬럼 → 내부 상태 (정상=배송, 배송후 전체 교환=교환, 배송후 전체 취소=반품)
const normCS = raw => {
  if (!raw) return "배송";
  const v = String(raw).trim().toLowerCase().replace(/\s/g,"");
  // 배송 전 취소 = '취소' (반품 아님 — 배송 자체가 일어나지 않은 주문 무효화)
  if (v.includes("배송전") && v.includes("취소")) return "취소";
  // 그 외 취소(배송 후/미명시) = 반품 처리
  if (v.includes("취소")) return "반품";
  if (v.includes("교환")) return "교환";
  return "배송";
};

function fmtEokMan(n){
  const eok=Math.floor(n/1e8);
  const man=Math.round((n%1e8)/1e4);
  if(man===0) return eok+"억";
  const cheon=Math.floor(man/1000);
  const baek=Math.floor((man%1000)/100);
  const sip=Math.floor((man%100)/10);
  const il=man%10;
  let s="";
  if(cheon) s+=cheon+"천";
  if(baek) s+=baek+"백";
  if(sip) s+=sip+"십";
  if(il) s+=il;
  return eok+"억"+s+"만";
}
// 억 이상: 천만 단위까지만 표시 (테이블 셀용)
const fmtWonShort = n => {
  if (!n) return "—";
  if (n >= 1e8) {
    const eok = Math.floor(n / 1e8);
    const cheon = Math.floor((n % 1e8) / 1e7);
    return "₩" + eok + "억" + (cheon > 0 ? cheon + "천만" : "");
  }
  if (n >= 1e4) return "₩" + Math.round(n / 1e4) + "만";
  return "₩" + n.toLocaleString();
};
const fmtWon = n => {
  if (!n) return "—";
  if (n>=1e8) return "₩"+fmtEokMan(n);
  if (n>=1e4) return "₩"+(n/1e4).toFixed(0)+"만";
  return "₩"+n.toLocaleString();
};

const COLOR_HEX={
  // 한글
  "블랙":"#1C1C1C","화이트":"#F8F8F8","차콜":"#3D3D3D","베이지":"#E8DCC8","아이보리":"#F4F0E0",
  "크림":"#FDF6E3","브라운":"#7B4F2E","카키":"#7B7142","네이비":"#1A2A4A","그레이":"#9E9E9E",
  "핑크":"#F4A7B9","레드":"#D94F4F","블루":"#4A7EC7","그린":"#5A9E6F","옐로우":"#F5C842",
  "퍼플":"#8B5DA8","오렌지":"#E87D3E","민트":"#7EC8B8","스카이블루":"#87CEEB","버건디":"#7C2D3E",
  "와인":"#6B2737","코랄":"#E87060","올리브":"#6B7C3A","라벤더":"#B89FCC","머스타드":"#C8A020",
  "연그레이":"#C8C8C8","다크네이비":"#0D1A30","오트밀":"#D8C8AE","샌드":"#C8B090","모카":"#8B6348",
  "스트라이프":"linear-gradient(135deg,#333 25%,#fff 25%,#fff 50%,#333 50%,#333 75%,#fff 75%)",
  // 영문
  "BLACK":"#1C1C1C","WHITE":"#F8F8F8","CHARCOAL":"#3D3D3D","BEIGE":"#E8DCC8","IVORY":"#F4F0E0",
  "CREAM":"#FDF6E3","BROWN":"#7B4F2E","KHAKI":"#7B7142","NAVY":"#1A2A4A","GRAY":"#9E9E9E","GREY":"#9E9E9E",
  "PINK":"#F4A7B9","RED":"#D94F4F","BLUE":"#4A7EC7","GREEN":"#5A9E6F","YELLOW":"#F5C842",
  "PURPLE":"#8B5DA8","ORANGE":"#E87D3E","MINT":"#7EC8B8","SKY BLUE":"#87CEEB","SKY":"#87CEEB",
  "BURGUNDY":"#7C2D3E","WINE":"#6B2737","CORAL":"#E87060","OLIVE":"#6B7C3A","LAVENDER":"#B89FCC",
  "MUSTARD":"#C8A020","OATMEAL":"#D8C8AE","SAND":"#C8B090","MOCHA":"#8B6348",
  "STRIPE":"linear-gradient(135deg,#333 25%,#fff 25%,#fff 50%,#333 50%,#333 75%,#fff 75%)",
  "DARK NAVY":"#0D1A30","LIGHT GRAY":"#C8C8C8",
};
function colorSwatch(name){
  const hex=COLOR_HEX[name]||COLOR_HEX[name?.toUpperCase()];
  if(!hex) return null;
  const isLight=["#F8F8F8","#F4F0E0","#FDF6E3","#E8DCC8","#F4F0E0","#D8C8AE"].includes(hex);
  const isGradient=hex.startsWith("linear");
  return <span style={{
    display:"inline-block",width:10,height:10,borderRadius:"50%",flexShrink:0,
    background:hex,
    border:`1px solid ${isLight||isGradient?"#ccc":"transparent"}`,
    verticalAlign:"middle",marginRight:5,
  }}/>;
}

const KR_COLORS=new Set(Object.keys(COLOR_HEX));
const SIZE_RE=/^(XS|S|M|L|XL|2XL|3XL|XXL|Free|F|\d{3})(-\S+)?$/i;

function parseOption(productName, optionName) {
  let color=null, size=null;
  // 상품명 [COLOR] 추출
  ((productName||"").match(/\[([^\]]+)\]/g)||[]).forEach(b=>{
    const v=b.slice(1,-1).trim();
    if(!SIZE_RE.test(v)) color=color||v;
  });
  // 옵션명 파싱
  const raw=(optionName||"").replace(/^\[|\]$/g,"").trim();
  if(!raw) return {color,size};
  if(raw.includes("-")){
    raw.split("-").forEach(p=>{
      const v=p.trim();
      if(SIZE_RE.test(v)) size=size||v.toUpperCase();
      else if(KR_COLORS.has(v)) color=color||v;
    });
  } else if(SIZE_RE.test(raw)){
    size=raw.toUpperCase();
  } else {
    color=color||raw;
  }
  return {color,size};
}

function useWindowWidth(){
  const [w,setW]=useState(()=>window.innerWidth);
  useEffect(()=>{
    const h=()=>setW(window.innerWidth);
    window.addEventListener("resize",h);
    return()=>window.removeEventListener("resize",h);
  },[]);
  return w;
}

function detectFields(columns) {
  const lc = columns.map(c => c.toLowerCase().replace(/\s/g,""));
  const f = (...kws) => { const i=lc.findIndex(c=>kws.some(k=>c.includes(k))); return i>=0?columns[i]:null; };
  return {
    channel:      f("판매처","channel","플랫폼","채널","mall","store","platform"),
    product:      f("상품명","product","품명","item","name"),
    option:       f("옵션","option","size","color","사이즈","색상"),
    qty:          f("수량","qty","quantity","개수","판매수량","입고"),
    cs:           f("cs","처리","cs처리","cs상태"),
    date:         f("배송일","delivery_date")||f("주문일","날짜","date","order_date","주문날짜","reg_date"),
    orderId:      f("관리번호","order_id","주문번호","orderid"),
    memo:         f("메모","memo","비고","note"),
    revenue:      f("금액","revenue","sales","매출","price","가격","결제금액","주문금액"),
  };
}

// 기간 필터 유틸
// 로컬 타임존 기준 날짜 문자열 (UTC 기반 toISOString 대신 사용)
function localDate(offsetDays=0){
  const d=new Date();d.setDate(d.getDate()+offsetDays);
  return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
}

function filterByDate(rows, dateField, period, customStart, customEnd, upToYesterday=false) {
  if (period === "all") return rows;
  const today = localDate(0);
  const ceiling = upToYesterday ? localDate(-1) : today;
  if (period === "week") {
    const now = new Date();
    const dow = now.getDay() || 7;
    const monday = new Date(now);
    monday.setDate(now.getDate() - dow + 1);
    const cutStr = [monday.getFullYear(),String(monday.getMonth()+1).padStart(2,'0'),String(monday.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cutStr && r[dateField] <= ceiling);
  }
  if (period === "yd") {
    const yStr = localDate(-1);
    return rows.filter(r => r[dateField] === yStr);
  }
  if (period === "7d") {
    const c = new Date(); c.setDate(c.getDate()-7);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10) && r[dateField] <= ceiling);
  }
  if (period === "14d") {
    const c = new Date(); c.setDate(c.getDate()-14);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10) && r[dateField] <= ceiling);
  }
  if (period === "1m") {
    const d=new Date(); d.setMonth(d.getMonth()-1);
    const cut=[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cut && r[dateField] <= ceiling);
  }
  if (period === "tm") {
    const n=new Date();
    const cut=[n.getFullYear(),String(n.getMonth()+1).padStart(2,'0'),'01'].join('-');
    return rows.filter(r => r[dateField] >= cut && r[dateField] <= ceiling);
  }
  if (period === "3m") {
    const d=new Date(); d.setMonth(d.getMonth()-3);
    const cut=[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cut && r[dateField] <= ceiling);
  }
  if (period === "6m") {
    const d=new Date(); d.setMonth(d.getMonth()-6);
    const cut=[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cut && r[dateField] <= ceiling);
  }
  if (period === "custom" && customStart && customEnd) {
    return rows.filter(r => r[dateField] >= customStart && r[dateField] <= customEnd);
  }
  return rows;
}

function getPeriodStr(period, customStart, customEnd) {
  const pad=n=>String(n).padStart(2,'0');
  const fmt=d=>`${d.getFullYear()}.${pad(d.getMonth()+1)}.${pad(d.getDate())}`;
  const today=new Date();
  const todayStr=fmt(today);
  if(period==="custom"&&customStart&&customEnd)
    return `${customStart.replace(/-/g,'.')} ~ ${customEnd.replace(/-/g,'.')}`;
  if(period==="all") return "전체";
  if(period==="yd"){const d=new Date(today);d.setDate(d.getDate()-1);return fmt(d);}
  if(period==="tm"){const s=new Date(today.getFullYear(),today.getMonth(),1);return `${fmt(s)} ~ ${todayStr}`;}
  const s=new Date(today);
  if(period==="7d") s.setDate(s.getDate()-7);
  else if(period==="14d") s.setDate(s.getDate()-14);
  else if(period==="1m") s.setMonth(s.getMonth()-1);
  else if(period==="3m") s.setMonth(s.getMonth()-3);
  else if(period==="6m") s.setMonth(s.getMonth()-6);
  else if(period==="week"){const dow=s.getDay()||7;s.setDate(s.getDate()-dow+1);}
  else return "";
  return `${fmt(s)} ~ ${todayStr}`;
}

// ─────────────────────────────────────────────
// SHARED UI
function InfoTip({ text, children }) {
  const [show,setShow]=useState(false);
  return (
    <span style={{position:"relative",display:"inline-flex",alignItems:"center"}}
      onMouseEnter={()=>setShow(true)} onMouseLeave={()=>setShow(false)}>
      {children}
      {show&&(
        <span style={{position:"absolute",bottom:"calc(100% + 6px)",left:"50%",transform:"translateX(-50%)",
          background:"#333",color:"#fff",fontSize:11,lineHeight:1.5,padding:"6px 10px",borderRadius:6,
          whiteSpace:"pre-wrap",maxWidth:240,width:"max-content",zIndex:999,
          boxShadow:"0 2px 8px rgba(0,0,0,0.18)",pointerEvents:"none"}}>
          {text}
          <span style={{position:"absolute",top:"100%",left:"50%",transform:"translateX(-50%)",
            border:"5px solid transparent",borderTopColor:"#333"}}/>
        </span>
      )}
    </span>
  );
}

// ─────────────────────────────────────────────
const Card=React.forwardRef(function Card({ children, style={} },ref) {
  return (
    <div ref={ref} style={{ background:D.surface, border:`1px solid ${D.border}`,
      borderRadius:10, padding:"16px 18px", ...style }}>
      {children}
    </div>
  );
});
function SecTitle({ children, ts }) {
  return (
    <div style={{ display:"flex", alignItems:"baseline", gap:6, marginBottom:12 }}>
      <span style={{ color:D.textSub, fontSize:12, letterSpacing:"0.1em", textTransform:"uppercase", fontFamily:"'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif", fontWeight:600 }}>{children}</span>
      <UpdatedAt ts={ts}/>
    </div>
  );
}
function KPI({ label, value, sub, accent="#111", onClick }) {
  return (
    <div onClick={onClick} style={{ background:D.surface, border:`1px solid ${D.border}`,
      borderRadius:9, padding:"14px 16px", flex:1, minWidth:110,
      cursor:onClick?"pointer":"default", transition:"box-shadow 0.15s" }}
      onMouseEnter={e=>{if(onClick)e.currentTarget.style.boxShadow=`0 0 0 2px ${accent}44`;}}
      onMouseLeave={e=>{e.currentTarget.style.boxShadow="";}}>
      <div style={{ color:D.textMeta, fontSize:10, letterSpacing:"0.09em", textTransform:"uppercase", marginBottom:5 }}>
        {label}{onClick&&<span style={{marginLeft:4,fontSize:9,opacity:0.5}}>↗</span>}
      </div>
      <div style={{ color:accent, fontSize:20, fontWeight:600 }}>{value}</div>
      {sub&&<div style={{ color:D.textMeta, fontSize:10, marginTop:3 }}>{sub}</div>}
    </div>
  );
}
function InfoBtn({ onClick }) {
  return (
    <button onClick={onClick}
      style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:"50%",
        width:20,height:20,display:"inline-flex",alignItems:"center",justifyContent:"center",
        fontSize:10,cursor:"pointer",color:D.textSub,marginLeft:6,verticalAlign:"middle"}}>
      i
    </button>
  );
}
function Tip({ active, payload, label }) {
  if (!active||!payload?.length) return null;
  return (
    <div style={{ background:D.surface, border:`1px solid ${D.border}`,
      borderRadius:7, padding:"8px 12px", fontSize:11, boxShadow:"0 2px 8px #0001" }}>
      <div style={{ color:D.textMeta, marginBottom:3 }}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{ color:p.color||D.text }}>
          {p.name}: <strong>{typeof p.value==="number"?p.value.toLocaleString():p.value}</strong>
        </div>
      ))}
    </div>
  );
}
function TabBar({ tabs, active, onChange }) {
  return (
    <div style={{ display:"flex", borderBottom:`1px solid ${D.border}`, marginBottom:18 }}>
      {tabs.map(t=>(
        <button key={t.key} onClick={()=>onChange(t.key)}
          style={{ background:"transparent", border:"none",
            borderBottom:active===t.key?`2px solid ${D.black}`:"2px solid transparent",
            color:active===t.key?D.black:D.textSub,
            padding:"9px 16px", fontWeight:active===t.key?600:400,
            fontSize:13, cursor:"pointer", marginBottom:-1, transition:"all 0.12s" }}>
          {t.label}
        </button>
      ))}
    </div>
  );
}
function Btn({ children, onClick, variant="primary", disabled, style={} }) {
  const styles = {
    primary: { bg:D.black,    cl:"#fff",     bd:"none" },
    ghost:   { bg:"transparent", cl:D.textSub, bd:`1px solid ${D.border}` },
    danger:  { bg:D.red,     cl:"#fff",     bd:"none" },
  };
  const s = styles[variant]||styles.primary;
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ background:disabled?"#ddd":s.bg, color:disabled?D.textMeta:s.cl,
        border:disabled?`1px solid ${D.border}`:s.bd, borderRadius:7, padding:"8px 18px",
        fontWeight:500, fontSize:13, cursor:disabled?"not-allowed":"pointer",
        transition:"all 0.12s", ...style }}>
      {children}
    </button>
  );
}
function Alert({ type, msg, ts }) {
  if (!msg) return null;
  const c = {success:D.green,error:D.red,warn:D.amber,info:D.blue}[type]||D.textSub;
  const i = {success:"✓",error:"✕",warn:"⚠",info:"i"}[type];
  return (
    <div style={{ background:`${c}0d`, border:`1px solid ${c}30`, borderRadius:7,
      padding:"9px 13px", color:c, fontSize:12, marginTop:9, lineHeight:1.5,
      whiteSpace:"pre-line" }}>
      {i} {msg}<UpdatedAt ts={ts}/>
    </div>
  );
}
function Steps({ current, steps }) {
  return (
    <div style={{ display:"flex", alignItems:"center", marginBottom:18 }}>
      {steps.map((s,i)=>(
        <div key={s} style={{ display:"flex", alignItems:"center" }}>
          <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:3 }}>
            <div style={{ width:22, height:22, borderRadius:"50%",
              background:i<current?D.black:i===current?D.black:"transparent",
              color:i<=current?"#fff":D.textMeta,
              border:`1px solid ${i<=current?D.black:D.border}`,
              display:"flex", alignItems:"center", justifyContent:"center",
              fontSize:10, fontWeight:600 }}>
              {i<current?"✓":i+1}
            </div>
            <div style={{ fontSize:9, color:i===current?D.black:D.textMeta,
              fontWeight:i===current?600:400, whiteSpace:"nowrap" }}>{s}</div>
          </div>
          {i<steps.length-1&&(
            <div style={{ width:28, height:1, background:i<current?D.black:D.border,
              margin:"0 4px", marginBottom:13 }}/>
          )}
        </div>
      ))}
    </div>
  );
}
function DateRange({ start, end, onStart, onEnd }) {
  const diff = start&&end&&start<=end?Math.round((new Date(end)-new Date(start))/86400000)+1:null;
  const inp = { background:D.surface, border:`1px solid ${D.border}`,
    borderRadius:6, padding:"7px 10px", fontSize:12, color:D.text,
    width:"100%", boxSizing:"border-box" };
  return (
    <div style={{ marginBottom:14 }}>
      <div style={{ color:D.textMeta, fontSize:10, textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:6 }}>데이터 기간</div>
      <div style={{ display:"flex", gap:6, alignItems:"center" }}>
        <input type="date" value={start} onChange={e=>onStart(e.target.value)} style={inp}/>
        <span style={{ color:D.textMeta }}>—</span>
        <input type="date" value={end} onChange={e=>onEnd(e.target.value)} style={inp}/>
      </div>
      {start&&end&&start>end&&<div style={{color:D.red,fontSize:10,marginTop:3}}>시작일이 종료일보다 늦습니다</div>}
      {diff&&<div style={{color:D.textMeta,fontSize:10,marginTop:3}}>{start} ~ {end} · {diff}일</div>}
    </div>
  );
}

// CalendarPicker — inline calendar for single or range date selection (light D theme)
// single mode: value + onChange(dateStr)
// range mode:  rangeStart + rangeEnd + onRangeChange({start,end})
// availableDates (optional Set<string>): only these dates are clickable
function CalendarPicker({ mode="single", value, onChange, rangeStart, rangeEnd, onRangeChange, availableDates, DC:dc }) {
  const C = dc || { bg:"transparent", surface:D.surface, border:D.border, text:D.text, sub:D.textSub, dim:D.textMeta, green:"#1a7a4f", greenBg:"rgba(26,122,79,0.12)" };
  const today = new Date().toISOString().slice(0,10);
  const initMonth = () => {
    const base = mode==="range" ? (rangeStart||value||today) : (value||today);
    const d = new Date(base||today);
    return { y: d.getFullYear(), m: d.getMonth() };
  };
  const [cal, setCal] = useState(initMonth);
  // range picking state: "start" = waiting for start click, "end" = waiting for end click
  const [picking, setPicking] = useState("start");

  const firstDay = new Date(cal.y, cal.m, 1).getDay();
  const totalDays = new Date(cal.y, cal.m+1, 0).getDate();
  const monthStr = `${cal.y}.${String(cal.m+1).padStart(2,"0")}`;
  const prevMonth = () => setCal(p => { let {y,m}=p; m--; if(m<0){m=11;y--;} return{y,m}; });
  const nextMonth = () => setCal(p => { let {y,m}=p; m++; if(m>11){m=0;y++;} return{y,m}; });
  const ds = (d) => `${cal.y}-${String(cal.m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;

  const handleClick = (dateStr) => {
    if(availableDates && !availableDates.has(dateStr)) return;
    if(mode==="single") { onChange&&onChange(dateStr); return; }
    // range mode
    if(picking==="start" || !rangeStart) {
      onRangeChange&&onRangeChange({start:dateStr,end:""});
      setPicking("end");
    } else {
      if(dateStr < rangeStart) {
        onRangeChange&&onRangeChange({start:dateStr,end:rangeStart});
      } else {
        onRangeChange&&onRangeChange({start:rangeStart,end:dateStr});
      }
      setPicking("start");
    }
  };

  const inRange = (dateStr) => rangeStart && rangeEnd && dateStr > rangeStart && dateStr < rangeEnd;
  const isStart = (dateStr) => dateStr === rangeStart;
  const isEnd   = (dateStr) => dateStr === rangeEnd;
  const isSelected = (dateStr) => mode==="single" ? dateStr===value : isStart(dateStr)||isEnd(dateStr);

  const diff = mode==="range" && rangeStart && rangeEnd && rangeStart<=rangeEnd
    ? Math.round((new Date(rangeEnd)-new Date(rangeStart))/86400000)+1 : null;

  const btnBase = { border:"none", borderRadius:5, padding:"5px 0", fontSize:11,
    textAlign:"center", cursor:"pointer", width:28, height:26, lineHeight:"16px" };

  return (
    <div style={{ display:"inline-block" }}>
      {/* Month nav */}
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8, gap:4 }}>
        <button onClick={prevMonth} style={{ background:"none", border:"none", color:C.sub, cursor:"pointer", fontSize:16, padding:"0 4px", lineHeight:1 }}>‹</button>
        <span style={{ fontSize:12, fontWeight:600, color:C.text }}>{monthStr}</span>
        <button onClick={nextMonth} style={{ background:"none", border:"none", color:C.sub, cursor:"pointer", fontSize:16, padding:"0 4px", lineHeight:1 }}>›</button>
      </div>
      {/* Weekday headers */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(7,28px)", gap:2, fontSize:10, color:C.sub, marginBottom:3, textAlign:"center" }}>
        {["일","월","화","수","목","금","토"].map(d=><span key={d}>{d}</span>)}
      </div>
      {/* Days grid */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(7,28px)", gap:2 }}>
        {Array.from({length:firstDay}).map((_,i)=><span key={`e${i}`}/>)}
        {Array.from({length:totalDays}).map((_,i)=>{
          const day=i+1; const dateStr=ds(day);
          const avail = !availableDates || availableDates.has(dateStr);
          const sel = isSelected(dateStr);
          const rng = inRange(dateStr);
          const bg = sel ? C.green : rng ? C.greenBg : "transparent";
          const col = sel ? "#fff" : avail ? C.text : C.dim;
          return (
            <button key={day} onClick={()=>handleClick(dateStr)} disabled={!avail}
              style={{ ...btnBase, background:bg, color:col,
                fontWeight:sel?700:avail?500:400,
                cursor:avail?"pointer":"default",
                outline: (isStart(dateStr)||isEnd(dateStr)) ? `2px solid ${C.green}` : "none",
                outlineOffset: -1,
              }}>
              {day}
            </button>
          );
        })}
      </div>
      {/* Status line */}
      {mode==="range" && (
        <div style={{ marginTop:6, fontSize:10, color:C.sub, textAlign:"center", minHeight:14 }}>
          {!rangeStart ? "시작일을 선택하세요"
            : !rangeEnd ? <span style={{color:C.green}}>{rangeStart} 선택됨 · 종료일을 선택하세요</span>
            : <span>{rangeStart} ~ {rangeEnd} · {diff}일</span>}
        </div>
      )}
      {mode==="single" && value && (
        <div style={{ marginTop:6, fontSize:10, color:C.sub, textAlign:"center" }}>{value}</div>
      )}
    </div>
  );
}

const parseHtmlTable=(text,opts)=>{
  // Parse HTML-disguised-as-XLS (common in Korean e-commerce exports)
  const parser=new DOMParser();
  const doc=parser.parseFromString(text,"text/html");
  const rows=Array.from(doc.querySelectorAll("tr"));
  if(!rows.length) return null;
  const toCell=td=>td.textContent.trim();
  const headers=Array.from(rows[0].querySelectorAll("th,td")).map(toCell);
  if(!headers.length) return null;
  const th=opts?.transformHeader||(h=>h);
  const mappedHeaders=headers.map(th);
  const data=[];
  for(let i=1;i<rows.length;i++){
    const cells=Array.from(rows[i].querySelectorAll("td,th")).map(toCell);
    if(cells.every(c=>!c)) continue;
    const row={};
    mappedHeaders.forEach((h,j)=>{row[h]=cells[j]??""});
    data.push(row);
  }
  return data;
};

// ─────────────────────────────────────────────
// 업로더 공통 — 누구나 알아볼 수 있는 에러 메시지 빌더
// ─────────────────────────────────────────────
// 컬럼 누락 에러: 파일에 있는 헤더 + 필요한 컬럼 + 누락된 것 + 해결법까지 함께 표시
function uploadErrColumns({ missing, required, headers }) {
  const headerLines = headers?.length
    ? headers.map(h=>`  • ${h}`).join("\n")
    : "  (헤더가 비어있음)";
  return [
    `필요한 컬럼을 파일에서 찾지 못했습니다.`,
    ``,
    `누락된 컬럼: ${missing.join(", ")}`,
    ``,
    `현재 파일의 컬럼 헤더 (첫 행):`,
    headerLines,
    ``,
    `이 업로더가 필요로 하는 컬럼:`,
    required.map(r=>`  • ${r}`).join("\n"),
    ``,
    `해결 방법: 파일의 첫 행 헤더에 위 컬럼명을 포함시키세요. 일부 유사 이름은 자동 매칭됩니다(예: '관리번호'→주문번호).`,
  ].join("\n");
}
// 파일 파싱 실패 에러 (포맷 불일치, 헤더 누락 등)
function uploadErrParse(detail="") {
  return [
    `파일을 읽을 수 없습니다.`,
    detail?`\n원인: ${detail}`:"",
    ``,
    `확인 사항:`,
    `  • CSV 또는 Excel(.xlsx / .xls) 파일이 맞는지`,
    `  • 첫 행이 컬럼 헤더이고, 그 아래에 1개 이상의 데이터 행이 있는지`,
    `  • 파일이 손상되지 않았는지 (Excel로 한 번 열어서 다시 저장해보세요)`,
  ].join("\n");
}
// 데이터 갈음 안내 — 업로드 직전에 보여줘야 함
function uploadReplaceWarn(prevCount, scope) {
  if(!prevCount||prevCount<=0) return "";
  return `⚠ 기존 ${prevCount.toLocaleString()}건이 삭제되고 새 데이터로 교체됩니다 (${scope}).`;
}

const parseAnyFile=(file,opts,completeCb,errorCb)=>{
  const ext=file.name.split(".").pop().toLowerCase();
  if(ext==="xlsx"||ext==="xls"){
    const reader=new FileReader();
    reader.onload=async e=>{
      try{
        // Detect HTML-disguised-as-XLS by peeking at first 100 bytes
        const peek=new TextDecoder("utf-8").decode(new Uint8Array(e.target.result,0,Math.min(200,e.target.result.byteLength))).trimStart().toLowerCase();
        const isHtml=peek.startsWith("<")||peek.includes("<html")||peek.includes("<meta")||peek.includes("<!doctype");

        let data=null;
        if(!isHtml){
          // True binary Excel — use XLSX.js
          try{
            const XLSX=await getXLSX();
            const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:false});
            const ws=wb.Sheets[wb.SheetNames[0]];
            data=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false,dateNF:"yyyy-mm-dd"});
            if(opts.transformHeader) data=data.map(row=>{const nr={};Object.keys(row).forEach(k=>{nr[opts.transformHeader(k)]=row[k];});return nr;});
          }catch(_){}
        }

        if(!data||!data.length){
          // HTML-as-XLS (이지어드민 등) — parse as HTML table
          const text=new TextDecoder("utf-8").decode(e.target.result);
          data=parseHtmlTable(text,opts);
        }

        if(!data||!data.length) throw new Error(uploadErrParse("Excel/HTML 모두 시도했으나 행을 추출하지 못했습니다"));
        completeCb({data});
      }catch(err){if(errorCb)errorCb(err);}
    };
    reader.readAsArrayBuffer(file);
  }else{
    getPapa().then(Papa=>Papa.parse(file,{...opts,complete:completeCb,error:errorCb}));
  }
};

function DropZone({ onFile, label="파일을 드래그 앤 드롭 또는 클릭하여 선택", fileName="", columns="" }) {
  const [hover,setHover]=useState(false);
  const handle=useCallback(e=>{
    e.preventDefault();
    const file=e.dataTransfer?.files?.[0]||e.target.files?.[0];
    if(file) onFile(file);
  },[onFile]);
  const cols=columns.split("·").map(s=>s.trim()).filter(Boolean);
  return (
    <label onDragOver={e=>{e.preventDefault();setHover(true);}}
      onDragLeave={()=>setHover(false)} onDrop={e=>{setHover(false);handle(e);}}
      style={{ display:"flex", flexDirection:"column", alignItems:"center",
        justifyContent:"center", minHeight:100, padding:"10px 14px",
        border:`1.5px dashed ${hover?D.black:D.border}`, borderRadius:9,
        cursor:"pointer", background:hover?D.surfaceAlt:D.surface, transition:"all 0.13s" }}>
      <input type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={handle}/>
      <div style={{ color:D.textSub, fontSize:13 }}>{label}</div>
      {fileName
        ?<div style={{color:D.textMeta,fontSize:11,marginTop:3}}>{fileName}</div>
        :<div style={{color:D.textMeta,fontSize:11,marginTop:3}}>드래그 앤 드롭 또는 클릭하여 파일 선택</div>}
      {!fileName&&cols.length>0&&(
        <div style={{marginTop:6,fontSize:10,textAlign:"center",lineHeight:1.8}}>
          <span style={{color:D.textMeta,fontWeight:600}}>필요 컬럼 </span>
          <span style={{color:D.textMeta}}>{cols.join(" · ")}</span>
        </div>
      )}
    </label>
  );
}
function PreviewTable({ rows, cols, outIdx=new Set(), maxRows=80 }) {
  return (
    <div style={{ overflowY:"auto", maxHeight:340 }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
        <thead style={{ position:"sticky", top:0, background:D.surface }}>
          <tr>
            <th style={{ padding:"5px 7px", color:D.textMeta, fontWeight:400,
              borderBottom:`1px solid ${D.border}`, width:22, textAlign:"center" }}>#</th>
            {cols.map(c=>(
              <th key={c.key} style={{ padding:"5px 7px", textAlign:"left",
                color:D.textMeta, fontWeight:400, borderBottom:`1px solid ${D.border}` }}>{c.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.slice(0,maxRows).map((r,i)=>{
            const isOut=outIdx.has(i);
            return (
              <tr key={i} style={{ borderBottom:`1px solid ${D.border}`,
                background:isOut?"#fef2f2":"transparent" }}>
                <td style={{ padding:"5px 7px", color:isOut?D.red:D.textMeta, textAlign:"center" }}>
                  {isOut?"⚠":i+1}</td>
                {cols.map(c=>(
                  <td key={c.key} style={{ padding:"5px 7px",
                    color:isOut?D.red:c.color||D.text, fontWeight:c.bold?600:400,
                    maxWidth:c.maxW, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                    {c.fmt?c.fmt(r[c.key],r):r[c.key]||"—"}
                  </td>
                ))}
              </tr>
            );
          })}
          {rows.length>maxRows&&(
            <tr><td colSpan={cols.length+1} style={{padding:7,color:D.textMeta,textAlign:"center",fontSize:10}}>
              + {rows.length-maxRows}건 더</td></tr>
          )}
        </tbody>
      </table>
    </div>
  );
}
function StatRow({ items }) {
  return (
    <div style={{ display:"grid", gridTemplateColumns:`repeat(${items.length},1fr)`, gap:7, marginBottom:12 }}>
      {items.map(it=>(
        <div key={it.label} style={{ background:D.surfaceAlt, border:`1px solid ${D.border}`,
          borderRadius:7, padding:"9px 10px", textAlign:"center" }}>
          <div style={{ color:it.color||D.black, fontWeight:600, fontSize:17 }}>{it.value}</div>
          <div style={{ color:D.textMeta, fontSize:10, marginTop:2 }}>{it.label}</div>
        </div>
      ))}
    </div>
  );
}
function InfoModal({ show, onClose, title, children }) {
  if (!show) return null;
  return (
    <div style={{ position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.5)",
      zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center" }}
      onClick={onClose}>
      <div style={{ background:D.surface,borderRadius:12,padding:24,maxWidth:520,
        width:"90%",boxShadow:"0 8px 32px #0003" }}
        onClick={e=>e.stopPropagation()}>
        <div style={{ fontWeight:700,fontSize:15,marginBottom:14 }}>{title}</div>
        <div style={{ color:D.textSub,fontSize:13,lineHeight:1.9 }}>{children}</div>
        <button onClick={onClose}
          style={{ marginTop:18,width:"100%",background:D.surfaceAlt,
            border:`1px solid ${D.border}`,borderRadius:7,padding:"9px",
            fontSize:13,cursor:"pointer",color:D.textSub }}>
          닫기
        </button>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA DELETE SECTION
// ─────────────────────────────────────────────
function DataDeleteSection({ table, dateField, label, onDone }) {
  const today = new Date().toISOString().slice(0,10);
  const [start,setStart]=useState(today);
  const [end,setEnd]=useState(today);
  const [step,setStep]=useState(0); // 0=idle, 1=confirm1, 2=confirm2
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const dateValid = start&&end&&start<=end;

  const handleDelete = async () => {
    setLoading(true);
    const db = await getSupabase();
    // .select()를 붙여서 실제로 삭제된 행을 받아옴 (RLS 등으로 0건이 삭제되면 알 수 있음)
    const { data, error } = await db.from(table).delete()
      .gte(dateField,start).lte(dateField,end).select();
    setLoading(false);
    if (error) { setResult({type:"error",msg:`삭제 실패: ${error.message}`}); setStep(0); return; }
    const affected = (data||[]).length;
    if (affected===0) {
      setResult({type:"warn",msg:`삭제된 행 0건 — 해당 기간에 데이터가 없거나 Supabase RLS 정책에 의해 차단되었을 수 있습니다.`});
      setStep(0);
      return;
    }
    setResult({type:"success",msg:`${start} ~ ${end} · ${affected.toLocaleString()}건 삭제 완료`});
    setStep(0);
    onDone?.();
  };

  return (
    <Card style={{ border:`1px solid ${D.red}30`, background:"#fff9f9" }}>
      <div style={{ fontWeight:600,fontSize:13,color:D.red,marginBottom:12 }}>
        🗑 데이터 삭제 — {label}
      </div>
      <DateRange start={start} end={end} onStart={setStart} onEnd={setEnd}/>
      {step===0&&(
        <Btn variant="danger" onClick={()=>setStep(1)} disabled={!dateValid} style={{width:"100%"}}>
          이 기간 데이터 삭제
        </Btn>
      )}
      {step===1&&(
        <div>
          <div style={{color:D.amber,fontSize:12,marginBottom:10,padding:"8px 10px",
            background:`${D.amber}12`,borderRadius:6}}>
            ⚠ {start} ~ {end} 기간의 {label} 데이터를 삭제하시겠습니까?
          </div>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="ghost" onClick={()=>setStep(0)} style={{flex:1}}>취소</Btn>
            <Btn variant="danger" onClick={()=>setStep(2)} style={{flex:1}}>예, 삭제하겠습니다</Btn>
          </div>
        </div>
      )}
      {step===2&&(
        <div>
          <div style={{color:D.red,fontWeight:600,fontSize:12,marginBottom:10,padding:"8px 10px",
            background:`${D.red}12`,borderRadius:6}}>
            ⚠ 정말로 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.
          </div>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="ghost" onClick={()=>setStep(0)} style={{flex:1}}>취소</Btn>
            <Btn variant="danger" onClick={handleDelete} disabled={loading} style={{flex:1}}>
              {loading?"삭제 중...":"최종 삭제"}
            </Btn>
          </div>
        </div>
      )}
      {result&&<Alert type={result.type} msg={result.msg}/>}
    </Card>
  );
}

// ─────────────────────────────────────────────
// MULTI-COLUMN SANKEY  (상품명 → 판매처 → 반품)  입고수=블록 상하높이
// ─────────────────────────────────────────────
const SVG_W = 1400;

function ProductSankey({ stockRows, orderRows, period="3m", customStart, customEnd, limit=20 }) {
  const containerRef = useRef(null);
  const [ctnSize, setCtnSize] = useState({w: window.innerWidth, h: window.innerHeight});
  const [sel, setSel] = useState(null); // { type:"prod"|"ch"|"ret"|"exch", key:string }
  useEffect(() => {
    if (!containerRef.current) return;
    const ro = new ResizeObserver(([e]) => {
      setCtnSize({w: e.contentRect.width, h: e.contentRect.height});
    });
    ro.observe(containerRef.current);
    return () => ro.disconnect();
  }, []);

  const filteredOrders = useMemo(() => {
    return filterByDate(orderRows, "order_date", period, customStart, customEnd);
  }, [orderRows, period, customStart, customEnd]);

  // 입고: 기간 내 upload_date 기준으로 필터링 후 상품별 수량 합산
  const filteredStocks = useMemo(() => {
    return filterByDate(stockRows, "upload_date", period, customStart, customEnd);
  }, [stockRows, period, customStart, customEnd]);

  const data = useMemo(() => {
    const prodMap = {};
    filteredStocks.forEach(r => {
      const key = r.product_name || "미분류";
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, shipped:0, returned:0, exchanged:0, byChannel:{} };
      prodMap[key].stock += (r.qty||0);
    });
    // 주문 날짜 범위 수집 (상품별)
    const prodDates = {};
    filteredOrders.forEach(r => {
      const key = r.product_name || "미분류";
      const ch = r.channel || "미분류";
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, shipped:0, returned:0, exchanged:0, byChannel:{} };
      if (!prodMap[key].byChannel[ch]) prodMap[key].byChannel[ch] = { shipped:0, returned:0, exchanged:0 };
      if (r.status==="배송")  { prodMap[key].shipped++;   prodMap[key].byChannel[ch].shipped++; }
      if (r.status==="반품")  { prodMap[key].returned++;  prodMap[key].byChannel[ch].returned++; }
      if (r.status==="교환")  { prodMap[key].exchanged++; prodMap[key].byChannel[ch].exchanged++; }
      if (!prodDates[key]) prodDates[key] = { min: r.order_date, max: r.order_date };
      else {
        if (r.order_date < prodDates[key].min) prodDates[key].min = r.order_date;
        if (r.order_date > prodDates[key].max) prodDates[key].max = r.order_date;
      }
    });
    const prods = Object.values(prodMap)
      .filter(p => p.shipped>0||p.stock>0)
      .sort((a,b)=>(b.stock||0)-(a.stock||0)||(b.shipped||0)-(a.shipped||0))
      .slice(0, limit);
    const chanMap = {};
    filteredOrders.forEach(r => {
      const ch = r.channel||"미분류";
      if (!chanMap[ch]) chanMap[ch] = { name:ch, shipped:0, returned:0, exchanged:0, byProd:{} };
      if (!chanMap[ch].byProd[r.product_name||"미분류"]) chanMap[ch].byProd[r.product_name||"미분류"] = { shipped:0, returned:0, exchanged:0 };
      if (r.status==="배송") { chanMap[ch].shipped++; chanMap[ch].byProd[r.product_name||"미분류"].shipped++; }
      if (r.status==="반품") { chanMap[ch].returned++; chanMap[ch].byProd[r.product_name||"미분류"].returned++; }
      if (r.status==="교환") { chanMap[ch].exchanged++; chanMap[ch].byProd[r.product_name||"미분류"].exchanged++; }
    });
    const channels = Object.values(chanMap).sort((a,b)=>b.shipped-a.shipped);
    const totalReturned  = filteredOrders.filter(r=>r.status==="반품").length;
    const totalExchanged = filteredOrders.filter(r=>r.status==="교환").length;
    // 반품/교환 채널별 분포
    const retByCh = {}, exchByCh = {};
    filteredOrders.forEach(r => {
      const ch = r.channel||"미분류";
      if (r.status==="반품") retByCh[ch] = (retByCh[ch]||0)+1;
      if (r.status==="교환") exchByCh[ch] = (exchByCh[ch]||0)+1;
    });
    return { prods, channels, totalReturned, totalExchanged, prodDates, retByCh, exchByCh };
  }, [filteredStocks, filteredOrders, limit]);

  if (!data.prods.length) return (
    <div style={{ textAlign:"center", padding:80, color:D.textMeta, fontSize:13 }}>
      입고 또는 주문·배송 업로드 데이터를 등록하면<br/>상품별 물류 흐름이 표시됩니다
    </div>
  );

  const { prods, channels, totalReturned, totalExchanged, prodDates, retByCh, exchByCh } = data;
  const n = prods.length;

  // ── 레이아웃 상수 ──
  const PAD_T=36, PAD_H=8, ROW_GAP=4, MIN_H=10;
  const NODE_W=200;
  // 좌우 끝까지 사용: col0=왼쪽 끝, col1=중앙, col2=오른쪽 끝
  const COLS_X = [
    PAD_H,
    Math.round((SVG_W - NODE_W) / 2),
    SVG_W - NODE_W - PAD_H,
  ];

  const totalStock   = prods.reduce((s,p)=>s+p.stock,0)||1;
  const totalShipped = prods.reduce((s,p)=>s+p.shipped,0)||1;
  const chanTotal    = channels.reduce((s,c)=>s+c.shipped,0)||1;

  // 상품마다 충분한 높이 확보 — 압축 없이 총 높이를 늘림
  const ROW_H = 22;
  const TARGET_H = n * ROW_H;
  const rawH = prods.map(p => p.stock>0 ? Math.max(MIN_H, (p.stock/totalStock)*TARGET_H) : MIN_H);
  const rawSum = rawH.reduce((s,h)=>s+h,0);
  const prodH = rawH.map(h => Math.max(MIN_H, Math.round(h * TARGET_H / rawSum)));

  const yPos = [];
  let cumY = PAD_T+16;
  prodH.forEach(h => { yPos.push(cumY); cumY += h+ROW_GAP; });
  const blockTotalH = cumY - ROW_GAP - (PAD_T+16);
  const totalSvgH   = cumY + 30;

  // 컬럼1 높이: blockTotalH × (배송수 / 입고수) — 입고 대비 배송 비율
  const totalStockQty = filteredStocks.reduce((s,r)=>s+(r.qty||0),0)||1;
  const col1H = Math.max(MIN_H * channels.length, Math.round(blockTotalH * Math.min(1, chanTotal / totalStockQty)));

  // 컬럼2 높이: col1H × (반품+교환 / 배송) — 배송 대비 반품 비율
  const totalRE = (totalReturned + totalExchanged) || 1;
  const col2H = chanTotal > 0
    ? Math.max(0, Math.round(col1H * Math.min(1, (totalReturned + totalExchanged) / chanTotal)))
    : 0;

  const chanYOf = {};
  let cy = PAD_T+16;
  channels.forEach(ch=>{
    const h = Math.max(MIN_H, (ch.shipped/chanTotal)*col1H - ROW_GAP);
    chanYOf[ch.name] = cy + h/2;
    cy += h + ROW_GAP;
  });

  // 컬럼2: 반품/교환 블록 높이 분할
  const retBlockH  = totalReturned  > 0 ? Math.max(MIN_H, Math.round((totalReturned /totalRE)*col2H) - ROW_GAP) : 0;
  const exchBlockH = totalExchanged > 0 ? Math.max(MIN_H, Math.round((totalExchanged/totalRE)*col2H) - ROW_GAP) : 0;
  const retBlockY  = PAD_T+16;
  const exchBlockY = retBlockY + retBlockH + ROW_GAP;
  const retCenterY  = retBlockY  + retBlockH/2;
  const exchCenterY = exchBlockY + exchBlockH/2;

  const maxStroke    = Math.min(20, Math.max(4, Math.round(400/n))) * 2/3;
  const maxRetStroke = Math.min(18, Math.max(3, Math.round(360/n))) * 2/3;

  // SVG는 너비 기준으로만 스케일 → 폰트가 목표 px로 보이도록 SVG 좌표 단위로 역산
  const svgScale = (ctnSize.w / SVG_W) || 1;
  const _fs = px => px / svgScale;
  const hdrFs = _fs(13);
  const lblFs = _fs(11);
  const subFs = _fs(9);

  const headers = ["입고","판매처별 배송","반품/교환"];

  return (
    <div ref={containerRef} style={{ width:"100%" }}>
      <svg width="100%" viewBox={`0 0 ${SVG_W} ${totalSvgH}`}
        style={{ display:"block" }}>

        {/* 컬럼 헤더 */}
        {headers.map((h,ci)=>(
          <text key={h} x={COLS_X[ci]+NODE_W/2} y={PAD_T-4}
            textAnchor="middle" fill={D.textSub} fontSize={hdrFs} fontWeight="600">{h}</text>
        ))}

        {/* 상품 → 판매처 연결선 */}
        {prods.map((p,i)=>{
          if (!p.shipped) return null;
          const x1=COLS_X[0]+NODE_W, y1=yPos[i]+prodH[i]/2;
          return Object.entries(p.byChannel).map(([ch,v])=>{
            if (!v.shipped) return null;
            const x2=COLS_X[1], y2=chanYOf[ch]||PAD_T+20;
            const thick=Math.max(1,(v.shipped/totalShipped)*maxStroke);
            const mx=(x1+x2)/2;
            return <path key={`p${i}c${ch}`} d={`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`}
              fill="none" stroke={D.SANKEY[i%D.SANKEY.length]} strokeWidth={thick} opacity={0.14}/>;
          });
        })}

        {/* 판매처 → 반품 연결선 */}
        {channels.map((ch,ci)=>{
          const x1=COLS_X[1]+NODE_W; const y1=chanYOf[ch.name]||PAD_T+20;
          const mx=(x1+COLS_X[2])/2;
          return (
            <g key={`rc${ci}`}>
              {ch.returned>0&&<path d={`M${x1},${y1} C${mx},${y1} ${mx},${retCenterY} ${COLS_X[2]},${retCenterY}`}
                fill="none" stroke={D.red} strokeWidth={Math.max(1,(ch.returned/(totalReturned||1))*maxRetStroke)} opacity={0.22}/>}
              {ch.exchanged>0&&<path d={`M${x1},${y1} C${mx},${y1} ${mx},${exchCenterY} ${COLS_X[2]},${exchCenterY}`}
                fill="none" stroke={D.amber} strokeWidth={Math.max(1,(ch.exchanged/(totalExchanged||1))*maxRetStroke)} opacity={0.22}/>}
            </g>
          );
        })}

        {/* 컬럼0: 상품 블록 */}
        {prods.map((p,i)=>{
          const y=yPos[i]; const h=prodH[i]; const col=D.SANKEY[i%D.SANKEY.length];
          const mid=h/2;
          const isSelected = sel?.type==="prod"&&sel.key===p.name;
          return (
            <g key={p.name} style={{cursor:"pointer"}} onClick={()=>setSel(isSelected?null:{type:"prod",key:p.name})}>
              <rect x={COLS_X[0]} y={y} width={NODE_W} height={h} rx={3} fill={col} opacity={isSelected?0.25:0.09}/>
              <rect x={COLS_X[0]} y={y} width={3} height={h} rx={1} fill={col}/>
              {isSelected&&<rect x={COLS_X[0]} y={y} width={NODE_W} height={h} rx={3} fill="none" stroke={col} strokeWidth={1.5}/>}
              <text x={COLS_X[0]+12} y={y+mid-(h>40?10:0)} dominantBaseline="middle"
                fill={D.black} fontSize={lblFs}>
                {p.name.length>22?p.name.slice(0,22)+"…":p.name}
              </text>
              {h>=40&&<text x={COLS_X[0]+12} y={y+mid+lblFs+2} dominantBaseline="middle"
                fill={D.textMeta} fontSize={subFs}>
                입고 {p.stock} · 배송 {p.shipped}
                {p.returned>0?` · 반품 ${p.returned}`:""}
                {p.exchanged>0?` · 교환 ${p.exchanged}`:""}
              </text>}
            </g>
          );
        })}

        {/* 컬럼1: 판매처 블록 */}
        {(()=>{
          let ry=PAD_T+16;
          return channels.map((ch,ci)=>{
            const h=Math.max(MIN_H,(ch.shipped/chanTotal)*col1H-ROW_GAP);
            const y=ry; ry+=h+ROW_GAP;
            const col=D.SANKEY[(ci+5)%D.SANKEY.length];
            chanYOf[ch.name]=y+h/2;
            const isSelected = sel?.type==="ch"&&sel.key===ch.name;
            return (
              <g key={ch.name} style={{cursor:"pointer"}} onClick={()=>setSel(isSelected?null:{type:"ch",key:ch.name})}>
                <rect x={COLS_X[1]} y={y} width={NODE_W} height={h} rx={4} fill={col} opacity={isSelected?0.28:0.12}/>
                <rect x={COLS_X[1]} y={y} width={4} height={h} rx={2} fill={col}/>
                {isSelected&&<rect x={COLS_X[1]} y={y} width={NODE_W} height={h} rx={4} fill="none" stroke={col} strokeWidth={1.5}/>}
                <text x={COLS_X[1]+12} y={y+h/2-(h>40?10:0)} dominantBaseline="middle"
                  fill={col} fontSize={lblFs}>{ch.name}</text>
                {h>=40&&<text x={COLS_X[1]+12} y={y+h/2+lblFs+2} dominantBaseline="middle"
                  fill={D.textMeta} fontSize={subFs}>{ch.shipped.toLocaleString()}건</text>}
              </g>
            );
          });
        })()}

        {/* 컬럼2: 반품 블록 */}
        {totalReturned>0&&(()=>{
          const isSelected = sel?.type==="ret";
          return (
            <g style={{cursor:"pointer"}} onClick={()=>setSel(isSelected?null:{type:"ret",key:"반품"})}>
              <rect x={COLS_X[2]} y={retBlockY} width={NODE_W} height={retBlockH} rx={4} fill={D.red} opacity={isSelected?0.22:0.1}/>
              <rect x={COLS_X[2]} y={retBlockY} width={4} height={retBlockH} rx={2} fill={D.red}/>
              {isSelected&&<rect x={COLS_X[2]} y={retBlockY} width={NODE_W} height={retBlockH} rx={4} fill="none" stroke={D.red} strokeWidth={1.5}/>}
              <text x={COLS_X[2]+12} y={retCenterY} dominantBaseline="middle"
                fill={D.red} fontSize={lblFs}>반품 {totalReturned}건</text>
            </g>
          );
        })()}

        {/* 컬럼2: 교환 블록 */}
        {totalExchanged>0&&(()=>{
          const isSelected = sel?.type==="exch";
          return (
            <g style={{cursor:"pointer"}} onClick={()=>setSel(isSelected?null:{type:"exch",key:"교환"})}>
              <rect x={COLS_X[2]} y={exchBlockY} width={NODE_W} height={exchBlockH} rx={4} fill={D.amber} opacity={isSelected?0.28:0.12}/>
              <rect x={COLS_X[2]} y={exchBlockY} width={4} height={exchBlockH} rx={2} fill={D.amber}/>
              {isSelected&&<rect x={COLS_X[2]} y={exchBlockY} width={NODE_W} height={exchBlockH} rx={4} fill="none" stroke={D.amber} strokeWidth={1.5}/>}
              <text x={COLS_X[2]+12} y={exchCenterY} dominantBaseline="middle"
                fill={D.amber} fontSize={lblFs}>교환 {totalExchanged}건</text>
            </g>
          );
        })()}
      </svg>

      {/* 선택 노드 세부 정보 패널 */}
      {sel&&(()=>{
        let title="", rows=[], note=null;
        if (sel.type==="prod") {
          const p = prods.find(x=>x.name===sel.key);
          if (!p) return null;
          const dt = prodDates[sel.key];
          title = p.name;
          note = dt ? `주문 데이터: ${dt.min} ~ ${dt.max}` : "주문 데이터 없음";
          rows = Object.entries(p.byChannel)
            .sort((a,b)=>b[1].shipped-a[1].shipped)
            .map(([ch,v])=>({ label:ch, cols:[`배송 ${v.shipped}`, v.returned?`반품 ${v.returned}`:"", v.exchanged?`교환 ${v.exchanged}`:""].filter(Boolean).join(" · ") }));
        } else if (sel.type==="ch") {
          const ch = channels.find(x=>x.name===sel.key);
          if (!ch) return null;
          title = ch.name;
          rows = Object.entries(ch.byProd||{})
            .sort((a,b)=>b[1].shipped-a[1].shipped)
            .slice(0,15)
            .map(([prod,v])=>({ label:prod, cols:[`배송 ${v.shipped}`, v.returned?`반품 ${v.returned}`:"", v.exchanged?`교환 ${v.exchanged}`:""].filter(Boolean).join(" · ") }));
        } else if (sel.type==="ret") {
          title = "반품 채널별 분포";
          rows = Object.entries(retByCh).sort((a,b)=>b[1]-a[1]).map(([ch,cnt])=>({ label:ch, cols:`${cnt}건 (${(cnt/totalReturned*100).toFixed(1)}%)` }));
        } else if (sel.type==="exch") {
          title = "교환 채널별 분포";
          rows = Object.entries(exchByCh).sort((a,b)=>b[1]-a[1]).map(([ch,cnt])=>({ label:ch, cols:`${cnt}건 (${(cnt/totalExchanged*100).toFixed(1)}%)` }));
        }
        return (
          <div style={{margin:"12px 0 0",padding:"14px 18px",background:D.surface,border:`1px solid ${D.border}`,borderRadius:10,fontSize:13}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <span style={{fontWeight:700,color:D.black,fontSize:14}}>{title}</span>
              {note&&<span style={{color:D.primary,fontSize:12}}>{note}</span>}
              <button onClick={()=>setSel(null)} style={{background:"none",border:"none",cursor:"pointer",color:D.textMeta,fontSize:16,lineHeight:1,padding:0}}>✕</button>
            </div>
            {rows.length===0&&<div style={{color:D.textMeta,fontSize:12}}>데이터 없음</div>}
            {rows.map((r,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:i<rows.length-1?`1px solid ${D.border}`:"none",gap:12}}>
                <span style={{color:D.textSub,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.label}</span>
                <span style={{color:D.black,whiteSpace:"nowrap"}}>{r.cols}</span>
              </div>
            ))}
          </div>
        );
      })()}
    </div>
  );
}

const getCSData=()=>{try{return JSON.parse(localStorage.getItem("cs_data")||"[]");}catch{return[];}};
const saveCSData=d=>localStorage.setItem("cs_data",JSON.stringify(d));
const getPromosCache=()=>{try{return JSON.parse(localStorage.getItem("promotions")||"[]").map(p=>({...p,files:p.files||(p.file?[p.file]:[]),file:undefined}));}catch{return[];}};
const setPromosCache=d=>localStorage.setItem("promotions",JSON.stringify(d));

// ─────────────────────────────────────────────
// ANALYTICS ENGINE
// ─────────────────────────────────────────────
function analyze(orderRows, stockRows, revenueRows, storeRows=[]) {
  // ═══════════════════════════════════════════════════════
  // 데이터 소스 가이드
  //   - orderRows   : 이지어드민 주문·배송 CSV (loadData에서 store_sales도 channel="오프라인 스토어"로 머지됨)
  //   - storeRows   : 매장 판매 CSV (StoreUploader) — 오프라인 전용 원본
  //   - revenueRows : 채널별 일자 매출 CSV (RevenueForm)
  //   - stockRows   : 입고 CSV
  // 매장(오프라인) 행은 '배송 KPI'에서 제외하고 별도 storeMetrics로 수집한다.
  // ═══════════════════════════════════════════════════════

  // 오프라인 채널 식별: 이지어드민 판교점/일산점 행, store_sales 머지 행 모두 포함
  const OFFLINE_CHS=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
  const isOffline = r => OFFLINE_CHS.has(r.channel||"");
  // MERRYON OVERSEA 채널은 판매처 상세/판매처별 매출/매출 점유율에서 제외
  const EXCL_CHS=new Set(["MERRYONOVERSEA","MERRYON OVERSEA","Merryon Oversea"]);
  const isExcl = r => EXCL_CHS.has(String(r.channel||"").trim());

  // ── [총 매출] ──────────────────────────────────────────
  // 소스: revenues(온라인 채널 일자 매출) + storeSales(매장 실판매금액)
  // 계산: SUM(amount - refund_amount)  +  매장 (배송 amount 합 - 반품 amount 합)
  const onlineRevenue   = revenueRows.reduce((s,r)=>s+(r.amount||0)-(r.refund_amount||0),0);
  const offlineRevenue  = storeRows.reduce((s,r)=>r.status==="배송"?s+(r.amount||0):r.status==="반품"?s-(r.amount||0):s,0);
  const totalRevenue    = onlineRevenue + offlineRevenue;

  // ── [매출 입력 기반 보조 지표] ────────────────────────
  // 소스: revenues CSV
  // 계산: order_count / refund_amount / refund_count 컬럼 단순 합산
  const totalOrderCount = revenueRows.reduce((s,r)=>s+(r.order_count||0),0);
  const totalRefundAmt  = revenueRows.reduce((s,r)=>s+(r.refund_amount||0),0);
  const totalRefundCount= revenueRows.reduce((s,r)=>s+(r.refund_count||0),0);

  // ── [배송·반품 KPI] (온라인만, 매장 제외) ─────────────
  // 소스: 이지어드민 orders CSV — 오프라인 채널(매장 판매 머지 행 포함)은 제외
  // 계산:
  //   - 배송 수    = COUNT(DISTINCT order_no||order_id) where status="배송"
  //   - 반품 수    = COUNT(DISTINCT order_no||order_id) where status="반품"
  //   - 배송 장수  = SUM(qty) where 배송
  //   - 반품 장수  = SUM(qty) where 반품
  //   - 반품률     = 반품 장수 / 배송 장수 * 100 (장수 기준 — 부분 반품/교환 반영)
  const onlineRows       = orderRows.filter(r=>!isOffline(r));
  const shippedRows      = onlineRows.filter(r=>r.status==="배송");
  const returnedRows     = onlineRows.filter(r=>r.status==="반품");
  const totalShipped     = new Set(shippedRows.map(r=>r.order_no||r.order_id).filter(Boolean)).size;
  const totalReturned    = new Set(returnedRows.map(r=>r.order_no||r.order_id).filter(Boolean)).size;
  const totalDeliveredQty= shippedRows.reduce((s,r)=>s+(r.qty||1),0);
  const totalReturnedQty = returnedRows.reduce((s,r)=>s+(r.qty||1),0);
  // returnRate는 storeQty/storeReturnedQty 정의 후 아래에서 매장 포함 통합 계산

  // ── [주문 KPI] (온라인 전체, 배송 완료 여부 무관) ─────
  // 소스: 이지어드민 orders CSV (매장 제외)
  // 계산:
  //   - 주문 수    = COUNT(DISTINCT order_no||order_id) 모든 상태 포함
  //   - 주문 장수  = SUM(qty) 모든 상태 포함
  const totalUniqueOrders = new Set(onlineRows.map(r=>r.order_no||r.order_id).filter(Boolean)).size;
  const totalOrderedQty   = onlineRows.reduce((s,r)=>s+(r.qty||1),0);

  // ── [매장 KPI] (store_sales CSV 별도 수집, 배송 카운트로 합산 안 함) ──
  // 소스: store_sales (StoreUploader 업로드, 기존 파싱 컬럼 유지)
  // 계산:
  //   - 매장 매출    = SUM(실판매금액)배송 - SUM(실판매금액)반품  ← offlineRevenue와 동일
  //   - 매장 주문 수 = COUNT(DISTINCT order_id) where 배송   (일자별 ID 집계의 유니크)
  //   - 매장 반품 수 = COUNT(DISTINCT order_id) where 반품
  //   - 판매 장수    = SUM(qty) where 배송
  //   - 반품 장수    = SUM(qty) where 반품
  //   - 객단가       = 매장 매출 / 매장 주문 수
  const storeShippedRows  = storeRows.filter(r=>r.status==="배송");
  const storeReturnedRows = storeRows.filter(r=>r.status==="반품");
  const storeRevenue      = offlineRevenue;
  const storeOrderCount   = new Set(storeShippedRows.map(r=>r.order_id).filter(Boolean)).size;
  const storeReturnedCount= new Set(storeReturnedRows.map(r=>r.order_id).filter(Boolean)).size;
  const storeQty          = storeShippedRows.reduce((s,r)=>s+(r.qty||1),0);
  const storeReturnedQty  = storeReturnedRows.reduce((s,r)=>s+(r.qty||1),0);
  const storeAOV          = storeOrderCount>0?Math.round(storeRevenue/storeOrderCount):0;
  const storeMetrics      = {storeRevenue,storeOrderCount,storeReturnedCount,storeQty,storeReturnedQty,storeAOV};

  // ── [반품률] (온라인 전용) ─────────────────────────────
  // 반품률 = 온라인 반품 qty ÷ 온라인 배송 qty × 100
  // 매장 배송 qty 는 분모에서 제외 (매장은 주문 수량 KPI 에만 합산)
  //   - 매장 반품도 집계 제외 (loadData 에서 status='반품' 필터링됨)
  //   - 노출용 두 합산 값(totalDeliveredQtyAll/totalReturnedQtyAll)은 매장 0 이라 totalDeliveredQty/totalReturnedQty 와 동일
  const totalDeliveredQtyAll = totalDeliveredQty + storeQty;       // 모달 등 노출용 (값은 변동 가능)
  const totalReturnedQtyAll  = totalReturnedQty  + storeReturnedQty;
  const returnRate           = totalDeliveredQty>0?(totalReturnedQty/totalDeliveredQty*100).toFixed(1):"0.0";

  // ── [재고] ──────────────────────────────────────────
  // 소스: stock_uploads — 입고 CSV 누적 (필터링은 호출부에서)
  const totalStock = stockRows.reduce((s,r)=>s+(r.qty||0),0);

  // ─────────────────────────────────────────────────────
  // 판매처별(채널별) 집계
  // 1) 매출/주문수/환불수: revenues CSV 단순 합산
  // 2) 배송/반품 고유 주문번호: 이지어드민 orders CSV에서 channel별 Set 누적
  // 3) 객단가 소스 금액 맵: 자사몰=payment_amount MAX, 29CM·무신사=sale_price SUM
  // 4) 오프라인 행(판교점/일산점/매장 머지)은 "오프라인 스토어"로 통합 후 storeMetrics로 덮어씀
  // ─────────────────────────────────────────────────────
  const byChannel={};
  // (1) 매출 입력 합산 — 채널별 매출/주문수/환불수 (MERRYON OVERSEA 제외)
  revenueRows.forEach(r=>{
    if(isExcl(r)) return;
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    byChannel[ch].revenue+=(r.amount||0)-(r.refund_amount||0);
    byChannel[ch].orderCount+=(r.order_count||0);
    byChannel[ch].refundCount+=(r.refund_count||0);
  });
  // (2),(3) 이지어드민 행 순회 — 배송/반품/전체 고유 주문번호 + 장수 + 객단가 금액 맵
  const chOrderAmt={};      // 채널별 {oid: amount} — 객단가 분자
  const chOrderIds={};      // 채널별 Set(oid) where 배송 — 객단가 분모 / 배송 카운트
  const chReturnedIds={};   // 채널별 Set(oid) where 반품 — 반품 카운트
  const chAllOrderIds={};   // 채널별 Set(oid) 모든 상태 — 판매처 상세의 '주문 수' 컬럼용
  const chOrderedQty={};    // 채널별 SUM(qty) 모든 상태 — '주문 장수'
  const chShippedQty={};    // 채널별 SUM(qty) 배송   — '배송 장수' + 반품률 분모
  const chReturnedQty={};   // 채널별 SUM(qty) 반품   — '반품 장수' + 반품률 분자
  const PAYMENT_CH=new Set(["자사몰"]); // payment_amount(MAX) 사용 채널
  orderRows.forEach(r=>{
    if(isExcl(r)) return; // MERRYON OVERSEA 제외
    const ch=r.channel||"미분류";
    // 주문번호 키: 신규 order_no 필드 우선, 없으면 order_id 전체(이전 데이터 호환)
    const oid=r.order_no||r.order_id||"";
    // CORD prefix = 29CM 취소 주문 → 반품으로 강제 (실제 컨디션 파악 시 재조정)
    const status=(r.status==="배송"&&/^CORD/i.test(oid))?"반품":r.status;
    const qty=(r.qty||1);
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    // 전체 주문 Set: 모든 상태(배송/반품/교환 등) 포함
    if(oid){
      if(!chAllOrderIds[ch]) chAllOrderIds[ch]=new Set();
      chAllOrderIds[ch].add(oid);
    }
    chOrderedQty[ch]=(chOrderedQty[ch]||0)+qty;
    if(status==="반품"){
      if(oid){
        if(!chReturnedIds[ch]) chReturnedIds[ch]=new Set();
        chReturnedIds[ch].add(oid);
      }
      chReturnedQty[ch]=(chReturnedQty[ch]||0)+qty;
    }
    if(status!=="배송") return;
    if(!chOrderIds[ch]) chOrderIds[ch]=new Set();
    chOrderIds[ch].add(oid);
    chShippedQty[ch]=(chShippedQty[ch]||0)+qty;
    if(!chOrderAmt[ch]) chOrderAmt[ch]={};
    if(PAYMENT_CH.has(ch)){
      // 자사몰: 결제금액은 동일 주문의 각 행에 중복 기록 → 덮어쓰기(= 사실상 MAX)
      const pa=r.payment_amount||r.amount||0;
      if(pa>0) chOrderAmt[ch][oid]=pa;
    } else {
      // 29CM·무신사: 판매가는 상품별로 다름 → 같은 주문 안에서 상품별 합산
      const sp=r.sale_price||r.amount||0;
      chOrderAmt[ch][oid]=(chOrderAmt[ch][oid]||0)+sp;
    }
  });
  // 채널 shipped/returned는 라인 수 X → 고유 주문번호 수로 채움
  Object.keys(byChannel).forEach(ch=>{
    byChannel[ch].shipped  = (chOrderIds[ch]||new Set()).size;
    byChannel[ch].returned = (chReturnedIds[ch]||new Set()).size;
  });

  // (4-a) 오프라인 채널(판교점/일산점/매장 머지 행) → "오프라인 스토어"로 통합
  // 단, 이 단계에서 만든 값들은 storeRows가 있으면 (4-b)에서 storeMetrics로 덮어씀
  const offlineAgg={name:"오프라인 스토어",revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
  const offlineOrderIds=new Set();
  const offlineReturnedIds=new Set();
  const offlineAllOrderIds=new Set();
  let offlineOrderedQty=0, offlineShippedQty=0, offlineReturnedQty=0;
  const offlineBreakdown={};  // 판교점/일산점 별도 보존(채널별 표 sub-row용)
  let hasOffline=false;
  Object.keys(byChannel).forEach(ch=>{
    if(OFFLINE_CHS.has(ch)){
      hasOffline=true;
      offlineAgg.revenue+=byChannel[ch].revenue;
      offlineAgg.orderCount+=byChannel[ch].orderCount;
      offlineAgg.refundCount+=byChannel[ch].refundCount;
      (chOrderIds[ch]||new Set()).forEach(id=>offlineOrderIds.add(id));
      (chReturnedIds[ch]||new Set()).forEach(id=>offlineReturnedIds.add(id));
      (chAllOrderIds[ch]||new Set()).forEach(id=>offlineAllOrderIds.add(id));
      offlineOrderedQty +=(chOrderedQty[ch] ||0);
      offlineShippedQty +=(chShippedQty[ch] ||0);
      offlineReturnedQty+=(chReturnedQty[ch]||0);
      if(ch!=="오프라인스토어"&&ch!=="오프라인"&&ch!=="오프라인 스토어")
        offlineBreakdown[ch]={...byChannel[ch]};
      delete byChannel[ch];
      delete chOrderIds[ch];
      delete chReturnedIds[ch];
      delete chAllOrderIds[ch];
      delete chOrderedQty[ch];
      delete chShippedQty[ch];
      delete chReturnedQty[ch];
    }
  });
  if(hasOffline){
    offlineAgg.shipped  = offlineOrderIds.size;
    offlineAgg.returned = offlineReturnedIds.size;
    byChannel["오프라인 스토어"]=offlineAgg;
    chOrderIds["오프라인 스토어"]=offlineOrderIds;
    chReturnedIds["오프라인 스토어"]=offlineReturnedIds;
    chAllOrderIds["오프라인 스토어"]=offlineAllOrderIds;
    chOrderedQty["오프라인 스토어"] =offlineOrderedQty;
    chShippedQty["오프라인 스토어"] =offlineShippedQty;
    chReturnedQty["오프라인 스토어"]=offlineReturnedQty;
  }
  // (4-b) 매장 판매 CSV가 있으면 storeMetrics로 오프라인 스토어 행 덮어쓰기
  // 이지어드민 머지 데이터보다 store_sales 원본이 권위 있음
  if(storeRows.length){
    if(!byChannel["오프라인 스토어"]) byChannel["오프라인 스토어"]={name:"오프라인 스토어",revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    byChannel["오프라인 스토어"].revenue  = storeRevenue;       // 실판매금액 합 (배송 - 반품)
    // 매장은 배송/반품 카운트에 흘러들어가지 않음 — 모두 0으로 고정
    byChannel["오프라인 스토어"].shipped  = 0;
    byChannel["오프라인 스토어"].returned = 0;
    // 객단가 분모(c.uniqueOrders)는 그대로 매장 배송 order_id로 사용 (storeAOV 계산용)
    chOrderIds["오프라인 스토어"] = new Set(storeShippedRows.map(r=>r.order_id).filter(Boolean));
    // 주문 수 집합: 매장 전체 order_id (배송+반품)로 덮어쓰기 — 매장 데이터가 권위
    chAllOrderIds["오프라인 스토어"] = new Set(storeRows.map(r=>r.order_id).filter(Boolean));
    // 주문 장수는 매장 권위 데이터(배송+반품 qty 합), 배송/반품 장수는 0 (배송과 무관)
    chOrderedQty["오프라인 스토어"]  = storeQty + storeReturnedQty;
    chShippedQty["오프라인 스토어"]  = 0;
    chReturnedQty["오프라인 스토어"] = 0;
    // 매장별(판교점/일산점) 매출 breakdown
    const storeByStore={};
    storeRows.forEach(r=>{
      const st=r.store_name||"오프라인 스토어";
      if(!storeByStore[st]) storeByStore[st]={revenue:0};
      if(r.status==="배송") storeByStore[st].revenue+=(r.amount||0);
      else if(r.status==="반품") storeByStore[st].revenue-=(r.amount||0);
    });
    Object.entries(storeByStore).forEach(([st,d])=>{
      if(st!=="오프라인 스토어") offlineBreakdown[st]={name:st,revenue:d.revenue};
    });
  }
  // ── [채널별 정렬 + 점유율/반품률/객단가] ──────────────
  // 정렬: 매출 desc, 동률이면 배송 수 desc
  // share         = 채널 매출 / 전체 채널 매출 합 * 100
  // returnRate    = 채널 반품(주문수) / 채널 배송(주문수) * 100
  // uniqueOrders  = chOrderIds[ch].size (오프라인은 매장 order_id, 온라인은 배송 oid)
  // avgOrderValue:
  //   - 오프라인 스토어: 매장 매출(storeRevenue) / 매장 주문수(uq)
  //   - 온라인: 채널 chOrderAmt 합 / 주문번호 수 (revenues CSV가 아닌 orders CSV에서 직접 계산)
  const channelList=Object.values(byChannel).sort((a,b)=>b.revenue-a.revenue||b.shipped-a.shipped);
  const totalRev=channelList.reduce((s,c)=>s+c.revenue,0)||1;
  channelList.forEach(c=>{
    c.share=((c.revenue||0)/totalRev*100).toFixed(1);
    // 장수 트래커 → 채널 객체에 노출
    c.orderedQty  = chOrderedQty[c.name] ||0;  // 주문 장수
    c.shippedQty  = chShippedQty[c.name] ||0;  // 배송 장수
    c.returnedQty = chReturnedQty[c.name]||0;  // 반품 장수
    // 반품률: 장수 기준 (부분 반품/교환 반영)
    c.returnRate  = c.shippedQty>0?(c.returnedQty/c.shippedQty*100).toFixed(1):"0.0";
    const uq=(chOrderIds[c.name]||new Set()).size||c.shipped;
    c.uniqueOrders=uq;
    // totalOrders: 모든 상태 포함 고유 주문번호 수 (판매처 상세의 '주문 수' 컬럼)
    c.totalOrders=(chAllOrderIds[c.name]||new Set()).size||uq;
    if(c.name==="오프라인 스토어"){
      c.avgOrderValue=storeAOV;  // storeRevenue / storeOrderCount
    } else {
      const orderMap=chOrderAmt[c.name]||{};
      const totalAmt=Object.values(orderMap).reduce((s,a)=>s+a,0);
      const orderCount=Object.keys(orderMap).length;
      c.avgOrderValue=(orderCount>0&&totalAmt>0)?Math.round(totalAmt/orderCount):0;
    }
  });

  // ── [월별 배송/반품 추이] ─────────────────────────────
  // 소스: 이지어드민 orders CSV (매장 포함 — 차트에서 라인이 채널별로 분리됨)
  // 계산: order_date의 YYYY-MM 단위 그룹 → 배송/반품 라인 카운트, 반품률
  // (라인 카운트 유지: 차트는 추세 시각화 목적이므로 행 단위로 충분)
  const byMonth={};
  orderRows.forEach(r=>{
    const ym=r.order_date?r.order_date.slice(0,7):null;
    if(!ym) return;
    if(!byMonth[ym]) byMonth[ym]={month:ym,shipped:0,returned:0};
    if(r.status==="배송") byMonth[ym].shipped++;
    if(r.status==="반품") byMonth[ym].returned++;
  });
  const monthlyData=Object.values(byMonth)
    .sort((a,b)=>a.month>b.month?1:-1)
    .map(m=>({...m,returnRate:m.shipped>0?(m.returned/m.shipped*100).toFixed(1):"0.0"}));

  // ── [판매·반품 Top — 주간 상품 랭킹] ──────────────────
  // 소스: 이지어드민 orders CSV + store_sales(loadData 머지) — 매장 상품도 포함
  // 계산:
  //   - 최신 주(latestWeek) 기준 행만 추림
  //   - 상품명 기준 그룹화(옵션 합산):
  //       qty      = SUM(qty)               ← 주문 기준(반품·취소 제외)
  //       orders   = COUNT(행)
  //       returned = COUNT(상태=반품 행)    ← 반품 Top 정렬용 (별도 집계)
  //   - 판매 Top = qty desc 상위 20 (주문 기준)
  //   - 반품 Top = returned>0 중 returned desc 상위 20
  const getWeek=ds=>{
    if(!ds) return null;
    const dt=new Date(ds); if(isNaN(dt)) return null;
    const tmp=new Date(dt);
    tmp.setDate(tmp.getDate()+3-((tmp.getDay()+6)%7));
    const w1=new Date(tmp.getFullYear(),0,4);
    const wn=1+Math.round(((tmp-w1)/86400000-3+((w1.getDay()+6)%7))/7);
    return `${dt.getFullYear()}-W${String(wn).padStart(2,"0")}`;
  };
  const weeks=[...new Set(orderRows.filter(r=>r.order_date).map(r=>getWeek(r.order_date)).filter(Boolean))].sort();
  const latestWeek=weeks[weeks.length-1];
  const weekRows=latestWeek?orderRows.filter(r=>getWeek(r.order_date)===latestWeek):orderRows;

  const byProd={};
  weekRows.forEach(r=>{
    const key=r.product_name||"미분류";
    if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
    // 반품 카운트는 모든 행에서 집계 (반품 Top 용)
    if(r.status==="반품") byProd[key].returned++;
    // 판매 Top 의 qty/orders 는 주문 기준 — 반품·취소 제외
    if(r.status==="반품"||r.status==="취소") return;
    byProd[key].qty+=(r.qty||0);
    byProd[key].orders++;
  });
  const prodList=Object.values(byProd);
  const weekBest=[...prodList].sort((a,b)=>b.qty-a.qty).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));
  const weekWorst=[...prodList].filter(p=>p.returned>0).sort((a,b)=>b.returned-a.returned).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));

  // ── [전체 합산 KPI (온라인 + 매장)] ────────────────────
  // KPI 카드의 '주문 수' / '주문 장수' 표기용
  // 배송 KPI는 매장 제외(totalShipped/totalDeliveredQty)지만,
  // 주문 KPI는 사용자 요청으로 매장까지 합쳐서 단일 숫자로 표시
  const totalUniqueOrdersAll = totalUniqueOrders + storeOrderCount;
  const totalOrderedQtyAll   = totalOrderedQty   + storeQty;

  return {
    totalRevenue,totalOrderCount,totalRefundAmt,totalRefundCount,returnRate,
    totalShipped,totalReturned,totalStock,
    totalUniqueOrders,totalOrderedQty,totalDeliveredQty,totalReturnedQty,
    totalUniqueOrdersAll,totalOrderedQtyAll,                 // 매장 포함 합산 (주문)
    totalDeliveredQtyAll,totalReturnedQtyAll,                // 매장 포함 합산 (반품률 분모/분자)
    storeMetrics,  // 매장 별도 KPI: revenue/orderCount/returnedCount/qty/returnedQty/aov
    channelList,offlineBreakdown,monthlyData,weekBest,weekWorst,latestWeek,weekRows,
    chOrderAmt,chAllOrderIds, // 객단가/주문수 모달 소스
  };
}

// ─────────────────────────────────────────────
// SUPABASE CLIENT
// ─────────────────────────────────────────────
const SUPA_URL = typeof import.meta!=="undefined"&&import.meta.env?.VITE_SUPABASE_URL||"";
const SUPA_KEY = typeof import.meta!=="undefined"&&import.meta.env?.VITE_SUPABASE_ANON_KEY||"";
let _supabase=null;
async function getSupabase() {
  if(_supabase) return _supabase;
  if(!SUPA_URL||!SUPA_KEY){
    _supabase={
      from:()=>({
        select:()=>({order:()=>({limit:()=>Promise.resolve({data:[],error:null}),ascending:false}),
          gte:()=>({lte:()=>({order:()=>Promise.resolve({data:[],error:null})})}),
          in:()=>Promise.resolve({data:[],error:null}),
        }),
        insert:rows=>Promise.resolve({data:rows,error:null}),
        upsert:(rows,o)=>Promise.resolve({data:rows,error:null}),
        update:d=>({eq:()=>Promise.resolve({error:null})}),
        delete:()=>({gte:()=>({lte:()=>Promise.resolve({error:null})}),eq:()=>Promise.resolve({error:null})}),
      }),
    };
    return _supabase;
  }
  const { createClient }=await import("@supabase/supabase-js");
  _supabase=createClient(SUPA_URL,SUPA_KEY);
  return _supabase;
}

// ─────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────
const PERIOD_TABS=[
  {key:"yd",label:"어제"},
  {key:"7d",label:"최근 7일"},
  {key:"1m",label:"최근 한달"},
  {key:"3m",label:"최근 3개월"},
  {key:"all",label:"전체"},
  {key:"custom",label:"기간 선택"},
];

const CH_ORDER=["자사몰","29CM","무신사","오프라인 스토어"];
const chRank=ch=>{const i=CH_ORDER.indexOf(ch);return i>=0?i:99;};

function getPriorPeriod(period,customStart,customEnd){
  const fmt=d=>d.toISOString().slice(0,10);
  if(period==="yd"){
    return{start:localDate(-2),end:localDate(-2)};
  }
  if(period==="7d"){
    const e=new Date();e.setDate(e.getDate()-7);
    const s=new Date();s.setDate(s.getDate()-14);
    return{start:fmt(s),end:fmt(e)};
  }
  if(period==="14d"){
    const e=new Date();e.setDate(e.getDate()-14);
    const s=new Date();s.setDate(s.getDate()-28);
    return{start:fmt(s),end:fmt(e)};
  }
  if(period==="1m"){
    const e=new Date();e.setMonth(e.getMonth()-1);
    const s=new Date();s.setMonth(s.getMonth()-2);
    return{start:fmt(s),end:fmt(e)};
  }
  if(period==="3m"){
    const e=new Date();e.setMonth(e.getMonth()-3);
    const s=new Date();s.setMonth(s.getMonth()-6);
    return{start:fmt(s),end:fmt(e)};
  }
  if(period==="6m"){
    const e=new Date();e.setMonth(e.getMonth()-6);
    const s=new Date();s.setMonth(s.getMonth()-12);
    return{start:fmt(s),end:fmt(e)};
  }
  if(period==="custom"&&customStart&&customEnd){
    const s0=new Date(customStart),e0=new Date(customEnd);
    const days=Math.round((e0-s0)/(864e5));
    const pe=new Date(s0);pe.setDate(pe.getDate()-1);
    const ps=new Date(pe);ps.setDate(ps.getDate()-days);
    return{start:fmt(ps),end:fmt(pe)};
  }
  return null;
}

// Compact calendar range dropdown — used in PromoFlow, DataCompare, CSData
function CalRangeDrop({id,start,end,onRange,openId,setOpenId,surface,borderColor,textActive,textInactive}){
  const isOpen=openId===id;
  const label=start&&end?`${start.slice(5)}~${end.slice(5)}`:"기간 선택";
  const active=!!(start&&end);
  const bg=surface||D.surface;
  const bc=borderColor||D.border;
  const ta=textActive||D.black;
  const ti=textInactive||D.textSub;
  return(
    <div style={{position:"relative",display:"inline-flex",gap:4,alignItems:"center"}}>
      <button onClick={()=>setOpenId(isOpen?null:id)}
        style={{background:active?ta:"transparent",color:active?"#fff":ti,
          border:`1px solid ${active?ta:bc}`,borderRadius:5,
          padding:"3px 8px",fontSize:10,cursor:"pointer",fontWeight:active?600:400}}>
        {label}
      </button>
      {active&&(
        <button onClick={()=>onRange("","")}
          style={{background:"none",border:"none",color:ti,cursor:"pointer",fontSize:12,padding:"0 2px",lineHeight:1}}>✕</button>
      )}
      {isOpen&&(
        <div style={{position:"absolute",right:0,top:"calc(100% + 6px)",zIndex:300,
          background:bg,border:`1px solid ${bc}`,borderRadius:10,
          padding:"14px 14px 10px",boxShadow:"0 4px 24px rgba(0,0,0,0.25)"}}>
          <CalendarPicker mode="range" rangeStart={start} rangeEnd={end}
            onRangeChange={({start:s,end:e})=>{onRange(s,e||"");if(s&&e)setOpenId(null);}}/>
        </div>
      )}
    </div>
  );
}

// CalDrop must be defined at module level (not inside Dashboard) so React
// preserves its identity across re-renders — prevents CalendarPicker's internal
// `picking` state from resetting between first and second date clicks.
function CalDrop({id,period,setPeriod,presets,start,setStart,end,setEnd,calOpenFor,setCalOpenFor,dark}){
  const isOpen=calOpenFor===id;
  const customLabel=period==="custom"&&start&&end?`${start.slice(5)}~${end.slice(5)}`:"직접 선택";
  const aC=dark?"rgba(240,237,232,0.9)":D.black;
  const aTxt=dark?"#111":"#fff";
  const iTxt=dark?"rgba(240,237,232,0.55)":D.textSub;
  const bd=dark?"rgba(240,237,232,0.22)":D.border;
  return(
    <div style={{position:"relative",display:"inline-flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
      {presets.map(([v,l])=>(
        <button key={v} data-hf onClick={()=>{setPeriod(v);setCalOpenFor(null);}}
          style={{background:period===v?aC:"transparent",color:period===v?aTxt:iTxt,
            border:`1px solid ${period===v?aC:bd}`,borderRadius:5,padding:"3px 8px",fontSize:10,cursor:"pointer",fontWeight:period===v?600:400}}>
          {l}
        </button>
      ))}
      <button data-hf onClick={()=>{setPeriod("custom");setCalOpenFor(isOpen?null:id);}}
        style={{background:period==="custom"?aC:"transparent",color:period==="custom"?aTxt:iTxt,
          border:`1px solid ${period==="custom"?aC:bd}`,borderRadius:5,padding:"3px 8px",fontSize:10,cursor:"pointer",fontWeight:period==="custom"?600:400}}>
        {customLabel}
      </button>
      {isOpen&&(
        <div style={{position:"absolute",right:0,top:"calc(100% + 6px)",zIndex:300,
          background:D.surface,border:`1px solid ${D.border}`,borderRadius:10,padding:"14px 14px 10px",
          boxShadow:"0 4px 24px rgba(0,0,0,0.4)"}}>
          <CalendarPicker mode="range" rangeStart={start} rangeEnd={end}
            onRangeChange={({start:s,end:e})=>{setStart(s);setEnd(e);if(s&&e)setCalOpenFor(null);}}/>
        </div>
      )}
    </div>
  );
}
function DateDrop({id,value,onChange,calOpenFor,setCalOpenFor,placeholder="날짜 선택"}){
  const isOpen=calOpenFor===id;
  return(
    <div style={{position:"relative",display:"inline-block"}}>
      <button data-hf onClick={()=>setCalOpenFor(isOpen?null:id)}
        style={{background:value?D.black:"transparent",color:value?"#fff":D.textSub,
          border:`1px solid ${value?D.black:D.border}`,borderRadius:5,padding:"3px 8px",fontSize:10,cursor:"pointer",fontWeight:value?600:400}}>
        {value?value.slice(5):placeholder}
      </button>
      {isOpen&&(
        <div style={{position:"absolute",left:0,top:"calc(100% + 6px)",zIndex:300,
          background:D.surface,border:`1px solid ${D.border}`,borderRadius:10,padding:"14px",
          boxShadow:"0 4px 24px rgba(0,0,0,0.4)"}}>
          <CalendarPicker mode="single" value={value} onChange={v=>{onChange(v);setCalOpenFor(null);}}/>
        </div>
      )}
    </div>
  );
}

function Dashboard({ orders, stocks, revenues, storeSales=[], ts, onRefresh }) {
  const isMobile=useWindowWidth()<=1080;
  const [period,setPeriod]=useState("1m");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [deleteAll,setDeleteAll]=useState(false);
  const [shippingPeriod,setShippingPeriod]=useState("1m");
  const [returnPeriod,setReturnPeriod]=useState("1m");
  const [rankBestPeriod,setRankBestPeriod]=useState("7d");
  const [rankBestChannel,setRankBestChannel]=useState("전체");
  const [rankBestCustomStart,setRankBestCustomStart]=useState("");
  const [rankBestCustomEnd,setRankBestCustomEnd]=useState("");
  const [rankWorstPeriod,setRankWorstPeriod]=useState("1m");
  const [rankWorstChannel,setRankWorstChannel]=useState("전체");
  const [rankWorstCustomStart,setRankWorstCustomStart]=useState("");
  const [rankWorstCustomEnd,setRankWorstCustomEnd]=useState("");
  const [chSort,setChSort]=useState({key:"revenue",dir:"desc"});
  const [optionPeriod,setOptionPeriod]=useState("1m");
  const [returnOptionPeriod,setReturnOptionPeriod]=useState("1m");
  const [shippingCustomStart,setShippingCustomStart]=useState("");
  const [shippingCustomEnd,setShippingCustomEnd]=useState("");
  const [returnCustomStart,setReturnCustomStart]=useState("");
  const [returnCustomEnd,setReturnCustomEnd]=useState("");
  const [optionCustomStart,setOptionCustomStart]=useState("");
  const [optionCustomEnd,setOptionCustomEnd]=useState("");
  const [returnOptionCustomStart,setReturnOptionCustomStart]=useState("");
  const [returnOptionCustomEnd,setReturnOptionCustomEnd]=useState("");
  const [calOpenFor,setCalOpenFor]=useState(null);
  const [offlineExpanded,setOfflineExpanded]=useState(false);
  const [kpiModal,setKpiModal]=useState(null); // "revenue"|"order"|"shipped"|"returnRate"|"stock"
  const [aovModal,setAovModal]=useState(null); // channel name
  const [chOrderModal,setChOrderModal]=useState(null); // channel name — 주문 수 소스 모달
  const [delPeriod,setDelPeriod]=useState("all");
  const [delStart,setDelStart]=useState("");
  const [delEnd,setDelEnd]=useState("");
  const [delCalOpen,setDelCalOpen]=useState(null);
  const salesByChCardRef=useRef(null);
  const chDetailCardRef=useRef(null);
  const shippingCardRef=useRef(null);
  const returnTrendCardRef=useRef(null);
  const salesTopCardRef=useRef(null);
  const optionPrefCardRef=useRef(null);
  const returnTopCardRef=useRef(null);
  const returnOptCardRef=useRef(null);

  const axTick={fill:D.textMeta,fontSize:10};
  const NoWrapTick=({x,y,payload})=>(
    <text x={x} y={y} dy={4} textAnchor="end" fill={D.textMeta} fontSize={10} style={{whiteSpace:"nowrap"}}>
      {payload.value?.length>22?payload.value.slice(0,22)+"…":payload.value}
    </text>
  );

  const filteredOrders=useMemo(()=>filterByDate(orders,"order_date",period,customStart,customEnd),[orders,period,customStart,customEnd]);
  const filteredRevenues=useMemo(()=>filterByDate(revenues,"date",period,customStart,customEnd),[revenues,period,customStart,customEnd]);
  const filteredStoreSales=useMemo(()=>filterByDate(storeSales,"sale_date",period,customStart,customEnd),[storeSales,period,customStart,customEnd]);
  const filteredStocks=useMemo(()=>{
    const periodRows=filterByDate(stocks,"upload_date",period,customStart,customEnd);
    const latest={};
    periodRows.forEach(r=>{
      const key=(r.product_name||"")+"__"+(r.option_name||"");
      if(!latest[key]||r.upload_date>latest[key].upload_date) latest[key]=r;
    });
    return Object.values(latest);
  },[stocks,period,customStart,customEnd]);
  const stats=useMemo(()=>analyze(filteredOrders,filteredStocks,filteredRevenues,filteredStoreSales),[filteredOrders,filteredStocks,filteredRevenues,filteredStoreSales]);

  // 직전 동일 기간 채널별 순매출
  const prevPeriod=useMemo(()=>getPriorPeriod(period,customStart,customEnd),[period,customStart,customEnd]);
  const prevChRevenue=useMemo(()=>{
    if(!prevPeriod) return {};
    const map={};
    const OFFL=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
    revenues.filter(r=>r.date>=prevPeriod.start&&r.date<=prevPeriod.end).forEach(r=>{
      const ch=OFFL.has(r.channel||"")?"오프라인 스토어":(r.channel||"미분류");
      map[ch]=(map[ch]||0)+(r.amount||0)-(r.refund_amount||0);
    });
    // store_sales 기반 오프라인 스토어 동기간 매출 합산
    storeSales.filter(r=>r.sale_date>=prevPeriod.start&&r.sale_date<=prevPeriod.end).forEach(r=>{
      const amt=r.amount||0;
      map["오프라인 스토어"]=(map["오프라인 스토어"]||0)+(r.status==="배송"?amt:r.status==="반품"?-amt:0);
    });
    return map;
  },[revenues,storeSales,prevPeriod]);

  // 플랫폼별 선호 옵션 (컬러/사이즈) — 독립 기간 필터
  const optionFilteredOrders=useMemo(()=>filterByDate(orders,"order_date",optionPeriod,optionCustomStart,optionCustomEnd),[orders,optionPeriod,optionCustomStart,optionCustomEnd]);
  const optionStats=useMemo(()=>{
    const map={};
    optionFilteredOrders.filter(r=>r.status==="배송").forEach(r=>{
      const ch=r.channel||"미분류";
      if(!map[ch]) map[ch]={colors:{},sizes:{}};
      const {color,size}=parseOption(r.product_name,r.option_name);
      if(color) map[ch].colors[color]=(map[ch].colors[color]||0)+1;
      if(size)  map[ch].sizes[size] =(map[ch].sizes[size] ||0)+1;
    });
    return Object.entries(map).map(([ch,d])=>({
      ch,
      colors:Object.entries(d.colors).sort((a,b)=>b[1]-a[1]).slice(0,7),
      sizes: Object.entries(d.sizes ).sort((a,b)=>b[1]-a[1]).slice(0,7),
    })).sort((a,b)=>chRank(a.ch)-chRank(b.ch));
  },[optionFilteredOrders,chRank]);

  // 플랫폼별 반품률 높은 옵션 — 독립 기간 필터
  const returnOptionFilteredOrders=useMemo(()=>filterByDate(orders,"order_date",returnOptionPeriod,returnOptionCustomStart,returnOptionCustomEnd),[orders,returnOptionPeriod,returnOptionCustomStart,returnOptionCustomEnd]);
  const returnOptionStats=useMemo(()=>{
    const map={};
    returnOptionFilteredOrders.forEach(r=>{
      const ch=r.channel||"미분류";
      if(!map[ch]) map[ch]={};
      const {color,size}=parseOption(r.product_name,r.option_name);
      const isShipped=r.status==="배송";
      const isReturned=r.status==="반품";
      if(!isShipped&&!isReturned) return;
      [[color,"c"],[size,"s"]].forEach(([val,type])=>{
        if(!val) return;
        const key=`${type}:${val}`;
        if(!map[ch][key]) map[ch][key]={name:val,type,shipped:0,returned:0};
        if(isShipped) map[ch][key].shipped++;
        if(isReturned) map[ch][key].returned++;
      });
    });
    return Object.entries(map).map(([ch,opts])=>{
      const all=Object.values(opts).filter(o=>(o.shipped+o.returned)>=3);
      const toRows=type=>all
        .filter(o=>o.type===type)
        .map(o=>({name:o.name,rate:o.returned/(o.shipped+o.returned)*100,returned:o.returned,total:o.shipped+o.returned}))
        .sort((a,b)=>b.rate-a.rate).slice(0,5);
      return {ch,colors:toRows("c"),sizes:toRows("s")};
    }).filter(d=>d.colors.length>0||d.sizes.length>0)
      .sort((a,b)=>chRank(a.ch)-chRank(b.ch));
  },[returnOptionFilteredOrders,chRank]);

  // 판매처 채널 목록 (전체 orders 기준, 오프라인스토어 제외)
  const activeChannels=useMemo(()=>{
    const dynamic=["자사몰","29CM","무신사"];
    const inData=new Set(orders.map(r=>r.channel||"미분류").filter(Boolean));
    return [...dynamic.filter(c=>inData.has(c)),"오프라인 스토어"];
  },[orders]);

  // 판매 Top 랭킹
  const bestFilteredOrders=useMemo(()=>
    filterByDate(orders,"order_date",rankBestPeriod,rankBestCustomStart,rankBestCustomEnd,true),
    [rankBestPeriod,orders,rankBestCustomStart,rankBestCustomEnd]);

  const bestRows=useMemo(()=>{
    const base=bestFilteredOrders;
    const rows=rankBestChannel==="전체"?base:base.filter(r=>r.channel===rankBestChannel);
    const byProd={};
    rows.forEach(r=>{
      const key=r.product_name||"미분류";
      if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
      // 반품 카운트는 모든 행에서 집계 (반품률용)
      if(r.status==="반품") byProd[key].returned++;
      // 판매 Top qty/orders 는 주문 기준 — 반품·취소 제외
      if(r.status==="반품"||r.status==="취소") return;
      byProd[key].qty+=(r.qty||0); byProd[key].orders++;
    });
    const totalQty=Object.values(byProd).reduce((s,p)=>s+p.qty,0)||1;
    return Object.values(byProd).sort((a,b)=>b.qty-a.qty).slice(0,20)
      .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0",
        share:(p.qty/totalQty*100).toFixed(1)}));
  },[bestFilteredOrders,rankBestChannel]);

  // 반품 Top 랭킹
  const worstFilteredOrders=useMemo(()=>
    filterByDate(orders,"order_date",rankWorstPeriod,rankWorstCustomStart,rankWorstCustomEnd,true),
    [rankWorstPeriod,orders,rankWorstCustomStart,rankWorstCustomEnd]);

  const worstRows=useMemo(()=>{
    const base=worstFilteredOrders;
    const rows=rankWorstChannel==="전체"?base:base.filter(r=>r.channel===rankWorstChannel);
    const byProd={};
    rows.forEach(r=>{
      const key=r.product_name||"미분류";
      if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,shipped:0,returned:0};
      byProd[key].qty+=(r.qty||0); byProd[key].orders++;
      if(r.status==="배송") byProd[key].shipped++;
      if(r.status==="반품") byProd[key].returned++;
    });
    const csData=getCSData();
    const csMap={};
    csData.forEach(r=>{
      if(!csMap[r.product_name])csMap[r.product_name]={};
      csMap[r.product_name][r.return_reason]=(csMap[r.product_name][r.return_reason]||0)+1;
    });
    const topReason=name=>{
      const m=csMap[name];
      if(!m)return"-";
      return Object.entries(m).sort((a,b)=>b[1]-a[1])[0]?.[0]||"-";
    };
    return Object.values(byProd).filter(p=>p.returned>0&&p.shipped>=3)
      .sort((a,b)=>(b.returned/b.orders)-(a.returned/a.orders)).slice(0,20)
      .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0",
        topReason:topReason(p.name)}));
  },[worstFilteredOrders,rankWorstChannel]);

  // 월별 배송량 차트 데이터
  const shippingChartData=useMemo(()=>{
    const today=localDate(-1);
    if(shippingPeriod==="yd"){
      const yStr=localDate(-1);
      const byDay={};
      orders.filter(r=>r.order_date===yStr).forEach(r=>{
        if(!byDay[yStr]) byDay[yStr]={date:yStr.slice(5),shipped:0};
        if(r.status==="배송") byDay[yStr].shipped++;
      });
      return Object.values(byDay);
    }
    if(shippingPeriod==="7d"||shippingPeriod==="1m"){
      const c=new Date();
      if(shippingPeriod==="7d") c.setDate(c.getDate()-7);
      else c.setMonth(c.getMonth()-1);
      const cut=c.toISOString().slice(0,10);
      const byDay={};
      orders.filter(r=>r.order_date>=cut&&r.order_date<=today).forEach(r=>{
        const d=r.order_date;
        if(!byDay[d]) byDay[d]={date:d.slice(5),shipped:0};
        if(r.status==="배송") byDay[d].shipped++;
      });
      return Object.values(byDay).sort((a,b)=>a.date>b.date?1:-1);
    }
    if(shippingPeriod==="custom"&&shippingCustomStart&&shippingCustomEnd){
      const diff=(new Date(shippingCustomEnd)-new Date(shippingCustomStart))/86400000;
      const src=orders.filter(r=>r.order_date>=shippingCustomStart&&r.order_date<=shippingCustomEnd);
      if(diff<=60){
        const byDay={};
        src.forEach(r=>{const d=r.order_date;if(!byDay[d])byDay[d]={date:d.slice(5),shipped:0};if(r.status==="배송")byDay[d].shipped++;});
        return Object.values(byDay).sort((a,b)=>a.date>b.date?1:-1);
      }
      const byMonth={};
      src.forEach(r=>{const ym=r.order_date?.slice(0,7);if(!ym)return;if(!byMonth[ym])byMonth[ym]={date:ym,shipped:0};if(r.status==="배송")byMonth[ym].shipped++;});
      return Object.values(byMonth).sort((a,b)=>a.date>b.date?1:-1);
    }
    const c=new Date(); c.setMonth(c.getMonth()-3);
    const cut=c.toISOString().slice(0,10);
    const byMonth={};
    orders.filter(r=>r.order_date>=cut).forEach(r=>{
      const ym=r.order_date?.slice(0,7); if(!ym) return;
      if(!byMonth[ym]) byMonth[ym]={date:ym,shipped:0};
      if(r.status==="배송") byMonth[ym].shipped++;
    });
    return Object.values(byMonth).sort((a,b)=>a.date>b.date?1:-1);
  },[orders,shippingPeriod,shippingCustomStart,shippingCustomEnd]);

  // 일별 반품 by 채널 차트
  const returnChartData=useMemo(()=>{
    let start,end=localDate(-1);
    if(returnPeriod==="custom"&&returnCustomStart&&returnCustomEnd){
      start=returnCustomStart; end=returnCustomEnd;
    } else if(returnPeriod==="yd"){
      start=localDate(-1); end=start;
    } else if(returnPeriod==="7d"){
      const d=new Date(); d.setDate(d.getDate()-7);
      start=d.toISOString().slice(0,10);
    } else {
      const d=new Date(); d.setMonth(d.getMonth()-(returnPeriod==="1m"?1:3));
      start=d.toISOString().slice(0,10);
    }
    const filteredRet=orders.filter(r=>r.order_date>=start&&r.order_date<=end&&r.channel!=="오프라인 스토어");
    const retByCh={};
    filteredRet.forEach(r=>{
      if(r.status==="반품"){
        const ch=r.channel||"미분류";
        retByCh[ch]=(retByCh[ch]||0)+1;
      }
    });
    // 반품 건수 내림차순, 0건 채널 제외
    const chs=Object.entries(retByCh).sort((a,b)=>b[1]-a[1]).map(([ch])=>ch).slice(0,5);
    const byDate={};
    filteredRet.forEach(r=>{
      const d=r.order_date; const ch=r.channel||"미분류";
      if(!d||!chs.includes(ch)) return;
      if(!byDate[d]){byDate[d]={date:d.slice(5)};chs.forEach(c=>byDate[d][c]=0);}
      if(r.status==="반품") byDate[d][ch]=(byDate[d][ch]||0)+1;
    });
    return {data:Object.values(byDate).sort((a,b)=>a.date>b.date?1:-1),channels:chs};
  },[orders,returnPeriod,returnCustomStart,returnCustomEnd]);

  function RankTable({data,cols}){
    return(
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <thead><tr>
          <th style={{padding:"5px 7px",color:D.textMeta,fontWeight:400,borderBottom:`1px solid ${D.border}`,width:22}}>#</th>
          {cols.map(c=><th key={c.key} style={{padding:"5px 7px",textAlign:c.right?"right":"left",color:D.textMeta,fontWeight:400,borderBottom:`1px solid ${D.border}`}}>{c.label}</th>)}
        </tr></thead>
        <tbody>
          {data.map((row,i)=>(
            <tr key={i} style={{borderBottom:`1px solid ${D.border}`}}>
              <td style={{padding:"5px 7px",color:i<3?D.black:D.textMeta,fontWeight:i<3?600:400}}>{i+1}</td>
              {cols.map(c=><td key={c.key} style={{padding:"5px 7px",textAlign:c.right?"right":"left",
                color:c.color||D.text,fontWeight:c.bold?600:400,maxWidth:c.maxW,
                overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                {c.fmt?c.fmt(row[c.key],row):row[c.key]}
              </td>)}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  const getPeriodLabel=p=>{
    const fmt=d=>`${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")}`;
    const today=new Date(); const todayStr=fmt(today);
    if(p==="7d"){const c=new Date();c.setDate(c.getDate()-7);return`${fmt(c)} ~ ${todayStr}`;}
    if(p==="14d"){const c=new Date();c.setDate(c.getDate()-14);return`${fmt(c)} ~ ${todayStr}`;}
    if(p==="1m"){const c=new Date();c.setMonth(c.getMonth()-1);return`${fmt(c)} ~ ${todayStr}`;}
    if(p==="tm"){const c=new Date(today.getFullYear(),today.getMonth(),1);return`${fmt(c)} ~ ${todayStr}`;}
    if(p==="3m"){const c=new Date();c.setMonth(c.getMonth()-3);return`${fmt(c)} ~ ${todayStr}`;}
    if(p==="6m"){const c=new Date();c.setMonth(c.getMonth()-6);return`${fmt(c)} ~ ${todayStr}`;}
    return null;
  };

  const PeriodBtn=({k,l})=>(
    <button onClick={()=>setPeriod(k)}
      style={{background:"transparent",border:`1px solid ${period===k?D.black:D.border}`,
        color:period===k?D.black:D.textSub,borderRadius:6,padding:"4px 11px",
        fontSize:11,cursor:"pointer",fontWeight:period===k?600:400}}>
      {l}
    </button>
  );

  const SmPeriodBtn=({val,cur,onChange,label})=>(
    <button onClick={()=>onChange(val)}
      style={{background:"transparent",border:`1px solid ${cur===val?D.black:D.border}`,
        color:cur===val?D.black:D.textSub,borderRadius:5,padding:"3px 8px",
        fontSize:10,cursor:"pointer",fontWeight:cur===val?600:400}}>
      {label}
    </button>
  );


  return (
    <div style={{padding:"20px 24px",maxWidth:1400,margin:"0 auto"}}>
      {/* 상단 기간 선택 + 새로고침 */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",flexDirection:"column",gap:5}}>
          <CalDrop id="kpi" period={period} setPeriod={setPeriod}
            presets={[["yd","어제"],["7d","최근 7일"],["tm","이번 한달"],["1m","최근 한달"],["3m","최근 3개월"],["all","전체"]]}
            start={customStart} setStart={setCustomStart}
            end={customEnd} setEnd={setCustomEnd}
            calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
          <div style={{fontSize:10,color:D.textMeta,paddingLeft:2,minHeight:14,lineHeight:"14px"}}>
            {getPeriodLabel(period)||""}
          </div>
        </div>
        <button onClick={onRefresh}
          style={{background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:7,
            padding:"5px 13px",fontSize:12,cursor:"pointer",color:D.textSub,
            display:"flex",alignItems:"center",gap:5}}>
          ↺ 새로고침
        </button>
      </div>

      {/* KPI 카드 - 총 매출/주문/반품은 매출입력, 배송·주문은 이지어드민 */}
      <div style={{display:"flex",gap:9,marginBottom:20,flexWrap:"wrap",minHeight:82}}>
        <KPI label="총 매출" value={fmtWonShort(stats.totalRevenue)} accent={D.black} onClick={()=>setKpiModal("revenue")}/>
        <KPI label="주문 건" value={stats.totalUniqueOrdersAll.toLocaleString()+"건"} sub={stats.totalOrderedQtyAll.toLocaleString()+"장"} accent={D.green} onClick={()=>setKpiModal("order")}/>
        <KPI label="배송 건" value={stats.totalShipped.toLocaleString()+"건"} sub={stats.totalDeliveredQty.toLocaleString()+"장"} accent={D.green} onClick={()=>setKpiModal("shipped")}/>
        {!["yd","7d"].includes(period)&&<KPI label="반품률" value={stats.returnRate+"%"}
          sub={`${stats.totalReturnedQty.toLocaleString()}장 / ${stats.totalDeliveredQty.toLocaleString()}장`}
          accent={parseFloat(stats.returnRate)>10?D.red:D.textSub}
          onClick={()=>setKpiModal("returnRate")}/>}
        <KPI label="입고 수량" value={stats.totalStock.toLocaleString()+"개"} accent={D.blue} onClick={()=>setKpiModal("stock")}/>
      </div>

      {/* 매출 점유율 + 판매처별 매출 + 판매처 상세 — 동일 일정 필터 사용 → 하나의 카드, 점선 구분 */}
      <Card style={{marginBottom:20}}>
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"280px 1fr",gap:0,minHeight:220}}>
        <div style={{paddingRight:isMobile?0:14,paddingBottom:isMobile?14:0,
          borderRight:isMobile?"none":`1px dashed ${D.border}`,
          borderBottom:isMobile?`1px dashed ${D.border}`:"none"}}>
          <SecTitle ts={ts.orders}>매출 점유율</SecTitle>
          {(()=>{
            const sorted=[...stats.channelList.slice(0,6)].sort((a,b)=>b.revenue-a.revenue);
            const total=sorted.reduce((s,c)=>s+c.revenue,0)||1;
            return (
              <>
                <ResponsiveContainer width="100%" height={150}>
                  <PieChart>
                    <Pie data={sorted.map(c=>({name:c.name,value:c.revenue}))}
                      dataKey="value" nameKey="name" cx="50%" cy="50%"
                      innerRadius={38} outerRadius={58} paddingAngle={2}>
                      {sorted.map((c,i)=>(<Cell key={i} fill={chColor(c.name)}/>))}
                    </Pie>
                    <Tooltip formatter={(v,n)=>{
                      const pct=(v/total*100).toFixed(1);
                      return [`₩${v.toLocaleString()} (${pct}%)`,n];
                    }} contentStyle={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:7,fontSize:11}}/>
                  </PieChart>
                </ResponsiveContainer>
                <div style={{display:"flex",flexWrap:"wrap",gap:"4px 10px",justifyContent:"center",paddingTop:6}}>
                  {sorted.map((c,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:D.textSub}}>
                      <span style={{width:7,height:7,borderRadius:"50%",background:chColor(c.name),flexShrink:0,display:"inline-block"}}/>
                      <span>{c.name}</span>
                      <span style={{color:D.textMeta}}>({(c.revenue/total*100).toFixed(1)}%)</span>
                    </div>
                  ))}
                </div>
              </>
            );
          })()}
        </div>
        <div ref={salesByChCardRef} style={{paddingLeft:isMobile?0:14,paddingTop:isMobile?14:0}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:2}}>
            <SecTitle ts={ts.orders}>판매처별 매출</SecTitle>
            <CaptureBtn cardRef={salesByChCardRef} filename="판매처별매출" DC={{border:D.border,sub:D.textMeta}}/>
          </div>
          <ResponsiveContainer width="100%" height={160}>
            {(()=>{
              // 오프라인 스토어 → 판교점/일산점 분리 스택
              const PANKYO_COLOR="#8b5cf6"; const ILSAN_COLOR="#c084fc";
              const chartData=stats.channelList.slice(0,7).map(c=>{
                if(c.name==="오프라인 스토어"){
                  const bd=stats.offlineBreakdown||{};
                  const pankyo=bd["판교점"]?.revenue||0;
                  const ilsan=bd["일산점"]?.revenue||0;
                  const hasBd=(pankyo+ilsan)>0;
                  return{name:c.name,
                    revenue:hasBd?0:c.revenue,
                    판교점:pankyo,
                    일산점:ilsan};
                }
                return{name:c.name,revenue:c.revenue};
              });
              const ChTip=({active,payload,label})=>{
                if(!active||!payload?.length) return null;
                const isOffline=label==="오프라인 스토어";
                const entries=payload.filter(p=>(p.value||0)>0);
                const total=entries.reduce((s,p)=>s+(p.value||0),0);
                return(
                  <div style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:7,padding:"8px 12px",fontSize:11,boxShadow:"0 2px 8px #0001"}}>
                    <div style={{color:D.textMeta,marginBottom:3}}>{label}</div>
                    {isOffline&&entries.length>1&&(
                      <div style={{color:D.text,marginBottom:2}}>합계: <strong>{fmtWonShort(total)}</strong></div>
                    )}
                    {entries.map((p,i)=>(
                      <div key={i} style={{color:p.color||D.text}}>
                        {p.name}: <strong>{fmtWonShort(p.value||0)}</strong>
                      </div>
                    ))}
                  </div>
                );
              };
              return(
                <BarChart data={chartData} layout="vertical" barCategoryGap="28%">
                  <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
                  <XAxis type="number" tick={axTick} tickFormatter={v=>v>=1e8?fmtEokMan(v):v>=1e4?(v/1e4).toFixed(0)+"만":v}/>
                  <YAxis type="category" dataKey="name" width={76} tick={axTick}/>
                  <Tooltip content={<ChTip/>}/>
                  <Bar dataKey="revenue" name="매출" stackId="a" radius={[0,3,3,0]}>
                    {chartData.map((c,i)=>(<Cell key={i} fill={chColor(c.name)}/>))}
                  </Bar>
                  <Bar dataKey="판교점" name="판교점" stackId="a" fill={PANKYO_COLOR} radius={[0,0,0,0]}/>
                  <Bar dataKey="일산점" name="일산점" stackId="a" fill={ILSAN_COLOR} radius={[0,3,3,0]}/>
                </BarChart>
              );
            })()}
          </ResponsiveContainer>
          {getPeriodStr(period,customStart,customEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:6}}>{getPeriodStr(period,customStart,customEnd)}</div>}
        </div>
      </div>

      {/* 점선 구분 */}
      <div style={{borderTop:`1px dashed ${D.border}`,margin:"20px 0"}}/>

      {/* 판매처 상세 */}
      <div ref={chDetailCardRef} style={{minHeight:380}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <SecTitle ts={ts.orders}>판매처 상세</SecTitle>
          <div style={{display:"flex",gap:4,flexWrap:"wrap",alignItems:"center"}}>
            {[["revenue","매출"],["share","점유율"],["orders","주문 건"],["shipped","배송 건"],...(!["yd","7d"].includes(period)?[["returned","반품 수량"],["rate","반품률"]]:[]),["aov","객단가"]].map(([k,l])=>(
              <button key={k} onClick={()=>setChSort({key:k,dir:"desc"})}
                style={{background:chSort.key===k?D.black:"transparent",
                  color:chSort.key===k?"#fff":D.textSub,
                  border:`1px solid ${chSort.key===k?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer",
                  fontWeight:600,minWidth:36,boxSizing:"border-box"}}>
                {l}
              </button>
            ))}
            <CaptureBtn cardRef={chDetailCardRef} filename="판매처상세" DC={{border:D.border,sub:D.textMeta}}/>
          </div>
        </div>
        {(()=>{
          const fmtChg=(cur,prev)=>{
            if(!prev||prev===0) return null;
            const pct=((cur-prev)/Math.abs(prev)*100);
            const up=pct>=0;
            return <span style={{fontSize:10,fontWeight:700,color:D.black,marginLeft:3}}>{up?"▲":"▼"}{Math.abs(pct).toFixed(1)}%</span>;
          };
          const hasRet=!["yd","7d"].includes(period);
          const ORDERS_TIP="주문 이후 취소된 수량도 집계한 데이터입니다. 총 인입 주문 수 집계용으로 매출 데이터와 별개의 로직입니다.";
          const cols=[
            {key:"name",   label:"판매처",     left:true, val:c=>c.name,               w:hasRet?"17%":"20%"},
            {key:"share",  label:"점유율",                val:c=>parseFloat(c.share),  w:"7%"},
            {key:"revenue",label:"매출",                  val:c=>c.revenue,            w:hasRet?"13%":"16%"},
            {key:"cmp",    label:"동기간 비교",            val:c=>0,                    w:hasRet?"12%":"15%"},
            // 주문 수: 모든 상태 고유 주문번호 (건수) + 주문 장수 (병기, 셀 안에 sub)
            {key:"orders", label:"주문 건",               val:c=>c.totalOrders||0,     w:"10%", tooltip:ORDERS_TIP},
            {key:"shipped",label:"배송 건",               val:c=>c.shipped,            w:"10%"},
            ...(hasRet?[
              {key:"returned",label:"반품 수량(장)",       val:c=>c.returnedQty||0,     w:"10%"},
              {key:"rate",   label:"반품률",              val:c=>c.shippedQty>0?c.returnedQty/c.shippedQty:0, w:"8%"},
            ]:[]),
            {key:"aov",    label:"객단가",                val:c=>c.avgOrderValue||0,   w:hasRet?"14%":"17%"},
          ];
          const sorted=[...stats.channelList].sort((a,b)=>{
            const col=cols.find(c=>c.key===chSort.key);
            if(!col||col.key==="cmp") return 0;
            const va=col.val(a), vb=col.val(b);
            return chSort.dir==="desc"?(vb>va?1:vb<va?-1:0):(va>vb?1:va<vb?-1:0);
          });
          const bd=stats.offlineBreakdown||{};
          const hasBd=Object.keys(bd).length>0;
          const subRows=(offlineExpanded&&hasBd)?Object.entries(bd).map(([n,d])=>({...d,name:n,isSubRow:true})):[];
          const allRows=[];
          sorted.forEach(c=>{
            allRows.push(c);
            if(c.name==="오프라인 스토어") subRows.forEach(sr=>allRows.push(sr));
          });
          return (
            <div style={{minHeight:300,overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,tableLayout:"fixed"}}>
              <colgroup>{cols.map(({key,w})=><col key={key} style={{width:w}}/>)}</colgroup>
              <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
                {cols.map(({key,label,left,tooltip,wip})=>(
                  <th key={key} onClick={key!=="cmp"?()=>setChSort({key,dir:"desc"}):undefined}
                    style={{padding:"7px 9px",textAlign:left?"left":"right",
                    color:chSort.key===key?D.black:D.textMeta,
                    fontWeight:600,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",
                    cursor:key!=="cmp"?"pointer":"default"}}>
                    {label}
                    {wip&&<span style={{marginLeft:4,fontSize:9,fontWeight:500,color:"#fff",background:"#bbb",borderRadius:3,padding:"1px 4px",verticalAlign:"middle"}}>개발중</span>}
                    {tooltip&&<InfoTip text={tooltip}><span style={{marginLeft:3,fontSize:10,color:D.textMeta,fontWeight:400,cursor:"default"}}>ⓘ</span></InfoTip>}
                  </th>
                ))}
              </tr></thead>
              <tbody>
                {allRows.map((c,idx)=>{
                  const prev=prevChRevenue[c.name]||0;
                  const isOffline=c.name==="오프라인 스토어";
                  return(
                    <tr key={c.name+idx} style={{borderBottom:`1px solid ${D.border}`,
                      background:c.isSubRow?"#faf8ff":"transparent",
                      opacity:c.isSubRow?0.85:1}}>
                      <td style={{padding:"7px 9px",fontWeight:c.isSubRow?400:600,
                        paddingLeft:c.isSubRow?20:9,color:c.isSubRow?D.textSub:D.black,
                        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {c.name}
                        {isOffline&&hasBd&&(
                          <button onClick={()=>setOfflineExpanded(v=>!v)}
                            style={{marginLeft:5,background:"none",border:"none",
                              padding:"0 2px",fontSize:10,cursor:"pointer",
                              color:D.textMeta}}>{offlineExpanded?"▾":"▸"}</button>
                        )}
                      </td>
                      <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>{c.share||"—"}%</td>
                      <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.revenue>0?fmtWonShort(c.revenue):"—"}</td>
                      <td style={{textAlign:"right",padding:"7px 9px"}}>
                        {!c.isSubRow&&prevPeriod?fmtChg(c.revenue,prev)||<span style={{color:D.textMeta,fontSize:10}}>—</span>:<span style={{color:D.textMeta,fontSize:10}}>—</span>}
                        {!c.isSubRow&&prevPeriod&&<div style={{fontSize:9,color:"#bbb",marginTop:1}}>{prevPeriod.start}~{prevPeriod.end}</div>}
                      </td>
                      {/* 주문 수: 클릭하면 채널별 주문 소스 모달, hover 시 안내 툴팁 */}
                      <td onClick={!c.isSubRow&&c.totalOrders>0?()=>setChOrderModal(c.name):undefined}
                        style={{textAlign:"right",padding:"7px 9px",
                          color:!c.isSubRow&&c.totalOrders>0?D.blue:D.textSub,
                          cursor:!c.isSubRow&&c.totalOrders>0?"pointer":"default",
                          textDecoration:!c.isSubRow&&c.totalOrders>0?"underline":"none"}}>
                        <InfoTip text={ORDERS_TIP}>
                          <span>{(c.totalOrders||0).toLocaleString()}</span>
                        </InfoTip>
                        <div style={{fontSize:9,color:D.textMeta,fontWeight:400}}>{(c.orderedQty||0).toLocaleString()}장</div>
                      </td>
                      <td style={{textAlign:"right",padding:"7px 9px",color:D.green}}>
                        {isOffline||c.isSubRow?(
                          <span style={{color:D.textMeta,fontSize:10}}>—</span>
                        ):(<>
                          {(c.shipped||0).toLocaleString()}
                          <div style={{fontSize:9,color:D.textMeta,fontWeight:400}}>{(c.shippedQty||0).toLocaleString()}장</div>
                        </>)}
                      </td>
                      {hasRet&&<td style={{textAlign:"right",padding:"7px 9px",color:D.red}}>
                        {isOffline||c.isSubRow?(
                          <span style={{fontSize:9,fontWeight:500,color:"#888",background:"#f2f2f2",borderRadius:3,padding:"2px 5px"}}>구현 중</span>
                        ):(<>
                          {(c.returnedQty||0).toLocaleString()}<span style={{fontSize:9,color:D.textMeta,marginLeft:2}}>장</span>
                        </>)}
                      </td>}
                      {hasRet&&<td style={{textAlign:"right",padding:"7px 9px",fontWeight:600}}>
                        {c.shippedQty>0?(c.returnedQty/c.shippedQty*100).toFixed(1):"0.0"}%</td>}
                      <td onClick={c.avgOrderValue>0?()=>setAovModal(c.name):undefined}
                        style={{textAlign:"right",padding:"7px 9px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                          color:c.avgOrderValue>0?D.blue:"—"===c.avgOrderValue?"transparent":D.textMeta,
                          cursor:c.avgOrderValue>0?"pointer":"default",
                          textDecoration:c.avgOrderValue>0?"underline":"none"}}>
                        {c.avgOrderValue>0?fmtWonShort(c.avgOrderValue):"—"}</td>
                    </tr>
                  );
                })}
                <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                  <td style={{padding:"7px 9px"}}>합계</td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>100%</td>
                  <td style={{textAlign:"right",padding:"7px 9px"}}>{fmtWonShort(stats.totalRevenue)}</td>
                  <td/>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>
                    <InfoTip text={ORDERS_TIP}>
                      <span>{stats.totalUniqueOrdersAll.toLocaleString()}</span>
                    </InfoTip>
                    <div style={{fontSize:9,color:D.textMeta,fontWeight:400}}>{stats.totalOrderedQtyAll.toLocaleString()}장</div>
                  </td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.green}}>
                    {stats.totalShipped.toLocaleString()}
                    <div style={{fontSize:9,color:D.textMeta,fontWeight:400}}>{stats.totalDeliveredQty.toLocaleString()}장</div>
                  </td>
                  {hasRet&&<td style={{textAlign:"right",padding:"7px 9px",color:D.red}}>
                    {stats.totalReturnedQty.toLocaleString()}<span style={{fontSize:9,color:D.textMeta,marginLeft:2}}>장</span>
                  </td>}
                  {hasRet&&<td style={{textAlign:"right",padding:"7px 9px"}}>{stats.returnRate}%</td>}
                  <td/>
                </tr>
              </tbody>
            </table>
            </div>
          );
        })()}
        {getPeriodStr(period,customStart,customEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:8}}>{getPeriodStr(period,customStart,customEnd)}</div>}
      </div>
      </Card>

      {/* 월별 배송량 (독립 기간) */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10,marginBottom:20}}>
        <Card ref={shippingCardRef}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <SecTitle ts={ts.orders}>배송량</SecTitle>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <CalDrop id="shipping" period={shippingPeriod} setPeriod={setShippingPeriod}
                presets={[["yd","어제"],["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]]}
                start={shippingCustomStart} setStart={setShippingCustomStart}
                end={shippingCustomEnd} setEnd={setShippingCustomEnd}
                calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
              <CaptureBtn cardRef={shippingCardRef} filename="배송량" DC={{border:D.border,sub:D.textMeta}}/>
            </div>
          </div>
          <ResponsiveContainer width="100%" height={170}>
            <BarChart data={shippingChartData}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="date" tick={axTick}/>
              <YAxis tick={axTick}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey="shipped" name="배송" fill="#7EADD4" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
          {getPeriodStr(shippingPeriod,shippingCustomStart,shippingCustomEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:6}}>{getPeriodStr(shippingPeriod,shippingCustomStart,shippingCustomEnd)}</div>}
        </Card>

        {/* 판매처별 일자 반품 */}
        <Card ref={returnTrendCardRef}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <SecTitle ts={ts.orders}>판매처별 반품 추이</SecTitle>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <CalDrop id="return" period={returnPeriod} setPeriod={setReturnPeriod}
                presets={[["yd","어제"],["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]]}
                start={returnCustomStart} setStart={setReturnCustomStart}
                end={returnCustomEnd} setEnd={setReturnCustomEnd}
                calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
              <CaptureBtn cardRef={returnTrendCardRef} filename="반품추이" DC={{border:D.border,sub:D.textMeta}}/>
            </div>
          </div>
          <ResponsiveContainer width="100%" height={170}>
            <LineChart data={returnChartData.data}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="date" tick={axTick}/>
              <YAxis tick={axTick}/>
              <Tooltip content={<Tip/>}/>
              {returnChartData.channels.map((ch,i)=>(
                <Line key={ch} type="monotone" dataKey={ch} name={ch}
                  stroke={D.SANKEY[(i+3)%D.SANKEY.length]} strokeWidth={1.5}
                  dot={false}/>
              ))}
              <Legend iconSize={8} wrapperStyle={{fontSize:10}}/>
            </LineChart>
          </ResponsiveContainer>
          {getPeriodStr(returnPeriod,returnCustomStart,returnCustomEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:6}}>{getPeriodStr(returnPeriod,returnCustomStart,returnCustomEnd)}</div>}
        </Card>
      </div>

      {/* 판매 Top */}
      <Card ref={salesTopCardRef} style={{marginBottom:20,minHeight:660}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <SecTitle ts={ts.orders}>판매 Top</SecTitle>
            {["전체",...activeChannels].map(ch=>(
              <button key={ch} onClick={()=>setRankBestChannel(ch)}
                style={{background:rankBestChannel===ch?D.black:"transparent",
                  color:rankBestChannel===ch?"#fff":D.textSub,
                  border:`1px solid ${rankBestChannel===ch?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>{ch}</button>
            ))}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <CalDrop id="best" period={rankBestPeriod} setPeriod={setRankBestPeriod}
              presets={[["yd","어제"],["7d","최근 7일"],["14d","최근 14일"],["1m","최근 한달"],["3m","최근 3개월"]]}
              start={rankBestCustomStart} setStart={setRankBestCustomStart}
              end={rankBestCustomEnd} setEnd={setRankBestCustomEnd}
              calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
            <CaptureBtn cardRef={salesTopCardRef} filename="판매Top" DC={{border:D.border,sub:D.textMeta}}/>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:14,alignItems:"start"}}>
          <div style={{minHeight:546,overflowY:"auto"}}>
            <RankTable data={bestRows} cols={[
              {key:"name",label:"상품명",maxW:190,bold:true,color:"#2d2d2d"},
              {key:"qty",label:"주문량",right:true,bold:true,fmt:v=>v.toLocaleString()},
              {key:"share",label:"주문 점유율",right:true,color:D.textMeta,fmt:v=>v+"%"},
            ]}/>
          </div>
          <ResponsiveContainer width="100%" height={546}>
            <BarChart data={bestRows.slice(0,12)} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
              <XAxis type="number" tick={axTick}/>
              <YAxis type="category" dataKey="name" width={180} tick={<NoWrapTick/>}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey="qty" name="주문량" radius={[0,3,3,0]}>
                {bestRows.slice(0,12).map((_,i)=>(
                  <Cell key={i} fill={D.SANKEY[i%D.SANKEY.length]}/>
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        {getPeriodStr(rankBestPeriod,rankBestCustomStart,rankBestCustomEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:6}}>{getPeriodStr(rankBestPeriod,rankBestCustomStart,rankBestCustomEnd)}</div>}
        <div style={{fontSize:10,color:D.textMeta,marginTop:4}}>주문일 기준 · 소스: 주문·배송 업로드 데이터 (오프라인 제외)</div>
      </Card>

      {/* 플랫폼별 선호 옵션 */}
      {optionStats.length>0&&(
        <Card ref={optionPrefCardRef} style={{marginBottom:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <SecTitle ts={ts.orders}>플랫폼별 선호 옵션</SecTitle>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <CalDrop id="option" period={optionPeriod} setPeriod={setOptionPeriod}
                presets={[["1m","1달"],["3m","3달"],["6m","6달"]]}
                start={optionCustomStart} setStart={setOptionCustomStart}
                end={optionCustomEnd} setEnd={setOptionCustomEnd}
                calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
              <CaptureBtn cardRef={optionPrefCardRef} filename="선호옵션" DC={{border:D.border,sub:D.textMeta}}/>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":`repeat(${optionStats.length},1fr)`,gap:16}}>
            {optionStats.map(({ch,colors,sizes})=>(
              <div key={ch}>
                <div style={{fontWeight:700,fontSize:12,color:D.textSub,marginBottom:10,letterSpacing:"0.04em"}}>{ch}</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <div>
                    <div style={{fontSize:10,color:D.textMeta,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em"}}>컬러</div>
                    {colors.length===0&&<div style={{fontSize:11,color:D.textMeta}}>데이터 없음</div>}
                    {colors.map(([name,cnt],i)=>{
                      const max=colors[0]?.[1]||1;
                      return (
                        <div key={name} style={{marginBottom:5}}>
                          <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2,alignItems:"center"}}>
                            <span style={{display:"flex",alignItems:"center",color:i===0?D.black:D.textSub,fontWeight:i===0?700:400}}>
                              {colorSwatch(name)}{name}
                            </span>
                            <span style={{color:D.textMeta}}>{cnt}</span>
                          </div>
                          <div style={{height:4,borderRadius:2,background:D.border}}>
                            <div style={{height:4,borderRadius:2,background:COLOR_HEX[name]||COLOR_HEX[name?.toUpperCase()]||(i===0?D.black:D.blue),width:`${(cnt/max*100).toFixed(0)}%`}}/>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  <div>
                    <div style={{fontSize:10,color:D.textMeta,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em"}}>사이즈</div>
                    {sizes.length===0&&<div style={{fontSize:11,color:D.textMeta}}>데이터 없음</div>}
                    {sizes.map(([name,cnt],i)=>{
                      const max=sizes[0]?.[1]||1;
                      return (
                        <div key={name} style={{marginBottom:5}}>
                          <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2}}>
                            <span style={{color:i===0?D.black:D.textSub,fontWeight:i===0?700:400}}>{name}</span>
                            <span style={{color:D.textMeta}}>{cnt}</span>
                          </div>
                          <div style={{height:4,borderRadius:2,background:D.border}}>
                            <div style={{height:4,borderRadius:2,background:i===0?D.black:D.green,width:`${(cnt/max*100).toFixed(0)}%`}}/>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            ))}
          </div>
          {getPeriodStr(optionPeriod,optionCustomStart,optionCustomEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:8}}>{getPeriodStr(optionPeriod,optionCustomStart,optionCustomEnd)}</div>}
        </Card>
      )}

      {/* 반품 탑 */}
      <Card ref={returnTopCardRef} style={{marginBottom:20,minHeight:660}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <SecTitle ts={ts.orders}>반품 Top</SecTitle>
            {["전체",...activeChannels].map(ch=>(
              <button key={ch} onClick={()=>setRankWorstChannel(ch)}
                style={{background:rankWorstChannel===ch?D.black:"transparent",
                  color:rankWorstChannel===ch?"#fff":D.textSub,
                  border:`1px solid ${rankWorstChannel===ch?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>{ch}</button>
            ))}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <CalDrop id="worst" period={rankWorstPeriod} setPeriod={setRankWorstPeriod}
              presets={[["1m","최근 한달"],["3m","최근 3개월"]]}
              start={rankWorstCustomStart} setStart={setRankWorstCustomStart}
              end={rankWorstCustomEnd} setEnd={setRankWorstCustomEnd}
              calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
            <CaptureBtn cardRef={returnTopCardRef} filename="반품Top" DC={{border:D.border,sub:D.textMeta}}/>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:14,alignItems:"start"}}>
          <div style={{minHeight:546,overflowY:"auto"}}>
            <RankTable data={worstRows} cols={[
              {key:"name",label:"상품명",maxW:160,bold:true,color:"#2d2d2d"},
              {key:"returnRate",label:"반품률",right:true,bold:true,color:D.red,fmt:v=>v+"%"},
              {key:"returned",label:"반품",right:true,color:D.red,fmt:v=>v.toLocaleString()},
              {key:"topReason",label:"주요 사유",right:false,color:D.textMeta,maxW:130},
              {key:"qty",label:"배송량",right:true,color:D.textSub,fmt:v=>v.toLocaleString()},
            ]}/>
          </div>
          <ResponsiveContainer width="100%" height={546}>
            <BarChart data={worstRows.slice(0,12)} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
              <XAxis type="number" tick={axTick} tickFormatter={v=>v+"%"}/>
              <YAxis type="category" dataKey="name" width={180} tick={<NoWrapTick/>}/>
              <Tooltip content={<Tip/>} formatter={(v)=>[v+"%","반품률"]}/>
              <Bar dataKey="returnRate" name="반품률" radius={[0,3,3,0]}>
                {worstRows.slice(0,12).map((_,i)=>{
                  const palette=["#4e79a7","#f28e2b","#e15759","#76b7b2","#59a14f","#edc948","#b07aa1","#ff9da7","#9c755f","#bab0ac","#6b6ecf","#8ca252"];
                  return <Cell key={i} fill={palette[i%palette.length]}/>;
                })}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        {getPeriodStr(rankWorstPeriod,rankWorstCustomStart,rankWorstCustomEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:6}}>{getPeriodStr(rankWorstPeriod,rankWorstCustomStart,rankWorstCustomEnd)}</div>}
        <div style={{fontSize:10,color:D.textMeta,marginTop:4}}>주문일 기준 · 소스: 주문·배송 업로드 데이터 (오프라인 제외)</div>
      </Card>

      {/* 플랫폼별 반품률 높은 옵션 */}
      {returnOptionStats.length>0&&(
        <Card ref={returnOptCardRef} style={{marginBottom:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <SecTitle ts={ts.orders}>플랫폼별 반품률 높은 옵션</SecTitle>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <CalDrop id="returnOption" period={returnOptionPeriod} setPeriod={setReturnOptionPeriod}
                presets={[["1m","1달"],["3m","3달"],["6m","6달"]]}
                start={returnOptionCustomStart} setStart={setReturnOptionCustomStart}
                end={returnOptionCustomEnd} setEnd={setReturnOptionCustomEnd}
                calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
              <CaptureBtn cardRef={returnOptCardRef} filename="반품률높은옵션" DC={{border:D.border,sub:D.textMeta}}/>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":`repeat(${returnOptionStats.length},1fr)`,gap:16}}>
            {returnOptionStats.map(({ch,colors,sizes})=>(
              <div key={ch}>
                <div style={{fontWeight:700,fontSize:12,color:D.textSub,marginBottom:10,letterSpacing:"0.04em"}}>{ch}</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <div>
                    <div style={{fontSize:10,color:D.textMeta,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em"}}>컬러</div>
                    {colors.length===0&&<div style={{fontSize:11,color:D.textMeta}}>데이터 없음</div>}
                    {colors.map(({name,rate,returned,total},i)=>(
                      <div key={name} style={{marginBottom:6}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2,alignItems:"center"}}>
                          <span style={{display:"flex",alignItems:"center",color:i===0?D.red:D.textSub,fontWeight:i===0?700:400}}>
                            {colorSwatch(name)}{name}
                          </span>
                          <span style={{color:i===0?D.red:D.textMeta,fontWeight:i===0?600:400}}>{rate.toFixed(1)}%</span>
                        </div>
                        <div style={{height:4,borderRadius:2,background:D.border}}>
                          <div style={{height:4,borderRadius:2,
                            background:COLOR_HEX[name]||COLOR_HEX[name?.toUpperCase()]||D.red,
                            opacity:0.7,width:`${Math.min(100,rate*4).toFixed(0)}%`}}/>
                        </div>
                        <div style={{fontSize:9,color:D.textMeta,marginTop:1}}>{returned}건 반품 / 배송 {total}건</div>
                      </div>
                    ))}
                  </div>
                  <div>
                    <div style={{fontSize:10,color:D.textMeta,fontWeight:600,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.08em"}}>사이즈</div>
                    {sizes.length===0&&<div style={{fontSize:11,color:D.textMeta}}>데이터 없음</div>}
                    {sizes.map(({name,rate,returned,total},i)=>(
                      <div key={name} style={{marginBottom:6}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2}}>
                          <span style={{color:i===0?D.red:D.textSub,fontWeight:i===0?700:400}}>{name}</span>
                          <span style={{color:i===0?D.red:D.textMeta,fontWeight:i===0?600:400}}>{rate.toFixed(1)}%</span>
                        </div>
                        <div style={{height:4,borderRadius:2,background:D.border}}>
                          <div style={{height:4,borderRadius:2,background:D.red,opacity:0.5+(i===0?0.3:0),
                            width:`${Math.min(100,rate*4).toFixed(0)}%`}}/>
                        </div>
                        <div style={{fontSize:9,color:D.textMeta,marginTop:1}}>{returned}건 반품 / 배송 {total}건</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            ))}
          </div>
          {getPeriodStr(returnOptionPeriod,returnOptionCustomStart,returnOptionCustomEnd)&&<div style={{fontSize:10,color:D.textMeta,marginTop:8}}>{getPeriodStr(returnOptionPeriod,returnOptionCustomStart,returnOptionCustomEnd)}</div>}
        </Card>
      )}

      {/* 데이터 삭제 */}
      <div style={{marginTop:24,paddingTop:16,borderTop:`1px solid ${D.border}`,display:"flex",justifyContent:"flex-end",flexWrap:"wrap",gap:8,alignItems:"center"}}>
        {!deleteAll?(
          <>
            <CalDrop id="del" period={delPeriod} setPeriod={setDelPeriod}
              presets={[["all","전체"]]}
              start={delStart} setStart={setDelStart}
              end={delEnd} setEnd={setDelEnd}
              calOpenFor={delCalOpen} setCalOpenFor={setDelCalOpen}/>
            <button onClick={()=>setDeleteAll(true)}
              style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"6px 14px",fontSize:11,cursor:"pointer",color:D.textMeta}}>
              ⚠ 데이터 삭제
            </button>
          </>
        ):(
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <span style={{fontSize:11,color:D.red}}>
              {delPeriod==="custom"&&delStart&&delEnd
                ? `${delStart}~${delEnd} 기간 데이터가 삭제됩니다. 확인하시겠습니까?`
                : "모든 주문·입고·매출 데이터가 삭제됩니다. 확인하시겠습니까?"}
            </span>
            <button onClick={async()=>{
              const db=await getSupabase();
              const s=delPeriod==="custom"&&delStart?delStart:"2000-01-01";
              const e=delPeriod==="custom"&&delEnd?delEnd:"2099-12-31";
              await Promise.all([
                // order_headers 삭제 시 order_items도 FK CASCADE로 함께 정리됨
                db.from("order_headers").delete().gte("order_date",s).lte("order_date",e),
                db.from("stock_uploads").delete().gte("upload_date",s).lte("upload_date",e),
                db.from("revenues").delete().gte("date",s).lte("date",e),
                db.from("store_sales").delete().gte("sale_date",s).lte("sale_date",e),
                ...(delPeriod!=="custom"?[db.from("cs_data").delete().gte("id",1)]:[]),
              ]);
              if(delPeriod!=="custom"){try{localStorage.removeItem("cs_data");}catch{}}
              setDeleteAll(false);
              onRefresh();
            }} style={{background:D.red,color:"#fff",border:"none",borderRadius:6,
              padding:"6px 14px",fontSize:11,cursor:"pointer",fontWeight:600}}>
              삭제
            </button>
            <button onClick={()=>setDeleteAll(false)}
              style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"6px 14px",fontSize:11,cursor:"pointer",color:D.text}}>
              취소
            </button>
          </div>
        )}
      </div>

      {/* KPI 소스 모달 */}
      {kpiModal&&(()=>{
        const OFFL=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
        const normCh=ch=>OFFL.has(ch||"")?"오프라인 스토어":(ch||"미분류");

        /* ── 총 매출 ── */
        let modalTitle="", modalContent=null;
        if(kpiModal==="revenue"){
          modalTitle="총 매출 소스";
          // by channel
          const byCh={};
          filteredRevenues.forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={revenue:0,refund:0,count:0};
            byCh[ch].revenue+=(r.amount||0);
            byCh[ch].refund+=(r.refund_amount||0);
            byCh[ch].count++;
          });
          // 오프라인 스토어: store_sales CSV 기반
          if(filteredStoreSales.length){
            if(!byCh["오프라인 스토어"]) byCh["오프라인 스토어"]={revenue:0,refund:0,count:0};
            filteredStoreSales.forEach(r=>{
              if(r.status==="배송"){ byCh["오프라인 스토어"].revenue+=(r.amount||0); byCh["오프라인 스토어"].count++; }
              else if(r.status==="반품") byCh["오프라인 스토어"].refund+=(r.amount||0);
            });
          }
          const chRows=Object.entries(byCh).sort((a,b)=>(b[1].revenue-b[1].refund)-(a[1].revenue-a[1].refund));
          // by date (last 30 entries or all)
          const byDate={};
          filteredRevenues.forEach(r=>{
            const d=r.date||"—";
            if(!byDate[d]) byDate[d]={revenue:0,refund:0};
            byDate[d].revenue+=(r.amount||0);
            byDate[d].refund+=(r.refund_amount||0);
          });
          const dateRows=Object.entries(byDate).sort((a,b)=>b[0]>a[0]?1:-1).slice(0,30);
          modalContent=(
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:16,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: <b>매출 입력 업로더</b> (revenues CSV) — 채널별 일자 매출/환불금<br/>
                계산: 채널별 = SUM(amount) − SUM(refund_amount) · 순매출 = 매출 − 반품<br/>
                오프라인 스토어 매출/반품은 <b>매장 판매 CSV</b> (store_sales)의 실판매금액 (배송 − 반품)
              </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>채널별</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>채널</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>매출</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>순매출</th>
                  </tr></thead>
                  <tbody>
                    {chRows.map(([ch,d])=>(
                      <tr key={ch} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",fontWeight:600}}>{ch}</td>
                        <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWonShort(d.revenue)}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{d.refund>0?fmtWonShort(d.refund):"—"}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",fontWeight:700}}>{fmtWonShort(d.revenue-d.refund)}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWonShort(chRows.reduce((s,[,d])=>s+d.revenue,0))}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{fmtWonShort(chRows.reduce((s,[,d])=>s+d.refund,0))}</td>
                      <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWonShort(stats.totalRevenue)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>날짜별 (최근 30일)</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>날짜</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>매출</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>순매출</th>
                  </tr></thead>
                  <tbody>
                    {dateRows.map(([d,v])=>(
                      <tr key={d} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{d}</td>
                        <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWonShort(v.revenue)}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{v.refund>0?fmtWonShort(v.refund):"—"}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",fontWeight:600}}>{fmtWonShort(v.revenue-v.refund)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            </div>
            </div>
          );
        }

        /* ── 주문 (KPI 주문 건) ── */
        else if(kpiModal==="order"){
          modalTitle="주문 소스";
          // 주문 건: 모든 상태(배송/반품/교환) 포함, 고유 주문번호 단위
          // 온라인: 이지어드민 orders + 매장: store_sales (둘 다 filteredOrders에 머지됨)
          const byCh={};
          filteredOrders.forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={oids:new Set(),qty:0};
            const oid=r.order_no||r.order_id||"";
            if(oid) byCh[ch].oids.add(oid);
            byCh[ch].qty+=(r.qty||1);
          });
          const chRows=Object.entries(byCh).sort((a,b)=>b[1].oids.size-a[1].oids.size);
          const byDate={};
          filteredOrders.forEach(r=>{
            const d=r.order_date||"—";
            if(!byDate[d]) byDate[d]={oids:new Set(),qty:0};
            const oid=r.order_no||r.order_id||"";
            if(oid) byDate[d].oids.add(oid);
            byDate[d].qty+=(r.qty||1);
          });
          const dateRows=Object.entries(byDate).sort((a,b)=>b[0]>a[0]?1:-1).slice(0,30);
          modalContent=(
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:16,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: <b>주문·배송 업로드 데이터</b> (orders, 모든 상태) + <b>매장 판매 CSV</b> (store_sales, 모든 상태)<br/>
                <b>주문 건</b> = COUNT(DISTINCT 주문번호) — 배송/반품/교환 등 상태 무관<br/>
                <b>주문 수량(장)</b> = SUM(qty) — 모든 상태 포함
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>채널별</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>채널</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>주문 건</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>주문 수량(장)</th>
                  </tr></thead>
                  <tbody>
                    {chRows.map(([ch,d])=>(
                      <tr key={ch} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",fontWeight:600}}>{ch}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green,fontWeight:600}}>{d.oids.size.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{d.qty.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{stats.totalUniqueOrdersAll.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{stats.totalOrderedQtyAll.toLocaleString()}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>주문일별 (최근 30일)</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>주문일</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>주문 건</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>주문 수량(장)</th>
                  </tr></thead>
                  <tbody>
                    {dateRows.map(([d,v])=>(
                      <tr key={d} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{d}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green,fontWeight:600}}>{v.oids.size.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{v.qty.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
              </div>
            </div>
          );
        }

        /* ── 배송 (KPI 배송 건) ── */
        else if(kpiModal==="shipped"){
          modalTitle="배송 소스";
          // 배송 건: 이지어드민 orders 중 status="배송", 매장 제외
          const OFFL2=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
          const isOff=r=>OFFL2.has(r.channel||"");
          const shipped=filteredOrders.filter(r=>r.status==="배송"&&!isOff(r));
          const byCh={};
          shipped.forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={oids:new Set(),qty:0};
            const oid=r.order_no||r.order_id||"";
            if(oid) byCh[ch].oids.add(oid);
            byCh[ch].qty+=(r.qty||1);
          });
          const chRows=Object.entries(byCh).sort((a,b)=>b[1].oids.size-a[1].oids.size);
          const byDate={};
          shipped.forEach(r=>{
            const d=r.order_date||"—";
            if(!byDate[d]) byDate[d]={oids:new Set(),qty:0};
            const oid=r.order_no||r.order_id||"";
            if(oid) byDate[d].oids.add(oid);
            byDate[d].qty+=(r.qty||1);
          });
          const dateRows=Object.entries(byDate).sort((a,b)=>b[0]>a[0]?1:-1).slice(0,30);
          modalContent=(
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:16,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: <b>주문·배송 업로드 데이터</b> (orders, status="배송")<br/>
                <b>배송 건</b> = COUNT(DISTINCT 주문번호) where status="배송"<br/>
                <b>배송 수량(장)</b> = SUM(qty) where status="배송"
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>채널별</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>채널</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송 건</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송 수량(장)</th>
                  </tr></thead>
                  <tbody>
                    {chRows.map(([ch,d])=>(
                      <tr key={ch} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",fontWeight:600}}>{ch}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green,fontWeight:600}}>{d.oids.size.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{d.qty.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{stats.totalShipped.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{stats.totalDeliveredQty.toLocaleString()}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>주문일별 (최근 30일)</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>주문일</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송 건</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송 수량(장)</th>
                  </tr></thead>
                  <tbody>
                    {dateRows.map(([d,v])=>(
                      <tr key={d} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{d}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green,fontWeight:600}}>{v.oids.size.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{v.qty.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
              </div>
            </div>
          );
        }

        /* ── 반품률 ── */
        else if(kpiModal==="returnRate"){
          modalTitle="반품률 소스";
          // 반품률 = 반품 수량(장) / 배송 수량(장) * 100 — 동기간 내, 매장 제외
          const OFFL3=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
          const isOff=r=>OFFL3.has(r.channel||"");
          const byCh={};
          filteredOrders.filter(r=>!isOff(r)&&(r.status==="배송"||r.status==="반품")).forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={shippedQty:0,returnedQty:0};
            const q=r.qty||1;
            if(r.status==="배송") byCh[ch].shippedQty+=q;
            else if(r.status==="반품") byCh[ch].returnedQty+=q;
          });
          const chRows=Object.entries(byCh).sort((a,b)=>b[1].returnedQty-a[1].returnedQty);
          // 매장 반품률 — 매장 판매 데이터에서 별도 계산 (배송 카운트와 무관)
          const storeShippedQty =filteredStoreSales.filter(r=>r.status==="배송").reduce((s,r)=>s+(r.qty||1),0);
          const storeReturnedQty=filteredStoreSales.filter(r=>r.status==="반품").reduce((s,r)=>s+(r.qty||1),0);
          const storeRate=storeShippedQty>0?(storeReturnedQty/storeShippedQty*100):0;
          const hasStore=storeShippedQty>0||storeReturnedQty>0;
          // top return products — 반품 수량(장) 기준 (온라인 + 매장)
          const byProd={};
          filteredOrders.filter(r=>!isOff(r)&&r.status==="반품").forEach(r=>{
            const k=(r.product_name||"미분류")+(r.option_name?" / "+r.option_name:"");
            if(!byProd[k]) byProd[k]=0;
            byProd[k]+=(r.qty||1);
          });
          filteredStoreSales.filter(r=>r.status==="반품").forEach(r=>{
            const k=(r.product_name||"미분류")+(r.option_name?" / "+r.option_name:"");
            if(!byProd[k]) byProd[k]=0;
            byProd[k]+=(r.qty||1);
          });
          const prodRows=Object.entries(byProd).sort((a,b)=>b[1]-a[1]).slice(0,20);
          modalContent=(
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:14,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: <b>주문·배송 업로드 데이터</b> (온라인 채널, status="배송"·"반품") · <b>매장 판매 데이터</b> (오프라인 스토어, status="배송"·"반품")<br/>
                <b>반품률</b> = <b>반품 수량(장) ÷ 배송 수량(장)</b> × 100 (동기간 내, 장수 단위 · 온라인과 매장 각각 별도 계산)<br/>
                ※ 반품은 배송 완료 이후부터 접수되므로 단기간(어제·7일) 기준은 0%에 가깝게 보일 수 있습니다. 최근 한달·3개월이 더 정확합니다.
              </div>
            <div style={{display:"grid",gridTemplateColumns:["yd","7d"].includes(period)?"1fr":"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>채널별 반품률</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>채널</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송 수량(장)</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품 수량(장)</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품률</th>
                  </tr></thead>
                  <tbody>
                    {chRows.map(([ch,d])=>{
                      const rate=d.shippedQty>0?(d.returnedQty/d.shippedQty*100):0;
                      return(
                        <tr key={ch} style={{borderBottom:`1px solid ${D.border}`}}>
                          <td style={{padding:"5px 7px",fontWeight:600}}>{ch}</td>
                          <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{d.shippedQty.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{d.returnedQty.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"5px 7px",fontWeight:700,color:rate>10?D.red:D.textSub}}>{rate.toFixed(1)}%</td>
                        </tr>
                      );
                    })}
                    {hasStore&&(
                      <tr style={{borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt}}>
                        <td style={{padding:"5px 7px",fontWeight:600}}>오프라인 스토어</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{storeShippedQty.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{storeReturnedQty.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",fontWeight:700,color:storeRate>10?D.red:D.textSub}}>{storeRate.toFixed(1)}%</td>
                      </tr>
                    )}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계 (반품률 = 온라인만)</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{stats.totalDeliveredQty.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{stats.totalReturnedQty.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px"}}>{stats.returnRate}%</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              {!["yd","7d"].includes(period)&&<div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>반품 Top 상품 (장수 기준)</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>상품/옵션</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품 수량(장)</th>
                  </tr></thead>
                  <tbody>
                    {prodRows.map(([k,qty])=>(
                      <tr key={k} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red,fontWeight:600}}>{qty.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>}
            </div>
            </div>
          );
        }

        /* ── 입고 수량 ── */
        else if(kpiModal==="stock"){
          modalTitle="입고 수량 소스";
          const byDate={};
          filteredStocks.forEach(r=>{
            const d=r.upload_date||"—";
            if(!byDate[d]) byDate[d]={qty:0,skus:0};
            byDate[d].qty+=(r.qty||0);
            byDate[d].skus++;
          });
          const dateRows=Object.entries(byDate).sort((a,b)=>b[0]>a[0]?1:-1);
          // top stocked products
          const byProd={};
          filteredStocks.forEach(r=>{
            const k=(r.product_name||"미분류")+(r.option_name?" / "+r.option_name:"");
            if(!byProd[k]) byProd[k]=0;
            byProd[k]+=(r.qty||0);
          });
          const prodRows=Object.entries(byProd).sort((a,b)=>b[1]-a[1]).slice(0,20);
          modalContent=(
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:16,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: <b>입고 업로드 데이터</b> (stock_uploads)<br/>
                <b>입고 수량</b> = SUM(qty) 업로드 행 전체<br/>
                <b>SKU수</b> = COUNT(행) — 업로드된 상품·옵션 라인 수
              </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>업로드 날짜별</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>날짜</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>수량</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>SKU수</th>
                  </tr></thead>
                  <tbody>
                    {dateRows.map(([d,v])=>(
                      <tr key={d} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{d}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",fontWeight:600,color:D.blue}}>{v.qty.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.textSub}}>{v.skus.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.blue}}>{stats.totalStock.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.textSub}}>{filteredStocks.length.toLocaleString()}</td>
                    </tr>
                  </tbody>
                </table>
                <div style={{marginTop:10,padding:"8px 10px",background:D.surfaceAlt,borderRadius:6,fontSize:11,color:D.textMeta}}>
                  ※ 동일 SKU는 최신 업로드 기준으로 1건만 반영됩니다.
                </div>
              </div>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>입고 Top 상품</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>상품/옵션</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>수량</th>
                  </tr></thead>
                  <tbody>
                    {prodRows.map(([k,qty])=>(
                      <tr key={k} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.blue,fontWeight:600}}>{qty.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            </div>
            </div>
          );
        }

        return(
          <div onClick={()=>setKpiModal(null)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:2000,
              display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:D.surface,borderRadius:14,padding:"24px 28px",
                width:"min(860px,95vw)",maxHeight:"85vh",overflowY:"auto",
                boxShadow:"0 8px 40px rgba(0,0,0,0.22)",position:"relative"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
                <div style={{fontWeight:700,fontSize:15,color:D.black,letterSpacing:"0.04em"}}>{modalTitle}</div>
                <button onClick={()=>setKpiModal(null)}
                  style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
                    padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
              </div>
              {modalContent}
            </div>
          </div>
        );
      })()}

      {/* 객단가 소스 모달 */}
      {aovModal&&(()=>{
        const ch=aovModal;
        const isOffline=ch==="오프라인 스토어";
        let rows=[];
        if(isOffline){
          const offMap={};
          filteredStoreSales.filter(r=>r.status==="배송"&&r.order_id).forEach(r=>{
            if(!offMap[r.order_id]) offMap[r.order_id]={order_no:r.order_id,date:r.sale_date,amount:0,store:r.store_name||""};
            offMap[r.order_id].amount+=(r.amount||0);
          });
          rows=Object.values(offMap).sort((a,b)=>a.date<b.date?1:-1);
        } else {
          const orderMap=stats.chOrderAmt?.[ch]||{};
          // order_date lookup from filteredOrders
          const dateLookup={};
          filteredOrders.filter(r=>(r.channel||"")===(ch||"")).forEach(r=>{
            const oid=r.order_no||r.order_id;
            if(oid&&!dateLookup[oid]) dateLookup[oid]=r.order_date||"";
          });
          rows=Object.entries(orderMap).map(([oid,amt])=>({
            order_no:oid,date:dateLookup[oid]||"—",amount:amt
          })).sort((a,b)=>a.date<b.date?1:-1);
        }
        const totalAmt=rows.reduce((s,r)=>s+(r.amount||0),0);
        const aov=rows.length>0?Math.round(totalAmt/rows.length):0;
        const PAYMENT_CH=new Set(["자사몰"]);
        const amtLabel=isOffline?"실판매금액 합":PAYMENT_CH.has(ch)?"결제금액":"판매가 합";
        return(
          <div onClick={()=>setAovModal(null)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:2000,
              display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:D.surface,borderRadius:14,padding:"24px 28px",
                width:"min(600px,95vw)",maxHeight:"85vh",overflowY:"auto",
                boxShadow:"0 8px 40px rgba(0,0,0,0.22)",position:"relative"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{fontWeight:700,fontSize:15,color:D.black}}>{ch} 객단가 소스</div>
                <button onClick={()=>setAovModal(null)}
                  style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
                    padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
              </div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:14,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: 주문·배송 업로드 데이터{isOffline?" / 매장 판매 업로드 데이터":""}<br/>
                금액 기준: {amtLabel}<br/>
                객단가 = {totalAmt.toLocaleString()}원 ÷ {rows.length}건 = <b>{fmtWon(aov)}</b>
              </div>
              <table style={{width:"100%",fontSize:12,borderCollapse:"collapse"}}>
                <thead>
                  <tr style={{borderBottom:`1px solid ${D.border}`}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500,color:D.textMeta}}>주문번호</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500,color:D.textMeta}}>주문일</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500,color:D.textMeta}}>{amtLabel}</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0,200).map((r,i)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${D.border}`}}>
                      <td style={{padding:"5px 7px",color:D.textMeta,maxWidth:200,
                        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.order_no}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{r.date}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",fontWeight:600}}>{(r.amount||0).toLocaleString()}원</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700,background:D.bg}}>
                    <td style={{padding:"7px"}}>합계 {rows.length}건</td>
                    <td/>
                    <td style={{textAlign:"right",padding:"7px"}}>{totalAmt.toLocaleString()}원</td>
                  </tr>
                </tfoot>
              </table>
              {rows.length>200&&<div style={{fontSize:11,color:D.textMeta,marginTop:8}}>상위 200건만 표시 (전체 {rows.length}건)</div>}
            </div>
          </div>
        );
      })()}

      {/* 채널별 주문 수 소스 모달 — 판매처 상세에서 '주문 수' 셀 클릭 시 */}
      {chOrderModal&&(()=>{
        const ch=chOrderModal;
        const isOffline=ch==="오프라인 스토어";
        // 매장은 storeSales 원본, 온라인은 filteredOrders에서 채널 매칭
        const sourceLabel = isOffline ? "매장 판매 업로드 데이터" : "주문·배송 업로드 데이터";
        const idKey = r => r.order_no || r.order_id || "";
        const dateKey = r => isOffline ? (r.sale_date||"—") : (r.order_date||"—");
        const baseRows = isOffline
          ? filteredStoreSales
          : filteredOrders.filter(r=>(r.channel||"")===ch);
        // 주문번호별 그룹화 — 상태별 라인 수/장수/금액 합산
        const byOid={};
        baseRows.forEach(r=>{
          const oid=idKey(r);
          if(!oid) return;
          if(!byOid[oid]) byOid[oid]={oid,date:dateKey(r),shippedQty:0,returnedQty:0,otherQty:0,lines:0};
          byOid[oid].lines++;
          const q=r.qty||1;
          if(r.status==="배송") byOid[oid].shippedQty+=q;
          else if(r.status==="반품") byOid[oid].returnedQty+=q;
          else byOid[oid].otherQty+=q;
          // 최신 날짜 채택
          if(dateKey(r)>byOid[oid].date) byOid[oid].date=dateKey(r);
        });
        const rows=Object.values(byOid).sort((a,b)=>a.date<b.date?1:-1);
        const totalQty=rows.reduce((s,r)=>s+r.shippedQty+r.returnedQty+r.otherQty,0);
        return(
          <div onClick={()=>setChOrderModal(null)}
            style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:2000,
              display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
            <div onClick={e=>e.stopPropagation()}
              style={{background:D.surface,borderRadius:14,padding:"24px 28px",
                width:"min(720px,95vw)",maxHeight:"85vh",overflowY:"auto",
                boxShadow:"0 8px 40px rgba(0,0,0,0.22)",position:"relative"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{fontWeight:700,fontSize:15,color:D.black}}>{ch} 주문 수 소스</div>
                <button onClick={()=>setChOrderModal(null)}
                  style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
                    padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
              </div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:14,lineHeight:1.8,
                background:D.bg,borderRadius:6,padding:"8px 12px"}}>
                소스: {sourceLabel}<br/>
                계산: 모든 상태(배송/반품/교환 등) 포함, COUNT(DISTINCT 주문번호) = <b>{rows.length.toLocaleString()}건</b> · SUM(qty) = <b>{totalQty.toLocaleString()}장</b>
              </div>
              <table style={{width:"100%",fontSize:12,borderCollapse:"collapse"}}>
                <thead>
                  <tr style={{borderBottom:`1px solid ${D.border}`}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500,color:D.textMeta}}>주문번호</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500,color:D.textMeta}}>주문일</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500,color:D.green}}>배송 장수</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500,color:D.red}}>반품 장수</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0,200).map((r,i)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${D.border}`}}>
                      <td style={{padding:"5px 7px",color:D.textMeta,maxWidth:240,
                        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.oid}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.textMeta}}>{r.date}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{r.shippedQty||"—"}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:r.returnedQty>0?D.red:D.textMeta}}>{r.returnedQty||"—"}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700,background:D.bg}}>
                    <td style={{padding:"7px"}}>합계 {rows.length.toLocaleString()}건</td>
                    <td/>
                    <td style={{textAlign:"right",padding:"7px",color:D.green}}>{rows.reduce((s,r)=>s+r.shippedQty,0).toLocaleString()}</td>
                    <td style={{textAlign:"right",padding:"7px",color:D.red}}>{rows.reduce((s,r)=>s+r.returnedQty,0).toLocaleString()}</td>
                  </tr>
                </tfoot>
              </table>
              {rows.length>200&&<div style={{fontSize:11,color:D.textMeta,marginTop:8}}>상위 200건만 표시 (전체 {rows.length}건)</div>}
            </div>
          </div>
        );
      })()}

    </div>
  );
}

// ─────────────────────────────────────────────
// INVENTORY AGING BOARD
// ─────────────────────────────────────────────
const AGING_STAGES=[
  {key:"normal",  label:"정상",   min:0,  max:30,  color:"#22c55e", bg:"#f0fdf4"},
  {key:"caution", label:"주의",   min:31, max:90,  color:"#f59e0b", bg:"#fffbeb"},
  {key:"slow",    label:"저회전", min:91, max:180, color:"#f97316", bg:"#fff7ed"},
  {key:"dead",    label:"악성",   min:181,max:Infinity, color:"#ef4444", bg:"#fef2f2"},
];
function getAgingStage(days){
  return AGING_STAGES.find(s=>days>=s.min&&days<=s.max)||AGING_STAGES[3];
}

function InventoryAgingUploader({ onDone }){
  const today=new Date().toISOString().slice(0,10);
  const [diagDate,setDiagDate]=useState(today);
  const [fileName,setFileName]=useState("");
  const [preview,setPreview]=useState(null);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [infoOpen,setInfoOpen]=useState(false);

  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    parseAnyFile(file,{header:true,skipEmptyLines:true},({data})=>{
        try{
          if(!data.length) throw new Error(uploadErrParse("파일에 데이터 행이 없습니다"));
          const cols=Object.keys(data[0]);
          const lc=cols.map(c=>c.toLowerCase().replace(/[\s_]/g,""));
          const find=(...kws)=>{const i=lc.findIndex(c=>kws.some(k=>c.includes(k)));return i>=0?cols[i]:null;};
          const prodCol=find("상품명","상품","product","품명");
          const optCol=find("옵션","option","사이즈","색상");
          const qtyCol=find("수량","qty","quantity","개수","재고");
          const recCol=find("처음입고일","입고일","최초입고","first","입고");
          const shipCol=find("마지막배송일","최종배송일","마지막출고","last","배송일","출고일");
          const missingCols=[];
          if(!prodCol) missingCols.push("상품명");
          if(!recCol)  missingCols.push("처음입고일");
          if(missingCols.length){
            throw new Error(uploadErrColumns({
              missing:missingCols,
              required:["상품명","옵션","수량","처음입고일","마지막배송일"],
              headers:cols,
            }));
          }
          const rows=data.map(r=>({
            product_name:String(r[prodCol]||"").trim(),
            option_name:optCol?String(r[optCol]||"").trim():"",
            qty:toNum(r[qtyCol]||"0"),
            first_received_date:toDate(r[recCol]),
            last_shipped_date:shipCol&&r[shipCol]?toDate(r[shipCol]):null,
            diagnosis_date:diagDate,
          })).filter(r=>r.product_name&&r.first_received_date);
          if(!rows.length) throw new Error("파싱된 행이 0건입니다. '상품명'과 '처음입고일' 둘 다 값이 있는 행이 1개 이상 있어야 합니다.");
          setPreview(rows);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e.message?.message||String(e.message||e)}));
  },[diagDate]);

  const [existingCount,setExistingCount]=useState(0);
  // 파일 파싱 후 기존 데이터 수 조회 (갈음 안내용)
  useEffect(()=>{
    if(!preview) {setExistingCount(0);return;}
    (async()=>{
      const db=await getSupabase();
      const {count}=await db.from("inventory_aging").select("*",{count:"exact",head:true});
      setExistingCount(count||0);
    })();
  },[preview]);

  const handleUpload=async()=>{
    if(!preview?.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    const {error:delErr}=await db.from("inventory_aging").delete().gte("diagnosis_date","2000-01-01");
    if(delErr){setResult({type:"error",msg:"삭제 실패: "+delErr.message});setLoading(false);return;}
    for(let i=0;i<preview.length;i+=500){
      const {error}=await db.from("inventory_aging").insert(preview.slice(i,i+500));
      if(error){setResult({type:"error",msg:"삽입 실패: "+error.message});setLoading(false);return;}
    }
    try{localStorage.setItem("merryon_aging_date",diagDate);}catch{}
    setResult({type:"success",msg:`${preview.length}건 저장 완료 (기존 ${existingCount}건 대체)`});
    setPreview(null); setFileName(""); setExistingCount(0);
    onDone?.();
    setLoading(false);
  };

  return (
    <Card style={{marginBottom:20}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
        <SecTitle>재고 현황 업로드</SecTitle>
        <InfoBtn onClick={()=>setInfoOpen(true)}/>
        <span style={{fontSize:10,color:D.textMeta,marginLeft:4}}>
          진단 날짜: <input type="date" value={diagDate} onChange={e=>setDiagDate(e.target.value)}
            style={{fontSize:10,border:`1px solid ${D.border}`,borderRadius:4,padding:"1px 5px",color:D.black}}/>
        </span>
      </div>
      <div style={{fontSize:11,color:D.textMeta,marginBottom:12,lineHeight:1.7,padding:"8px 10px",
        background:D.bg,borderRadius:6,border:`1px solid ${D.border}`}}>
        ⚠ 해당 데이터는 현재고를 기반으로 진단해야하므로 누적 데이터가 유의미하지 않습니다.<br/>
        필요 시 <strong>정상재고 1 이상</strong>으로 검색한 애널리틱스용 파일을 업로드해주세요.
      </div>
      <DropZone onFile={handleFile} fileName={fileName} label="재고 파일 업로드"
        columns="상품명 · 옵션 · 수량 · 처음입고일 · 마지막배송일"/>
      {(preview||loading)&&(
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginTop:8}}>
          {preview&&<span style={{fontSize:11,color:D.textMeta}}>{preview.length}건 파싱됨</span>}
          {preview&&(
            <button onClick={handleUpload} disabled={loading}
              style={{background:D.blue,color:"#fff",border:"none",borderRadius:6,
                padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>
              {loading?"저장 중…":"저장"}
            </button>
          )}
        </div>
      )}
      {preview&&existingCount>0&&<Alert type="warn" msg={uploadReplaceWarn(existingCount,"재고 에이징 전체")}/>}
      {result&&<Alert type={result.type==="error"?"error":"success"} msg={result.msg}/>}
      {infoOpen&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.35)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center"}}
          onClick={()=>setInfoOpen(false)}>
          <div style={{background:"#fff",borderRadius:12,padding:28,maxWidth:340,width:"90%"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:16}}>재고 에이징 기준</div>
            {AGING_STAGES.map(s=>(
              <div key={s.key} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                <div style={{width:12,height:12,borderRadius:"50%",background:s.color,flexShrink:0}}/>
                <div style={{fontSize:13}}>
                  <strong style={{color:s.color}}>{s.label}</strong>
                  <span style={{color:D.textMeta,marginLeft:6}}>
                    {s.max===Infinity?`${s.min}일 이상`:`${s.min}~${s.max}일`}
                  </span>
                </div>
              </div>
            ))}
            <div style={{fontSize:11,color:D.textMeta,marginTop:12,lineHeight:1.6}}>
              기준: <strong>마지막 배송일</strong> 기준, 없는 경우 <strong>처음 입고일</strong> 기준
            </div>
            <button onClick={()=>setInfoOpen(false)}
              style={{marginTop:16,width:"100%",padding:"8px",background:D.black,color:"#fff",
                border:"none",borderRadius:6,fontSize:13,cursor:"pointer"}}>닫기</button>
          </div>
        </div>
      )}
    </Card>
  );
}

function InventoryAgingBoard({ diagDate }){
  const [rows,setRows]=useState([]);
  const [selectedStage,setSelectedStage]=useState(null);
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    (async()=>{
      setLoading(true);
      const db=await getSupabase();
      const {data}=await db.from("inventory_aging").select("*");
      setRows(data||[]);
      setLoading(false);
    })();
  },[diagDate]);

  const enriched=useMemo(()=>{
    if(!diagDate) return [];
    const base=new Date(diagDate);
    return rows.map(r=>{
      const refDate=r.last_shipped_date||r.first_received_date;
      const days=refDate?Math.max(0,Math.round((base-new Date(refDate))/(1000*60*60*24))):0;
      return {...r,days,stage:getAgingStage(days)};
    });
  },[rows,diagDate]);

  const stageSummary=useMemo(()=>
    AGING_STAGES.map(s=>({
      ...s,
      items:enriched.filter(r=>r.stage.key===s.key),
      totalQty:enriched.filter(r=>r.stage.key===s.key).reduce((sum,r)=>sum+(r.qty||0),0),
    }))
  ,[enriched]);

  const shown=selectedStage
    ? enriched.filter(r=>r.stage.key===selectedStage).sort((a,b)=>b.days-a.days)
    : [];

  if(loading) return <div style={{padding:20,color:D.textMeta,fontSize:12}}>불러오는 중…</div>;
  if(!enriched.length) return null;

  return (
    <Card>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
        <SecTitle>재고 에이징 보드</SecTitle>
        {diagDate&&<span style={{fontSize:11,color:D.textMeta}}>진단 기준일: {diagDate}</span>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
        {stageSummary.map(s=>(
          <div key={s.key} onClick={()=>setSelectedStage(selectedStage===s.key?null:s.key)}
            style={{padding:"12px 14px",borderRadius:8,background:selectedStage===s.key?s.color:s.bg,
              border:`1.5px solid ${selectedStage===s.key?s.color:s.color+"44"}`,
              cursor:"pointer",transition:"all 0.15s"}}>
            <div style={{fontSize:11,fontWeight:700,color:selectedStage===s.key?"#fff":s.color,marginBottom:4}}>
              {s.label}
              <span style={{fontSize:10,fontWeight:400,marginLeft:4}}>
                {s.max===Infinity?`${s.min}일+`:`${s.min}~${s.max}일`}
              </span>
            </div>
            <div style={{fontSize:18,fontWeight:800,color:selectedStage===s.key?"#fff":s.color}}>{s.items.length}<span style={{fontSize:11,fontWeight:400,marginLeft:2}}>종</span></div>
            <div style={{fontSize:11,color:selectedStage===s.key?"rgba(255,255,255,0.85)":D.textMeta,marginTop:2}}>
              재고 {s.totalQty.toLocaleString()}개
            </div>
          </div>
        ))}
      </div>
      {selectedStage&&shown.length>0&&(
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{borderBottom:`2px solid ${D.border}`}}>
              {["상품명","옵션","수량","처음입고일","마지막배송일","경과일","단계"].map(h=>(
                <th key={h} style={{padding:"7px 9px",textAlign:h==="수량"||h==="경과일"?"right":"left",
                  color:D.textMeta,fontWeight:600,whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {shown.map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${D.border}`}}>
                  <td style={{padding:"7px 9px",fontWeight:500,maxWidth:200,overflow:"hidden",
                    textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.product_name}</td>
                  <td style={{padding:"7px 9px",color:D.textMeta,maxWidth:120,overflow:"hidden",
                    textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.option_name||"—"}</td>
                  <td style={{padding:"7px 9px",textAlign:"right",fontWeight:600}}>{(r.qty||0).toLocaleString()}</td>
                  <td style={{padding:"7px 9px",color:D.textMeta,whiteSpace:"nowrap"}}>{r.first_received_date||"—"}</td>
                  <td style={{padding:"7px 9px",color:D.textMeta,whiteSpace:"nowrap"}}>{r.last_shipped_date||"—"}</td>
                  <td style={{padding:"7px 9px",textAlign:"right",fontWeight:700,color:r.stage.color}}>{r.days}일</td>
                  <td style={{padding:"7px 9px"}}>
                    <span style={{background:r.stage.bg,color:r.stage.color,border:`1px solid ${r.stage.color}55`,
                      borderRadius:4,padding:"2px 7px",fontSize:10,fontWeight:700}}>{r.stage.label}</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

// ─────────────────────────────────────────────
// LOGISTICS FLOW PAGE
// ─────────────────────────────────────────────
function LogisticsFlow({ orders, stocks, ts }) {
  const [period,setPeriod]=useState("3m");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [calOpenFor,setCalOpenFor]=useState(null);
  const [sankeyFull,setSankeyFull]=useState(false);
  const [flowSort,setFlowSort]=useState("stock"); // "stock"|"shipped"|"returned"
  const [sankeyLimit,setSankeyLimit]=useState(30);
  const [tableLimit,setTableLimit]=useState(30);
  const [tablePeriod,setTablePeriod]=useState("1m");
  const [agingKey,setAgingKey]=useState(0);
  const [agingDiagDate,setAgingDiagDate]=useState(()=>{
    try{return localStorage.getItem("merryon_aging_date")||new Date().toISOString().slice(0,10);}
    catch{return new Date().toISOString().slice(0,10);}
  });

  const filteredOrders=useMemo(()=>filterByDate(orders,"order_date",period,customStart,customEnd),[orders,period,customStart,customEnd]);
  const filteredTableOrders=useMemo(()=>filterByDate(orders,"order_date",tablePeriod,"",""),[orders,tablePeriod]);
  const filteredTableStocks=useMemo(()=>
    filterByDate(stocks,"upload_date",tablePeriod,"","")
  ,[stocks,tablePeriod]);

  return (
    <div style={{padding:"20px 24px",maxWidth:1600,margin:"0 auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
        <div>
          <div style={{color:D.black,fontWeight:600,fontSize:15,display:"flex",alignItems:"center",gap:8}}>
            물류 플로우 <UpdatedAt ts={ts.orders||ts.stock}/>
          </div>
        </div>
        <CalDrop id="logistics" period={period} setPeriod={setPeriod}
          presets={[["1m","1개월"],["3m","3개월"],["6m","6개월"],["all","전체"]]}
          start={customStart} setStart={setCustomStart}
          end={customEnd} setEnd={setCustomEnd}
          calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
      </div>

      <div style={{display:"flex",gap:12,marginBottom:14,flexWrap:"wrap"}}>
        {[{label:"배송",color:D.green},{label:"반품",color:D.red},{label:"교환",color:D.amber}]
          .map(({label,color})=>(
            <div key={label} style={{display:"flex",alignItems:"center",gap:5,fontSize:11}}>
              <div style={{width:8,height:8,borderRadius:"50%",background:color}}/>
              <span style={{color:D.textSub}}>{label}</span>
            </div>
          ))}
      </div>

      <Card style={{marginBottom:12,padding:"12px 16px"}}>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:8,gap:4}}>
          {[30,50,100].map(n=>(
            <button key={n} onClick={()=>setSankeyLimit(n)}
              style={{background:sankeyLimit===n?D.black:"transparent",
                color:sankeyLimit===n?"#fff":D.textSub,
                border:`1px solid ${sankeyLimit===n?D.black:D.border}`,
                borderRadius:5,padding:"3px 10px",fontSize:10,cursor:"pointer",
                fontWeight:sankeyLimit===n?600:400}}>
              {n}개
            </button>
          ))}
        </div>
        <ProductSankey stockRows={stocks} orderRows={orders} period={period} customStart={customStart} customEnd={customEnd} limit={sankeyLimit}/>
      </Card>

      {/* 산키 전체화면 오버레이 */}
      {sankeyFull&&(
        <div style={{position:"fixed",top:0,left:0,width:"100vw",height:"100vh",
          background:"#fff",zIndex:9999,overflow:"auto",boxSizing:"border-box"}}>
          <div style={{position:"sticky",top:0,background:"#fff",borderBottom:`1px solid ${D.border}`,
            padding:"10px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",zIndex:1}}>
            <span style={{fontWeight:600,fontSize:14,color:D.black}}>물류 플로우</span>
            <button onClick={()=>setSankeyFull(false)}
              style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"5px 14px",fontSize:11,cursor:"pointer",color:D.text}}>
              닫기 ✕
            </button>
          </div>
          <div style={{padding:"16px 24px"}}>
            <ProductSankey stockRows={stocks} orderRows={orders} period={period} customStart={customStart} customEnd={customEnd} limit={sankeyLimit}/>
          </div>
        </div>
      )}

      {/* 상품별 흐름 요약 */}
      {orders.length>0&&(
        <Card>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:8}}>
            <SecTitle ts={ts.orders}>상품별 흐름 요약</SecTitle>
            <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
              {[["1m","1개월"],["3m","3개월"],["6m","6개월"],["all","1년+"]].map(([k,l])=>(
                <button key={k} onClick={()=>setTablePeriod(k)}
                  style={{background:tablePeriod===k?D.black:"transparent",
                    color:tablePeriod===k?"#fff":D.textSub,
                    border:`1px solid ${tablePeriod===k?D.black:D.border}`,
                    borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer",
                    fontWeight:tablePeriod===k?600:400}}>
                  {l}
                </button>
              ))}
              <div style={{width:1,background:D.border,margin:"0 2px"}}/>
              {[["stock","입고 수 높은 순"],["shipped","배송 많은 순"],["returned","반품 많은 순"]].map(([k,l])=>(
                <button key={k} onClick={()=>setFlowSort(k)}
                  style={{background:flowSort===k?D.black:"transparent",
                    color:flowSort===k?"#fff":D.textSub,
                    border:`1px solid ${flowSort===k?D.black:D.border}`,
                    borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer",
                    fontWeight:flowSort===k?600:400}}>
                  {l}
                </button>
              ))}
              <div style={{width:1,background:D.border,margin:"0 2px"}}/>
              {[30,50,100].map(n=>(
                <button key={n} onClick={()=>setTableLimit(n)}
                  style={{background:tableLimit===n?D.black:"transparent",
                    color:tableLimit===n?"#fff":D.textSub,
                    border:`1px solid ${tableLimit===n?D.black:D.border}`,
                    borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer",
                    fontWeight:tableLimit===n?600:400}}>
                  {n}개
                </button>
              ))}
            </div>
          </div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
                {["#","상품명","입고","배송","반품","반품률"].map(h=>(
                  <th key={h} style={{padding:"7px 9px",textAlign:h==="상품명"?"left":"right",
                    color:D.textMeta,fontWeight:400,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {(()=>{
                  const prodMap={};
                  filteredTableStocks.forEach(r=>{
                    const k=r.product_name||"미분류";
                    if(!prodMap[k]) prodMap[k]={name:k,stock:0,shipped:0,returned:0};
                    prodMap[k].stock+=(r.qty||0);
                  });
                  filteredTableOrders.forEach(r=>{
                    const k=r.product_name||"미분류";
                    if(!prodMap[k]) prodMap[k]={name:k,stock:0,shipped:0,returned:0};
                    if(r.status==="배송") prodMap[k].shipped++;
                    if(r.status==="반품") prodMap[k].returned++;
                  });
                  const sortFn=flowSort==="stock"?(a,b)=>b.stock-a.stock
                    :flowSort==="returned"?(a,b)=>b.returned-a.returned
                    :(a,b)=>b.shipped-a.shipped;
                  return Object.values(prodMap)
                    .filter(p=>p.shipped>0||p.stock>0)
                    .sort(sortFn)
                    .slice(0,tableLimit)
                    .map((p,i)=>{
                      const total=p.shipped+p.returned;
                      const rr=total>0?(p.returned/total*100).toFixed(1):"0.0";
                      return(
                        <tr key={p.name} style={{borderBottom:`1px solid ${D.border}`}}>
                          <td style={{padding:"6px 9px",color:D.textMeta,textAlign:"right"}}>{i+1}</td>
                          <td style={{padding:"6px 9px",fontWeight:i<3?700:400,maxWidth:220,
                            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",
                            color:flowSort==="stock"?D.black:D.blue,
                            fontWeight:flowSort==="stock"?700:400}}>{p.stock.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",
                            color:flowSort==="shipped"?D.black:D.green,
                            fontWeight:flowSort==="shipped"?700:600}}>{p.shipped.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",
                            color:flowSort==="returned"?D.black:D.red,
                            fontWeight:flowSort==="returned"?700:400}}>{p.returned.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",fontWeight:600,
                            color:parseFloat(rr)>10?D.red:D.textSub}}>{rr}%</td>
                        </tr>
                      );
                    });
                })()}
              </tbody>
            </table>
          </div>
        </Card>
      )}

    </div>
  );
}

// ─────────────────────────────────────────────
// PROMO FLOW PAGE
// ─────────────────────────────────────────────
const PROMO_PLATFORMS=["자사몰","29CM","무신사","오프라인 스토어"];

function SmallDateButtonPicker({value,onChange}){
  const todayStr=new Date().toISOString().slice(0,10);
  const datePart=(value&&value.length>=10)?value.slice(0,10):todayStr;
  const [y,m,d]=datePart.split("-").map(Number);
  const [halfYear,setHalfYear]=useState(m<=6?0:1);
  const setDate=(ny,nm,nd)=>{
    const maxD=new Date(ny,nm,0).getDate();
    const cd=Math.min(nd,maxD);
    const ds=`${ny}-${String(nm).padStart(2,"0")}-${String(cd).padStart(2,"0")}`;
    const tp=(value&&value.includes("T"))?value.slice(10):"";
    onChange(ds+tp);
  };
  const daysInMonth=new Date(y,m,0).getDate();
  const monthSet=halfYear===0?[1,2,3,4,5,6]:[7,8,9,10,11,12];
  const bNone={border:"none",background:"transparent",cursor:"pointer",fontFamily:"inherit"};
  const circSel={background:D.black,color:"#fff",fontWeight:700,borderRadius:"50%"};
  const circDef={background:"transparent",color:D.textSub};
  const timePart=(value&&value.includes("T"))?value.slice(11,16):"";
  return(
    <div style={{userSelect:"none",minWidth:220}}>
      <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:4}}>
        <button onClick={()=>setDate(y-1,m,d)} style={{...bNone,fontSize:12,color:D.textSub,padding:"1px 4px"}}>◀</button>
        <span style={{fontSize:12,fontWeight:700,color:D.text,minWidth:36,textAlign:"center"}}>{y}년</span>
        <button onClick={()=>setDate(y+1,m,d)} style={{...bNone,fontSize:12,color:D.textSub,padding:"1px 4px"}}>▶</button>
        <button onClick={()=>setHalfYear(v=>v===0?1:0)}
          style={{...bNone,fontSize:11,color:D.textSub,background:D.surfaceAlt,borderRadius:4,padding:"2px 6px",marginLeft:"auto"}}>
          {halfYear===0?"1~6월":"7~12월"}
        </button>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:2,marginBottom:4}}>
        {monthSet.map(mo=>(
          <button key={mo} onClick={()=>setDate(y,mo,Math.min(d,new Date(y,mo,0).getDate()))}
            style={{...bNone,fontSize:11,padding:"2px 5px",borderRadius:"50%",...(mo===m?circSel:circDef)}}>{mo}월</button>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:1}}>
        {Array.from({length:daysInMonth},(_,i)=>i+1).map(dd=>(
          <button key={dd} onClick={()=>setDate(y,m,dd)}
            style={{...bNone,fontSize:11,padding:"2px 0",textAlign:"center",borderRadius:"50%",...(dd===d?circSel:circDef)}}>{dd}</button>
        ))}
      </div>
      <div style={{marginTop:6}}>
        <div style={{fontSize:11,color:D.textMeta,marginBottom:3}}>시간</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:2}}>
          {Array.from({length:24},(_,h)=>{
            const hStr=String(h).padStart(2,"0")+":00";
            const sel=timePart===hStr;
            return(
              <button key={h} onClick={()=>onChange(datePart+"T"+hStr)}
                style={{...bNone,fontSize:10,padding:"3px 0",textAlign:"center",borderRadius:3,
                  background:sel?D.black:"transparent",color:sel?"#fff":D.textSub,
                  border:`1px solid ${sel?D.black:D.border}`}}>
                {h}시
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function SubmitEodPicker({value,onChange}){
  const [open,setOpen]=useState(false);
  const ref=useRef(null);
  useEffect(()=>{
    if(!open)return;
    const handler=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};
    document.addEventListener("mousedown",handler);
    return()=>document.removeEventListener("mousedown",handler);
  },[open]);
  const display=value?value.replace("T"," "):"날짜/시간 선택";
  return(
    <div style={{position:"relative"}} ref={ref}>
      <button onClick={()=>setOpen(v=>!v)}
        style={{width:"100%",textAlign:"left",background:"transparent",border:`1px solid ${D.border}`,
          borderRadius:5,padding:"7px 10px",fontSize:13,color:value?D.text:D.textMeta,
          cursor:"pointer",fontFamily:"'Pretendard','Noto Sans KR',sans-serif",boxSizing:"border-box"}}>
        {display}
      </button>
      {open&&(
        <div style={{position:"absolute",zIndex:200,background:D.surface,border:`1px solid ${D.border}`,
          borderRadius:8,padding:12,top:"calc(100% + 4px)",left:0,
          boxShadow:"0 4px 20px rgba(0,0,0,0.12)"}}>
          <SmallDateButtonPicker value={value||""} onChange={v=>{onChange(v);}}/>
        </div>
      )}
    </div>
  );
}

function DateButtonPicker({value,onChange}){
  const todayStr=new Date().toISOString().slice(0,10);
  const datePart=(value&&value.length>=10)?value.slice(0,10):todayStr;
  const [y,m,d]=datePart.split("-").map(Number);
  const [halfYear,setHalfYear]=useState(m<=6?0:1);
  const setDate=(ny,nm,nd)=>{
    const maxD=new Date(ny,nm,0).getDate();
    const cd=Math.min(nd,maxD);
    const ds=`${ny}-${String(nm).padStart(2,"0")}-${String(cd).padStart(2,"0")}`;
    const tp=(value&&value.includes("T"))?value.slice(10):"";
    onChange(ds+tp);
  };
  const daysInMonth=new Date(y,m,0).getDate();
  const monthSet=halfYear===0?[1,2,3,4,5,6]:[7,8,9,10,11,12];
  const bNone={border:"none",background:"transparent",cursor:"pointer",fontFamily:"inherit"};
  const circSel={background:D.black,color:"#fff",fontWeight:700,borderRadius:"50%"};
  const circDef={background:"transparent",color:D.textSub};
  return(
    <div style={{userSelect:"none"}}>
      <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:4}}>
        <button onClick={()=>setDate(y-1,m,d)} style={{...bNone,fontSize:13,color:D.textSub,padding:"1px 6px"}}>◀</button>
        <span style={{fontSize:13,fontWeight:700,color:D.text,minWidth:44,textAlign:"center"}}>{y}년</span>
        <button onClick={()=>setDate(y+1,m,d)} style={{...bNone,fontSize:13,color:D.textSub,padding:"1px 6px"}}>▶</button>
        <button onClick={()=>setHalfYear(v=>v===0?1:0)}
          style={{...bNone,fontSize:11,color:D.textSub,background:D.surfaceAlt,borderRadius:4,padding:"2px 6px",marginLeft:"auto"}}>
          {halfYear===0?"1~6월":"7~12월"}
        </button>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:2,marginBottom:4}}>
        {monthSet.map(mo=>(
          <button key={mo} onClick={()=>setDate(y,mo,Math.min(d,new Date(y,mo,0).getDate()))}
            style={{...bNone,fontSize:12,padding:"3px 5px",borderRadius:"50%",...(mo===m?circSel:circDef)}}>{mo}월</button>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
        {Array.from({length:daysInMonth},(_,i)=>i+1).map(dd=>(
          <button key={dd} onClick={()=>setDate(y,m,dd)}
            style={{...bNone,fontSize:12,padding:"3px 0",textAlign:"center",borderRadius:"50%",...(dd===d?circSel:circDef)}}>{dd}</button>
        ))}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 프로모션 할인율 그리드 — Editor + 표시 컴포넌트
// 저장 구조: { products: [{group,rate,start,end}], coupons: [{rate,start,end}] }
// ─────────────────────────────────────────────
function emptyProductRow(){return{group:"",rate:"",start:"",end:""};}
function emptyCouponRow(){return{rate:"",start:"",end:""};}
function normalizePlan(p){
  return {
    products: Array.isArray(p?.products)?p.products:[],
    coupons:  Array.isArray(p?.coupons)?p.coupons:[],
  };
}

function DiscountPlanEditor({ value, onChange }) {
  const plan=normalizePlan(value);
  // 기본: 상품 3행 / 쿠폰 1행 (실제 입력값이 없으면 보이게)
  const products=plan.products.length?plan.products:[emptyProductRow(),emptyProductRow(),emptyProductRow()];
  const coupons =plan.coupons.length ?plan.coupons :[emptyCouponRow()];

  const commit=(next)=>onChange({
    products: next.products.filter(r=>r.group||r.rate||r.start||r.end),
    coupons:  next.coupons.filter(r=>r.rate||r.start||r.end),
  });

  const setProducts=(arr)=>commit({products:arr,coupons});
  const setCoupons =(arr)=>commit({products,coupons:arr});

  const cellInp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
    padding:"5px 8px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box",
    fontFamily:"'Pretendard','Noto Sans KR',sans-serif"};
  const lbl={fontSize:11,color:D.textMeta,marginBottom:4,fontWeight:600};
  const head={fontSize:11,color:D.textMeta,fontWeight:600,padding:"4px 6px",textAlign:"left"};

  return (
    <div style={{border:`1px solid ${D.border}`,borderRadius:6,padding:"10px 12px",background:D.surfaceAlt}}>
      <div style={{fontWeight:700,fontSize:13,color:D.black,marginBottom:8}}>할인율</div>

      {/* 상품 할인 */}
      <div style={{marginBottom:12}}>
        <div style={lbl}>상품 할인 <span style={{color:D.textMeta,fontWeight:400}}>· 상품군별 할인율 + 기간</span></div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>
            <th style={{...head,width:"34%"}}>상품군</th>
            <th style={{...head,width:"14%"}}>할인율(%)</th>
            <th style={{...head,width:"22%"}}>시작</th>
            <th style={{...head,width:"22%"}}>종료</th>
            <th style={{...head,width:"8%"}}/>
          </tr></thead>
          <tbody>
            {products.map((row,i)=>(
              <tr key={i}>
                <td style={{padding:"3px 4px"}}>
                  <input value={row.group} onChange={e=>{const n=[...products];n[i]={...row,group:e.target.value};setProducts(n);}}
                    style={cellInp} placeholder="예: 신상품, 전체"/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input type="number" value={row.rate} onChange={e=>{const n=[...products];n[i]={...row,rate:e.target.value};setProducts(n);}}
                    style={cellInp} placeholder="0" min="0" max="100"/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input type="date" value={row.start} onChange={e=>{const n=[...products];n[i]={...row,start:e.target.value};setProducts(n);}} style={cellInp}/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input type="date" value={row.end} onChange={e=>{const n=[...products];n[i]={...row,end:e.target.value};setProducts(n);}} style={cellInp}/>
                </td>
                <td style={{padding:"3px 4px",textAlign:"right"}}>
                  <button onClick={()=>{const n=products.filter((_,j)=>j!==i);setProducts(n.length?n:[emptyProductRow()]);}}
                    style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:14,lineHeight:1}}>✕</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <button onClick={()=>setProducts([...products,emptyProductRow()])}
          style={{marginTop:6,background:"transparent",border:`1px dashed ${D.border}`,borderRadius:5,
            padding:"4px 12px",fontSize:11,color:D.textMeta,cursor:"pointer"}}>+ 행 추가</button>
      </div>

      {/* 쿠폰 */}
      <div>
        <div style={lbl}>쿠폰 <span style={{color:D.textMeta,fontWeight:400}}>· 할인율 + 기간 (상품 할인 적용 후 추가 적용)</span></div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>
            <th style={{...head,width:"14%"}}>할인율(%)</th>
            <th style={{...head,width:"22%"}}>시작</th>
            <th style={{...head,width:"22%"}}>종료</th>
            <th style={{...head,width:"8%"}}/>
          </tr></thead>
          <tbody>
            {coupons.map((row,i)=>(
              <tr key={i}>
                <td style={{padding:"3px 4px"}}>
                  <input type="number" value={row.rate} onChange={e=>{const n=[...coupons];n[i]={...row,rate:e.target.value};setCoupons(n);}}
                    style={cellInp} placeholder="0" min="0" max="100"/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input type="date" value={row.start} onChange={e=>{const n=[...coupons];n[i]={...row,start:e.target.value};setCoupons(n);}} style={cellInp}/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input type="date" value={row.end} onChange={e=>{const n=[...coupons];n[i]={...row,end:e.target.value};setCoupons(n);}} style={cellInp}/>
                </td>
                <td style={{padding:"3px 4px",textAlign:"right"}}>
                  <button onClick={()=>{const n=coupons.filter((_,j)=>j!==i);setCoupons(n.length?n:[emptyCouponRow()]);}}
                    style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:14,lineHeight:1}}>✕</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <button onClick={()=>setCoupons([...coupons,emptyCouponRow()])}
          style={{marginTop:6,background:"transparent",border:`1px dashed ${D.border}`,borderRadius:5,
            padding:"4px 12px",fontSize:11,color:D.textMeta,cursor:"pointer"}}>+ 행 추가</button>
      </div>
    </div>
  );
}

// 표 셀에 표시되는 컴팩트 보기 — 상품 할인 행 + 쿠폰 행 + 총 할인율 범위
function DiscountPlanView({ plan }) {
  const p=normalizePlan(plan);
  const hasAny=p.products.length||p.coupons.length;
  if(!hasAny) return <span style={{color:D.textMeta,fontSize:11}}>—</span>;
  // 총 할인율 = 1 - (1 - 상품) × (1 - 쿠폰)
  // 기간별 compound 계산: 모든 일자에 대해 매일 적용 가능 할인 중 max 한 개씩 골라 compound
  const dateSet=new Set();
  [...p.products,...p.coupons].forEach(r=>{
    if(!r.start||!r.end) return;
    const s=new Date(r.start);const e=new Date(r.end);
    for(let dt=s;dt<=e;dt=new Date(dt.getTime()+86400000)){
      dateSet.add(dt.toISOString().slice(0,10));
    }
  });
  const dates=[...dateSet].sort();
  const perDay=dates.map(d=>{
    const pr=p.products.filter(r=>r.start<=d&&r.end>=d).map(r=>+r.rate||0);
    const cr=p.coupons .filter(r=>r.start<=d&&r.end>=d).map(r=>+r.rate||0);
    const maxP=pr.length?Math.max(...pr):0;
    const maxC=cr.length?Math.max(...cr):0;
    const total=Math.round((1-(1-maxP/100)*(1-maxC/100))*1000)/10; // %, 소수1자리
    return {date:d,product:maxP,coupon:maxC,total};
  });
  const maxTotal=perDay.length?Math.max(...perDay.map(d=>d.total)):0;
  return (
    <div style={{fontSize:11,lineHeight:1.55,minWidth:140}}>
      {p.products.map((r,i)=>(
        <div key={"p"+i} style={{whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
          · {r.group||"전체"} <b style={{color:D.text}}>{r.rate||0}%</b>{" "}
          <span style={{color:D.textMeta}}>{r.start?.slice(5)}~{r.end?.slice(5)}</span>
        </div>
      ))}
      {p.coupons.map((r,i)=>(
        <div key={"c"+i} style={{color:D.blue,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
          🎟 쿠폰 <b>{r.rate||0}%</b>{" "}
          <span style={{color:D.textMeta}}>{r.start?.slice(5)}~{r.end?.slice(5)}</span>
        </div>
      ))}
      {maxTotal>0&&(
        <div style={{marginTop:4,padding:"2px 6px",background:`${D.red}10`,color:D.red,
          borderRadius:4,fontWeight:600,display:"inline-block"}}>
          최대 총 {maxTotal}%
        </div>
      )}
    </div>
  );
}

function PromoFlow({ revenues, storeSales=[], orders=[] }) {
  const [promos,setPromos]=useState(getPromosCache);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({name:"",platform:"자사몰",start_date:"",end_date:"",memo:"",content:"",files:[],discount_plan:{products:[],coupons:[]}});
  const today=new Date().toISOString().slice(0,10);
  const [impactModal,setImpactModal]=useState(null);
  const [viewStart,setViewStart]=useState(()=>{const d=new Date();d.setDate(d.getDate()-30);return d.toISOString().slice(0,10);});
  const [viewEnd,setViewEnd]=useState(()=>{const d=new Date();d.setDate(d.getDate()+30);return d.toISOString().slice(0,10);});
  const [viewPeriod,setViewPeriod]=useState("2m");
  const [calOpenFor,setCalOpenFor]=useState(null);
  const handleViewPeriod=v=>{
    setViewPeriod(v);
    if(v==="custom") return;
    const days=v==="1m"?15:v==="2m"?30:45;
    const s=new Date();s.setDate(s.getDate()-days);
    const e=new Date();e.setDate(e.getDate()+days);
    setViewStart(s.toISOString().slice(0,10));
    setViewEnd(e.toISOString().slice(0,10));
  };

  const [hoveredPromo,setHoveredPromo]=useState(null);
  const [fileAddTarget,setFileAddTarget]=useState(null);
  const fileInputRef=useRef(null);
  const [isDragging,setIsDragging]=useState(false);
  const dragRef=useRef(null);
  const formFileRef=useRef(null);
  const [formFileDragOver,setFormFileDragOver]=useState(false);
  const [tableFileDragOver,setTableFileDragOver]=useState(null);
  // Hidden promo log (localStorage only — no schema change needed)
  const getHiddenLog=()=>{try{return JSON.parse(localStorage.getItem("hidden_promo_log")||"[]");}catch{return[];}};
  const [hiddenLog,setHiddenLog]=useState(getHiddenLog);
  const hiddenIds=useMemo(()=>new Set(hiddenLog.map(h=>h.id)),[hiddenLog]);
  const [selHiddenIds,setSelHiddenIds]=useState(new Set());
  // Promo search
  const [searchStart,setSearchStart]=useState("");
  const [searchEnd,setSearchEnd]=useState("");
  const [searchCh,setSearchCh]=useState("");
  const addFilesFromList=(fileList,currentCount,onFile)=>{
    const remaining=3-currentCount;
    Array.from(fileList).slice(0,remaining).forEach(file=>readFileData(file,onFile));
  };
  useEffect(()=>{
    const onMove=e=>{
      if(!dragRef.current) return;
      const{startX,startVS,startVE,width}=dragRef.current;
      const totalDays=(new Date(startVE)-new Date(startVS))/86400000;
      const delta=Math.round(-(e.clientX-startX)/width*totalDays);
      const s=new Date(startVS);s.setDate(s.getDate()+delta);
      const en=new Date(startVE);en.setDate(en.getDate()+delta);
      setViewStart(s.toISOString().slice(0,10));
      setViewEnd(en.toISOString().slice(0,10));
    };
    const onUp=()=>{dragRef.current=null;setIsDragging(false);};
    window.addEventListener("mousemove",onMove);
    window.addEventListener("mouseup",onUp);
    return()=>{window.removeEventListener("mousemove",onMove);window.removeEventListener("mouseup",onUp);};
  },[]);
  const onDragStart=useCallback((e,w)=>{
    if(e.button!==0) return;
    e.preventDefault();
    const width=w||e.currentTarget.getBoundingClientRect().width;
    dragRef.current={startX:e.clientX,startVS:viewStart,startVE:viewEnd,width};
    setIsDragging(true);
  },[viewStart,viewEnd]);
  const [promoCalOpen,setPromoCalOpen]=useState(null);
  const [editingPromoId,setEditingPromoId]=useState(null);
  const [editPromoForm,setEditPromoForm]=useState({});
  const startEditPromo=p=>{setEditingPromoId(p.id);setEditPromoForm({name:p.name,platform:p.platform,start_date:p.start_date,end_date:p.end_date,content:p.content||p.memo||"",discount_plan:p.discount_plan||{products:[],coupons:[]}});};
  const savePromoEdit=()=>{
    patchPromo(editingPromoId,{...editPromoForm,memo:editPromoForm.content});
    setEditingPromoId(null);
  };
  const nowStr=new Date().toISOString().slice(0,16);
  const isEnded=p=>p.end_date&&String(p.end_date)<nowStr;
  const readFileData=(file,cb)=>{const r=new FileReader();r.onload=e=>cb({name:file.name,type:file.type,data:e.target.result});r.readAsDataURL(file);};

  const patchPromo=useCallback(async(id,updates)=>{
    setPromos(prev=>{const next=prev.map(p=>p.id===id?{...p,...updates}:p);setPromosCache(next);return next;});
    const db=await getSupabase();
    await db.from("promotions").update(updates).eq("id",id);
  },[]);

  // Load from Supabase — localStorage는 Supabase에 실제 데이터 있을 때만 덮어씀
  useEffect(()=>{
    (async()=>{
      const local=getPromosCache(); // 먼저 읽어둠 (덮어쓰기 방지)
      const db=await getSupabase();
      const{data,error}=await db.from("promotions").select("*").order("start_date",{ascending:true});
      if(!error&&data){
        const rows=data.map(p=>({...p,files:p.files||(p.file?[p.file]:[]),discount_plan:p.discount_plan||{products:[],coupons:[]}}));
        if(rows.length>0){
          setPromos(rows);setPromosCache(rows);
        } else if(local.length>0){
          // Supabase 비어있고 로컬에 데이터 있으면 마이그레이션
          const{error:e}=await db.from("promotions").insert(local);
          if(!e){setPromos(local);setPromosCache(local);}
          // 실패해도 useState 초기값(local) 그대로 표시됨
        }
      }
      // error 시 localStorage 초기값 유지 (useState에서 이미 설정됨)
    })();
  },[]);

  const addPromo=async()=>{
    if(!form.name||!form.start_date||!form.end_date)return;
    const newP={...form,id:Date.now()};
    setPromos(prev=>{const next=[...prev,newP];setPromosCache(next);return next;});
    const db=await getSupabase();
    await db.from("promotions").insert(newP);
    setForm({name:"",platform:"자사몰",start_date:"",end_date:"",memo:"",content:"",files:[],discount_plan:{products:[],coupons:[]}});
    setShowForm(false);
  };
  const delPromo=async id=>{
    setPromos(prev=>{const next=prev.filter(p=>p.id!==id);setPromosCache(next);return next;});
    const db=await getSupabase();
    await db.from("promotions").delete().eq("id",id);
  };
  const addFileToPromo=(id,f)=>{
    const promo=promos.find(p=>p.id===id);
    patchPromo(id,{files:[...(promo?.files||[]),f].slice(0,3)});
  };
  const removeFileFromPromo=(id,idx)=>{
    const promo=promos.find(p=>p.id===id);
    patchPromo(id,{files:(promo?.files||[]).filter((_,i)=>i!==idx)});
  };
  const hidePromo=p=>{
    const entry={...p,hidden_at:new Date().toISOString()};
    const next=[...hiddenLog.filter(h=>h.id!==p.id),entry];
    setHiddenLog(next);localStorage.setItem("hidden_promo_log",JSON.stringify(next));
  };
  const delFromHiddenLog=ids=>{
    const next=hiddenLog.filter(h=>!ids.has(h.id));
    setHiddenLog(next);setSelHiddenIds(new Set());
    localStorage.setItem("hidden_promo_log",JSON.stringify(next));
  };

  const getSubmitPromos=()=>{try{return JSON.parse(localStorage.getItem("submit_promos")||"[]");}catch{return [];}};
  const saveSubmitPromosLocal=data=>localStorage.setItem("submit_promos",JSON.stringify(data));
  const [submitPromos,setSubmitPromos]=useState(getSubmitPromos);
  const [showSubmitForm,setShowSubmitForm]=useState(false);
  const [submitForm,setSubmitForm]=useState({title:"",content:"",eod:""});
  const [editingSubmitId,setEditingSubmitId]=useState(null);
  const [editSubmitForm,setEditSubmitForm]=useState({title:"",content:"",eod:""});

  // Load submit_promotions from Supabase on mount
  useEffect(()=>{
    (async()=>{
      const local=getSubmitPromos();
      const db=await getSupabase();
      const{data,error}=await db.from("submit_promotions").select("*").order("id",{ascending:true});
      if(!error&&data){
        if(data.length>0){
          setSubmitPromos(data);saveSubmitPromosLocal(data);
        } else if(local.length>0){
          const{error:e}=await db.from("submit_promotions").insert(local);
          if(!e){setSubmitPromos(local);saveSubmitPromosLocal(local);}
        }
      }
    })();
  },[]);

  const addSubmitPromo=async()=>{
    if(!submitForm.title)return;
    const newS={id:Date.now(),...submitForm};
    const next=[...submitPromos,newS];
    setSubmitPromos(next);saveSubmitPromosLocal(next);
    const db=await getSupabase();
    await db.from("submit_promotions").insert(newS);
    setSubmitForm({title:"",content:"",eod:""});
  };
  const saveSubmitEdit=async()=>{
    const next=submitPromos.map(s=>s.id===editingSubmitId?{...s,...editSubmitForm}:s);
    setSubmitPromos(next);saveSubmitPromosLocal(next);
    const db=await getSupabase();
    await db.from("submit_promotions").update(editSubmitForm).eq("id",editingSubmitId);
    setEditingSubmitId(null);
  };
  const delSubmitPromo=async id=>{
    const next=submitPromos.filter(s=>s.id!==id);
    setSubmitPromos(next);saveSubmitPromosLocal(next);
    const db=await getSupabase();
    await db.from("submit_promotions").delete().eq("id",id);
  };

  const startMs=new Date(viewStart).getTime();
  const endMs=new Date(viewEnd).getTime();
  const totalMs=Math.max(1,endMs-startMs);
  const datePct=d=>{const ms=new Date(d).getTime();return Math.min(100,Math.max(0,(ms-startMs)/totalMs*100));};

  const revenueData=useMemo(()=>{
    if(!viewStart||!viewEnd||viewStart>viewEnd) return [];
    // viewStart~viewEnd 전체 날짜를 먼저 생성 (프로모션 캘린더와 동일 범위)
    const byDate={};
    const cur=new Date(viewStart);
    const end=new Date(viewEnd);
    while(cur<=end){
      const key=cur.toISOString().slice(0,10);
      byDate[key]={date:key.slice(5),fullDate:key,...Object.fromEntries(PROMO_PLATFORMS.map(p=>[p,null]))};
      cur.setDate(cur.getDate()+1);
    }
    // 실제 매출 데이터 채우기
    revenues.filter(r=>r.date>=viewStart&&r.date<=viewEnd).forEach(r=>{
      if(!byDate[r.date]) return;
      byDate[r.date][r.channel]=(byDate[r.date][r.channel]||0)+(r.amount||0);
    });
    // 오프라인 스토어 순매출 (배송 - 반품)
    storeSales.filter(r=>r.sale_date>=viewStart&&r.sale_date<=viewEnd).forEach(r=>{
      if(!byDate[r.sale_date]) return;
      const ch="오프라인 스토어";
      const cur=byDate[r.sale_date][ch]||0;
      if(r.status==="배송") byDate[r.sale_date][ch]=cur+(r.amount||0);
      else if(r.status==="반품") byDate[r.sale_date][ch]=cur-(r.amount||0);
    });
    Object.values(byDate).forEach(row=>{
      if((row["오프라인 스토어"]||0)<0) row["오프라인 스토어"]=0;
    });
    return Object.values(byDate).sort((a,b)=>a.date>b.date?1:-1);
  },[revenues,storeSales,viewStart,viewEnd]);

  const inp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
    padding:"8px 12px",fontSize:16,color:D.text,width:"100%",boxSizing:"border-box",
    fontFamily:"'Pretendard','Noto Sans KR',sans-serif"};

  return (
    <div style={{padding:"20px 24px",maxWidth:1600,margin:"0 auto"}}>
      <style>{`
        @keyframes promoShimmer {
          0%   { transform: translateX(-100%); }
          100% { transform: translateX(400%); }
        }
        .promo-shimmer {
          position: absolute;
          top: 0; bottom: 0;
          width: 50%;
          background: linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.55) 40%, rgba(255,255,255,0.75) 50%, rgba(255,255,255,0.55) 60%, transparent 100%);
          animation: promoShimmer 10s linear infinite;
          pointer-events: none;
        }
      `}</style>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <div style={{fontWeight:600,fontSize:17,color:D.black}}>프로모션 플로우</div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <CalDrop id="promoView" period={viewPeriod} setPeriod={handleViewPeriod}
            presets={[["1m","1개월"],["2m","2개월"],["3m","3개월"]]}
            start={viewStart} setStart={setViewStart}
            end={viewEnd} setEnd={setViewEnd}
            calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}/>
          <button onClick={()=>setShowForm(v=>!v)}
            style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
              padding:"6px 14px",fontSize:13,cursor:"pointer",fontWeight:600}}>
            {showForm?"취소":"+ 프로모션 추가"}
          </button>
        </div>
      </div>

      {showForm&&(
        <Card style={{marginBottom:20}}>
          <div style={{fontWeight:600,fontSize:14,marginBottom:16}}>프로모션 추가</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr auto",gap:8,alignItems:"start"}}>
            <div>
              <div style={{fontSize:12,color:D.textMeta,marginBottom:4}}>프로모션명</div>
              <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={inp} placeholder="예: 오픈 기념 할인"/>
            </div>
            <div>
              <div style={{fontSize:12,color:D.textMeta,marginBottom:4}}>플랫폼</div>
              <div style={{display:"flex",gap:4}}>
                {PROMO_PLATFORMS.map(p=>(
                  <button key={p} onClick={()=>setForm(f=>({...f,platform:p}))}
                    style={{flex:1,background:form.platform===p?chColor(p):"transparent",
                      color:form.platform===p?"#fff":D.textSub,
                      border:`1px solid ${form.platform===p?chColor(p):D.border}`,
                      borderRadius:5,padding:"6px 4px",fontSize:13,cursor:"pointer"}}>
                    {p}
                  </button>
                ))}
              </div>
            </div>
            {[["시작일시","start_date"],["종료일시","end_date"]].map(([label,field])=>(
              <div key={field}>
                <div style={{fontSize:12,color:D.textMeta,marginBottom:4}}>{label}</div>
                <DateDrop id={`promo_${field}`}
                  value={form[field]?.slice(0,10)||""}
                  onChange={v=>{const time=form[field]?.slice(10)||"";setForm(f=>({...f,[field]:v+time}));}}
                  calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}
                  placeholder="날짜 선택"/>
                <div style={{display:"flex",gap:3,marginTop:4}}>
                  {[["T10:00","오전 10시"],["T11:00","오전 11시"],["T23:59","오후 23:59"]].map(([time,tl])=>(
                    <button key={time} onClick={()=>{
                      const base=form[field]?form[field].slice(0,10):new Date().toISOString().slice(0,10);
                      setForm(f=>({...f,[field]:`${base}${time}`}));
                    }} style={{flex:1,fontSize:11,padding:"3px 2px",
                      background:form[field]&&form[field].includes(time)?D.black:D.surfaceAlt,
                      color:form[field]&&form[field].includes(time)?"#fff":D.textSub,
                      border:`1px solid ${form[field]&&form[field].includes(time)?D.black:D.border}`,
                      borderRadius:3,cursor:"pointer",whiteSpace:"nowrap"}}>
                      {tl}
                    </button>
                  ))}
                  <button onClick={()=>{
                    const base=form[field]?form[field].slice(0,10):"";
                    setForm(f=>({...f,[field]:base}));
                  }} style={{fontSize:11,padding:"3px 6px",background:"transparent",
                    border:`1px solid ${D.border}`,borderRadius:3,cursor:"pointer",color:D.textMeta,whiteSpace:"nowrap"}}>
                    시간 삭제
                  </button>
                </div>
              </div>
            ))}
            <button onClick={addPromo}
              style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
                padding:"9px 16px",fontSize:14,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap",marginTop:18}}>
              저장
            </button>
          </div>
          <div style={{marginTop:10,display:"grid",gridTemplateColumns:"1.5fr 1fr",gap:8}}>
            <div>
              <div style={{fontSize:12,color:D.textMeta,marginBottom:4}}>프로모션 내용</div>
              <textarea value={form.content||form.memo||""} onChange={e=>setForm(f=>({...f,content:e.target.value,memo:e.target.value}))}
                onKeyDown={e=>{
                  if(e.key==="Enter"){
                    e.preventDefault();
                    const ta=e.target;
                    const{selectionStart:s,selectionEnd:en,value:v}=ta;
                    const next=v.slice(0,s)+"\n○ "+v.slice(en);
                    setForm(f=>({...f,content:next,memo:next}));
                    requestAnimationFrame(()=>{ta.selectionStart=ta.selectionEnd=s+3;});
                  }
                }}
                style={{...inp,resize:"vertical",minHeight:144,lineHeight:1.5}} placeholder="할인율, 대상 상품, 조건 등 (선택)"/>
            </div>
            <div>
              <div style={{fontSize:12,color:D.textMeta,marginBottom:4}}>첨부 파일 <span style={{opacity:.6}}>(최대 3개)</span></div>
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {(form.files||[]).map((f,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:4,
                    background:D.surfaceAlt,borderRadius:4,padding:"4px 8px",fontSize:13}}>
                    <span>📎</span>
                    <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:D.textSub}}>{f.name}</span>
                    <button onClick={()=>setForm(fm=>({...fm,files:(fm.files||[]).filter((_,j)=>j!==i)}))}
                      style={{background:"none",border:"none",color:D.textMeta,cursor:"pointer",padding:0,fontSize:15,lineHeight:1}}>✕</button>
                  </div>
                ))}
                {(form.files||[]).length<3&&(
                  <div
                    onDragOver={e=>{e.preventDefault();setFormFileDragOver(true);}}
                    onDragLeave={e=>{if(!e.currentTarget.contains(e.relatedTarget))setFormFileDragOver(false);}}
                    onDrop={e=>{e.preventDefault();setFormFileDragOver(false);
                      addFilesFromList(e.dataTransfer.files,(form.files||[]).length,
                        f=>setForm(fm=>({...fm,files:[...(fm.files||[]),f].slice(0,3)})));
                    }}
                    onClick={()=>formFileRef.current?.click()}
                    style={{border:`1px dashed ${formFileDragOver?D.blue:D.border}`,borderRadius:6,
                      padding:"14px 12px",textAlign:"center",cursor:"pointer",
                      color:formFileDragOver?D.blue:D.textMeta,fontSize:13,userSelect:"none",
                      background:formFileDragOver?"#eef3ff":"transparent",transition:"all 0.15s"}}>
                    {formFileDragOver?"여기에 놓기 ↓":"파일을 드래그하거나 클릭하여 선택"}
                    <input ref={formFileRef} type="file" multiple style={{display:"none"}} onChange={e=>{
                      addFilesFromList(e.target.files,(form.files||[]).length,
                        f=>setForm(fm=>({...fm,files:[...(fm.files||[]),f].slice(0,3)})));
                      e.target.value="";
                    }}/>
                  </div>
                )}
              </div>
            </div>
          </div>
          <div style={{marginTop:14}}>
            <DiscountPlanEditor value={form.discount_plan}
              onChange={v=>setForm(f=>({...f,discount_plan:v}))}/>
          </div>
        </Card>
      )}

      {/* 플랫폼별 가로 캘린더 바 */}
      <Card style={{marginBottom:0,borderBottomLeftRadius:0,borderBottomRightRadius:0,borderBottom:"none"}}>
        <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>플랫폼별 프로모션 일정</div>
        <div onMouseDown={onDragStart}
          style={{cursor:isDragging?"grabbing":"grab",userSelect:"none",position:"relative",overflow:"hidden"}}>
        <div style={{position:"absolute",inset:0,pointerEvents:"none",zIndex:1,overflow:"hidden",borderRadius:4}}>
          <div className="promo-shimmer"/>
        </div>
        {/* 날짜 눈금 */}
        <div style={{position:"relative",height:16,marginBottom:4,paddingLeft:70}}>
          {[0,25,50,75,100].map(pct=>{
            const ms=startMs+(endMs-startMs)*pct/100;
            const d=new Date(ms);
            const label=`${d.getMonth()+1}/${d.getDate()}`;
            return <span key={pct} style={{position:"absolute",left:`${pct}%`,transform:"translateX(-50%)",
              fontSize:11,color:D.textMeta}}>{label}</span>;
          })}
          {(()=>{const tp=datePct(today);return tp>=0&&tp<=100?(
            <span style={{position:"absolute",left:`${tp}%`,transform:"translateX(-50%)",
              fontSize:11,color:D.primary,fontWeight:700}}>오늘</span>
          ):null;})()}
        </div>
        {PROMO_PLATFORMS.map(plat=>{
          const bars=promos.filter(p=>p.platform===plat&&p.end_date>=viewStart&&p.start_date<=viewEnd)
            .sort((a,b)=>a.start_date>b.start_date?1:-1);
          // lane 배정: 겹치지 않도록 최소 lane에 배치
          const lanes=[];
          const laned=bars.map(promo=>{
            let lane=lanes.findIndex(endDate=>endDate<=promo.start_date);
            if(lane===-1){lane=lanes.length;lanes.push("");}
            lanes[lane]=promo.end_date;
            return{promo,lane};
          });
          const numLanes=Math.max(1,lanes.length);
          const laneH=26;const laneGap=2;
          const trackH=numLanes*laneH+(numLanes-1)*laneGap;
          return (
            <div key={plat} style={{display:"flex",alignItems:"center",marginBottom:8,gap:8}}>
              <div style={{width:62,fontSize:13,color:D.textSub,flexShrink:0,textAlign:"right",lineHeight:1.3}}>{plat}</div>
              <div style={{flex:1,position:"relative",height:trackH,background:D.surfaceAlt,borderRadius:4}}>
                {laned.map(({promo,lane})=>{
                  const l=datePct(promo.start_date);
                  const r=datePct(promo.end_date);
                  const w=Math.max(0.5,r-l);
                  const ended=isEnded(promo);
                  const top=lane*(laneH+laneGap);
                  return (
                    <div key={promo.id}
                      onMouseEnter={e=>{const rect=e.currentTarget.getBoundingClientRect();setHoveredPromo({promo,rect});}}
                      onMouseLeave={()=>setHoveredPromo(null)}
                      style={{position:"absolute",left:`${l}%`,width:`${w}%`,height:laneH,top,
                        background:ended?"#aaa":chColor(plat),borderRadius:4,display:"flex",alignItems:"center",
                        padding:"0 6px",fontSize:12,color:"#fff",overflow:"hidden",
                        boxSizing:"border-box",cursor:"pointer",minWidth:4,opacity:ended?0.6:1}}>
                      <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                        textDecoration:ended?"line-through":"none"}}>{promo.name}</span>
                    </div>
                  );
                })}
                {(()=>{const tp=datePct(today);return tp>=0&&tp<=100?(
                  <div style={{position:"absolute",left:`${tp}%`,top:0,bottom:0,width:1.5,
                    background:"rgba(0,0,0,0.25)",pointerEvents:"none"}}/>
                ):null;})()}
              </div>
              <div style={{width:20,flexShrink:0}}/>
            </div>
          );
        })}
        </div>{/* end draggable */}
        {/* 테이블 파일 추가용 hidden input */}
        <input type="file" ref={fileInputRef} style={{display:"none"}} onChange={e=>{
          const file=e.target.files?.[0];
          if(!file||!fileAddTarget) return;
          readFileData(file,f=>addFileToPromo(fileAddTarget,f));
          setFileAddTarget(null);
        }}/>
        {/* 호버 팝업 */}
        {hoveredPromo&&(()=>{
          const {promo,rect}=hoveredPromo;
          const above=rect.bottom+140>window.innerHeight;
          return (
            <div style={{position:"fixed",left:Math.min(rect.left,window.innerWidth-260),
              top:above?rect.top-8:rect.bottom+6,transform:above?"translateY(-100%)":"none",
              zIndex:1000,background:"#fff",border:`1px solid ${D.border}`,borderRadius:8,
              padding:"10px 14px",boxShadow:"0 4px 20px rgba(0,0,0,0.13)",
              minWidth:220,maxWidth:300,pointerEvents:"none",fontSize:13,lineHeight:1.6}}>
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:chColor(promo.platform),flexShrink:0}}/>
                <span style={{fontWeight:600,fontSize:14}}>{promo.name}</span>
                <span style={{fontSize:12,color:D.textMeta,marginLeft:"auto"}}>{promo.platform}</span>
              </div>
              <div style={{color:D.textMeta,fontSize:12,marginBottom:4}}>
                {promo.start_date} ~ {promo.end_date}
              </div>
              {(promo.content||promo.memo)&&(
                <div style={{color:D.text,fontSize:13,marginBottom:4,wordBreak:"break-all",whiteSpace:"pre-wrap"}}>{promo.content||promo.memo}</div>
              )}
              {(promo.files||[]).map((f,i)=>(
                <div key={i} style={{fontSize:12,color:D.textMeta}}>📎 {f.name}</div>
              ))}
              {isEnded(promo)&&(
                <div style={{marginTop:5,fontSize:12,color:"#e55",fontWeight:600}}>종료된 프로모션</div>
              )}
            </div>
          );
        })()}
      </Card>

      {/* 기간별 플랫폼 매출 그래프 */}
      <Card style={{marginBottom:20,borderTopLeftRadius:0,borderTopRightRadius:0}}>
        <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>기간별 플랫폼 매출</div>
        {revenueData.length>0&&revenues.some(r=>r.date>=viewStart&&r.date<=viewEnd)?(
          <div style={{position:"relative",overflow:"hidden"}} onMouseDown={onDragStart}>
          <div style={{position:"absolute",inset:0,pointerEvents:"none",zIndex:1,overflow:"hidden"}}>
            <div className="promo-shimmer"/>
          </div>
          {isDragging&&<div style={{position:"absolute",inset:0,zIndex:10,cursor:"grabbing"}}/>}
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={revenueData}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="date" tick={{fill:D.textMeta,fontSize:11}}/>
              <YAxis tick={{fill:D.textMeta,fontSize:11}} tickFormatter={v=>v>=10000?(v/10000).toFixed(0)+"만":v}/>
              <Tooltip content={({active,payload})=>{
                if(!active||!payload?.length) return null;
                const fullDate=payload[0]?.payload?.fullDate||"";
                const label=payload[0]?.payload?.date||"";
                // 진행 중 / 진행했던 분리 (hidden 포함)
                const allForTip=[...promos,...hiddenLog.filter(h=>!promos.find(p=>p.id===h.id))];
                const inRange=allForTip.filter(p=>p.start_date.slice(0,10)<=fullDate&&(p.end_date||"9999").slice(0,10)>=fullDate);
                const runningPromos=inRange.filter(p=>!isEnded(p));
                const pastPromos=inRange.filter(p=>isEnded(p));
                return (
                  <div style={{background:"#fff",border:`1px solid ${D.border}`,borderRadius:8,
                    padding:"10px 14px",fontSize:13,boxShadow:"0 4px 16px rgba(0,0,0,0.1)",minWidth:180}}>
                    <div style={{fontWeight:600,marginBottom:6,color:D.text}}>{fullDate||label}</div>
                    {payload.filter(p=>p.value!=null).sort((a,b)=>(b.value||0)-(a.value||0)).map((p,i)=>(
                      <div key={i} style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
                        <div style={{width:10,height:3,background:p.stroke,borderRadius:2,flexShrink:0}}/>
                        <span style={{color:D.textSub,flex:1}}>{p.name}</span>
                        <span style={{fontWeight:600}}>₩{(p.value||0).toLocaleString()}</span>
                      </div>
                    ))}
                    {runningPromos.length>0&&(
                      <div style={{marginTop:8,paddingTop:6,borderTop:`1px solid ${D.border}`}}>
                        <div style={{fontSize:11,color:D.textMeta,marginBottom:4,letterSpacing:"0.05em"}}>진행 중인 프로모션</div>
                        {runningPromos.map(p=>(
                          <div key={p.id} style={{display:"flex",alignItems:"center",gap:5,marginBottom:2}}>
                            <div style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),flexShrink:0}}/>
                            <span style={{color:D.textSub,fontSize:12}}>{p.platform}</span>
                            <span style={{fontWeight:600,fontSize:12,marginLeft:2}}>{p.name}</span>
                          </div>
                        ))}
                      </div>
                    )}
                    {pastPromos.length>0&&(
                      <div style={{marginTop:6,paddingTop:6,borderTop:`1px solid ${D.border}`}}>
                        <div style={{fontSize:11,color:"#aaa",marginBottom:4,letterSpacing:"0.05em"}}>진행했던 프로모션</div>
                        {pastPromos.map(p=>(
                          <div key={p.id} style={{display:"flex",alignItems:"center",gap:5,marginBottom:2,opacity:0.7}}>
                            <div style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),flexShrink:0}}/>
                            <span style={{color:D.textSub,fontSize:12}}>{p.platform}</span>
                            <span style={{fontWeight:600,fontSize:12,marginLeft:2,textDecoration:"line-through"}}>{p.name}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              }}/>
              <Legend iconSize={8} wrapperStyle={{fontSize:12}}/>
              {PROMO_PLATFORMS.map(p=>(
                <Line key={p} type="monotone" dataKey={p} name={p}
                  stroke={chColor(p)} strokeWidth={2} dot={false} connectNulls={false}/>
              ))}
            </LineChart>
          </ResponsiveContainer>
          </div>
        ):(
          <div style={{textAlign:"center",padding:40,color:D.textMeta,fontSize:14}}>
            해당 기간에 매출 데이터가 없습니다
          </div>
        )}
      </Card>

      {/* 제출해야하는 프로모션 */}
      <Card>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:showSubmitForm?12:0}}>
          <div style={{fontWeight:600,fontSize:14,color:D.black}}>제출해야하는 프로모션</div>
          <button onClick={()=>setShowSubmitForm(v=>!v)}
            style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
              padding:"5px 12px",fontSize:12,cursor:"pointer",fontWeight:600}}>
            {showSubmitForm?"닫기":"+ 추가"}
          </button>
        </div>
        {showSubmitForm&&<div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:12}}>
          <div>
            <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>프로모션명</div>
            <input value={submitForm.title} onChange={e=>setSubmitForm(f=>({...f,title:e.target.value}))}
              placeholder="프로모션명 입력"
              style={{width:"100%",boxSizing:"border-box",background:"transparent",border:`1px solid ${D.border}`,
                borderRadius:5,padding:"7px 10px",fontSize:14,color:D.text,fontFamily:"'Pretendard','Noto Sans KR',sans-serif"}}/>
          </div>
          <div>
            <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>내용</div>
            <textarea value={submitForm.content} onChange={e=>setSubmitForm(f=>({...f,content:e.target.value}))}
              placeholder="프로모션 내용 간략히 입력"
              rows={2}
              style={{width:"100%",boxSizing:"border-box",background:"transparent",border:`1px solid ${D.border}`,
                borderRadius:5,padding:"7px 10px",fontSize:13,color:D.text,resize:"vertical",
                fontFamily:"'Pretendard','Noto Sans KR',sans-serif"}}/>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
            <div style={{flex:1}}>
              <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>EOD</div>
              <SubmitEodPicker value={submitForm.eod} onChange={v=>setSubmitForm(f=>({...f,eod:v}))}/>
            </div>
            <button onClick={addSubmitPromo}
              style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                padding:"7px 18px",fontSize:13,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap",height:36}}>저장</button>
          </div>
        </div>}
        {submitPromos.length>0&&(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:D.surfaceAlt}}>
                {["프로모션명","내용","EOD","",""].map((h,i)=>(
                  <th key={i} style={{padding:"5px 8px",textAlign:"left",fontWeight:600,
                    color:D.textSub,borderBottom:`1px solid ${D.border}`,fontSize:12,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {submitPromos.map(s=>{
                const isE=editingSubmitId===s.id;
                const tdS={padding:"7px 8px",borderBottom:`1px solid ${D.border}`,color:D.text};
                if(isE) return (
                  <tr key={s.id}>
                    <td style={tdS}>
                      <input value={editSubmitForm.title} onChange={e=>setEditSubmitForm(f=>({...f,title:e.target.value}))}
                        style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
                          padding:"4px 8px",fontSize:13,color:D.text,width:"100%",boxSizing:"border-box",
                          fontFamily:"'Pretendard','Noto Sans KR',sans-serif"}}/>
                    </td>
                    <td style={tdS}>
                      <textarea value={editSubmitForm.content||""} onChange={e=>setEditSubmitForm(f=>({...f,content:e.target.value}))}
                        rows={2}
                        style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
                          padding:"4px 8px",fontSize:13,color:D.text,width:"100%",boxSizing:"border-box",resize:"vertical",
                          fontFamily:"'Pretendard','Noto Sans KR',sans-serif"}}/>
                    </td>
                    <td style={tdS}>
                      <SubmitEodPicker value={editSubmitForm.eod||""} onChange={v=>setEditSubmitForm(f=>({...f,eod:v}))}/>
                    </td>
                    <td style={{...tdS,whiteSpace:"nowrap"}}>
                      <button onClick={saveSubmitEdit}
                        style={{background:D.black,color:"#fff",border:"none",borderRadius:4,
                          padding:"4px 12px",fontSize:12,cursor:"pointer",fontWeight:600,marginRight:4}}>저장</button>
                      <button onClick={()=>setEditingSubmitId(null)}
                        style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
                          padding:"4px 8px",fontSize:12,cursor:"pointer",color:D.textSub}}>취소</button>
                    </td>
                    <td style={tdS}/>
                  </tr>
                );
                return (
                  <tr key={s.id}>
                    <td style={{...tdS,fontWeight:600}}>{s.title}</td>
                    <td style={{...tdS,color:D.textSub,maxWidth:200,wordBreak:"break-all"}}>{s.content||"—"}</td>
                    <td style={{...tdS,color:D.textSub}}>{s.eod?s.eod.replace("T"," "):"—"}</td>
                    <td style={{...tdS,whiteSpace:"nowrap"}}>
                      <button onClick={()=>{setEditingSubmitId(s.id);setEditSubmitForm({title:s.title,content:s.content||"",eod:s.eod||""});}}
                        style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
                          padding:"3px 10px",fontSize:12,cursor:"pointer",color:D.textSub,marginRight:4}}>수정</button>
                      <button onClick={()=>delSubmitPromo(s.id)}
                        style={{background:D.black,color:"#fff",border:"none",borderRadius:4,
                          padding:"3px 10px",fontSize:12,cursor:"pointer",fontWeight:600}}>제출 완료</button>
                    </td>
                    <td style={tdS}/>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
        {submitPromos.length===0&&(
          <div style={{textAlign:"center",padding:"16px 0",color:D.textMeta,fontSize:13}}>제출할 프로모션을 추가하세요</div>
        )}
      </Card>

      {/* 등록된 프로모션 목록 표 */}
      {promos.filter(p=>!hiddenIds.has(p.id)).length>0&&(
        <Card>
          <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>등록된 프로모션</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:D.surfaceAlt}}>
                {["채널","프로모션명","기간","상세 내용","할인율","첨부 파일","","",""].map((h,i)=>(
                  <th key={i} style={{padding:"5px 8px",textAlign:"left",fontWeight:600,
                    color:D.textSub,borderBottom:`1px solid ${D.border}`,fontSize:12,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...promos].filter(p=>!hiddenIds.has(p.id)).sort((a,b)=>a.start_date>b.start_date?1:-1).map(p=>{
                const ended=isEnded(p);
                const isEditing=editingPromoId===p.id;
                const td={style:{padding:"6px 8px",borderBottom:`1px solid ${D.border}`,
                  color:ended?"#ccc":D.text,textDecoration:"none"}};
                const inp3={background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                  padding:"7px 10px",fontSize:15,color:D.text,width:"100%",boxSizing:"border-box",
                  fontFamily:"'Pretendard','Noto Sans KR',sans-serif"};
                if(isEditing) return (
                  <tr key={p.id}>
                    <td colSpan={8} style={{padding:"10px 8px",borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt}}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,marginBottom:8}}>
                        <div>
                          <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>프로모션명</div>
                          <input value={editPromoForm.name} onChange={e=>setEditPromoForm(f=>({...f,name:e.target.value}))} style={inp3}/>
                        </div>
                        <div>
                          <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>플랫폼</div>
                          <div style={{display:"flex",gap:3}}>
                            {PROMO_PLATFORMS.map(pl=>(
                              <button key={pl} onClick={()=>setEditPromoForm(f=>({...f,platform:pl}))}
                                style={{flex:1,background:editPromoForm.platform===pl?chColor(pl):"transparent",
                                  color:editPromoForm.platform===pl?"#fff":D.textSub,
                                  border:`1px solid ${editPromoForm.platform===pl?chColor(pl):D.border}`,
                                  borderRadius:4,padding:"5px 3px",fontSize:12,cursor:"pointer"}}>{pl}</button>
                            ))}
                          </div>
                        </div>
                        {[["시작일시","start_date"],["종료일시","end_date"]].map(([label,field])=>(
                          <div key={field}>
                            <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>{label}</div>
                            <DateButtonPicker value={editPromoForm[field]||""} onChange={v=>setEditPromoForm(f=>({...f,[field]:v}))}/>
                            <div style={{display:"flex",gap:3,marginTop:3}}>
                              {[["T10:00","10시"],["T11:00","11시"],["T23:59","23:59"]].map(([time,tl])=>(
                                <button key={time} onClick={()=>{
                                  const base=(editPromoForm[field]||new Date().toISOString().slice(0,10)).slice(0,10);
                                  setEditPromoForm(f=>({...f,[field]:`${base}${time}`}));
                                }} style={{flex:1,fontSize:11,padding:"2px",background:D.surfaceAlt,
                                  border:`1px solid ${D.border}`,borderRadius:3,cursor:"pointer",color:D.textSub}}>{tl}</button>
                              ))}
                            </div>
                          </div>
                        ))}
                      </div>
                      <div style={{marginBottom:8}}>
                        <div style={{fontSize:12,color:D.textMeta,marginBottom:3}}>프로모션 내용</div>
                        <textarea value={editPromoForm.content} onChange={e=>setEditPromoForm(f=>({...f,content:e.target.value}))}
                          style={{...inp3,resize:"vertical",minHeight:60,lineHeight:1.5}} placeholder="할인율, 대상 상품, 조건 등"/>
                      </div>
                      <div style={{marginBottom:8}}>
                        <DiscountPlanEditor value={editPromoForm.discount_plan}
                          onChange={v=>setEditPromoForm(f=>({...f,discount_plan:v}))}/>
                      </div>
                      <div style={{display:"flex",gap:6}}>
                        <button onClick={savePromoEdit}
                          style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                            padding:"6px 16px",fontSize:13,cursor:"pointer",fontWeight:600}}>저장</button>
                        <button onClick={()=>setEditingPromoId(null)}
                          style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                            padding:"6px 12px",fontSize:13,cursor:"pointer",color:D.textSub}}>취소</button>
                      </div>
                    </td>
                  </tr>
                );
                return (
                  <tr key={p.id}>
                    <td {...td}>
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <div style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),flexShrink:0}}/>
                        {p.platform}
                      </div>
                    </td>
                    <td {...td} style={{...td.style,fontWeight:600}}>
                      {p.name}
                      {ended&&<span style={{marginLeft:6,fontSize:11,fontWeight:500,color:D.red}}>종료된 프로모션</span>}
                      {/* 시작일이 지난(진행중 또는 종료된) 프로모션은 임팩트 분석 진입 가능 */}
                      {p.start_date&&p.start_date.slice(0,10)<=today&&(
                        <button onClick={()=>setImpactModal(p)}
                          style={{marginLeft:8,background:D.black,color:"#fff",border:"none",borderRadius:5,
                            padding:"2px 9px",fontSize:11,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap"}}>
                          임팩트 분석
                        </button>
                      )}
                    </td>
                    <td {...td} style={{...td.style,whiteSpace:"nowrap",textDecoration:"none"}}>
                      {[p.start_date,p.end_date].map((dt,i)=>{
                        const [d,t]=(dt||"").split("T");
                        return (
                          <div key={i} style={{lineHeight:1.4}}>
                            <span style={{fontWeight:700,fontSize:14}}>{d}</span>
                            {t&&<span style={{fontWeight:500,fontSize:13,color:D.textSub,marginLeft:4}}>{t}</span>}
                          </div>
                        );
                      })}
                    </td>
                    <td {...td} style={{...td.style,maxWidth:200,color:ended?"#bbb":D.textSub,whiteSpace:"pre-wrap"}}>{p.content||p.memo||"—"}</td>
                    <td {...td} style={{...td.style,verticalAlign:"top",minWidth:170,maxWidth:220}}>
                      <DiscountPlanView plan={p.discount_plan}/>
                    </td>
                    <td {...td} style={{...td.style,minWidth:160}}>
                      <div
                        onDragOver={e=>{e.preventDefault();setTableFileDragOver(p.id);}}
                        onDragLeave={e=>{if(!e.currentTarget.contains(e.relatedTarget))setTableFileDragOver(null);}}
                        onDrop={e=>{e.preventDefault();setTableFileDragOver(null);
                          addFilesFromList(e.dataTransfer.files,(p.files||[]).length,f=>addFileToPromo(p.id,f));
                        }}
                        style={{display:"flex",flexDirection:"column",gap:3,
                          border:`1px dashed ${tableFileDragOver===p.id?D.blue:"transparent"}`,
                          borderRadius:4,padding:tableFileDragOver===p.id?4:0,
                          background:tableFileDragOver===p.id?"#eef3ff":"transparent",
                          minHeight:24,transition:"all 0.15s"}}>
                        {(p.files||[]).map((f,i)=>(
                          <div key={i} style={{display:"flex",alignItems:"center",gap:3}}>
                            <a href={f.data} download={f.name}
                              style={{fontSize:12,color:ended?"#bbb":D.textSub,textDecoration:"none",
                                wordBreak:"break-all",flex:1}}
                              title={f.name}>📎 {f.name}</a>
                            <button onClick={()=>removeFileFromPromo(p.id,i)}
                              style={{background:"none",border:"none",color:D.textMeta,cursor:"pointer",
                                padding:0,fontSize:14,lineHeight:1,flexShrink:0}}>✕</button>
                          </div>
                        ))}
                        {(p.files||[]).length<3&&(
                          tableFileDragOver===p.id
                            ?<span style={{fontSize:11,color:D.blue,textAlign:"center",padding:"2px 0"}}>여기에 놓기 ↓</span>
                            :<button onClick={()=>{setFileAddTarget(p.id);fileInputRef.current.value="";fileInputRef.current.click();}}
                              style={{background:"transparent",border:`1px dashed ${D.border}`,borderRadius:3,
                                padding:"2px 6px",fontSize:12,color:D.textMeta,cursor:"pointer",
                                whiteSpace:"nowrap",alignSelf:"flex-start"}}>+ 파일 추가</button>
                        )}
                        {!(p.files||[]).length&&tableFileDragOver!==p.id&&<span style={{color:D.textMeta}}>—</span>}
                      </div>
                    </td>
                    <td style={{padding:"6px 8px",borderBottom:`1px solid ${D.border}`}}>
                      <button onClick={()=>startEditPromo(p)} title="수정"
                        style={{background:"transparent",border:"none",color:D.textMeta,
                          cursor:"pointer",padding:"2px 4px",fontSize:15,filter:"grayscale(1)"}}>✎</button>
                    </td>
                    <td style={{padding:"6px 8px",borderBottom:`1px solid ${D.border}`}}>
                      {ended&&(
                        <button onClick={()=>hidePromo(p)} title="가리기 (종료 프로모션 로그)"
                          style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
                            color:D.textMeta,cursor:"pointer",padding:"2px 8px",fontSize:11,whiteSpace:"nowrap",marginRight:4}}>
                          가리기
                        </button>
                      )}
                    </td>
                    <td style={{padding:"6px 8px",borderBottom:`1px solid ${D.border}`}}>
                      <button onClick={()=>delPromo(p.id)}
                        style={{background:"transparent",border:"none",color:D.textMeta,
                          cursor:"pointer",padding:0,fontSize:14}}>✕</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      )}

      {/* 프로모션 검색 */}
      <Card style={{marginTop:12}}>
        <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>프로모션 검색</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",marginBottom:14}}>
          <DateDrop id="searchStart" value={searchStart} onChange={setSearchStart}
            calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="시작일"/>
          <span style={{color:D.textMeta,fontSize:11}}>~</span>
          <DateDrop id="searchEnd" value={searchEnd} onChange={setSearchEnd}
            calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="종료일"/>
          <span style={{color:D.borderMid,fontSize:14}}>|</span>
          {["",...PROMO_PLATFORMS].map(ch=>(
            <button key={ch||"all"} onClick={()=>setSearchCh(ch)}
              style={{background:searchCh===ch?D.black:"transparent",color:searchCh===ch?"#fff":D.textSub,
                border:`1px solid ${searchCh===ch?D.black:D.border}`,borderRadius:5,
                padding:"4px 10px",fontSize:12,cursor:"pointer"}}>
              {ch||"전체 채널"}
            </button>
          ))}
          {(searchStart||searchEnd||searchCh)&&(
            <button onClick={()=>{setSearchStart("");setSearchEnd("");setSearchCh("");}}
              style={{background:"none",border:"none",color:D.textMeta,cursor:"pointer",fontSize:13}}>✕ 초기화</button>
          )}
        </div>
        {(()=>{
          const s=searchStart||"0000-01-01";const e=searchEnd||"9999-12-31";
          const allP=[...promos,...hiddenLog.filter(h=>!promos.find(p=>p.id===h.id))];
          const matched=allP.filter(p=>{
            const overlap=p.start_date.slice(0,10)<=e&&(p.end_date||"9999-12-31").slice(0,10)>=s;
            const chMatch=!searchCh||p.platform===searchCh;
            return overlap&&chMatch;
          }).sort((a,b)=>a.start_date>b.start_date?1:-1);
          // Revenue for period + channel
          const revTotal=(()=>{
            let t=0;
            revenues.filter(r=>r.date>=s&&r.date<=e&&(!searchCh||r.channel===searchCh)).forEach(r=>t+=(r.amount||0));
            if(!searchCh||searchCh==="오프라인 스토어"){
              storeSales.filter(r=>r.sale_date>=s&&r.sale_date<=e).forEach(r=>{
                if(r.status==="배송") t+=(r.amount||0);
                else if(r.status==="반품") t-=(r.amount||0);
              });
            }
            return Math.max(0,t);
          })();
          if(!matched.length&&!(searchStart||searchEnd||searchCh)){
            return <div style={{textAlign:"center",padding:"20px 0",color:D.textMeta,fontSize:13}}>기간 또는 채널을 선택하면 검색 결과가 표시됩니다</div>;
          }
          return(
            <>
              <div style={{display:"flex",gap:12,marginBottom:12,flexWrap:"wrap"}}>
                <div style={{background:D.surfaceAlt,borderRadius:7,padding:"10px 16px",minWidth:140}}>
                  <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>검색된 프로모션</div>
                  <div style={{fontSize:18,fontWeight:700,color:D.black}}>{matched.length}개</div>
                </div>
                {(searchStart||searchEnd)&&(
                  <div style={{background:D.surfaceAlt,borderRadius:7,padding:"10px 16px",minWidth:140}}>
                    <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>기간 매출{searchCh?` (${searchCh})`:" (전체)"}</div>
                    <div style={{fontSize:18,fontWeight:700,color:D.black}}>₩{revTotal.toLocaleString()}</div>
                  </div>
                )}
              </div>
              {matched.length>0&&(
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:D.surfaceAlt}}>
                      {["채널","프로모션명","기간","상태"].map(h=>(
                        <th key={h} style={{padding:"5px 8px",textAlign:"left",fontWeight:600,
                          color:D.textSub,borderBottom:`1px solid ${D.border}`,fontSize:12}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {matched.map(p=>{
                      const ended=isEnded(p);const hidden=hiddenIds.has(p.id);
                      return(
                        <tr key={p.id} style={{borderBottom:`1px solid ${D.border}`,opacity:hidden?0.55:1}}>
                          <td style={{padding:"5px 8px"}}>
                            <span style={{display:"inline-flex",alignItems:"center",gap:5}}>
                              <span style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),display:"inline-block"}}/>
                              <span style={{color:D.textSub}}>{p.platform}</span>
                            </span>
                          </td>
                          <td style={{padding:"5px 8px",fontWeight:600,color:D.text}}>
                            {p.name}
                            {hidden&&<span style={{marginLeft:6,fontSize:10,color:D.textMeta,fontWeight:400}}>(가려짐)</span>}
                          </td>
                          <td style={{padding:"5px 8px",color:D.textSub,whiteSpace:"nowrap",fontSize:12}}>
                            {p.start_date?.slice(0,10)} ~ {p.end_date?.slice(0,10)}
                          </td>
                          <td style={{padding:"5px 8px"}}>
                            {ended
                              ?<span style={{fontSize:11,color:D.red,fontWeight:600}}>종료</span>
                              :<span style={{fontSize:11,color:D.green,fontWeight:600}}>진행 중</span>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </>
          );
        })()}
      </Card>

      {/* 가려진 종료 프로모션 — 페이지 가장 하단. 가리기 버튼으로 숨긴 종료 프로모션만 모아 표시 + 임팩트 분석 진입 */}
      {hiddenLog.length>0&&(
        <Card style={{marginTop:12}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
            <div style={{fontWeight:600,fontSize:14,color:D.black}}>
              가려진 종료 프로모션
              <span style={{marginLeft:8,fontSize:11,color:D.textMeta,fontWeight:400}}>{hiddenLog.length}건</span>
            </div>
            {selHiddenIds.size>0&&(
              <button onClick={()=>delFromHiddenLog(selHiddenIds)}
                style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                  padding:"4px 12px",fontSize:12,cursor:"pointer",fontWeight:600}}>
                선택 삭제 ({selHiddenIds.size})
              </button>
            )}
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:D.surfaceAlt}}>
                <th style={{padding:"4px 6px",width:22}}/>
                {["채널","프로모션명","기간","할인율","가린 시각",""].map((h,i)=>(
                  <th key={i} style={{padding:"4px 8px",textAlign:"left",fontWeight:600,
                    color:D.textSub,borderBottom:`1px solid ${D.border}`,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...hiddenLog].sort((a,b)=>b.hidden_at>a.hidden_at?1:-1).map(h=>(
                <tr key={h.id} style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                  <td style={{padding:"4px 6px"}}>
                    <input type="checkbox" checked={selHiddenIds.has(h.id)}
                      onChange={ev=>{const s=new Set(selHiddenIds);ev.target.checked?s.add(h.id):s.delete(h.id);setSelHiddenIds(s);}}
                      style={{cursor:"pointer"}}/>
                  </td>
                  <td style={{padding:"4px 8px"}}>
                    <span style={{display:"inline-flex",alignItems:"center",gap:4}}>
                      <span style={{width:5,height:5,borderRadius:"50%",background:chColor(h.platform),display:"inline-block"}}/>
                      {h.platform}
                    </span>
                  </td>
                  <td style={{padding:"4px 8px",color:D.text,fontWeight:500}}>{h.name}</td>
                  <td style={{padding:"4px 8px",whiteSpace:"nowrap"}}>{h.start_date?.slice(0,10)} ~ {h.end_date?.slice(0,10)}</td>
                  <td style={{padding:"4px 8px",verticalAlign:"top",minWidth:170,maxWidth:220}}>
                    <DiscountPlanView plan={h.discount_plan}/>
                  </td>
                  <td style={{padding:"4px 8px",fontSize:11}}>{h.hidden_at?new Date(h.hidden_at).toLocaleString("ko-KR",{month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"}):""}</td>
                  <td style={{padding:"4px 8px"}}>
                    <button onClick={()=>setImpactModal(h)}
                      style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                        padding:"2px 9px",fontSize:11,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap"}}>
                      임팩트 분석
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      )}

      {impactModal&&<PromoImpactModal promo={impactModal} onClose={()=>setImpactModal(null)} revenues={revenues} storeSales={storeSales} orders={orders}/>}
    </div>
  );
}

// ─────────────────────────────────────────────
// 프로모션 임팩트 분석 모달 — 종료된 프로모션 임팩트 분석
//   - 일별 매출: 직전 동일 기간(점선 전) → 프로모션 기간(점선 후)
//   - Top 20: 프로모션 기간 + 해당 채널의 배송 완료된 상품 수량 랭킹
// ─────────────────────────────────────────────
function PromoImpactModal({ promo, onClose, revenues=[], storeSales=[], orders=[] }) {
  const ch=promo.platform;
  const dayMs=86400000;
  const todayStr=new Date().toISOString().slice(0,10);
  const yesterdayStr=new Date(Date.now()-dayMs).toISOString().slice(0,10);
  const promoStart=String(promo.start_date||"").slice(0,10);
  const promoEndRaw=String(promo.end_date||"").slice(0,10);
  // 진행중 프로모션: 종료일이 미래 → 분석 종료일은 어제로 클램프
  //   - 어제가 시작일보다 이르면(시작 당일) 시작일로 클램프
  const isOngoing=promoEndRaw>todayStr;
  const promoEnd=isOngoing
    ?(yesterdayStr>=promoStart?yesterdayStr:promoStart)
    :promoEndRaw;
  const dur=Math.max(0,(new Date(promoEnd)-new Date(promoStart))/dayMs); // 일 수 (포함 길이 = dur+1)
  const lenDays=dur+1;
  const prevStart=new Date(new Date(promoStart).getTime()-lenDays*dayMs).toISOString().slice(0,10);
  const prevEnd=new Date(new Date(promoStart).getTime()-dayMs).toISOString().slice(0,10);
  const modalCardRef=useRef(null);

  // 채널 매출 소스: 자사몰/29CM/무신사 → revenues, 오프라인 스토어 → storeSales
  const dailyRevenue=useMemo(()=>{
    const map={};
    const init=d=>{if(!map[d]) map[d]={date:d,revenue:0};};
    // 전체 두 기간 일자 채우기 (0으로 시작)
    let cur=new Date(prevStart);
    const last=new Date(promoEnd);
    while(cur<=last){ init(cur.toISOString().slice(0,10)); cur=new Date(cur.getTime()+dayMs); }
    if(ch==="오프라인 스토어"){
      storeSales.forEach(r=>{
        const d=r.sale_date;
        if(!d||d<prevStart||d>promoEnd) return;
        init(d);
        if(r.status==="배송") map[d].revenue+=(r.amount||0);
        else if(r.status==="반품") map[d].revenue-=(r.amount||0);
      });
    } else {
      revenues.forEach(r=>{
        if(r.channel!==ch) return;
        const d=r.date;
        if(!d||d<prevStart||d>promoEnd) return;
        init(d);
        map[d].revenue+=((r.amount||0)-(r.refund_amount||0));
      });
    }
    return Object.values(map).sort((a,b)=>a.date>b.date?1:-1).map(p=>({
      ...p,
      // 경계일(promoStart)에 prev/promo 둘 다 값을 줘서 두 라인을 시각적으로 잇는다
      prev:p.date<=promoStart?p.revenue:null,
      promo:p.date>=promoStart?p.revenue:null,
    }));
  },[ch,prevStart,prevEnd,promoStart,promoEnd,revenues,storeSales]);

  // Top 20 상품: 프로모션 기간 + 해당 채널의 배송 status 행
  const top20=useMemo(()=>{
    const OFFLINE=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
    const matchesCh=r=>{
      if(ch==="오프라인 스토어") return OFFLINE.has(r.channel||"");
      return (r.channel||"")===ch;
    };
    const m={};
    orders.forEach(r=>{
      if(r.status!=="배송") return;
      if(!matchesCh(r)) return;
      const d=r.order_date;
      if(!d||d<promoStart||d>promoEnd) return;
      const k=r.product_name||"미분류";
      if(!m[k]) m[k]={name:k,qty:0,orders:new Set()};
      m[k].qty+=(r.qty||1);
      const oid=r.order_no||r.order_id;
      if(oid) m[k].orders.add(oid);
    });
    return Object.values(m).map(p=>({...p,orders:p.orders.size}))
      .sort((a,b)=>b.qty-a.qty).slice(0,20);
  },[ch,promoStart,promoEnd,orders]);

  const prevTotal=dailyRevenue.filter(p=>p.date<promoStart).reduce((s,p)=>s+(p.revenue||0),0);
  const promoTotal=dailyRevenue.filter(p=>p.date>=promoStart).reduce((s,p)=>s+(p.revenue||0),0);
  const chg=prevTotal>0?((promoTotal-prevTotal)/prevTotal*100):null;

  return (
    <div onClick={onClose}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,
        display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div ref={modalCardRef} onClick={e=>e.stopPropagation()}
        style={{background:D.surface,borderRadius:14,padding:"24px 28px",
          width:"min(900px,95vw)",maxHeight:"90vh",overflowY:"auto",
          boxShadow:"0 8px 40px rgba(0,0,0,0.22)"}}>
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:8,gap:10}}>
          <div>
            <div style={{fontWeight:700,fontSize:16,color:D.black}}>
              {promo.name}
              <span style={{fontSize:12,color:D.textMeta,fontWeight:500,marginLeft:6}}>· 임팩트 분석</span>
              {isOngoing&&(
                <span style={{marginLeft:8,fontSize:10,fontWeight:700,color:"#fff",
                  background:D.green,padding:"2px 7px",borderRadius:10,verticalAlign:"middle"}}>
                  진행중
                </span>
              )}
            </div>
            {/* 기간 표기 — 진행중일 경우 분석기간이 어제까지로 클램프됨을 명확히 표시 */}
            <div style={{fontSize:11,color:D.textMeta,marginTop:5,lineHeight:1.7}}>
              <div>
                <span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:chColor(ch),verticalAlign:"middle",marginRight:5}}/>
                <b style={{color:D.text,fontWeight:600}}>{ch}</b>
                {isOngoing&&<span style={{marginLeft:6,color:D.textSub}}>· 프로모션 종료일 {promoEndRaw} (현재 진행 중)</span>}
              </div>
              <div>
                <span style={{display:"inline-block",minWidth:80,color:D.textSub,fontWeight:600}}>분석 기간</span>
                {promoStart} ~ {promoEnd} <span style={{color:D.textSub,fontWeight:500}}>({lenDays}일{isOngoing?", 시작일 ~ 어제":""})</span>
              </div>
              <div>
                <span style={{display:"inline-block",minWidth:80,color:D.textSub,fontWeight:600}}>직전 동일기간</span>
                {prevStart} ~ {prevEnd} <span style={{color:D.textSub,fontWeight:500}}>({lenDays}일)</span>
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:6,flexShrink:0}}>
            <CaptureBtn cardRef={modalCardRef} filename={`임팩트분석_${promo.name}_${promoStart}_${promoEnd}`} DC={{border:D.border,sub:D.textMeta}}/>
            <button onClick={onClose}
              style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
          </div>
        </div>

        {/* 매출 요약 */}
        <div style={{display:"flex",gap:14,marginTop:14,marginBottom:10,fontSize:12,flexWrap:"wrap"}}>
          <div style={{padding:"7px 12px",background:D.surfaceAlt,borderRadius:6}}>
            <span style={{color:D.textMeta}}>직전 매출</span> <b style={{marginLeft:6}}>{fmtWonShort(prevTotal)}</b>
          </div>
          <div style={{padding:"7px 12px",background:D.surfaceAlt,borderRadius:6}}>
            <span style={{color:D.textMeta}}>프로모션 매출</span> <b style={{marginLeft:6}}>{fmtWonShort(promoTotal)}</b>
          </div>
          {chg!==null&&(
            <div style={{padding:"7px 12px",background:chg>=0?`${D.green}12`:`${D.red}12`,borderRadius:6,color:chg>=0?D.green:D.red}}>
              <span>증감</span> <b style={{marginLeft:6}}>{chg>=0?"+":""}{chg.toFixed(1)}%</b>
            </div>
          )}
        </div>

        {/* 일별 매출 추이 */}
        <div style={{marginTop:6,marginBottom:18}}>
          <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"}}>
            일별 매출 — 직전 동일기간(점선) → 프로모션 기간(실선)
          </div>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={dailyRevenue} margin={{top:30,right:24,left:0,bottom:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="date" tick={{fontSize:10,fill:D.textMeta}} interval="preserveStartEnd"/>
              <YAxis tick={{fontSize:10,fill:D.textMeta}} tickFormatter={v=>fmtWonShort(v)}/>
              <Tooltip formatter={(v)=>v==null?"":fmtWon(v)} labelFormatter={d=>d}/>
              <ReferenceLine x={promoStart} stroke={D.red} strokeDasharray="4 3"
                label={{value:"프로모션 시작",position:"insideTop",offset:-18,fill:D.red,fontSize:11,fontWeight:600}}/>
              <Line type="monotone" dataKey="prev"  name="직전 매출"  stroke={D.textMeta} strokeWidth={2} strokeDasharray="5 4" dot={false} connectNulls={false}/>
              <Line type="monotone" dataKey="promo" name="프로모션 매출" stroke={chColor(ch)||D.blue} strokeWidth={2} dot={false} connectNulls={false}/>
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* Top 20 */}
        <div>
          <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"}}>
            프로모션 기간 판매 Top 20 ({ch}, 배송 완료 기준)
          </div>
          {top20.length===0?(
            <div style={{color:D.textMeta,fontSize:12,padding:"30px 0",textAlign:"center",background:D.surfaceAlt,borderRadius:6}}>
              해당 기간·채널의 배송 데이터가 없습니다.
            </div>
          ):(
            <div style={{maxHeight:360,overflowY:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta,position:"sticky",top:0,background:D.surface}}>
                  <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500,width:30}}>#</th>
                  <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500}}>상품명</th>
                  <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>판매 수량(장)</th>
                  <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>주문 건</th>
                </tr></thead>
                <tbody>
                  {top20.map((p,i)=>(
                    <tr key={p.name+i} style={{borderBottom:`1px solid ${D.border}`}}>
                      <td style={{padding:"5px 7px",color:D.textMeta}}>{i+1}</td>
                      <td style={{padding:"5px 7px",color:D.text,maxWidth:380,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={p.name}>{p.name}</td>
                      <td style={{padding:"5px 7px",textAlign:"right",color:D.blue,fontWeight:600}}>{p.qty.toLocaleString()}</td>
                      <td style={{padding:"5px 7px",textAlign:"right",color:D.textSub}}>{p.orders.toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// CS DATA INPUT
// ─────────────────────────────────────────────
function CSDataInput() {
  const [csData,setCSData]=useState(getCSData);
  const today=new Date().toISOString().slice(0,10);
  const [date,setDate]=useState(today);
  const [filterProd,setFilterProd]=useState("");
  const [csvResult,setCsvResult]=useState(null);
  const [editCell,setEditCell]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [selected,setSelected]=useState(new Set());
  const [delConfirm,setDelConfirm]=useState(false);
  const [csFilterStart,setCsFilterStart]=useState("");
  const [csFilterEnd,setCsFilterEnd]=useState("");
  const [csCalOpen,setCsCalOpen]=useState(null);

  useEffect(()=>{
    (async()=>{
      const local=getCSData();
      const db=await getSupabase();
      const{data,error}=await db.from("cs_data").select("*").order("id",{ascending:false});
      if(!error&&data){
        if(data.length>0){saveCSData(data);setCSData(data);}
        else if(local.length>0){await db.from("cs_data").insert(local);}
      }
    })();
  },[]);

  const handleCSVFile=useCallback(file=>{
    if(!file)return;
    setCsvResult(null);
    parseAnyFile(file,{header:true,skipEmptyLines:true},async({data})=>{
        try{
          const cols=Object.keys(data[0]||{});
          const lc=cols.map(c=>c.toLowerCase().replace(/[\s\[\]()]/g,""));
          const findCol=(...kws)=>{const i=lc.findIndex(c=>kws.some(k=>c.includes(k)));return i>=0?cols[i]:null;};
          const prodCol=findCol("상품명","상품","product","item");
          const reasonCol=findCol("반품사유","반품","사유","reason","취소");
          const dateCol=findCol("날짜","date","일자","접수일","처리일");
          const chCol=findCol("판매처","채널","channel","플랫폼","mall");
          const missingCs=[];
          if(!prodCol)   missingCs.push("상품명");
          if(!reasonCol) missingCs.push("반품사유");
          if(missingCs.length){
            throw new Error(uploadErrColumns({
              missing:missingCs,
              required:["날짜","판매처","상품명","반품사유"],
              headers:cols,
            }));
          }

          const extractReason=raw=>{
            const s=String(raw||"").toLowerCase();
            if(s.includes("사이즈")||s.includes("size")||s.includes("미스")) return "사이즈 미스";
            if(s.includes("퀄리티")||s.includes("불량")||s.includes("품질")) return "퀄리티";
            if(s.includes("배송")&&!s.includes("배송비")&&!s.includes("회수")) return "배송";
            if(s.includes("단순변심")||s.includes("변심")) return "단순변심";
            return "단순변심";
          };

          const splitProducts=raw=>{
            const parts=[];let cur="";let depth=0;
            for(const ch of String(raw||"")){
              if(ch==="["){depth++;cur+=ch;}
              else if(ch==="]"){depth=Math.max(0,depth-1);cur+=ch;}
              else if(ch===","&&depth===0){const t=cur.trim().replace(/\t/g," ");if(t)parts.push(t);cur="";}
              else{cur+=ch;}
            }
            const t=cur.trim().replace(/\t/g," ");if(t)parts.push(t);
            return parts.length?parts:[String(raw||"").trim()];
          };

          let lastDate=today;
          const newEntries=[];
          for(const r of data){
            const rawDate=dateCol?String(r[dateCol]||"").trim():"";
            if(rawDate.includes("반품취소")) continue;
            const parsedDate=toDate(rawDate);
            if(parsedDate) lastDate=parsedDate;
            const rawProd=prodCol?String(r[prodCol]||"").trim():"";
            const rawReason=reasonCol?String(r[reasonCol]||"").trim():"";
            if(!rawProd&&!rawReason) continue;
            const reason=extractReason(rawReason);
            const rawCh=chCol?String(r[chCol]||"").trim():"";
            const channel=rawCh?normChannel(rawCh):"자사몰";
            for(const prod of splitProducts(rawProd)){
              if(!prod) continue;
              newEntries.push({id:Date.now()+Math.random(),date:lastDate,product_name:prod,return_reason:reason,channel});
            }
          }

          if(!newEntries.length)throw new Error("유효한 데이터 행이 없습니다");
          const next=[...newEntries,...csData];
          saveCSData(next);setCSData(next);
          setCsvResult({type:"success",msg:`${newEntries.length}건 추가 완료`});
          const db=await getSupabase();
          await db.from("cs_data").insert(newEntries);
        }catch(e){setCsvResult({type:"error",msg:e.message});}
      },e=>setCsvResult({type:"error",msg:e.message}));
  },[csData,today]);

  const del=async id=>{
    const next=csData.filter(r=>r.id!==id);
    saveCSData(next);setCSData(next);setSelected(s=>{const n=new Set(s);n.delete(id);return n;});
    const db=await getSupabase();
    await db.from("cs_data").delete().eq("id",id);
  };
  const toggleSel=id=>setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>{
    const ids=filtered.map(r=>r.id);
    const allSel=ids.length>0&&ids.every(id=>selected.has(id));
    setSelected(s=>{const n=new Set(s);ids.forEach(id=>allSel?n.delete(id):n.add(id));return n;});
  };
  const delSelected=async()=>{
    const ids=[...selected];
    const next=csData.filter(r=>!selected.has(r.id));
    saveCSData(next);setCSData(next);setSelected(new Set());setDelConfirm(false);
    const db=await getSupabase();
    await Promise.all(ids.map(id=>db.from("cs_data").delete().eq("id",id)));
  };

  const startCsEdit=(id,field,val)=>{setEditCell({id,field});setEditVal(String(val??""));};
  const saveCsEdit=async()=>{
    if(!editCell) return;
    const next=csData.map(r=>r.id===editCell.id?{...r,[editCell.field]:editVal}:r);
    saveCSData(next);setCSData(next);
    const db=await getSupabase();
    await db.from("cs_data").update({[editCell.field]:editVal}).eq("id",editCell.id);
    setEditCell(null);
  };

  const filtered=csData.filter(r=>
    (!filterProd||(r.product_name||"").includes(filterProd)||(r.date||"").includes(filterProd)||(r.channel||"").includes(filterProd)||(r.return_reason||"").includes(filterProd))&&
    (!csFilterStart||r.date>=csFilterStart)&&
    (!csFilterEnd||r.date<=csFilterEnd)
  );

  return (
    <div style={{display:"grid",gridTemplateColumns:"280px 1fr",gap:14}}>
      <Card>
        <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>CS 반품 데이터 업로드</div>
        <div style={{fontSize:10,color:D.textMeta,marginBottom:8}}>날짜 없는 행의 기준일</div>
        <CalendarPicker mode="single" value={date} onChange={setDate}/>
        <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${D.border}`}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6,letterSpacing:"0.06em",textTransform:"uppercase"}}>CSV 업로드</div>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:8,lineHeight:1.6}}>
            필수: <strong>[상품]</strong> · <strong>[반품 사유]</strong><br/>
            선택: [날짜] [판매처]
          </div>
          <DropZone onFile={handleCSVFile} label="반품 CS 파일 업로드"
            columns="날짜 · 판매처 · 상품명 · 반품사유"/>
          {csvResult&&<Alert type={csvResult.type} msg={csvResult.msg}/>}
        </div>
      </Card>
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,gap:8,flexWrap:"wrap"}}>
          <div style={{fontWeight:600,fontSize:13}}>반품 사유 내역</div>
          <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
            <CalRangeDrop id="csFilter" start={csFilterStart} end={csFilterEnd}
              onRange={(s,e)=>{setCsFilterStart(s);setCsFilterEnd(e);setDelConfirm(false);}}
              openId={csCalOpen} setOpenId={setCsCalOpen}/>
          <input value={filterProd} onChange={e=>{setFilterProd(e.target.value);setDelConfirm(false);}}
            style={{...inp,width:200,fontSize:11,padding:"5px 8px"}} placeholder="날짜·상품명·판매처 검색"/>
          </div>
        </div>
        {selected.size>0&&(
          <div style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
            <span style={{fontSize:11,color:D.textMeta}}>{selected.size}개 선택</span>
            {!delConfirm
              ?<button onClick={()=>setDelConfirm(true)}
                 style={{background:"#e55",color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>삭제</button>
              :<>
                <button onClick={delSelected}
                  style={{background:"#e55",color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>확인 삭제</button>
                <button onClick={()=>setDelConfirm(false)}
                  style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>취소</button>
              </>}
          </div>
        )}
        <div style={{overflowY:"auto",maxHeight:480}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
              <th style={{padding:"6px 8px",width:28}}>
                <input type="checkbox" checked={filtered.length>0&&filtered.every(r=>selected.has(r.id))}
                  onChange={toggleAll} style={{cursor:"pointer"}}/>
              </th>
              {["날짜","판매처","상품명","반품 사유",""].map(h=>(
                <th key={h} style={{padding:"6px 8px",textAlign:"left",color:D.textMeta,fontWeight:400}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {filtered.slice(0,200).map(r=>{
                const isSel=selected.has(r.id);
                const cell=(field,content,style={})=>{
                  const isEd=editCell?.id===r.id&&editCell?.field===field;
                  return(
                    <td style={{padding:"5px 8px",cursor:"pointer",...style}}
                      title="더블클릭하여 수정"
                      onDoubleClick={()=>startCsEdit(r.id,field,r[field])}>
                      {isEd
                        ?<input autoFocus value={editVal} onChange={e=>setEditVal(e.target.value)}
                            onBlur={saveCsEdit} onKeyDown={e=>{if(e.key==="Enter")saveCsEdit();if(e.key==="Escape")setEditCell(null);}}
                            style={{width:"100%",border:`1px solid ${D.primary}`,borderRadius:3,padding:"1px 4px",fontSize:11}}/>
                        :content}
                    </td>
                  );
                };
                return(
                  <tr key={r.id} style={{borderBottom:`1px solid ${D.border}`,background:isSel?D.surfaceAlt:"transparent"}}>
                    <td style={{padding:"5px 8px"}}>
                      <input type="checkbox" checked={isSel} onChange={()=>toggleSel(r.id)} style={{cursor:"pointer"}}/>
                    </td>
                    {cell("date",<span style={{color:D.textMeta,whiteSpace:"nowrap"}}>{r.date}</span>)}
                    {cell("channel",<span style={{color:chColor(r.channel),fontWeight:600}}>{r.channel}</span>)}
                    {cell("product_name",r.product_name,{maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"})}
                    {cell("return_reason",<span style={{color:D.textSub}}>{r.return_reason}</span>)}
                    <td style={{padding:"5px 8px"}}>
                      <button onClick={()=>del(r.id)}
                        style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:10}}>✕</button>
                    </td>
                  </tr>
                );
              })}
              {filtered.length===0&&<tr><td colSpan={6} style={{padding:24,textAlign:"center",color:D.textMeta}}>데이터 없음</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT — 매출 입력
// ─────────────────────────────────────────────
const REVENUE_CHANNELS = ["자사몰","29CM","무신사"];

function RevenueForm({ onUpdate, histRefreshKey=0 }) {
  const today=new Date().toISOString().slice(0,10);
  const [date,setDate]=useState(today);
  const [dateMode,setDateMode]=useState("single"); // "single"|"range"
  const [dateEnd,setDateEnd]=useState(today);
  const [ch,setCh]=useState(REVENUE_CHANNELS[0]);
  const [amt,setAmt]=useState("");
  const [orderCnt,setOrderCnt]=useState("");
  const [refundAmt,setRefundAmt]=useState("");
  const [refundCnt,setRefundCnt]=useState("");
  // CSV 업로드 상태
  const [csvPreview,setCsvPreview]=useState(null); // {rows, overlaps}
  const [csvConflictChoice,setCsvConflictChoice]=useState(null); // null | "new" | "keep"
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [history,setHistory]=useState([]);
  const [histTs,setHistTs]=useState(null);
  const [editId,setEditId]=useState(null);
  const [editData,setEditData]=useState({});
  const [deleteConfirm,setDeleteConfirm]=useState(null);
  const [histChFilter,setHistChFilter]=useState("전체");
  const [chDeleteConfirm,setChDeleteConfirm]=useState(false);

  const loadHistory=useCallback(async()=>{
    const db=await getSupabase();
    let all=[];let from=0;const PAGE=1000;
    while(true){
      const{data,error}=await db.from("revenues").select("*").order("date",{ascending:false}).range(from,from+PAGE-1);
      if(error||!data||data.length===0) break;
      all=all.concat(data);
      if(data.length<PAGE) break;
      from+=PAGE;
    }
    // 대시보드와 동일하게 dedup (date+channel 기준 최신 id만)
    const revMap={};
    all.forEach(r=>{const k=`${r.date}__${r.channel}`;if(!revMap[k]||r.id>revMap[k].id)revMap[k]=r;});
    setHistory(Object.values(revMap).sort((a,b)=>b.date>a.date?1:b.date<a.date?-1:0));
    setHistTs(nowStr());
  },[]);

  const handleSave=async()=>{
    const num=Number(amt.replace(/,/g,""));
    if(!amt||isNaN(num)){setResult({type:"error",msg:"매출 금액을 입력해주세요."});return;}
    setLoading(true);setResult(null);
    const db=await getSupabase();
    // Build date list
    const dates=[];
    if(dateMode==="range"&&dateEnd>date){
      const cur=new Date(date);
      const end=new Date(dateEnd);
      while(cur<=end){dates.push(cur.toISOString().slice(0,10));cur.setDate(cur.getDate()+1);}
    } else {
      dates.push(date);
    }
    // 갈음 안내용: 삭제 직전 기존 건수 조회
    const {count:prevCount}=await db.from("revenues").select("*",{count:"exact",head:true}).in("date",dates).eq("channel",ch);
    // DELETE 후 INSERT — UNIQUE 제약 없이도 중복 방지
    const {error:de}=await db.from("revenues").delete().in("date",dates).eq("channel",ch);
    if(de){setResult({type:"error",msg:"기존 삭제 실패: "+de.message});setLoading(false);return;}
    for(const d of dates){
      const {error}=await db.from("revenues").insert({
        date:d,channel:ch,
        amount:Math.round(num/dates.length),
        order_count:Math.round((Number(orderCnt)||0)/dates.length),
        refund_amount:Math.round((Number(refundAmt.replace(/,/g,""))||0)/dates.length),
        refund_count:Math.round((Number(refundCnt)||0)/dates.length),
      });
      if(error){setResult({type:"error",msg:error.message});setLoading(false);return;}
    }
    const ts2=nowStr();
    const replaceNote=prevCount>0?` (기존 ${prevCount}건 대체됨)`:"";
    setResult({type:"success",msg:`${dates.length}일 저장 완료${replaceNote}`,ts:ts2});
    setAmt("");setOrderCnt("");setRefundAmt("");setRefundCnt("");
    onUpdate(ts2);if(history.length)loadHistory();
    setLoading(false);
  };

  const startEdit=r=>{
    setEditId(r.id);
    setEditData({
      date:r.date,channel:r.channel,
      amount:r.amount||0,order_count:r.order_count||0,
      refund_amount:r.refund_amount||0,refund_count:r.refund_count||0,
    });
    setDeleteConfirm(null);
  };

  const saveEdit=async()=>{
    const db=await getSupabase();
    const {error}=await db.from("revenues").update({
      amount:Number(editData.amount)||0,
      order_count:Number(editData.order_count)||0,
      refund_amount:Number(editData.refund_amount)||0,
      refund_count:Number(editData.refund_count)||0,
    }).eq("id",editId);
    if(!error){setEditId(null);loadHistory();}
  };

  const handleDelete=async id=>{
    if(deleteConfirm!==id){setDeleteConfirm(id);return;}
    const db=await getSupabase();
    await db.from("revenues").delete().eq("id",id);
    setDeleteConfirm(null); loadHistory();
  };

  const handleChannelDelete=async()=>{
    if(!chDeleteConfirm){setChDeleteConfirm(true);return;}
    const db=await getSupabase();
    await db.from("revenues").delete().eq("channel",histChFilter);
    setChDeleteConfirm(false);
    const ts2=nowStr();
    onUpdate(ts2);
    loadHistory();
  };

  const handleCsvFile=useCallback(file=>{
    parseAnyFile(file,{header:true,skipEmptyLines:true},async({data})=>{
      const cols=Object.keys(data[0]||{});
      const lc=cols.map(c=>c.toLowerCase().replace(/[\s\[\]()_]/g,""));
      const find=(...kws)=>{const i=lc.findIndex(c=>kws.some(k=>c.includes(k)));return i>=0?cols[i]:null;};
      const dateCol=find("날짜","date","일자");
      const chCol=find("판매처","채널","channel","플랫폼");
      const amtCol=find("매출","amount","금액");
      const ordCol=find("주문수","주문건","ordercount","order");
      const refAmtCol=find("환불금","refundamount","환불액");
      const refCntCol=find("환불수","환불건","refundcount");
      if(!dateCol||!amtCol){
        const missingRev=[];
        if(!dateCol) missingRev.push("날짜");
        if(!amtCol)  missingRev.push("매출금액");
        setCsvPreview({error:uploadErrColumns({
          missing:missingRev,
          required:["날짜","판매처","매출금액","주문수","환불금","환불수"],
          headers:cols,
        })});return;
      }
      const rows=data.filter(r=>r[dateCol]&&toDate(r[dateCol])).map(r=>({
        date:toDate(r[dateCol]),
        channel:chCol?normChannel(r[chCol]):ch,
        amount:Number(String(r[amtCol]||"0").replace(/[^0-9.-]/g,""))||0,
        order_count:ordCol?Number(r[ordCol]||0):0,
        refund_amount:refAmtCol?Math.abs(Number(String(r[refAmtCol]||"0").replace(/[^0-9.-]/g,""))):0,
        refund_count:refCntCol?Math.abs(Number(r[refCntCol]||0)):0,
      }));
      // 기존 데이터와 겹치는 (date, channel) 쌍 확인
      const db=await getSupabase();
      const dates=[...new Set(rows.map(r=>r.date))];
      const {data:existing}=await db.from("revenues").select("date,channel").in("date",dates);
      const existSet=new Set((existing||[]).map(r=>`${r.date}__${r.channel}`));
      const overlaps=rows.filter(r=>existSet.has(`${r.date}__${r.channel}`));
      setCsvPreview({rows,overlaps});
      setCsvConflictChoice(null);
    });
  },[]);

  const handleCsvUpload=async(choice)=>{
    if(!csvPreview?.rows) return;
    const db=await getSupabase();
    let toUpload=csvPreview.rows;
    if(choice==="keep"){
      const overlapKeys=new Set(csvPreview.overlaps.map(r=>`${r.date}__${r.channel}`));
      toUpload=toUpload.filter(r=>!overlapKeys.has(`${r.date}__${r.channel}`));
    }
    // upsert 대신 DELETE→INSERT: UNIQUE 제약 없이도 중복 방지
    const delDates=[...new Set(toUpload.map(r=>r.date))];
    const delChs=[...new Set(toUpload.map(r=>r.channel))];
    for(let i=0;i<delDates.length;i+=100){
      const batch=delDates.slice(i,i+100);
      const {error:de}=await db.from("revenues").delete().in("date",batch).in("channel",delChs);
      if(de){setResult({type:"error",msg:"기존 데이터 삭제 실패: "+de.message});return;}
    }
    for(let i=0;i<toUpload.length;i+=200){
      const {error}=await db.from("revenues").insert(toUpload.slice(i,i+200));
      if(error){setResult({type:"error",msg:error.message});return;}
    }
    const ts2=nowStr();
    const overlapCount=choice==="new"?csvPreview.overlaps?.length||0:0;
    const replaceNote=overlapCount>0?` (기존 ${overlapCount}건 대체됨)`:"";
    setResult({type:"success",msg:`${toUpload.length}건 저장 완료${replaceNote}`,ts:ts2});
    setCsvPreview(null);setCsvConflictChoice(null);
    onUpdate(ts2);loadHistory();
  };

  const inp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
    padding:"7px 10px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box"};
  const numInp=(v,fn)=>(
    <input type="text" value={v} onChange={e=>fn(e.target.value.replace(/[^0-9,]/g,""))} style={inp}/>
  );

  return (
    <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>

      {/* CSV 충돌 다이얼로그 */}
      {csvPreview&&!csvPreview.error&&csvPreview.overlaps?.length>0&&!csvConflictChoice&&(
        <div style={{gridColumn:"1/-1",background:"#fff9e6",border:`1px solid ${D.amber}`,
          borderRadius:8,padding:"14px 18px"}}>
          <div style={{fontWeight:600,marginBottom:6,color:D.amber}}>
            ⚠ 기존 데이터와 겹치는 항목 {csvPreview.overlaps.length}건 발견
          </div>
          <div style={{fontSize:11,color:D.textSub,marginBottom:10}}>
            {csvPreview.overlaps.slice(0,5).map(r=>`${r.date} · ${r.channel}`).join(" / ")}
            {csvPreview.overlaps.length>5&&` 외 ${csvPreview.overlaps.length-5}건`}
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>{setCsvConflictChoice("new");handleCsvUpload("new");}}
              style={{background:D.red,color:"#fff",border:"none",borderRadius:6,
                padding:"7px 16px",fontSize:12,cursor:"pointer",fontWeight:600}}>
              새 데이터로 덮어쓰기
            </button>
            <button onClick={()=>{setCsvConflictChoice("keep");handleCsvUpload("keep");}}
              style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
                padding:"7px 16px",fontSize:12,cursor:"pointer",fontWeight:600}}>
              겹치는 항목 건너뛰기
            </button>
            <button onClick={()=>setCsvPreview(null)}
              style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"7px 16px",fontSize:12,cursor:"pointer",color:D.textSub}}>
              취소
            </button>
          </div>
        </div>
      )}

      <Card>
        <div style={{fontWeight:600,marginBottom:14,fontSize:13}}>매출 입력</div>

        <div style={{marginBottom:10}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6}}>날짜</div>
          <div style={{display:"flex",gap:4,marginBottom:8}}>
            {[["single","단일"],["range","기간"]].map(([k,l])=>(
              <button key={k} onClick={()=>setDateMode(k)}
                style={{flex:1,background:dateMode===k?D.black:"transparent",
                  color:dateMode===k?"#fff":D.textSub,
                  border:`1px solid ${dateMode===k?D.black:D.border}`,
                  borderRadius:5,padding:"5px 4px",fontSize:11,cursor:"pointer"}}>
                {l}
              </button>
            ))}
          </div>
          {dateMode==="single"
            ? <CalendarPicker mode="single" value={date} onChange={setDate}/>
            : <CalendarPicker mode="range" rangeStart={date} rangeEnd={dateEnd}
                onRangeChange={({start,end})=>{setDate(start);setDateEnd(end||start);}}/>
          }
        </div>

        <div style={{marginBottom:10}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6}}>판매처</div>
          <div style={{display:"flex",gap:6}}>
            {REVENUE_CHANNELS.map(c=>(
              <button key={c} onClick={()=>setCh(c)}
                style={{flex:1,background:ch===c?D.black:"transparent",
                  color:ch===c?"#fff":D.textSub,
                  border:`1px solid ${ch===c?D.black:D.border}`,
                  borderRadius:6,padding:"7px 4px",fontSize:12,cursor:"pointer",fontWeight:ch===c?600:400}}>
                {c}
              </button>
            ))}
          </div>
        </div>

        {[
          {label:"매출 금액",val:amt,fn:setAmt,ph:"1500000"},
          {label:"주문 수",val:orderCnt,fn:setOrderCnt,ph:"0",num:true},
          {label:"환불 금액",val:refundAmt,fn:setRefundAmt,ph:"0"},
          {label:"환불 수",val:refundCnt,fn:setRefundCnt,ph:"0",num:true},
        ].map(({label,val,fn,ph,num})=>(
          <div key={label} style={{marginBottom:10}}>
            <div style={{color:D.textMeta,fontSize:10,marginBottom:4}}>{label}</div>
            <input type="text" value={val} placeholder={ph}
              onChange={e=>fn(e.target.value.replace(num?/[^0-9]/g:/[^0-9,]/g,""))}
              style={inp}/>
          </div>
        ))}

        {(amt||refundAmt)&&(()=>{
          const gross=Number(amt.replace(/,/g,"")||0);
          const refundNum=Number(refundAmt.replace(/,/g,"")||0);
          const net=gross-refundNum;
          return(
            <div style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:7,padding:"8px 12px",marginBottom:10,fontSize:11,color:D.textSub}}>
              {gross>0&&<div style={{display:"flex",justifyContent:"space-between"}}>
                <span>매출</span><span>₩{gross.toLocaleString()}</span>
              </div>}
              {refundNum>0&&<div style={{display:"flex",justifyContent:"space-between",color:D.red,marginTop:2}}>
                <span>환불</span><span>-₩{refundNum.toLocaleString()}</span>
              </div>}
              <div style={{display:"flex",justifyContent:"space-between",fontWeight:700,color:D.text,marginTop:4,paddingTop:4,borderTop:`1px solid ${D.border}`}}>
                <span>순매출</span><span>₩{net.toLocaleString()}</span>
              </div>
            </div>
          );
        })()}
        <Btn onClick={handleSave} disabled={loading} style={{width:"100%"}}>
          {loading?"저장 중...":"저장"}
        </Btn>
        {result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}

        <div style={{marginTop:14,borderTop:`1px solid ${D.border}`,paddingTop:14}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6}}>CSV 일괄 업로드</div>
          <DropZone onFile={handleCsvFile} label="매출 파일 업로드"
            columns="날짜 · 판매처 · 매출금액 · 주문수 · 환불금 · 환불수"/>
          {csvPreview?.error&&<div style={{color:D.red,fontSize:10,marginTop:4}}>{csvPreview.error}</div>}
          {csvPreview&&!csvPreview.error&&(csvPreview.overlaps?.length===0||csvConflictChoice)&&(
            <div style={{marginTop:6,display:"flex",gap:8,alignItems:"center"}}>
              <span style={{fontSize:11,color:D.textSub}}>{csvPreview.rows.length}건 파싱됨 · {[...new Set(csvPreview.rows.map(r=>r.channel))].join(", ")}</span>
              {!csvConflictChoice&&(
                <button onClick={()=>handleCsvUpload("new")}
                  style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                    padding:"5px 12px",fontSize:11,cursor:"pointer"}}>
                  저장
                </button>
              )}
            </div>
          )}
        </div>
      </Card>

      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{display:"flex",alignItems:"baseline",gap:6}}>
            <span style={{fontWeight:600,fontSize:13}}>입력 내역</span>
            <UpdatedAt ts={histTs}/>
          </div>
          <Btn onClick={loadHistory} variant="ghost" style={{padding:"4px 11px",fontSize:11}}>불러오기</Btn>
        </div>
        {history.length>0&&(
          <div style={{display:"flex",gap:4,marginBottom:10,flexWrap:"wrap",alignItems:"center"}}>
            {["전체",...REVENUE_CHANNELS].map(c=>{
              const cnt=c==="전체"?history.length:history.filter(r=>r.channel===c).length;
              const sum=c==="전체"
                ?history.reduce((s,r)=>s+(r.amount||0),0)
                :history.filter(r=>r.channel===c).reduce((s,r)=>s+(r.amount||0),0);
              return(
                <button key={c} onClick={()=>{setHistChFilter(c);setChDeleteConfirm(false);}}
                  style={{background:histChFilter===c?D.black:"transparent",
                    color:histChFilter===c?"#fff":D.textSub,
                    border:`1px solid ${histChFilter===c?D.black:D.border}`,
                    borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",lineHeight:1.4}}>
                  {c} <span style={{opacity:0.7,fontSize:10}}>({cnt}건 · ₩{(sum/1e4).toFixed(0)}만)</span>
                </button>
              );
            })}
            {histChFilter!=="전체"&&(
              <button onClick={handleChannelDelete}
                style={{marginLeft:"auto",background:chDeleteConfirm?D.red:"transparent",
                  color:chDeleteConfirm?"#fff":D.red,
                  border:`1px solid ${D.red}`,borderRadius:6,
                  padding:"4px 12px",fontSize:11,cursor:"pointer",whiteSpace:"nowrap"}}>
                {chDeleteConfirm?`'${histChFilter}' 전체 삭제 확인`:`'${histChFilter}' 전체 삭제`}
              </button>
            )}
          </div>
        )}
        {history.length>0?(()=>{
          const filtered=histChFilter==="전체"?history:history.filter(r=>r.channel===histChFilter);
          const totalAmt=filtered.reduce((s,r)=>s+(r.amount||0),0);
          return(
          <div style={{overflowY:"auto",maxHeight:520}}>
            <div style={{fontSize:11,color:D.textMeta,marginBottom:6}}>
              {filtered.length}건 · 합계 ₩{totalAmt.toLocaleString()}
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
                {["날짜","판매처","매출","주문","환불금","환불수",""].map(h=>(
                  <th key={h} style={{padding:"5px 7px",textAlign:h===""?"center":"left",color:D.textMeta,fontWeight:400}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {filtered.map(r=>(
                  <tr key={r.id} style={{borderBottom:`1px solid ${D.border}`,
                    background:editId===r.id?D.surfaceAlt:"transparent"}}>
                    {editId===r.id?(
                      <>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{r.date}</td>
                        <td style={{padding:"5px 7px"}}>{r.channel}</td>
                        {["amount","order_count","refund_amount","refund_count"].map(k=>(
                          <td key={k} style={{padding:"4px 5px"}}>
                            <input type="text" value={editData[k]}
                              onChange={e=>setEditData(prev=>({...prev,[k]:e.target.value}))}
                              style={{width:70,border:`1px solid ${D.border}`,borderRadius:4,
                                padding:"3px 5px",fontSize:11}}/>
                          </td>
                        ))}
                        <td style={{padding:"4px 5px",whiteSpace:"nowrap"}}>
                          <button onClick={saveEdit} style={{background:D.green,color:"#fff",border:"none",
                            borderRadius:4,padding:"3px 7px",fontSize:10,cursor:"pointer",marginRight:3}}>저장</button>
                          <button onClick={()=>setEditId(null)} style={{background:"transparent",
                            border:`1px solid ${D.border}`,borderRadius:4,padding:"3px 7px",fontSize:10,cursor:"pointer"}}>취소</button>
                        </td>
                      </>
                    ):(
                      <>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{r.date}</td>
                        <td style={{padding:"5px 7px"}}>{r.channel}</td>
                        <td style={{padding:"5px 7px",fontWeight:600}}>₩{(r.amount||0).toLocaleString()}</td>
                        <td style={{padding:"5px 7px",color:D.textSub}}>{r.order_count||0}</td>
                        <td style={{padding:"5px 7px",color:D.textSub}}>₩{(r.refund_amount||0).toLocaleString()}</td>
                        <td style={{padding:"5px 7px",color:D.textSub}}>{r.refund_count||0}</td>
                        <td style={{padding:"4px 5px",whiteSpace:"nowrap"}}>
                          <button onClick={()=>startEdit(r)} style={{background:"transparent",
                            border:`1px solid ${D.border}`,borderRadius:4,padding:"3px 7px",
                            fontSize:10,cursor:"pointer",marginRight:3,color:D.textSub}}>수정</button>
                          <button onClick={()=>handleDelete(r.id)}
                            style={{background:deleteConfirm===r.id?D.red:"transparent",
                              color:deleteConfirm===r.id?"#fff":D.red,
                              border:`1px solid ${deleteConfirm===r.id?D.red:D.red}`,
                              borderRadius:4,padding:"3px 7px",fontSize:10,cursor:"pointer"}}>
                            {deleteConfirm===r.id?"확인":"삭제"}
                          </button>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          );
        })():<div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>불러오기를 눌러주세요</div>}
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT — 입고 CSV
// ─────────────────────────────────────────────
function StockUploader({ onUpdate, histRefreshKey=0 }) {
  const today=new Date().toISOString().slice(0,10);
  const [startDate,setStartDate]=useState(today);
  const [endDate,setEndDate]=useState(today);
  const [step,setStep]=useState(0);
  const [fileName,setFileName]=useState("");
  const [preview,setPreview]=useState(null);
  const [existing,setExisting]=useState(null);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [history,setHistory]=useState(null);
  const [histTs,setHistTs]=useState(null);
  const [selected,setSelected]=useState(new Set());
  const [deleteConfirm,setDeleteConfirm]=useState(false);
  const [histFilter,setHistFilter]=useState("");
  const dateValid=startDate&&endDate&&startDate<=endDate;

  const loadHistory=useCallback(async()=>{
    const db=await getSupabase();
    let all=[];let from=0;const PAGE=1000;
    while(true){
      const{data,error}=await db.from("stock_uploads").select("*").order("upload_date",{ascending:false}).order("product_name").range(from,from+PAGE-1);
      if(error||!data||data.length===0) break;
      all=all.concat(data);
      if(data.length<PAGE) break;
      from+=PAGE;
    }
    setHistory(all); setHistTs(nowStr()); setSelected(new Set()); setDeleteConfirm(false);
  },[]);

  const toggleSelect=id=>setSelected(prev=>{const s=new Set(prev);s.has(id)?s.delete(id):s.add(id);return s;});
  const toggleAll=rows=>{
    const ids=rows.map(r=>r.id);
    setSelected(prev=>ids.every(id=>prev.has(id))?new Set():new Set(ids));
  };

  const handleDeleteSelected=async()=>{
    if(!deleteConfirm){setDeleteConfirm(true);return;}
    const db=await getSupabase();
    const ids=[...selected];
    for(let i=0;i<ids.length;i+=100){
      await db.from("stock_uploads").delete().in("id",ids.slice(i,i+100));
    }
    const ts2=nowStr(); onUpdate(ts2); loadHistory(); setDeleteConfirm(false);
  };

  const confirmDate=async()=>{
    setLoading(true);
    const db=await getSupabase();
    const {data}=await db.from("stock_uploads").select("*").gte("upload_date",startDate).lte("upload_date",endDate).order("upload_date");
    setExisting(data||[]); setStep(1); setLoading(false);
  };
  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    parseAnyFile(file,{header:true,skipEmptyLines:true},({data})=>{
        try{
          if(!data?.length) throw new Error(uploadErrParse("파일에 데이터 행이 없습니다"));
          const headers=Object.keys(data[0]||{});
          const f=detectFields(headers);
          if(!f.product){
            throw new Error(uploadErrColumns({
              missing:["상품명"],
              required:["상품명","옵션","수량","메모"],
              headers,
            }));
          }
          const rows=data.filter(r=>r[f.product]).map(r=>({
            product_name:String(r[f.product]||"").trim(),
            option_name:String(r[f.option]||"").trim(),
            qty:toNum(r[f.qty]),
            memo:String(r[f.memo]||"").trim(),
          }));
          if(!rows.length) throw new Error("파싱된 행이 0건입니다. '상품명' 컬럼에 값이 있는 행이 1개 이상 있어야 합니다.");
          setPreview(rows); setStep(2);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e?.message||String(e)}));
  },[]);
  const handleUpload=async()=>{
    if(!preview?.length||!dateValid) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    // 업로드 파일 내 상품들의 기존 데이터를 날짜 무관 전부 삭제 (이전 배치와 날짜 충돌 방지)
    const productNames=[...new Set(preview.map(r=>r.product_name).filter(Boolean))];
    for(let pi=0;pi<productNames.length;pi+=100){
      const chunk=productNames.slice(pi,pi+100);
      const {error:delErr}=await db.from("stock_uploads").delete().in("product_name",chunk);
      if(delErr){setResult({type:"error",msg:"삭제 실패: "+delErr.message});setLoading(false);return;}
    }
    const rows=preview.map(r=>({...r,upload_date:endDate}));
    for(let i=0;i<rows.length;i+=500){
      const {error}=await db.from("stock_uploads").insert(rows.slice(i,i+500));
      if(error){setResult({type:"error",msg:"삽입 실패: "+error.message});setLoading(false);return;}
    }
    const ts2=nowStr();
    await db.from("upload_logs").insert({upload_type:"stock",file_name:fileName,row_count:preview.length,inserted:preview.length,deleted:existing?.length||0,date_start:startDate,date_end:endDate});
    setStep(3); setResult({type:"success",msg:`기존 ${existing?.length||0}건 삭제 → 새 ${preview.length}건 등록`,ts:ts2});
    onUpdate(ts2); setLoading(false);
  };
  const reset=()=>{setStep(0);setPreview(null);setExisting(null);setFileName("");setResult(null);};

  return (
    <div>
      <Steps current={step} steps={["기간 선택","파일 업로드","미리보기 확인","완료"]}/>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>입고 기간 선택</div>
            <CalendarPicker mode="range" rangeStart={startDate} rangeEnd={endDate}
              onRangeChange={({start,end})=>{setStartDate(start);setEndDate(end||start);}}/>
            <div style={{color:D.red,fontSize:10,marginBottom:20}}>⚠ 확정 시 해당 기간 DB 데이터 전체 교체</div>
            <Btn onClick={confirmDate} disabled={!dateValid||loading} style={{width:"100%"}}>
              {loading?"조회 중...":"기간 확정"}
            </Btn>
          </>}
          {step===1&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>파일 업로드</div>
            <StatRow items={[{label:"삭제 예정",value:`${existing?.length||0}건`,color:D.red}]}/>
            <DropZone onFile={handleFile} fileName={fileName} label="입고 파일 업로드"
              columns="상품명 · 옵션 · 수량 · 메모"/>
            <button onClick={()=>{setStep(0);setExisting(null);}}
              style={{width:"100%",background:"transparent",border:"none",color:D.textMeta,
                fontSize:11,cursor:"pointer",marginTop:8,padding:"5px"}}>← 기간 다시 선택</button>
          </>}
          {step===2&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>미리보기 확인</div>
            <StatRow items={[
              {label:"삭제 예정",value:`${existing?.length||0}건`,color:D.red},
              {label:"새 등록",value:`${preview?.length||0}건`,color:D.green},
            ]}/>
            <Btn onClick={handleUpload} disabled={loading} variant="danger" style={{width:"100%",marginBottom:7}}>
              {loading?"처리 중...":"확정 교체"}
            </Btn>
            <button onClick={()=>setStep(1)}
              style={{width:"100%",background:"transparent",border:"none",color:D.textMeta,
                fontSize:11,cursor:"pointer",padding:"5px"}}>← 파일 다시 선택</button>
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
          </>}
          {step===3&&<div style={{textAlign:"center"}}>
            <div style={{fontSize:36,marginBottom:8}}>✓</div>
            <div style={{color:D.green,fontWeight:600,marginBottom:10}}>교체 완료</div>
            {result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}
            <Btn onClick={reset} variant="ghost" style={{width:"100%",marginTop:12}}>새 업로드</Btn>
          </div>}
          {result?.type==="error"&&step!==2&&<Alert type="error" msg={result.msg}/>}
        </Card>
        <Card>
          {(step===1||step===2)?(
            <>
              <div style={{fontWeight:500,fontSize:12,marginBottom:20}}>
                {step<2?`기존 DB — ${startDate}~${endDate}`:`새 파일 — ${fileName}`}
              </div>
              {step===1&&(existing?.length?
                <PreviewTable rows={existing} cols={[
                  {key:"upload_date",label:"업로드일",color:D.textMeta},
                  {key:"product_name",label:"상품명",maxW:150},
                  {key:"option_name",label:"옵션",color:D.textMeta},
                  {key:"qty",label:"수량",bold:true},
                  {key:"memo",label:"메모",color:D.textMeta},
                ]}/>:
                <div style={{color:D.green,textAlign:"center",padding:60,fontSize:12}}>해당 기간 기존 데이터 없음</div>)}
              {step===2&&preview&&<PreviewTable rows={preview} cols={[
                {key:"product_name",label:"상품명",maxW:180},
                {key:"option_name",label:"옵션",color:D.textMeta},
                {key:"qty",label:"수량",bold:true},
                {key:"memo",label:"메모",color:D.textMeta},
              ]}/>}
            </>
          ):(
            <>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                  <span style={{fontWeight:600,fontSize:13}}>입고 내역</span>
                  <UpdatedAt ts={histTs}/>
                </div>
                <div style={{display:"flex",gap:6,alignItems:"center"}}>
                  {selected.size>0&&(
                    <button onClick={handleDeleteSelected}
                      style={{background:deleteConfirm?D.red:"transparent",color:deleteConfirm?"#fff":D.red,
                        border:`1px solid ${D.red}`,borderRadius:6,padding:"4px 12px",fontSize:11,cursor:"pointer"}}>
                      {deleteConfirm?`${selected.size}건 삭제 확인`:`${selected.size}건 삭제`}
                    </button>
                  )}
                  <Btn onClick={loadHistory} variant="ghost" style={{padding:"4px 11px",fontSize:11}}>불러오기</Btn>
                </div>
              </div>
              {history===null
                ?<div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>불러오기를 눌러주세요</div>
                :history.length===0
                  ?<div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>입고 데이터 없음</div>
                  :(()=>{
                    const filtered=histFilter
                      ?history.filter(r=>(r.product_name||"").includes(histFilter)||(r.option_name||"").includes(histFilter)||(r.upload_date||"").includes(histFilter))
                      :history;
                    const allSelected=filtered.length>0&&filtered.every(r=>selected.has(r.id));
                    return(
                      <div>
                        <div style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
                          <input type="text" placeholder="상품명 · 업로드일 검색" value={histFilter}
                            onChange={e=>{setHistFilter(e.target.value);setDeleteConfirm(false);}}
                            style={{flex:1,border:`1px solid ${D.border}`,borderRadius:6,padding:"5px 8px",
                              fontSize:12,background:"transparent",color:D.text}}/>
                          <span style={{fontSize:11,color:D.textMeta,whiteSpace:"nowrap"}}>{filtered.length}건</span>
                        </div>
                        <div style={{overflowY:"auto",maxHeight:520}}>
                          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                            <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
                              <th style={{padding:"5px 7px",textAlign:"center",width:28}}>
                                <input type="checkbox" checked={allSelected} onChange={()=>toggleAll(filtered)}/>
                              </th>
                              {["업로드일","상품명","옵션","수량","메모"].map(h=>(
                                <th key={h} style={{padding:"5px 7px",textAlign:"left",color:D.textMeta,fontWeight:400}}>{h}</th>
                              ))}
                            </tr></thead>
                            <tbody>
                              {filtered.map(r=>(
                                <tr key={r.id} style={{borderBottom:`1px solid ${D.border}`,
                                  background:selected.has(r.id)?"#f5f5f5":"transparent"}}>
                                  <td style={{padding:"4px 7px",textAlign:"center"}}>
                                    <input type="checkbox" checked={selected.has(r.id)} onChange={()=>{toggleSelect(r.id);setDeleteConfirm(false);}}/>
                                  </td>
                                  <td style={{padding:"5px 7px",color:D.textMeta}}>{r.upload_date}</td>
                                  <td style={{padding:"5px 7px",fontWeight:600,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.product_name}</td>
                                  <td style={{padding:"5px 7px",color:D.textSub}}>{r.option_name||"—"}</td>
                                  <td style={{padding:"5px 7px",fontWeight:600}}>{(r.qty||0).toLocaleString()}</td>
                                  <td style={{padding:"5px 7px",color:D.textMeta}}>{r.memo||""}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    );
                  })()
              }
            </>
          )}
        </Card>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// 공통 업로드 내역 패널 (Supabase 테이블 기반)
// ─────────────────────────────────────────────
const HIST_PAGE=200;
function DataHistoryPanel({
  table, dateField, searchFields, cols, editableCols=[], onChanged,
  placeholder="날짜·품목 검색", idField="id", refreshKey=0,
  // 보조 테이블 join (예: order_items + order_headers)
  joinTable=null, joinOn=null, joinFields=[],
  // 컬럼별 edit 대상 테이블/키 라우팅 (지정 안 된 컬럼은 기본 table+idField)
  editTableMap={}, editTableKey={},
}) {
  const [rows,setRows]=useState([]);
  const [loading,setLoading]=useState(true);
  const [filter,setFilter]=useState("");
  const [page,setPage]=useState(0);
  const [selected,setSelected]=useState(new Set());
  const [editCell,setEditCell]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [deleteConfirm,setDeleteConfirm]=useState(false);
  const [result,setResult]=useState(null);

  useEffect(()=>{
    (async()=>{
      setLoading(true);
      const db=await getSupabase();
      const PAGE=1000;
      const fetchAll=async(tbl,orderCol,asc)=>{
        let all=[];let from=0;
        while(true){
          const {data,error}=await db.from(tbl).select("*")
            .order(orderCol,{ascending:asc}).range(from,from+PAGE-1);
          if(error||!data||data.length===0) break;
          all=all.concat(data);
          if(data.length<PAGE) break;
          from+=PAGE;
        }
        return all;
      };
      // join이 있으면 main은 idField asc로 로딩 후 dateField로 client sort (dateField가 join쪽일 수 있어)
      const mainSortCol = joinTable ? idField : dateField;
      const mainAsc     = joinTable ? true : false;
      const mainAll=await fetchAll(table,mainSortCol,mainAsc);
      let merged=mainAll;
      if(joinTable&&joinOn){
        const jAll=await fetchAll(joinTable,joinOn,true);
        const jMap={};
        jAll.forEach(j=>{jMap[j[joinOn]]=j;});
        const fields=joinFields.length?joinFields:Object.keys(jAll[0]||{});
        merged=mainAll.map(r=>{
          const j=jMap[r[joinOn]]||{};
          const enriched={...r};
          fields.forEach(f=>{enriched[f]=j[f]!==undefined?j[f]:enriched[f];});
          return enriched;
        });
        // dateField 기준 client-side desc 정렬
        merged.sort((a,b)=>String(b[dateField]||"").localeCompare(String(a[dateField]||"")));
      }
      setRows(merged);
      setLoading(false);
    })();
  // refreshKey 변경 시에도 재로딩 (외부에서 삭제·업로드 발생 시)
  },[table,dateField,joinTable,joinOn,refreshKey]);

  const filtered=filter
    ?rows.filter(r=>[...searchFields,dateField].some(f=>String(r[f]||"").includes(filter)))
    :rows;
  const totalPages=Math.max(1,Math.ceil(filtered.length/HIST_PAGE));
  const safePage=Math.min(page,totalPages-1);
  const pageRows=filtered.slice(safePage*HIST_PAGE,(safePage+1)*HIST_PAGE);

  const toggleSelect=id=>setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>{
    const ids=pageRows.map(r=>r[idField]);
    const allSel=ids.every(id=>selected.has(id));
    setSelected(s=>{const n=new Set(s);ids.forEach(id=>allSel?n.delete(id):n.add(id));return n;});
  };

  const handleDelete=async()=>{
    const db=await getSupabase();
    const cnt=selected.size;
    const {error}=await db.from(table).delete().in(idField,[...selected]);
    if(error){setResult({type:"error",msg:error.message});return;}
    setRows(r=>r.filter(row=>!selected.has(row[idField])));
    setSelected(new Set()); setDeleteConfirm(false);
    setResult({type:"success",msg:`${cnt}건 삭제 완료`});
    onChanged?.();
  };

  const startEdit=(id,field,val)=>{setEditCell({id,field});setEditVal(String(val??""));};
  const saveEdit=async()=>{
    if(!editCell) return;
    const db=await getSupabase();
    // 컬럼별 edit 대상 테이블/키 라우팅
    const targetTable=editTableMap[editCell.field]||table;
    const targetKey  =targetTable===table?idField:(editTableKey[editCell.field]||joinOn);
    const row=rows.find(r=>r[idField]===editCell.id);
    const keyVal=row?row[targetKey]:editCell.id;
    const {error}=await db.from(targetTable).update({[editCell.field]:editVal}).eq(targetKey,keyVal);
    if(!error){
      // 같은 joinOn 값을 공유하는 모든 행에 반영 (header 편집 시)
      setRows(rs=>rs.map(r=>{
        const same=targetTable===table
          ?r[idField]===editCell.id
          :r[targetKey]===keyVal;
        return same?{...r,[editCell.field]:editVal}:r;
      }));
    }
    setEditCell(null);
  };

  const inp2={border:`1px solid ${D.border}`,borderRadius:6,padding:"5px 10px",fontSize:12,background:"transparent",color:D.text};

  return (
    <Card style={{marginTop:14}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:8,flexWrap:"wrap"}}>
        <span style={{fontWeight:600,fontSize:13}}>업로드 내역</span>
        <div style={{display:"flex",gap:8,alignItems:"center",flex:1,justifyContent:"flex-end"}}>
          <input placeholder={placeholder} value={filter} onChange={e=>{setFilter(e.target.value);setDeleteConfirm(false);setPage(0);}}
            style={{...inp2,minWidth:180,maxWidth:280}}/>
          <span style={{fontSize:11,color:D.textMeta,whiteSpace:"nowrap"}}>
            {loading?"로딩 중…":`${filtered.length}건`}
          </span>
        </div>
      </div>
      {selected.size>0&&(
        <div style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
          <span style={{fontSize:11,color:D.textMeta}}>{selected.size}개 선택</span>
          {!deleteConfirm
            ?<button onClick={()=>setDeleteConfirm(true)}
               style={{background:D.red,color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>삭제</button>
            :<>
              <button onClick={handleDelete}
                style={{background:D.red,color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>확인 삭제</button>
              <button onClick={()=>setDeleteConfirm(false)}
                style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>취소</button>
            </>}
        </div>
      )}
      {loading?(
        <div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>불러오는 중…</div>
      ):rows.length===0?(
        <div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>데이터 없음</div>
      ):(
        <div style={{overflowX:"auto",maxHeight:500,overflowY:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead style={{position:"sticky",top:0,background:D.surface,zIndex:1}}>
              <tr style={{borderBottom:`1px solid ${D.border}`}}>
                <th style={{padding:"5px 7px",width:28}}>
                  <input type="checkbox" checked={pageRows.length>0&&pageRows.every(r=>selected.has(r[idField]))} onChange={toggleAll}/>
                </th>
                {cols.map(c=>(
                  <th key={c.key} style={{padding:"5px 7px",textAlign:"left",color:D.textMeta,fontWeight:400,whiteSpace:"nowrap"}}>{c.label}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {pageRows.map(r=>(
                <tr key={r[idField]} style={{borderBottom:`1px solid ${D.border}`,background:selected.has(r[idField])?"#f5f5f5":"transparent"}}>
                  <td style={{padding:"4px 7px"}}><input type="checkbox" checked={selected.has(r[idField])} onChange={()=>toggleSelect(r[idField])}/></td>
                  {cols.map(c=>{
                    const isEditing=editCell?.id===r[idField]&&editCell?.field===c.key;
                    const editable=editableCols.includes(c.key);
                    return(
                      <td key={c.key} style={{padding:"4px 7px",color:c.color||D.black,fontWeight:c.bold?600:400,
                        maxWidth:c.maxW||200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                        cursor:editable?"pointer":"default",userSelect:editable?"none":"auto"}}
                        title={String(r[c.key]??"")+(editable?" (더블클릭하여 수정)":"")}
                        onDoubleClick={editable?()=>startEdit(r[idField],c.key,r[c.key]):undefined}>
                        {isEditing
                          ?<input autoFocus value={editVal} onChange={e=>setEditVal(e.target.value)}
                              onBlur={saveEdit} onKeyDown={e=>{if(e.key==="Enter")saveEdit();if(e.key==="Escape")setEditCell(null);}}
                              style={{width:"100%",border:`1px solid ${D.primary}`,borderRadius:3,padding:"1px 4px",fontSize:11}}/>
                          :c.fmt?c.fmt(r[c.key]):String(r[c.key]??"")}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {totalPages>1&&(
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8,marginTop:8,fontSize:11,color:D.textMeta}}>
          <button onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={safePage===0}
            style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 10px",cursor:safePage===0?"default":"pointer",color:safePage===0?D.textMeta:D.text}}>
            이전
          </button>
          <span>{safePage+1} / {totalPages}</span>
          <button onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))} disabled={safePage===totalPages-1}
            style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 10px",cursor:safePage===totalPages-1?"default":"pointer",color:safePage===totalPages-1?D.textMeta:D.text}}>
            다음
          </button>
        </div>
      )}
      {result&&<Alert type={result.type} msg={result.msg}/>}
    </Card>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT — 이지어드민 CSV (주문일 기준 · 배송일 선택)
// ─────────────────────────────────────────────
function EasyAdminUploader({ onUpdate, histRefreshKey=0 }) {
  const [startDate,setStartDate]=useState("");
  const [endDate,setEndDate]=useState("");
  const [step,setStep]=useState(0);
  const [fileName,setFileName]=useState("");
  const [parsedFile,setParsedFile]=useState(null);
  const [inRange,setInRange]=useState([]);
  const [outRows,setOutRows]=useState([]);
  const [dupInfo,setDupInfo]=useState(null);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);

  // Step 1: 파일 선택 (파싱만, 업로드 X)
  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    parseAnyFile(file,{header:true,skipEmptyLines:true,transformHeader:h=>h.trim()},({data})=>{
        try{
          if(!data.length) throw new Error("데이터가 없습니다");
          const f=detectFields(Object.keys(data[0]));

          // 컬럼 탐색 (Unicode 정규화 + exact match 우선)
          const allCols=Object.keys(data[0]);
          const nrm=s=>String(s).trim().normalize("NFC");
          const findCol=(...names)=>{
            // 1순위: exact match
            for(const n of names){ const c=allCols.find(h=>nrm(h)===nrm(n)); if(c) return c; }
            // 2순위: includes match
            for(const n of names){ const c=allCols.find(h=>nrm(h).includes(nrm(n))); if(c) return c; }
            return null;
          };
          const orderDateCol = findCol("주문일","주문일시","주문날짜","order_date","날짜","date") || f.date;
          const deliveryDateCol = findCol("배송일","배송일시","배송날짜","배송완료일","발송일","출고일","출고일시","출고완료일","배송(예정)일","예정배송일","delivery_date");
          const orderIdCol = findCol("주문번호","orderid") || findCol("관리번호","order_id") || f.orderId;
          const channelCol = findCol("판매처","channel","플랫폼","채널") || f.channel;
          const productCol = findCol("상품명","product","품명") || f.product;
          const optionCol  = findCol("옵션명","옵션","option") || f.option;
          const csCol          = findCol("CS","cs처리","cs상태","cs") || f.cs;
          const statusCol      = findCol("상태","status") || f.status;
          const qtyCol         = findCol("주문수량","수량","qty","quantity") || f.qty;
          // 판매가 (상품별, 29CM·무신사 AOV용 SUM)
          const salePriceCol   = findCol("판매가","상품금액","상품판매가","item_price");
          // 결제금액 (주문 단위, 자사몰 AOV용 — 합산 금지)
          const paymentAmtCol  = findCol("결제금액","주문금액","결제총액","실결제금액","payment_amount");
          // amount: 하위 호환용 (기존 차트 소스)
          const amtCol         = salePriceCol || paymentAmtCol
                               || findCol("판매금액","실판매가","금액","amount","price") || f.revenue;

          // 누락된 필수 컬럼을 한 번에 모아 안내 (한 개씩 throw 대신)
          const missingCols=[];
          if(!orderIdCol)      missingCols.push("주문번호");
          if(!orderDateCol)    missingCols.push("주문일");
          if(!deliveryDateCol) missingCols.push("배송일");
          if(missingCols.length){
            throw new Error(uploadErrColumns({
              missing:missingCols,
              required:["주문번호","주문일","배송일","판매처","상품명","옵션","수량","판매가","결제금액","CS처리"],
              headers:allCols,
            }));
          }

          // 주문일+주문번호+상품명+옵션 기준 중복 합산 (날짜별 주문 유니크화)
          const grouped={};
          data.filter(r=>{
            if(!r[orderIdCol]) return false;
            const chRaw=String(r[channelCol]||"").trim();
            if(chRaw==="MERRYONOVERSEA") return false;
            if(chRaw==="예약거래") return false; // 매장 CSV로 별도 집계
            return true;
          }).forEach(r=>{
            const oid=String(r[orderIdCol]).trim();
            const prod=String(r[productCol]||"").trim();
            const opt=String(r[optionCol]||"").trim();
            const ch=normChannel(r[channelCol]);
            const orderDateVal=toDate(r[orderDateCol]);
            const deliveryDateVal=toDate(r[deliveryDateCol]);
            const csRaw=csCol?String(r[csCol]||"").trim():"";
            const statusRaw=statusCol?String(r[statusCol]||"").trim():"";
            // 상태 추론 우선순위: CS → 상태 → 기본 "배송"
            //   normCS: 배송전+취소→"취소", 그 외 취소→"반품", 교환→"교환"
            const rawStatus=csRaw?normCS(csRaw):(statusRaw?normCS(statusRaw):"배송");
            let status=rawStatus;
            // 배송 전 취소/교환은 실제 배송 발생 여부에 따라 재분류
            //   - 배송일 있음: 결국 배송됨 → "배송" 으로 카운트
            //   - 배송일 없음 + 교환: "주문" (교환으로 잡지 않음)
            //   - 배송일 없음 + 취소: "취소" 유지
            const csLower=csRaw.toLowerCase().replace(/\s/g,"");
            const isPreShip=csLower.includes("배송전");
            if(isPreShip&&deliveryDateVal) status="배송";
            else if(isPreShip&&!deliveryDateVal&&rawStatus==="교환") status="주문";
            // 배송일 없으면 실제 배송 전 → "주문"
            if(status==="배송"&&!deliveryDateVal) status="주문";
            // CORD prefix = 29CM 취소 주문 → 반품으로 강제 (배송일 있는 경우만 fire)
            if(status==="배송"&&/^CORD/i.test(oid)) status="반품";
            const qty=toNum(r[qtyCol])||1;
            const salePriceVal  = salePriceCol  ? toNum(r[salePriceCol])  : 0;
            const paymentAmtVal = paymentAmtCol ? toNum(r[paymentAmtCol]) : 0;
            const amt = salePriceVal || paymentAmtVal || (amtCol?toNum(r[amtCol]):0);
            // 주문일+주문번호+상품명+옵션 조합을 DB key (날짜별+상품별 유니크)
            const dbKey=`${orderDateVal||""}||${oid}||${prod}||${opt}`;
            if(!grouped[dbKey]){
              grouped[dbKey]={order_id:dbKey,order_no:oid,order_date:orderDateVal,
                delivery_date:deliveryDateVal||null,channel:ch,
                product_name:prod,option_name:opt,
                qty:0,amount:0,sale_price:0,payment_amount:0,
                status,raw_status:csRaw||statusRaw};
            }
            grouped[dbKey].qty+=qty;
            grouped[dbKey].amount+=amt;                                        // 하위 호환
            grouped[dbKey].sale_price+=salePriceVal;                          // 상품별 합산 (29CM·무신사)
            if(paymentAmtVal) grouped[dbKey].payment_amount=paymentAmtVal;    // 주문 단위, 덮어쓰기 (합산 X)
            grouped[dbKey].status=status;
          });
          const parsed=Object.values(grouped);
          setParsedFile(parsed);
          // 주문일 min/max 자동 감지
          const dates=parsed.map(r=>r.order_date).filter(Boolean).sort();
          const autoStart=dates[0]||"";
          const autoEnd=dates[dates.length-1]||"";
          setStartDate(autoStart);
          setEndDate(autoEnd);
          const validRows=parsed.filter(r=>r.order_date);
          const noDateRows=parsed.filter(r=>!r.order_date);
          setInRange(validRows);
          setOutRows(noDateRows);
          setDupInfo({total:validRows.length,newCount:validRows.length,updateCount:0,sameCount:0});
          setResult({type:"info",msg:`주문일 ${autoStart} ~ ${autoEnd} · ${validRows.length}건 파싱 완료`+(noDateRows.length>0?` (주문일 없는 ${noDateRows.length}건 제외)`:"")+` | 주문번호: "${orderIdCol}" · 배송일: "${deliveryDateCol}"`});
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e.message}));
  },[]);

  // Step 0→1: 미리보기 (파싱 완료 후 확인) — 기존 헤더(주문 단위) 개수 조회
  const [existingCount,setExistingCount]=useState(0);
  const handlePreview=async()=>{
    if(!parsedFile?.length) {setResult({type:"error",msg:"파일을 먼저 선택해주세요"});return;}
    setLoading(true);
    try{
      const db=await getSupabase();
      const {count}=await db.from("order_headers").select("*",{count:"exact",head:true})
        .gte("order_date",startDate).lte("order_date",endDate);
      setExistingCount(count||0);
    }catch{}
    setLoading(false);
    setStep(1);
  };

  // Step 1→2: 확정 업로드 — order_headers + order_items 두 테이블에 분리 적재
  //   1) 기간 내 헤더 삭제 → CASCADE로 items 동반 삭제
  //   2) 기간 외 동일 order_no 충돌 헤더 삭제 → CASCADE
  //   3) headers insert → 4) items insert (FK 충족)
  const handleUpload=async()=>{
    if(!inRange.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    const {count:prevCount}=await db.from("order_headers").select("*",{count:"exact",head:true})
      .gte("order_date",startDate).lte("order_date",endDate);

    // 헤더/아이템 분리 — 스키마:
    //   order_headers: order_no, order_date, channel, payment_amount
    //   order_items  : order_no, product_name, option_name, qty, sale_price, status, delivery_date, raw_status
    const headersMap={};
    const itemsList=[];
    inRange.forEach(r=>{
      if(!headersMap[r.order_no]){
        headersMap[r.order_no]={
          order_no:r.order_no,
          order_date:r.order_date,
          channel:r.channel,
          payment_amount:r.payment_amount||0,
        };
      } else if((r.payment_amount||0)>0){
        // 동일 order_no 다중 행: payment_amount 양수 우선
        headersMap[r.order_no].payment_amount=r.payment_amount;
      }
      itemsList.push({
        order_no:r.order_no,
        product_name:r.product_name||"미분류",
        option_name:r.option_name||"",
        qty:r.qty||0,
        sale_price:r.sale_price||0,
        status:r.status,
        delivery_date:r.delivery_date||null,
        raw_status:r.raw_status||null,
      });
    });
    const newOrderNos=Object.keys(headersMap);

    // 1) 기간 내 헤더 삭제
    const {error:delErr1}=await db.from("order_headers").delete()
      .gte("order_date",startDate).lte("order_date",endDate);
    if(delErr1){setResult({type:"error",msg:"기간 삭제 실패: "+delErr1.message});setLoading(false);return;}
    // 2) 기간 밖 충돌(같은 order_no, 다른 날짜) 헤더 삭제
    if(newOrderNos.length){
      for(let i=0;i<newOrderNos.length;i+=500){
        const {error:delErr2}=await db.from("order_headers").delete()
          .in("order_no",newOrderNos.slice(i,i+500));
        if(delErr2){setResult({type:"error",msg:"충돌 삭제 실패: "+delErr2.message});setLoading(false);return;}
      }
    }
    // 3) headers insert
    const headersList=Object.values(headersMap);
    for(let i=0;i<headersList.length;i+=500){
      const {error}=await db.from("order_headers").insert(headersList.slice(i,i+500));
      if(error){setResult({type:"error",msg:"헤더 삽입 실패: "+error.message});setLoading(false);return;}
    }
    // 4) items insert (FK는 headers가 먼저 있어야 충족)
    for(let i=0;i<itemsList.length;i+=500){
      const {error}=await db.from("order_items").insert(itemsList.slice(i,i+500));
      if(error){setResult({type:"error",msg:"아이템 삽입 실패: "+error.message});setLoading(false);return;}
    }
    const ts2=nowStr();
    await db.from("upload_logs").insert({
      upload_type:"orders",file_name:fileName,
      row_count:parsedFile?.length||0,
      inserted:itemsList.length,updated:0,
      skipped:outRows.length,date_start:startDate,date_end:endDate,
    });
    setStep(2);
    const replaceNote=prevCount>0?` (기존 주문 ${prevCount}건 대체됨)`:"";
    setResult({type:"success",msg:`${itemsList.length}건 등록 완료 (${startDate} ~ ${endDate})${replaceNote}`,ts:ts2});
    onUpdate(ts2); setLoading(false);
  };

  const reset=()=>{setStep(0);setInRange([]);setOutRows([]);setDupInfo(null);setFileName("");setParsedFile(null);setResult(null);setStartDate("");setEndDate("");};

  return (
    <div>
      <Steps current={step} steps={["파일 선택","미리보기 확인","완료"]}/>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>파일 선택</div>
            <DropZone onFile={handleFile} fileName={fileName} label="주문·배송 파일 선택"
              columns="주문번호 · 주문일 · 배송일 · 판매처 · 상품명 · 옵션 · 수량 · 판매가 · 결제금액 · CS처리"/>
            {result&&<Alert type={result.type} msg={result.msg}/>}
            {startDate&&<div style={{color:D.blue,fontSize:11,marginTop:8,lineHeight:1.7}}>
              감지된 주문일: <b>{startDate}</b> ~ <b>{endDate}</b>
            </div>}
            <div style={{marginTop:12}}>
              <Btn onClick={handlePreview} disabled={!parsedFile||loading} style={{width:"100%"}}>
                {loading?"파싱 중...":"미리보기"}
              </Btn>
            </div>
          </>}
          {step===1&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>미리보기 확인</div>
            {dupInfo&&<StatRow items={[
              {label:"업로드 대상",value:`${dupInfo.total}건`,color:D.green},
              {label:"주문일 없음(제외)",value:`${outRows.length}건`,color:D.textMeta},
            ]}/>}
            <div style={{color:D.textMeta,fontSize:11,marginBottom:8}}>주문일 {startDate} ~ {endDate}</div>
            {existingCount>0&&<Alert type="warn" msg={uploadReplaceWarn(existingCount,`${startDate}~${endDate} 주문일`)}/>}
            {outRows.length>0&&<Alert type="warn" msg={`주문일 없는 ${outRows.length}건 제외`}/>}
            <div style={{display:"flex",flexDirection:"column",gap:7,marginTop:12}}>
              <Btn onClick={handleUpload} disabled={loading||!inRange.length} style={{width:"100%"}}>
                {loading?"처리 중...":`확정 업로드 (${dupInfo?.total||0}건)`}
              </Btn>
              <button onClick={()=>setStep(0)}
                style={{background:"transparent",border:"none",color:D.textMeta,
                  fontSize:11,cursor:"pointer",padding:"5px"}}>← 파일 다시 선택</button>
            </div>
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
          </>}
          {step===2&&<div style={{textAlign:"center"}}>
            <div style={{fontSize:36,marginBottom:8}}>✓</div>
            <div style={{color:D.green,fontWeight:600,marginBottom:10}}>업로드 완료</div>
            {result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}
            <Btn onClick={reset} variant="ghost" style={{width:"100%",marginTop:12}}>새 업로드</Btn>
          </div>}
        </Card>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
            <span style={{fontWeight:500,fontSize:12}}>파일 미리보기</span>
            {inRange.length>0&&<div style={{display:"flex",gap:7}}>
              <span style={{background:`${D.green}12`,color:D.green,fontSize:10,padding:"2px 9px",borderRadius:20}}>기간 내 {inRange.length}건</span>
              {outRows.length>0&&<span style={{background:`${D.red}12`,color:D.red,fontSize:10,padding:"2px 9px",borderRadius:20}}>기간 외 {outRows.length}건</span>}
            </div>}
          </div>
          {inRange.length>0||outRows.length>0?(
            <PreviewTable
              rows={[...inRange,...outRows]}
              outIdx={new Set(inRange.map((_,i)=>-1).concat(outRows.map((_,i)=>inRange.length+i)).filter(i=>i>=inRange.length))}
              cols={[
                {key:"order_no",label:"주문번호",color:D.textMeta,maxW:100},
                {key:"order_date",label:"주문일",color:D.textMeta},
                {key:"delivery_date",label:"배송일",color:D.textMeta},
                {key:"channel",label:"판매처",bold:true},
                {key:"product_name",label:"상품명",maxW:140},
                {key:"option_name",label:"옵션",color:D.textMeta},
                {key:"qty",label:"수량",bold:true},
                {key:"status",label:"상태",fmt:v=>(
                  <span style={{color:v==="반품"?D.red:v==="교환"?D.amber:(v==="주문"||v==="취소")?D.textMeta:D.green,fontWeight:500}}>{v}</span>
                )},
              ]}
            />
          ):<div style={{color:D.textMeta,textAlign:"center",padding:80,fontSize:12}}>
            기간 선택 후 CSV 파일을 선택하고 미리보기를 누르세요
          </div>}
        </Card>
      </div>
      <DataHistoryPanel
        table="order_items" dateField="order_date" idField="item_id"
        joinTable="order_headers" joinOn="order_no"
        joinFields={["order_date","channel","payment_amount"]}
        editTableMap={{channel:"order_headers"}}
        editTableKey={{channel:"order_no"}}
        refreshKey={histRefreshKey}
        searchFields={["product_name","channel","order_no","option_name"]}
        placeholder="날짜·상품명·판매처 검색"
        editableCols={["channel","status","product_name","option_name"]}
        cols={[
          {key:"order_no",label:"주문번호",color:D.textMeta,maxW:110},
          {key:"order_date",label:"주문일",color:D.textMeta},
          {key:"delivery_date",label:"배송일",color:D.textMeta},
          {key:"channel",label:"판매처",bold:true},
          {key:"product_name",label:"상품명",maxW:180},
          {key:"option_name",label:"옵션",color:D.textMeta},
          {key:"qty",label:"수량"},
          {key:"status",label:"상태",fmt:v=><span style={{color:v==="반품"?D.red:v==="교환"?D.amber:(v==="주문"||v==="취소")?D.textMeta:D.green,fontWeight:500}}>{v}</span>},
        ]}
        onChanged={()=>onUpdate(nowStr())}
      />
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT — 매장 판매 CSV
// ─────────────────────────────────────────────
function StoreUploader({ onUpdate, histRefreshKey=0 }) {
  const [step,setStep]=useState(0);
  const [fileName,setFileName]=useState("");
  const [preview,setPreview]=useState(null);
  const [dateRange,setDateRange]=useState({start:"",end:""});
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [conflictCount,setConflictCount]=useState(0);

  // 파일 파싱 완료 후 기존 데이터 수 조회
  useEffect(()=>{
    if(!preview||!dateRange.start||!dateRange.end){setConflictCount(0);return;}
    (async()=>{
      const db=await getSupabase();
      const{count}=await db.from("store_sales").select("*",{count:"exact",head:true})
        .gte("sale_date",dateRange.start).lte("sale_date",dateRange.end);
      setConflictCount(count||0);
    })();
  },[preview,dateRange.start,dateRange.end]);

  const parseKRW=s=>{
    const str=String(s||"").trim().replace(/[\s,]/g,"");
    if(!str||str==="0"||str==="NAN") return 0;
    const neg=str.startsWith("(")&&str.endsWith(")");
    const num=parseInt(str.replace(/[()]/g,""),10)||0;
    return neg?-num:num;
  };

  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    parseAnyFile(file,{header:true,skipEmptyLines:true},({data})=>{
        try{
          if(!data?.length) throw new Error(uploadErrParse("파일에 데이터 행이 없습니다"));
          const headers=Object.keys(data[0]||{});
          // 필요 컬럼 존재 여부 사전 검증 (StoreUploader는 정확 매칭 사용)
          const REQ=["구매일자","상품명","수량","실판매금액"];
          const missing=REQ.filter(c=>!headers.includes(c));
          if(missing.length){
            throw new Error(uploadErrColumns({
              missing,
              required:["구매일자","매장","상품명","옵션","수량","실판매금액","ID"],
              headers,
            }));
          }
          const rows=data.map(r=>{
            const qty=parseKRW(r["수량"]);
            const amount=parseKRW(r["실판매금액"]);
            // 음수 qty 또는 음수 amount(괄호·-) 둘 다 반품 신호
            const isReturn=qty<0||amount<0;
            return{
              sale_date:(r["구매일자"]||"").trim().slice(0,10),
              store_name:(r["매장"]||"").trim(),
              product_name:(r["상품명"]||"").trim(),
              option_name:(r["옵션"]||"").trim(),
              qty:Math.abs(qty),
              amount:Math.abs(amount),
              order_id:(r["ID"]||"").trim(),
              // DB값은 기존 호환을 위해 '배송' 유지 (분석/필터 코드 다수에서 사용)
              // 매장 UI 표시는 '판매'로 변환 (StoreUploader 화면 / DataHistoryPanel fmt)
              status:isReturn?"반품":"배송",
            };
          }).filter(r=>r.sale_date&&r.product_name&&r.qty>0&&r.amount>0);
          if(!rows.length) throw new Error("파싱된 행이 0건입니다. '구매일자', '상품명', '수량'(>0), '실판매금액'(>0) 모두 값이 있는 행이 1개 이상 있어야 합니다.");
          const dates=[...new Set(rows.map(r=>r.sale_date))].sort();
          setDateRange({start:dates[0]||"",end:dates[dates.length-1]||""});
          setPreview(rows); setStep(1);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e?.message||String(e)}));
  },[]);

  const handleUpload=async()=>{
    if(!preview?.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    if(dateRange.start&&dateRange.end){
      await db.from("store_sales").delete().gte("sale_date",dateRange.start).lte("sale_date",dateRange.end);
    }
    for(let i=0;i<preview.length;i+=500){
      const{error}=await db.from("store_sales").insert(preview.slice(i,i+500));
      if(error){setResult({type:"error",msg:error.message});setLoading(false);return;}
    }
    const ts2=nowStr();
    const replaceNote=conflictCount>0?` (기존 ${conflictCount}건 대체됨)`:"";
    setStep(2);setResult({type:"success",msg:`${preview.length}건 등록 완료${replaceNote}`,ts:ts2});
    onUpdate(ts2);setLoading(false);
  };

  const reset=()=>{setStep(0);setPreview(null);setFileName("");setResult(null);setConflictCount(0);};

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<>
            <div style={{fontWeight:600,marginBottom:10,fontSize:13}}>매장 판매 CSV 업로드</div>
            <div style={{fontSize:11,color:D.textMeta,marginBottom:16,lineHeight:1.7}}>
              POS 시스템 판매 데이터를 업로드합니다.<br/>
              인식 컬럼: <b>구매일자 · 매장 · 상품명 · 옵션 · 수량 · 실판매금액 · ID</b><br/>
              실판매금액=0인 행 자동 제외
            </div>
            <DropZone onFile={handleFile} fileName={fileName} label="매장 판매 파일 업로드"
              columns="구매일자 · 매장 · 상품명 · 옵션 · 수량 · 실판매금액 · ID"/>
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
          </>}
          {step===1&&<>
            <div style={{fontWeight:600,marginBottom:10,fontSize:13}}>업로드 확인</div>
            <StatRow items={[
              {label:"파일",value:fileName.slice(0,20)},
              {label:"기간",value:`${dateRange.start}~${dateRange.end}`},
              {label:"행수",value:`${preview?.length||0}건`,color:D.green},
            ]}/>
            {conflictCount>0&&(
              <div style={{background:"#fff9e6",border:`1px solid ${D.amber}`,borderRadius:7,
                padding:"10px 12px",marginBottom:10,fontSize:12,lineHeight:1.6}}>
                <div style={{fontWeight:700,color:D.amber,marginBottom:4}}>
                  ⚠ 기존 데이터 {conflictCount.toLocaleString()}건과 겹칩니다
                </div>
                <div style={{color:D.textSub}}>
                  {dateRange.start} ~ {dateRange.end} 기간의 기존 데이터가 모두 삭제되고 새 데이터로 교체됩니다.
                </div>
              </div>
            )}
            <Btn onClick={handleUpload} disabled={loading} style={{width:"100%",marginBottom:7}}>
              {loading?"처리 중...":conflictCount>0?"덮어쓰기 업로드":"업로드"}
            </Btn>
            <button onClick={reset} style={{width:"100%",background:"transparent",border:"none",color:D.textMeta,fontSize:11,cursor:"pointer",padding:"5px"}}>← 다시 선택</button>
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
          </>}
          {step===2&&<div style={{textAlign:"center"}}>
            <div style={{fontSize:36,marginBottom:8}}>✓</div>
            <div style={{color:D.green,fontWeight:600,marginBottom:10}}>업로드 완료</div>
            {result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}
            <Btn onClick={reset} variant="ghost" style={{width:"100%",marginTop:12}}>새 업로드</Btn>
          </div>}
        </Card>
        <Card>
          <div style={{fontWeight:500,fontSize:12,marginBottom:16}}>
            {step===0?"파일을 업로드하면 미리보기가 표시됩니다":fileName+" 파싱 결과"}
          </div>
          {step===0&&<div style={{color:D.textMeta,textAlign:"center",padding:60,fontSize:12}}>업로드 대기 중</div>}
          {step>=1&&preview&&(
            <>
              <div style={{display:"flex",gap:20,marginBottom:12,fontSize:11,color:D.textMeta}}>
                <span>총 <b style={{color:D.black}}>{preview.length}</b>행</span>
                <span>판매 <b style={{color:D.green}}>{preview.filter(r=>r.status==="배송").length}</b>건</span>
                <span>반품 <b style={{color:D.red}}>{preview.filter(r=>r.status==="반품").length}</b>건</span>
                <span>매출 <b style={{color:D.black}}>{fmtWon(preview.filter(r=>r.status==="배송").reduce((s,r)=>s+(r.amount||0),0))}</b></span>
                <span>주문번호 <b style={{color:D.black}}>{new Set(preview.filter(r=>r.status==="배송").map(r=>r.order_id)).size}</b>개</span>
              </div>
              <PreviewTable rows={preview.slice(0,50)} cols={[
                {key:"sale_date",label:"날짜",color:D.textMeta},
                {key:"store_name",label:"매장",color:D.textMeta},
                {key:"product_name",label:"상품명",maxW:180},
                {key:"option_name",label:"옵션",color:D.textMeta},
                {key:"qty",label:"수량",bold:true},
                {key:"amount",label:"금액"},
                // 매장 UI에서는 '배송' → '판매' 로 표시 (DB값은 유지)
                {key:"status",label:"상태",color:D.textMeta,fmt:v=>v==="배송"?"판매":v},
                {key:"order_id",label:"주문ID",color:D.textMeta},
              ]}/>
            </>
          )}
        </Card>
      </div>
      <DataHistoryPanel
        table="store_sales" dateField="sale_date"
        refreshKey={histRefreshKey}
        searchFields={["product_name","store_name"]}
        placeholder="날짜·상품명·매장 검색"
        editableCols={["store_name","product_name","option_name","amount","status"]}
        cols={[
          {key:"sale_date",label:"날짜",color:D.textMeta},
          {key:"store_name",label:"매장",bold:true},
          {key:"product_name",label:"상품명",maxW:180},
          {key:"option_name",label:"옵션",color:D.textMeta},
          {key:"qty",label:"수량"},
          {key:"amount",label:"금액"},
          {key:"status",label:"상태",fmt:v=><span style={{color:v==="반품"?D.red:D.green,fontWeight:500}}>{v==="배송"?"판매":v}</span>},
          {key:"order_id",label:"주문ID",color:D.textMeta,maxW:100},
        ]}
        onChanged={()=>onUpdate(nowStr())}
      />
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT (탭 컨테이너)
// ─────────────────────────────────────────────
// 가이드 영상 URL (YouTube embed: "https://www.youtube.com/embed/VIDEO_ID?autoplay=1")
const GUIDE_VIDEOS={revenue:"",stock:"",orders:"",store:"",cs:""};

function GuideVideoModal({show,onClose,videoUrl}){
  if(!show) return null;
  return(
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"#000a",zIndex:1000,
      display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div onClick={e=>e.stopPropagation()}
        style={{background:"#111",borderRadius:12,padding:20,maxWidth:760,width:"90%",
          boxShadow:"0 8px 40px #000c",position:"relative"}}>
        <button onClick={onClose}
          style={{position:"absolute",top:10,right:14,background:"none",border:"none",
            color:"#aaa",fontSize:22,cursor:"pointer",lineHeight:1}}>✕</button>
        {videoUrl
          ?<iframe src={videoUrl} width="100%" height="400" frameBorder="0"
              allow="autoplay; fullscreen" allowFullScreen
              style={{borderRadius:8,display:"block"}}/>
          :<div style={{height:200,display:"flex",alignItems:"center",justifyContent:"center",
              color:"#555",fontSize:13}}>영상 준비 중입니다</div>
        }
      </div>
    </div>
  );
}

function GuideSection({tabKey,desc,isDark}){
  const [open,setOpen]=useState(false);
  const bdrColor=isDark?"#fff":"#111";
  return(
    <div style={{marginBottom:20,padding:"14px 18px",
      border:`1.5px solid ${bdrColor}`,borderRadius:8,
      background:isDark?"#0e0e0e":"#fafafa"}}>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12}}>
        <div style={{fontSize:12,color:isDark?"#ccc":D.textSub,lineHeight:1.8,flex:1,whiteSpace:"pre-line"}}>
          {desc}
        </div>
        <button onClick={()=>setOpen(true)}
          style={{flexShrink:0,background:isDark?"#fff":"#111",color:isDark?"#111":"#fff",
            border:"none",borderRadius:6,padding:"5px 14px",fontSize:12,
            cursor:"pointer",fontWeight:700,letterSpacing:"0.04em"}}>
          Guide
        </button>
      </div>
      <GuideVideoModal show={open} onClose={()=>setOpen(false)} videoUrl={GUIDE_VIDEOS[tabKey]}/>
    </div>
  );
}

function DataInput({ onUpdate, onDataChange, orders=[], stocks=[], revenues=[], storeSales=[] }) {
  const [tab,setTab]=useState("revenue");
  // 업로드 내역 패널들에 강제 새로고침 신호 (삭제/업로드 발생 시 증가)
  const [histRefreshKey,setHistRefreshKey]=useState(0);
  const bumpHist=()=>setHistRefreshKey(k=>k+1);

  const lastDate=(arr,field)=>{
    const d=arr.map(r=>r[field]).filter(Boolean).sort().at(-1);
    return d?<span style={{fontSize:10,color:D.textMeta,fontWeight:400,marginLeft:4}}>({d})</span>:null;
  };

  const tabs=[
    {key:"revenue",name:"매출 입력",extra:lastDate(revenues,"date")},
    {key:"stock",name:"입고",extra:lastDate(stocks,"upload_date")},
    {key:"orders",name:"주문·배송",extra:lastDate(orders,"order_date")},
    {key:"store",name:"매장 판매",extra:lastDate(storeSales,"sale_date")},
    {key:"cs",name:"CS"},
    {key:"delete",name:"데이터 삭제"},
  ];

  const GUIDES={
    revenue:"KPI 카드의 매출, 매출 점유율, 판매처별 매출의 소스입니다.\n매출 금액은 취소/환불이 포함된 금액이며, 엑셀 다운로드 시 각 채널 어드민의 통계에서 확인하세요.\n*매일 전날의 데이터를 업로드하세요.",
    stock:"KPI 카드의 입고 수량, 물류 플로우 섹션 전체의 데이터 소스입니다.\n*매일 전날의 데이터를 업로드하세요.",
    orders:"KPI 카드의 배송·반품 수, 판매처 상세의 배송·반품 수, 판매·반품 TOP, 플랫폼 별 선호·반품 옵션 랭킹, 객단가 계산의 데이터 소스입니다.\n필요 컬럼: 주문번호 · 주문일 · 배송일 · 판매처 · 상품명 · 옵션 · 수량 · 판매가(29CM·무신사 AOV) · 결제금액(자사몰 AOV) · CS처리\n*매일 최근 한달 데이터(주문건 반품 정보 업데이트)를 업로드하세요.",
    store:"KPI 카드의 매출(오프라인 스토어) 합산, 랭크 지표 내 오프라인 스토어 항목의 데이터 소스입니다.\n*매일 최근 한달의 데이터를 업로드하세요.",
    cs:"반품 랭크 상품의 주요 반품 사유 데이터 소스로 매칭됩니다.\n*매일 전날 데이터를 업로드하세요.",
  };

  return (
    <div style={{padding:"20px 24px",maxWidth:1400,margin:"0 auto"}}>
      {/* 탭 바 — 섹션명에 테두리 박스 */}
      <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:20}}>
        {tabs.map(t=>(
          <button key={t.key} onClick={()=>setTab(t.key)}
            style={{background:tab===t.key?D.black:"transparent",
              color:tab===t.key?"#fff":D.textSub,
              border:`1.5px solid ${tab===t.key?D.black:D.borderMid}`,
              borderRadius:6,padding:"6px 14px",fontWeight:tab===t.key?700:400,
              fontSize:12,cursor:"pointer",transition:"all 0.12s",
              display:"flex",alignItems:"center",gap:4}}>
            {t.name}{t.extra}
          </button>
        ))}
      </div>

      {/* 가이드 섹션 (삭제 탭 제외) */}
      {tab!=="delete"&&GUIDES[tab]&&(
        <GuideSection tabKey={tab} desc={GUIDES[tab]} isDark={false}/>
      )}

      {tab==="revenue"&&<RevenueForm onUpdate={ts=>onUpdate("revenue",ts)} histRefreshKey={histRefreshKey}/>}
      {tab==="stock"&&<StockUploader onUpdate={ts=>onUpdate("stock",ts)} histRefreshKey={histRefreshKey}/>}
      {tab==="orders"&&<EasyAdminUploader onUpdate={ts=>{onUpdate("orders",ts);onDataChange?.();bumpHist();}} histRefreshKey={histRefreshKey}/>}
      {tab==="store"&&<StoreUploader onUpdate={ts=>{onUpdate("store",ts);onDataChange?.();bumpHist();}} histRefreshKey={histRefreshKey}/>}
      {tab==="cs"&&<CSDataInput/>}
      {tab==="delete"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14}}>
          <DataDeleteSection table="revenues" dateField="date" label="매출 입력" onDone={()=>{onDataChange?.();bumpHist();}}/>
          <DataDeleteSection table="stock_uploads" dateField="upload_date" label="입고" onDone={()=>{onDataChange?.();bumpHist();}}/>
          <DataDeleteSection table="order_headers" dateField="order_date" label="주문·배송" onDone={()=>{onDataChange?.();bumpHist();}}/>
          <DataDeleteSection table="store_sales" dateField="sale_date" label="매장 판매" onDone={()=>{onDataChange?.();bumpHist();}}/>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// LOADING SCREEN
// ─────────────────────────────────────────────
function LoadingScreen() {
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",
      justifyContent:"center",minHeight:"100vh",background:D.bg}}>
      <div style={{marginBottom:24}}>
        <div style={{fontWeight:800,fontSize:22,letterSpacing:"0.12em",color:D.black,textAlign:"center"}}>MERRYON</div>
        <div style={{fontSize:10,color:D.textMeta,letterSpacing:"0.06em",textAlign:"center",marginTop:2}}>COMMERCE ANALYTICS</div>
      </div>
      <div style={{width:36,height:36,border:`2px solid ${D.border}`,
        borderTop:`2px solid ${D.black}`,borderRadius:"50%",
        animation:"mry-spin 0.9s linear infinite"}}/>
      <style>{`@keyframes mry-spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

// ─────────────────────────────────────────────
// INVENTORY TREND — CONSTANTS & UTILS
// ─────────────────────────────────────────────
const INV_AGING_DEFS={
  HEALTHY:{ label:"Healthy",     color:"#7EC8A4", desc:"마지막 판매 30일 이내" },
  SLOW:   { label:"Slow-moving", color:"#7B9EC8", desc:"31~90일" },
  AGING:  { label:"Aging",       color:"#C8A87B", desc:"91~180일" },
  DEAD:   { label:"Dead Stock",  color:"#C87B7B", desc:"180일 초과" },
};
const INV_AGING_KEYS=["HEALTHY","SLOW","AGING","DEAD"];

function getAgingKey(noSalesDays){
  const d=Math.max(0,noSalesDays||0);
  if(d<=30) return "HEALTHY";
  if(d<=90) return "SLOW";
  if(d<=180) return "AGING";
  return "DEAD";
}

function calcInvRow(row){
  const snap=row.snapshot_date;
  const noSalesDays=row.last_delivery_date
    ?Math.max(0,dayjs(snap).diff(dayjs(row.last_delivery_date),"day"))
    :Math.max(0,dayjs(snap).diff(dayjs(row.first_inbound_date),"day"));
  const skuAge=Math.max(0,dayjs(snap).diff(dayjs(row.first_inbound_date),"day"));
  const postRestockDays=row.latest_inbound_date?Math.max(0,dayjs(snap).diff(dayjs(row.latest_inbound_date),"day")):0;
  const sellThroughProxy=Math.round((row.cumulative_delivery_qty/(row.current_stock_qty+1))*100)/100;
  const currentInventoryValue=(row.current_stock_qty||0)*(row.selling_price||0);
  const agingKey=getAgingKey(noSalesDays);
  return{...row,noSalesDays,skuAge,postRestockDays,sellThroughProxy,currentInventoryValue,agingKey};
}

const INV_COL_ALIASES={
  product_code:           ["상품코드","product_code","품번","바코드","barcode","sku코드","sku_code"],
  product_name:           ["상품명","product_name"],
  option_name:            ["옵션","option_name","옵션명"],
  selling_price:          ["판매가","selling_price","가격","price"],
  supply_price:           ["공급가","supply_price","원가","cost","cost_price","sup_price"],
  current_stock_qty:      ["현재고","current_stock_qty","재고","현재재고"],
  first_inbound_date:     ["처음입고일","first_inbound_date","최초입고일"],
  first_inbound_qty:      ["처음입고수량","first_inbound_qty","최초입고수량"],
  cumulative_inbound_qty: ["누적입고","cumulative_inbound_qty","누적입고수량"],
  latest_inbound_date:    ["마지막입고일","latest_inbound_date","최근입고일"],
  latest_inbound_qty:     ["마지막입고수량","latest_inbound_qty","최근입고수량"],
  last_delivery_date:     ["마지막배송일","last_delivery_date","최근배송일","최근출고일"],
  cumulative_delivery_qty:["누적배송수량","cumulative_delivery_qty","누적배송","누적출고"],
  snapshot_date:          ["데이터날짜","snapshot_date","날짜"],
  // Reorder-specific columns (optional — extracted separately, not saved to inventory_snapshot)
  _r_avail:               ["가용재고","available_stock"],
  _r_incoming:            ["입고대기","incoming_stock"],
  _r_weekly:              ["1주발주합계","weekly_sales","1주판매합계"],
  _r_monthly:             ["4주발주합계","monthly_sales","4주판매합계"],
};

// 헤더 → 필드 매핑: 1차로 정확 일치, 2차로 부분 포함 매칭.
// 정확 일치 우선이라 "상품코드" 헤더는 product_name("상품명") 으로 잘못 빨려 들어가지 않음.
function mapInvCols(headers){
  const result={};
  const claimedHeaders=new Set();
  const norm=s=>String(s||"").trim().toLowerCase().replace(/[\s_]/g,"");
  const normHeaders=headers.map(norm);
  // Pass 1: exact match
  Object.entries(INV_COL_ALIASES).forEach(([field,aliases])=>{
    if(result[field]!==undefined) return;
    const aliasNorms=aliases.map(norm);
    for(let i=0;i<normHeaders.length;i++){
      if(claimedHeaders.has(i)) continue;
      if(aliasNorms.includes(normHeaders[i])){
        result[field]=i; claimedHeaders.add(i); break;
      }
    }
  });
  // Pass 2: partial (includes) match — only on headers not yet claimed
  headers.forEach((h,i)=>{
    if(claimedHeaders.has(i)) return;
    const n=normHeaders[i];
    Object.entries(INV_COL_ALIASES).forEach(([field,aliases])=>{
      if(!result[field]&&aliases.some(a=>n.includes(norm(a)))){
        result[field]=i; claimedHeaders.add(i);
      }
    });
  });
  return result;
}

function parseInvFile(file,onResult,onError){
  const reader=new FileReader();
  reader.onload=async e=>{
    try{
      const XLSX=await getXLSX();
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,dateNF:"YYYY-MM-DD"});
      if(!raw||raw.length<2){onError(uploadErrParse("파일에 데이터 행이 없습니다 (헤더 행만 있거나 비어있음)"));return;}
      const headers=raw[0].map(h=>String(h||"").trim());
      const colMap=mapInvCols(headers);
      const FIELD_LABELS={product_name:"상품명",current_stock_qty:"현재고",snapshot_date:"데이터날짜",first_inbound_date:"처음입고일"};
      const required=["product_name","current_stock_qty","snapshot_date","first_inbound_date"];
      const missing=required.filter(f=>colMap[f]===undefined).map(f=>FIELD_LABELS[f]);
      if(missing.length>0){
        onError(uploadErrColumns({
          missing,
          required:Object.values(FIELD_LABELS).concat(["상품코드","옵션","판매가","공급가","처음입고수량","누적입고","마지막입고일","마지막입고수량","마지막배송일","누적배송수량"]),
          headers,
        }));
        return;
      }
      const hasReorder=colMap["_r_weekly"]!==undefined||colMap["_r_avail"]!==undefined;
      const rows=raw.slice(1).map(r=>{
        const get=(f)=>colMap[f]!==undefined?String(r[colMap[f]]||"").trim():"";
        // Remove commas then parse as float→round, so "1,234.5" → 1235, not "12345"
        const getNum=(f)=>Math.round(parseFloat(String(r[colMap[f]]||"0").replace(/,/g,""))||0);
        const getDate=(f)=>{const v=get(f);if(!v||v==="-") return null;return toDate(v)||null;};
        return{
          snapshot_date:getDate("snapshot_date"),
          product_code:get("product_code")||"",
          product_name:get("product_name"),
          option_name:get("option_name")||"",
          selling_price:parseInt(get("selling_price").replace(/[^0-9]/g,""),10)||0,
          supply_price:parseInt(get("supply_price").replace(/[^0-9]/g,""),10)||0,
          current_stock_qty:getNum("current_stock_qty"),
          first_inbound_date:getDate("first_inbound_date"),
          first_inbound_qty:getNum("first_inbound_qty"),
          cumulative_inbound_qty:getNum("cumulative_inbound_qty"),
          latest_inbound_date:getDate("latest_inbound_date"),
          latest_inbound_qty:getNum("latest_inbound_qty"),
          last_delivery_date:getDate("last_delivery_date"),
          cumulative_delivery_qty:getNum("cumulative_delivery_qty"),
          // Optional reorder fields (stripped before inventory_snapshot insert)
          ...(hasReorder?{
            _r_avail:getNum("_r_avail"),
            _r_incoming:getNum("_r_incoming"),
            _r_weekly:getNum("_r_weekly"),
            _r_monthly:getNum("_r_monthly"),
          }:{}),
        };
      }).filter(r=>r.product_name&&r.snapshot_date&&r.first_inbound_date);
      onResult(rows);
    }catch(err){onError(uploadErrParse(String(err?.message||err)));}
  };
  reader.onerror=()=>onError(uploadErrParse("파일 읽기 도중 시스템 오류가 발생했습니다"));
  reader.readAsArrayBuffer(file);
}

// ─────────────────────────────────────────────
// INVENTORY UPLOADER
// ─────────────────────────────────────────────
function InventoryUploader({DC,onUploaded,onReorderDone}){
  const [file,setFile]=useState(null);
  const [uploadStatus,setUploadStatus]=useState(null);
  const [statusMsg,setStatusMsg]=useState("");
  const [parsedRows,setParsedRows]=useState([]);
  const [snapDate,setSnapDate]=useState(null);
  const [history,setHistory]=useState([]);
  const [histLoading,setHistLoading]=useState(false);
  const [conflictInfo,setConflictInfo]=useState(null);
  const [showModal,setShowModal]=useState(false);
  const [histFilter,setHistFilter]=useState("");
  const [selDates,setSelDates]=useState(new Set());
  const [delConfirm,setDelConfirm]=useState(false);
  const [dragOver,setDragOver]=useState(false);

  const loadHistory=useCallback(async()=>{
    setHistLoading(true);
    const db=await getSupabase();
    let data=[];let from=0;const PAGE=1000;
    while(true){
      const{data:chunk,error}=await db.from("inventory_snapshot")
        .select("snapshot_date,created_at")
        .order("snapshot_date",{ascending:false})
        .range(from,from+PAGE-1);
      if(error||!chunk||chunk.length===0) break;
      data=data.concat(chunk);
      if(chunk.length<PAGE) break;
      from+=PAGE;
    }
    if(!data.length){setHistLoading(false);return;}
    const map={};
    data.forEach(r=>{
      if(!map[r.snapshot_date]) map[r.snapshot_date]={snapshot_date:r.snapshot_date,row_count:0,uploaded_at:r.created_at};
      map[r.snapshot_date].row_count++;
      if(r.created_at>map[r.snapshot_date].uploaded_at) map[r.snapshot_date].uploaded_at=r.created_at;
    });
    setHistory(Object.values(map).sort((a,b)=>b.snapshot_date.localeCompare(a.snapshot_date)));
    setHistLoading(false);
  },[]);

  useEffect(()=>{loadHistory();},[loadHistory]);

  const handleFile=useCallback(f=>{
    if(!f) return;
    setFile(f);setUploadStatus("parsing");setStatusMsg("파일 파싱 중...");setParsedRows([]);setSnapDate(null);
    parseInvFile(f,parsed=>{
      setUploadStatus(null);setStatusMsg("");
      if(!parsed.length){setUploadStatus("error");setStatusMsg("유효한 데이터 행이 없습니다");return;}
      const dates=[...new Set(parsed.map(r=>r.snapshot_date).filter(Boolean))].sort();
      setSnapDate(dates[dates.length-1]||null);
      setParsedRows(parsed);
    },err=>{setUploadStatus("error");setStatusMsg(err);});
  },[]);

  const doUpload=useCallback(async(replace=false)=>{
    if(!parsedRows.length||!snapDate) return;
    setUploadStatus("uploading");setStatusMsg("저장 중...");
    try{
      const db=await getSupabase();
      if(replace){
        const{error:de}=await db.from("inventory_snapshot").delete().eq("snapshot_date",snapDate);
        if(de) throw new Error(de.message);
      }
      // Strip reorder-specific fields before inserting to inventory_snapshot
      const invRows=parsedRows.map(({_r_avail,_r_incoming,_r_weekly,_r_monthly,...rest})=>rest);
      const CHUNK=500;
      let insertedAny=false;
      for(let i=0;i<invRows.length;i+=CHUNK){
        const{error}=await db.from("inventory_snapshot").insert(invRows.slice(i,i+CHUNK));
        if(error){
          // Clean up any rows already inserted for this snapDate to avoid partial data
          if(insertedAny) await db.from("inventory_snapshot").delete().eq("snapshot_date",snapDate);
          throw new Error(error.message);
        }
        insertedAny=true;
      }
      setUploadStatus("done");setStatusMsg(`${parsedRows.length.toLocaleString()}개 행 저장 완료`);
      setFile(null);setParsedRows([]);setSnapDate(null);setConflictInfo(null);setShowModal(false);
      await loadHistory();
      if(onUploaded) onUploaded();
      // Post-process: reorder calculation (fire-and-forget, never blocks upload)
      if(onReorderDone&&parsedRows.some(r=>r._r_weekly!=null)){
        computeAndSaveReorder(parsedRows,snapDate).then(()=>onReorderDone()).catch(()=>{});}
      else if(!parsedRows.some(r=>r._r_weekly!=null)&&onReorderDone) onReorderDone();
    }catch(err){setUploadStatus("error");setStatusMsg(String(err));}
  },[parsedRows,snapDate,loadHistory,onUploaded,onReorderDone]);

  const handleUploadClick=useCallback(async()=>{
    if(!parsedRows.length||!snapDate) return;
    const db=await getSupabase();
    const{count}=await db.from("inventory_snapshot").select("snapshot_date",{count:"exact",head:true}).eq("snapshot_date",snapDate);
    if(count&&count>0){
      setConflictInfo({date:snapDate,existingCount:count,newCount:parsedRows.length});
      setShowModal(true);
    } else {
      doUpload(false);
    }
  },[parsedRows,snapDate,doUpload]);

  const handleDelete=useCallback(async()=>{
    const db=await getSupabase();
    for(const d of selDates){
      await db.from("inventory_snapshot").delete().eq("snapshot_date",d);
    }
    setSelDates(new Set());setDelConfirm(false);
    await loadHistory();
    if(onUploaded) onUploaded();
  },[selDates,loadHistory,onUploaded]);

  const filteredHist=histFilter?history.filter(h=>h.snapshot_date.includes(histFilter)):history;
  const stClr={parsing:"#7B9EC8",validating:"#C8A87B",uploading:"#7EC8A4",done:"#7EC8A4",error:"#C87B7B"};

  return(
    <div>
      {/* Drop zone */}
      <div
        onDragOver={e=>{e.preventDefault();setDragOver(true);}}
        onDragLeave={()=>setDragOver(false)}
        onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)handleFile(f);}}
        onClick={()=>{const inp=document.createElement("input");inp.type="file";inp.accept=".xlsx,.xls,.csv";inp.onchange=ev=>handleFile(ev.target.files[0]);inp.click();}}
        style={{border:`1.5px dashed ${dragOver?"#7EC8A4":DC.border}`,borderRadius:10,padding:"22px 20px",textAlign:"center",
          background:dragOver?"rgba(126,200,164,0.06)":DC.card,cursor:"pointer",transition:"all .15s"}}
      >
        <div style={{fontSize:22,opacity:.45,marginBottom:6}}>⬆</div>
        <div style={{fontSize:13,fontWeight:600,color:DC.text,marginBottom:10}}>Excel / CSV 드래그 &amp; 드롭</div>
        <div style={{fontSize:12,lineHeight:1.9,textAlign:"left",display:"inline-block",width:"100%"}}>
          <div style={{marginBottom:6}}>
            <span style={{color:"#7EC8A4",fontWeight:700,fontSize:13}}>인벤토리 트렌드</span>
            <div style={{display:"flex",flexWrap:"wrap",gap:"2px 8px",marginTop:3}}>
              {["상품명","상품코드","옵션","판매가","공급가","현재고","처음입고일","처음입고수량","누적입고","마지막입고일","마지막입고수량","마지막배송일","누적배송수량","데이터날짜"].map(c=>(
                <span key={c} style={{background:"rgba(126,200,164,0.1)",border:"1px solid rgba(126,200,164,0.25)",
                  borderRadius:4,padding:"1px 6px",fontSize:12,color:"#7EC8A4",fontFamily:"monospace"}}>{c}</span>
              ))}
            </div>
          </div>
          <div style={{marginBottom:8}}>
            <span style={{color:"#7B9EC8",fontWeight:700,fontSize:13}}>리오더 계산기</span>
            <div style={{display:"flex",flexWrap:"wrap",gap:"2px 8px",marginTop:3}}>
              {["가용재고","입고대기","1주발주합계","4주발주합계"].map(c=>(
                <span key={c} style={{background:"rgba(123,158,200,0.1)",border:"1px solid rgba(123,158,200,0.25)",
                  borderRadius:4,padding:"1px 6px",fontSize:12,color:"#7B9EC8",fontFamily:"monospace"}}>{c}</span>
              ))}
            </div>
          </div>
          <div style={{color:DC.dim,fontSize:12}}>인벤토리 트렌드 / 리오더 계산기의 공통 데이터 소스가 됩니다.</div>
        </div>
        {file&&<div style={{marginTop:6,fontSize:13,color:"#7EC8A4"}}>{file.name}</div>}
      </div>

      {/* Status bar */}
      {uploadStatus&&(
        <div style={{marginTop:8,padding:"7px 12px",borderRadius:6,background:"rgba(255,255,255,0.04)",
          border:`1px solid ${DC.border}`,display:"flex",alignItems:"center",gap:8}}>
          {uploadStatus!=="done"&&uploadStatus!=="error"&&(
            <span style={{display:"inline-block",width:10,height:10,borderRadius:"50%",
              border:`2px solid ${stClr[uploadStatus]||"#888"}`,borderTopColor:"transparent",
              animation:"invSpin 0.7s linear infinite",flexShrink:0}}/>
          )}
          <span style={{fontSize:12,color:stClr[uploadStatus]||DC.text}}>{statusMsg}</span>
        </div>
      )}

      {/* Upload action */}
      {parsedRows.length>0&&uploadStatus!=="uploading"&&(
        <div style={{marginTop:10,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:DC.sub}}>
            데이터날짜: <span style={{color:DC.text,fontWeight:600}}>{snapDate}</span>
            {` — ${parsedRows.length.toLocaleString()}개 SKU`}
          </span>
          <button onClick={handleUploadClick}
            style={{background:"#7EC8A4",color:"#0a1a12",border:"none",borderRadius:6,
              padding:"6px 18px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
            업로드
          </button>
        </div>
      )}

      {/* Conflict modal */}
      {showModal&&conflictInfo&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#1a1a1a",border:"1px solid #333",borderRadius:12,padding:28,maxWidth:360,width:"90%"}}>
            <div style={{fontWeight:700,fontSize:15,color:"#F0F0F0",marginBottom:10}}>업로드 충돌 감지</div>
            <div style={{fontSize:13,color:"#888",lineHeight:1.9,marginBottom:20}}>
              <span style={{color:"#F0F0F0",fontWeight:600}}>{conflictInfo.date}</span> 날짜에<br/>
              기존 데이터 <span style={{color:"#F0F0F0"}}>{conflictInfo.existingCount.toLocaleString()}</span>행이 존재합니다.<br/>
              신규 <span style={{color:"#F0F0F0"}}>{conflictInfo.newCount.toLocaleString()}</span>행으로 교체하시겠습니까?
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button onClick={()=>setShowModal(false)}
                style={{background:"transparent",color:"#888",border:"1px solid #333",borderRadius:6,padding:"6px 16px",fontSize:12,cursor:"pointer"}}>취소</button>
              <button onClick={()=>doUpload(true)}
                style={{background:"#C87B7B",color:"#fff",border:"none",borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer"}}>교체</button>
            </div>
          </div>
        </div>
      )}

      {/* History */}
      <div style={{marginTop:20}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:8}}>
          <span style={{fontSize:13,fontWeight:600,color:DC.sub,letterSpacing:".04em"}}>업로드 이력</span>
          <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
            <input type="date" value={histFilter} onChange={e=>setHistFilter(e.target.value)}
              style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,padding:"3px 8px",
                fontSize:12,color:DC.text,colorScheme:"dark",fontFamily:"inherit"}}/>
            {histFilter&&<button onClick={()=>setHistFilter("")} style={{background:"none",border:"none",color:DC.sub,cursor:"pointer",fontSize:13,lineHeight:1}}>✕</button>}
            {selDates.size>0&&(
              delConfirm
                ?<><button onClick={handleDelete} style={{background:"#C87B7B",color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:12,cursor:"pointer",fontWeight:700}}>확인 삭제</button>
                   <button onClick={()=>setDelConfirm(false)} style={{background:"transparent",color:DC.sub,border:`1px solid ${DC.border}`,borderRadius:5,padding:"3px 10px",fontSize:12,cursor:"pointer"}}>취소</button></>
                :<button onClick={()=>setDelConfirm(true)} style={{background:"transparent",color:"#C87B7B",border:"1px solid #C87B7B",borderRadius:5,padding:"3px 10px",fontSize:12,cursor:"pointer"}}>
                  {selDates.size}개 삭제
                </button>
            )}
          </div>
        </div>
        {histLoading
          ?<div style={{color:DC.dim,fontSize:12,padding:"10px 0"}}>로딩 중...</div>
          :filteredHist.length===0
            ?<div style={{color:DC.dim,fontSize:12,padding:"10px 0"}}>이력 없음</div>
            :<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{color:DC.sub,borderBottom:`1px solid ${DC.border}`}}>
                  <th style={{padding:"5px 4px",textAlign:"left",fontWeight:500,width:22}}/>
                  <th style={{padding:"5px 8px",textAlign:"left",fontWeight:500}}>데이터 날짜</th>
                  <th style={{padding:"5px 8px",textAlign:"right",fontWeight:500}}>행 수</th>
                  <th style={{padding:"5px 8px",textAlign:"right",fontWeight:500}}>업로드 일시</th>
                </tr>
              </thead>
              <tbody>
                {filteredHist.map(h=>(
                  <tr key={h.snapshot_date} style={{borderBottom:`1px solid ${DC.border}`,color:selDates.has(h.snapshot_date)?DC.text:DC.sub}}>
                    <td style={{padding:"5px 4px"}}>
                      <input type="checkbox" checked={selDates.has(h.snapshot_date)}
                        onChange={ev=>{const s=new Set(selDates);ev.target.checked?s.add(h.snapshot_date):s.delete(h.snapshot_date);setSelDates(s);}}
                        style={{accentColor:"#C87B7B",cursor:"pointer"}}/>
                    </td>
                    <td style={{padding:"5px 8px",color:DC.text,fontWeight:500}}>{h.snapshot_date}</td>
                    <td style={{padding:"5px 8px",textAlign:"right"}}>{h.row_count.toLocaleString()}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",fontSize:12,color:DC.sub}}>
                      {h.uploaded_at?new Date(h.uploaded_at).toLocaleString("ko-KR",{month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit"}):"—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
        }
      </div>
      <style>{`@keyframes invSpin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

// ─────────────────────────────────────────────
// INV BUBBLE SCATTER PLOT
// ─────────────────────────────────────────────
function InvBubblePlot({DC,snapshotDates,stopRef}){
  const [dateMode,setDateMode]=useState("single"); // "single"|"range"
  const [selDate,setSelDate]=useState(null);
  const [selDateEnd,setSelDateEnd]=useState(null);
  const [data,setData]=useState([]);
  const [loading,setLoading]=useState(false);
  const [search,setSearch]=useState("");
  const [agingFilter,setAgingFilter]=useState(new Set(INV_AGING_KEYS));
  const [minStock,setMinStock]=useState(1);
  const [selectedSku,setSelectedSku]=useState(null);
  const [showSaleRec,setShowSaleRec]=useState(false);
  const [promoPage,setPromoPage]=useState(0);
  const [promoTableVisible,setPromoTableVisible]=useState(false);
  const promoTableRef=useRef(null);
  const [nextSecVisible,setNextSecVisible]=useState(false);
  useEffect(()=>{
    if(!stopRef?.current) return;
    const obs=new IntersectionObserver(([e])=>setNextSecVisible(e.isIntersecting),{threshold:0});
    obs.observe(stopRef.current);
    return()=>obs.disconnect();
  },[stopRef]);
  const showPromoSticky=useStickyTable("promoTable",showSaleRec&&promoTableVisible)&&!nextSecVisible;

  useEffect(()=>{ setSelDate(null); setSelDateEnd(null); setData([]); },[dateMode]);

  // Auto-select latest available date on first load
  useEffect(()=>{
    if(snapshotDates&&snapshotDates.length>0&&!selDate){
      const latest=[...snapshotDates].sort().pop();
      setSelDate(latest);
    }
  },[snapshotDates]);

  const loadDate=dateMode==="range"?(selDateEnd||selDate):selDate;
  useEffect(()=>{
    if(!loadDate){setData([]);return;}
    setLoading(true);
    (async()=>{
      const db=await getSupabase();
      let all=[];let from=0;const PAGE=1000;
      while(true){
        const{data:rows,error}=await db.from("inventory_snapshot").select("*")
          .eq("snapshot_date",loadDate).range(from,from+PAGE-1);
        if(error||!rows||rows.length===0) break;
        all=all.concat(rows);
        if(rows.length<PAGE) break;
        from+=PAGE;
      }
      setData(all.map(calcInvRow));
      setLoading(false);
    })();
  },[loadDate]);

  const filtered=useMemo(()=>
    data.filter(d=>
      d.current_stock_qty>=minStock&&
      agingFilter.has(d.agingKey)&&
      (!search||
        (d.product_name||"").toLowerCase().includes(search.toLowerCase())||
        (d.product_code||"").toLowerCase().includes(search.toLowerCase())||
        (d.option_name||"").toLowerCase().includes(search.toLowerCase()))
    )
  ,[data,agingFilter,minStock,search]);

  const {medX,medY}=useMemo(()=>{
    if(!filtered.length) return{medX:0,medY:0};
    const xs=[...filtered].map(d=>d.noSalesDays).sort((a,b)=>a-b);
    const ys=[...filtered].map(d=>d.current_stock_qty).sort((a,b)=>a-b);
    const mid=Math.floor(xs.length/2);
    return{medX:xs[mid]||0,medY:ys[mid]||0};
  },[filtered]);

  // In-range: right-upper quadrant (high unsold days AND high stock), with inbound constraint → 30~70%
  const saleRecs=useMemo(()=>{
    if(!loadDate||!filtered.length) return[];
    const snapM=dayjs(loadDate).month();
    const inRange=m=>{const diff=Math.abs(m-snapM);return diff<=3||diff>=9;};
    const raw=filtered
      .filter(d=>d.noSalesDays>medX&&d.current_stock_qty>medY)
      .filter(d=>d.latest_inbound_date&&inRange(dayjs(d.latest_inbound_date).month()));
    if(!raw.length) return[];
    const sortedX=[...raw].sort((a,b)=>a.noSalesDays-b.noSalesDays);
    const sortedY=[...raw].sort((a,b)=>a.current_stock_qty-b.current_stock_qty);
    const p30idx=Math.floor(raw.length*0.3);
    const p30X=Math.max(1,sortedX[p30idx]?.noSalesDays||medX);
    const p30Y=Math.max(1,sortedY[p30idx]?.current_stock_qty||medY);
    const candidates=raw
      .map(d=>{
        const xN=(d.noSalesDays-p30X)/p30X;
        const yN=(d.current_stock_qty-p30Y)/p30Y;
        return{...d,_dist:Math.sqrt(xN*xN+yN*yN)};
      })
      .sort((a,b)=>b._dist-a._dist);
    return candidates.map((d,i)=>{
      const n=Math.max(1,candidates.length-1);
      const rate=Math.max(30,Math.min(70,Math.round((70-(i/n)*40)/10)*10));
      return{...d,recommendedDiscount:rate,inZone:true};
    });
  },[filtered,medX,medY,loadDate]);

  // All SKUs with discount rate: in-zone 30~70%, out-of-zone 10~20%
  const allDiscountRecs=useMemo(()=>{
    if(!filtered.length) return[];
    const inZoneIds=new Set(saleRecs.map(d=>d.id));
    const outZone=filtered
      .filter(d=>!inZoneIds.has(d.id))
      .map(d=>({
        ...d,
        recommendedDiscount:(d.noSalesDays/Math.max(1,medX))<0.5?10:20,
        inZone:false,
      }));
    return[...saleRecs,...outZone].sort((a,b)=>b.recommendedDiscount-a.recommendedDiscount||b.noSalesDays-a.noSalesDays);
  },[filtered,saleRecs,medX]);

  const saleRecIds=useMemo(()=>new Set(saleRecs.map(d=>d.id)),[saleRecs]);

  useEffect(()=>{
    const el=promoTableRef.current;
    if(!el||!showSaleRec){setPromoTableVisible(false);return;}
    const obs=new IntersectionObserver(([e])=>setPromoTableVisible(e.isIntersecting),{threshold:0.05});
    obs.observe(el);
    return()=>obs.disconnect();
  },[showSaleRec]);

  useEffect(()=>{setPromoPage(0);},[showSaleRec]);

  const agingQtyStat=useMemo(()=>{
    const map={};INV_AGING_KEYS.forEach(k=>{map[k]={qty:0,val:0};});
    filtered.forEach(d=>{map[d.agingKey].qty+=(d.current_stock_qty||0);map[d.agingKey].val+=(d.currentInventoryValue||0);});
    const totalQty=INV_AGING_KEYS.reduce((s,k)=>s+map[k].qty,0);
    const totalVal=INV_AGING_KEYS.reduce((s,k)=>s+map[k].val,0);
    INV_AGING_KEYS.forEach(k=>{
      map[k].pct=totalQty?(map[k].qty/totalQty*100).toFixed(1):"0.0";
      map[k].valPct=totalVal?(map[k].val/totalVal*100).toFixed(1):"0.0";
    });
    return{map,totalQty,totalVal};
  },[filtered]);

  const maxZ=useMemo(()=>Math.max(...filtered.map(d=>d.currentInventoryValue),1),[filtered]);
  const minZ=useMemo(()=>Math.min(...filtered.filter(d=>d.currentInventoryValue>0).map(d=>d.currentInventoryValue),0),[filtered]);
  const getR=z=>{const n=maxZ===minZ?0.5:Math.max(0,(z-minZ)/(maxZ-minZ));return Math.max(5,Math.min(38,5+n*33));};

  const CustomDot=({cx,cy,payload})=>{
    if(cx==null||cy==null||!payload) return null;
    const r=getR(payload.currentInventoryValue);
    const col=INV_AGING_DEFS[payload.agingKey]?.color||"#888";
    const isSR=showSaleRec&&saleRecIds.has(payload.id);
    return(
      <circle cx={cx} cy={cy} r={r} fill={col} fillOpacity={0.62}
        stroke={isSR?"#fff":"none"} strokeWidth={isSR?2:0}
        style={{cursor:"pointer"}}
        onClick={()=>setSelectedSku(payload)}/>
    );
  };

  const BubbleTooltip=({active,payload})=>{
    if(!active||!payload?.length) return null;
    const d=payload[0]?.payload;
    if(!d) return null;
    const def=INV_AGING_DEFS[d.agingKey];
    return(
      <div style={{background:"#161616",border:"1px solid #2e2e2e",borderRadius:9,padding:"12px 14px",fontSize:12,minWidth:210,maxWidth:270,boxShadow:"0 4px 20px rgba(0,0,0,0.6)"}}>
        <div style={{fontWeight:700,color:"#F0F0F0",marginBottom:3,fontSize:14}}>{d.product_name}</div>
        {d.product_code&&<div style={{color:"#888",fontSize:11,fontFamily:"monospace",marginBottom:3}}>{d.product_code}</div>}
        {d.option_name&&<div style={{color:"#666",fontSize:12,marginBottom:8}}>{d.option_name}</div>}
        <div style={{display:"grid",gridTemplateColumns:"auto 1fr",gap:"3px 10px",color:"#888"}}>
          {[
            ["판매가",`${(d.selling_price||0).toLocaleString()}원`],
            ["현재고",`${(d.current_stock_qty||0).toLocaleString()}개`],
            ["현재 재고 금액",`${(d.currentInventoryValue||0).toLocaleString()}원`],
            ["미판매 일수",`${d.noSalesDays}일`],
            ["SKU 운영기간",`${d.skuAge}일`],
            ["최근입고 후",`${d.postRestockDays}일`],
            ["누적배송수량",`${(d.cumulative_delivery_qty||0).toLocaleString()}개`],
            ["판매효율 (STP)",`${d.sellThroughProxy}`],
          ].map(([l,v])=>(
            <><span key={`l${l}`}>{l}</span><span key={`v${l}`} style={{color:"#F0F0F0",textAlign:"right",fontWeight:500}}>{v}</span></>
          ))}
        </div>
        <div style={{marginTop:8,paddingTop:8,borderTop:"1px solid #2a2a2a",display:"flex",alignItems:"center",gap:5}}>
          <span style={{width:7,height:7,borderRadius:2,background:def?.color,display:"inline-block",flexShrink:0}}/>
          <span style={{color:def?.color,fontWeight:700,fontSize:12}}>{def?.label}</span>
        </div>
      </div>
    );
  };

  const downloadSaleRecs=async()=>{
    if(!allDiscountRecs.length) return;
    const XLSX=await getXLSX();
    const wb=XLSX.utils.book_new();
    // 사용자가 엑셀에서 「권장할인율(%)」 셀을 수정하면 「세일 후 가격」·「세일 후 원가율」 이 자동
    // 재계산되도록 수식 셀로 저장. (json_to_sheet 의 plain value 가 아니라 cell.f 사용)
    const header=[
      "상품코드","상품명","옵션","판매가","공급가","현재고","재고금액",
      "미판매일수","SKU기간","판매효율","Aging상태","제안범위",
      "권장할인율(%)","세일 후 가격","세일 후 원가율(%)",
    ];
    const aoa=[header];
    allDiscountRecs.forEach(d=>{
      aoa.push([
        d.product_code||"", d.product_name||"", d.option_name||"",
        d.selling_price||0, d.supply_price||0,
        d.current_stock_qty||0, d.currentInventoryValue||0,
        d.noSalesDays||0, d.skuAge||0,
        d.sellThroughProxy||0, INV_AGING_DEFS[d.agingKey]?.label||"",
        d.inZone?"범위 내":"범위 외",
        d.recommendedDiscount||0,
        null, // 세일 후 가격 → formula 로 채움
        null, // 원가율 → formula 로 채움
      ]);
    });
    const ws=XLSX.utils.aoa_to_sheet(aoa);
    // 각 데이터 행에 수식 주입 (1-based, 헤더가 1행)
    for(let i=0;i<allDiscountRecs.length;i++){
      const r=i+2; // Excel row number (2부터)
      const priceAddr=`D${r}`;        // 판매가
      const costAddr =`E${r}`;        // 공급가
      const discAddr =`M${r}`;        // 권장할인율(%)
      const saleAddr =`N${r}`;        // 세일 후 가격
      const rateAddr =`O${r}`;        // 세일 후 원가율(%)
      // 세일 후 가격 = 판매가 * (1 - 할인율/100)
      ws[saleAddr]={t:"n",f:`${priceAddr}*(1-${discAddr}/100)`,z:"#,##0"};
      // 세일 후 원가율 = 공급가 / 세일 후 가격 * 100  (분모 0 보호)
      ws[rateAddr]={t:"n",f:`IFERROR(${costAddr}/${saleAddr}*100,0)`,z:"0.0"};
    }
    // 컬럼 폭 보정
    ws["!cols"]=[
      {wch:12},{wch:24},{wch:14},{wch:10},{wch:10},{wch:8},{wch:12},
      {wch:9},{wch:8},{wch:8},{wch:9},{wch:9},
      {wch:13},{wch:13},{wch:16},
    ];
    XLSX.utils.book_append_sheet(wb,ws,"세일추천");
    XLSX.writeFile(wb,`sale_rec_${loadDate||"unknown"}.xlsx`);
  };

  const dcTheme={bg:"transparent",surface:"rgba(255,255,255,0.03)",border:DC.border,
    text:DC.text,sub:DC.sub,dim:DC.dim,green:"#7EC8A4",greenBg:"rgba(126,200,164,0.15)"};
  const availSet=useMemo(()=>new Set(snapshotDates),[snapshotDates]);

  // X-axis right edge = longest unsold days; Y-axis keeps headroom for bubble radius
  const xMax=useMemo(()=>filtered.length?Math.max(...filtered.map(d=>d.noSalesDays),1):100,[filtered]);
  const yMax=useMemo(()=>filtered.length?Math.max(...filtered.map(d=>d.current_stock_qty),1):100,[filtered]);
  const xDomain=useMemo(()=>[0,xMax],[xMax]);
  const yDomain=useMemo(()=>[0,Math.ceil(yMax*1.18)],[yMax]);

  const [showPromoInfo,setShowPromoInfo]=useState(false);

  return(
    <div style={{display:"flex",gap:16,alignItems:"flex-start"}}>
      {/* Left: Calendar + mode toggle + promo button */}
      <div style={{flexShrink:0}}>
        <div style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${DC.border}`,borderRadius:10,padding:14}}>
          {/* Single / Range toggle */}
          <div style={{display:"flex",gap:4,marginBottom:10}}>
            {[["single","단일"],["range","기간"]].map(([k,l])=>(
              <button key={k} data-hf onClick={()=>setDateMode(k)}
                style={{flex:1,background:dateMode===k?"#7EC8A4":"rgba(255,255,255,0.05)",
                  color:dateMode===k?"#0a1a12":DC.sub,
                  border:`1px solid ${dateMode===k?"#7EC8A4":DC.border}`,
                  borderRadius:6,padding:"5px 0",fontSize:12,cursor:"pointer",fontWeight:dateMode===k?700:400}}>
                {l}
              </button>
            ))}
          </div>
          <CalendarPicker
            mode={dateMode}
            value={selDate}
            onChange={setSelDate}
            rangeStart={selDate}
            rangeEnd={selDateEnd}
            onRangeChange={({start,end})=>{setSelDate(start);setSelDateEnd(end||"");}}
            availableDates={availSet}
            DC={dcTheme}
          />
          {/* 에이징 스테이터스 필터 */}
          <div style={{marginTop:10,paddingTop:10,borderTop:`1px solid ${DC.border}`}}>
            <div style={{fontSize:11,color:DC.sub,fontWeight:600,marginBottom:6}}>에이징 스테이터스</div>
            <div style={{display:"flex",flexDirection:"column",gap:4}}>
              {INV_AGING_KEYS.map(k=>{
                const def=INV_AGING_DEFS[k];const on=agingFilter.has(k);
                const stat=agingQtyStat.map[k];
                return(
                  <button key={k} data-hf onClick={()=>{const s=new Set(agingFilter);on?s.delete(k):s.add(k);setAgingFilter(s);}}
                    style={{background:on?`${def.color}22`:"transparent",color:on?def.color:DC.dim,
                      border:`1px solid ${on?def.color:DC.border}`,borderRadius:5,padding:"5px 8px",
                      fontSize:12,cursor:"pointer",textAlign:"left"}}>
                    <div style={{fontWeight:600,marginBottom:2}}>{def.label}</div>
                    <div style={{fontSize:10,opacity:.55,marginBottom:stat.qty>0?3:0}}>{def.desc}</div>
                    {stat.qty>0&&(
                      <div style={{display:"flex",flexDirection:"column",gap:1,fontSize:10,opacity:.75}}>
                        <div style={{display:"flex",justifyContent:"space-between"}}>
                          <span>수량 {stat.qty.toLocaleString()}개</span><span style={{fontWeight:600}}>{stat.pct}%</span>
                        </div>
                        <div style={{display:"flex",justifyContent:"space-between"}}>
                          <span>금액 {fmtWonShort(stat.val)}</span><span style={{fontWeight:600}}>{stat.valPct}%</span>
                        </div>
                      </div>
                    )}
                  </button>
                );
              })}
            </div>
            {agingQtyStat.totalQty>0&&(
              <div style={{marginTop:6,padding:"5px 6px",background:"rgba(255,255,255,0.03)",borderRadius:5,
                display:"flex",flexDirection:"column",gap:2,fontSize:10,color:DC.sub}}>
                <div style={{display:"flex",justifyContent:"space-between"}}>
                  <span>총 재고</span>
                  <span style={{color:DC.text,fontWeight:600}}>{agingQtyStat.totalQty.toLocaleString()}개</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between"}}>
                  <span>총 금액</span>
                  <span style={{color:DC.text,fontWeight:600}}>{fmtWonShort(agingQtyStat.totalVal)}</span>
                </div>
              </div>
            )}
          </div>
          {/* 프로모션 제안 버튼 */}
          <div style={{marginTop:10,paddingTop:10,borderTop:`1px solid ${DC.border}`}}>
            <button onClick={()=>setShowSaleRec(p=>!p)}
              style={{width:"100%",background:showSaleRec?"rgba(200,123,123,0.18)":"rgba(255,255,255,0.04)",
                color:showSaleRec?"#C87B7B":DC.sub,
                border:`1px solid ${showSaleRec?"#C87B7B":DC.border}`,
                borderRadius:7,padding:"9px 0",fontSize:13,fontWeight:600,cursor:"pointer",
                letterSpacing:"-0.2px",transition:"all .12s"}}>
              {"프로모션 제안"}{showSaleRec&&allDiscountRecs.length>0?` (${allDiscountRecs.length})`:""}
            </button>
            {/* 선정 기준 설명 토글 */}
            <button onClick={()=>setShowPromoInfo(p=>!p)}
              style={{width:"100%",marginTop:4,background:"transparent",border:"none",
                color:showPromoInfo?"#7EC8A4":DC.dim,fontSize:12,cursor:"pointer",
                textAlign:"left",padding:"4px 2px",lineHeight:1.4}}>
              {showPromoInfo?"▲ 선정 기준 닫기":"▼ 선정 기준 보기"}
            </button>
            {showPromoInfo&&(
              <div style={{marginTop:4,padding:"10px 11px",background:"rgba(255,255,255,0.03)",
                border:`1px solid ${DC.border}`,borderRadius:7,fontSize:12,color:DC.sub,lineHeight:1.75}}>
                <div style={{color:"#C87B7B",fontWeight:700,marginBottom:5,fontSize:13}}>프로모션 제안 선정 기준</div>
                <div style={{marginBottom:4}}>
                  <span style={{color:DC.text,fontWeight:600}}>① 위치 조건</span><br/>
                  미판매 일수 &gt; 전체 중앙값 <span style={{color:"#7B9EC8"}}>AND</span><br/>
                  현재고 &gt; 전체 중앙값<br/>
                  <span style={{color:DC.dim,fontSize:12}}>(차트 우상단 사분면 SKU)</span>
                </div>
                <div style={{marginBottom:4}}>
                  <span style={{color:DC.text,fontWeight:600}}>② 최근 입고 조건</span><br/>
                  최근 입고일이 스냅샷 기준 ±3개월 이내 (앞 3개월 ~ 뒤 3개월)<br/>
                  <span style={{color:DC.dim,fontSize:12}}>(오래된 재고보다 최근 발주 재고 우선)</span>
                </div>
                <div>
                  <span style={{color:DC.text,fontWeight:600}}>③ 할인율 산정</span><br/>
                  30번째 백분위수 대비 거리(미판매+재고) 기준 정렬<br/>
                  → 거리 클수록 <span style={{color:"#C87B7B"}}>최대 70%</span> 할인 권장<br/>
                  <span style={{color:DC.dim,fontSize:12}}>(최소 30%, 10% 단위)</span>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Right: Filters + Chart + Note */}
      <div style={{flex:1,minWidth:0}}>
        {/* Filters */}
        <div style={{border:`1px solid ${DC.border}`,borderRadius:8,padding:"10px 12px",marginBottom:10}}>
          <div style={{fontSize:12,color:DC.sub,fontWeight:600,marginBottom:7,display:"flex",alignItems:"center",gap:6}}>
            <span>Aging 필터</span>
            <span style={{fontWeight:400,color:DC.dim,fontSize:12}}>· 항목을 클릭해 해제/선택하면 구간별로 집중 분석할 수 있습니다</span>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {INV_AGING_KEYS.map(k=>{
              const def=INV_AGING_DEFS[k];const on=agingFilter.has(k);
              return(
                <button key={k} data-hf onClick={()=>{const s=new Set(agingFilter);on?s.delete(k):s.add(k);setAgingFilter(s);}}
                  style={{background:on?`${def.color}22`:"transparent",color:on?def.color:DC.dim,
                    border:`1px solid ${on?def.color:DC.border}`,borderRadius:5,padding:"4px 10px",fontSize:13,cursor:"pointer"}}>
                  {def.label}
                </button>
              );
            })}
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:8}}>
          <span style={{fontSize:13,color:DC.sub,fontWeight:600,flexShrink:0}}>검색</span>
          <input placeholder="상품명 / 옵션 검색" value={search} onChange={e=>setSearch(e.target.value)}
            style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,
              padding:"5px 10px",fontSize:12,color:DC.text,flex:1,minWidth:120,outline:"none",fontFamily:"inherit"}}/>
          <div style={{display:"flex",alignItems:"center",gap:6,fontSize:13,color:DC.sub,flexShrink:0}}>
            <span>최소재고</span>
            <input type="number" min={1} value={minStock} onChange={e=>setMinStock(Math.max(1,parseInt(e.target.value)||1))}
              style={{width:52,background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,
                padding:"4px 6px",fontSize:12,color:DC.text,textAlign:"center",fontFamily:"inherit"}}/>
          </div>
        </div>
        {selDate&&<div style={{fontSize:12,color:DC.sub,marginBottom:8}}>{selDate} 기준 · {filtered.length.toLocaleString()}개 SKU{showSaleRec?` · 프로모션 제안 ${allDiscountRecs.length}개`:""}</div>}

        {/* Chart */}
        {!selDate
          ?<div style={{textAlign:"center",padding:"64px 0",color:DC.dim,fontSize:13}}>날짜를 선택하면 버블 차트가 표시됩니다</div>
          :loading
            ?<div style={{textAlign:"center",padding:"64px 0",color:DC.dim,fontSize:13}}>데이터 로딩 중...</div>
            :filtered.length===0
              ?<div style={{textAlign:"center",padding:"64px 0",color:DC.dim,fontSize:13}}>해당 조건의 데이터 없음</div>
              :<div style={{width:"100%",height:460}}>
                <ResponsiveContainer width="100%" height="100%">
                  <ScatterChart margin={{top:24,right:32,bottom:44,left:16}}>
                    <CartesianGrid strokeDasharray="2 4" stroke="#1e1e1e"/>
                    <XAxis dataKey="noSalesDays" type="number" name="미판매 일수" domain={xDomain}
                      tick={{fill:DC.text,fontSize:11}} axisLine={{stroke:DC.border}} tickLine={false}
                      tickFormatter={v=>{const m=Math.round(v/30);if(m<=0)return"0";if(m<12)return`${m}개월`;const y=Math.floor(m/12),rm=m%12;return rm>0?`${y}년${rm}개월`:`${y}년`;}}
                      label={{value:"미판매 →",position:"insideBottom",offset:-28,fill:DC.text,fontSize:11}}/>
                    <YAxis dataKey="current_stock_qty" type="number" name="현재고" domain={yDomain}
                      tick={{fill:DC.text,fontSize:11}} axisLine={{stroke:DC.border}} tickLine={false}
                      label={{value:"현재고",angle:-90,position:"insideLeft",offset:14,fill:DC.text,fontSize:11}}/>
                    <ZAxis dataKey="currentInventoryValue" range={[16,1600]} name="재고금액"/>
                    <Tooltip content={<BubbleTooltip/>} cursor={false} wrapperStyle={{transition:"none"}}/>
                    <Scatter data={filtered} shape={<CustomDot/>}/>
                  </ScatterChart>
                </ResponsiveContainer>
              </div>
        }

        {/* Note */}
        {selDate&&filtered.length>0&&(
          <div style={{marginTop:10,padding:"9px 14px",background:"rgba(0,0,0,0.03)",borderRadius:7,
            border:`1px solid ${DC.border}`,fontSize:11,color:DC.text,lineHeight:1.8}}>
            <span style={{color:DC.text,fontWeight:600}}>해석:</span>
            {" "}오른쪽 위 = 장기 미판매 + 과재고 위험 SKU · 버블이 클수록 재고 금액 부담이 큼 · 버블 클릭 시 SKU 상세 확인
          </div>
        )}

        {/* 프로모션 제안 테이블 */}
        {showSaleRec&&allDiscountRecs.length>0&&(
          <div ref={promoTableRef} style={{marginTop:16,borderTop:`1px solid ${DC.border}`,paddingTop:14}}>
            {/* Header */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6,gap:8,flexWrap:"wrap"}}>
              <div>
                <span style={{fontSize:13,fontWeight:700,color:"#C87B7B"}}>전체 SKU 할인율 ({allDiscountRecs.length})</span>
                <span style={{fontSize:11,color:DC.dim,marginLeft:8}}>
                  제안 범위: 미판매 &gt; {fmtDays(medX)} AND 현재고 &gt; {medY.toLocaleString()}개 (중앙값)
                </span>
              </div>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <button onClick={downloadSaleRecs}
                  style={{background:"transparent",color:"#7EC8A4",border:"1px solid #7EC8A4",
                    borderRadius:6,padding:"4px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                  ↓ 엑셀
                </button>
                <button onClick={()=>setShowSaleRec(false)}
                  style={{background:"transparent",color:DC.dim,border:`1px solid ${DC.border}`,
                    borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>
                  닫기
                </button>
              </div>
            </div>
            {/* Range legend */}
            <div style={{display:"flex",gap:12,marginBottom:8,fontSize:11,color:DC.sub}}>
              <span><span style={{color:"#C87B7B",fontWeight:700}}>●</span> 제안 범위 내 30~70%</span>
              <span><span style={{color:DC.dim,fontWeight:700}}>●</span> 범위 외 10~20%</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead>
                  <tr style={{borderBottom:`1px solid ${DC.border}`}}>
                    {["순위","상품코드","상품명","옵션","현재고","재고금액","미판매일수","Aging","권장할인율","세일 후 원가율"].map(h=>(
                      <th key={h} style={{padding:"5px 7px",textAlign:h==="상품명"||h==="옵션"||h==="상품코드"?"left":"center",
                        fontWeight:600,color:DC.text,fontSize:12,whiteSpace:"nowrap"}}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {allDiscountRecs.slice(promoPage*50,(promoPage+1)*50).map((d,i)=>{
                    const def=INV_AGING_DEFS[d.agingKey];
                    const rank=promoPage*50+i+1;
                    return(
                      <tr key={d.id||i} style={{borderBottom:`1px solid ${DC.border}`,
                        background:i%2===0?"transparent":"rgba(0,0,0,0.02)",
                        opacity:d.inZone?1:0.75}}>
                        <td style={{padding:"5px 7px",textAlign:"center",color:d.inZone?"#C87B7B":DC.dim,fontWeight:700}}>{rank}</td>
                        <td style={{padding:"5px 7px",color:DC.sub,fontFamily:"monospace",fontSize:11,maxWidth:90,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                          title={d.product_code}>{d.product_code||"—"}</td>
                        <td style={{padding:"5px 7px",color:DC.text,fontWeight:500,maxWidth:140,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                          title={d.product_name}>{d.product_name}</td>
                        <td style={{padding:"5px 7px",color:DC.text,maxWidth:100,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                          title={d.option_name}>{d.option_name||"—"}</td>
                        <td style={{padding:"5px 7px",textAlign:"center",color:DC.text}}>{(d.current_stock_qty||0).toLocaleString()}</td>
                        <td style={{padding:"5px 7px",textAlign:"center",color:DC.text}}>{
                          (d.currentInventoryValue||0)>=10000
                            ?`${Math.round((d.currentInventoryValue||0)/10000)}만`
                            :(d.currentInventoryValue||0).toLocaleString()
                        }원</td>
                        <td style={{padding:"5px 7px",textAlign:"center",color:DC.text}}>{fmtDays(d.noSalesDays)}</td>
                        <td style={{padding:"5px 7px",textAlign:"center"}}>
                          <span style={{fontSize:12,fontWeight:600,color:def?.color||"#888"}}>{def?.label||"—"}</span>
                        </td>
                        <td style={{padding:"5px 7px",textAlign:"center"}}>
                          <span style={{fontWeight:800,fontSize:13,color:d.inZone?"#C87B7B":DC.sub}}>{d.recommendedDiscount}%</span>
                        </td>
                        <td style={{padding:"5px 7px",textAlign:"center",color:DC.text}}>
                          {(()=>{
                            const sp=d.selling_price||0;
                            const cp=d.supply_price||0;
                            const after=sp*(1-(d.recommendedDiscount||0)/100);
                            if(!cp||!after) return <span style={{color:DC.dim}}>—</span>;
                            const rate=cp/after*100;
                            const col=rate>=90?"#C87B7B":rate>=70?"#D9A53A":"#7EC8A4";
                            return <span style={{fontWeight:700,color:col}}>{rate.toFixed(1)}%</span>;
                          })()}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {/* Pagination */}
            {Math.ceil(allDiscountRecs.length/50)>1&&(
              <div style={{display:"flex",justifyContent:"center",gap:4,marginTop:10,flexWrap:"wrap"}}>
                {Array.from({length:Math.ceil(allDiscountRecs.length/50)}).map((_,i)=>(
                  <button key={i} onClick={()=>{setPromoPage(i);promoTableRef.current?.scrollIntoView({behavior:"smooth",block:"start"});}}
                    style={{background:promoPage===i?"#C87B7B":"transparent",
                      color:promoPage===i?"#fff":DC.sub,
                      border:`1px solid ${promoPage===i?"#C87B7B":DC.border}`,
                      borderRadius:5,padding:"3px 10px",fontSize:12,cursor:"pointer"}}>
                    {i+1}
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Sticky close button — appears when promo table is in viewport (hides if another table also enters) */}
        {showPromoSticky&&(
          <div style={{position:"fixed",bottom:"20vh",left:"50%",transform:"translateX(-50%)",zIndex:800,pointerEvents:"none"}}>
            <button onClick={()=>setShowSaleRec(false)}
              style={{pointerEvents:"auto",background:"rgba(30,30,30,0.92)",color:"#fff",
                border:"1px solid #444",borderRadius:20,padding:"8px 22px",
                fontSize:13,fontWeight:600,cursor:"pointer",backdropFilter:"blur(8px)",
                boxShadow:"0 4px 16px rgba(0,0,0,0.4)"}}>
              ✕ 표 닫기
            </button>
          </div>
        )}
      </div>

      {/* Side panel — position:fixed, 레이아웃 무관 */}
      {selectedSku&&(()=>{
        const d=selectedSku;
        const def=INV_AGING_DEFS[d.agingKey];
        const saleRec=saleRecs.find(s=>s.id===d.id);
        return(
          <div style={{position:"fixed",top:0,right:0,height:"50vh",width:300,background:"#141414",
            borderLeft:"1px solid #242424",zIndex:600,overflowY:"auto",padding:22,boxShadow:"-4px 0 20px rgba(0,0,0,0.5)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
              <span style={{fontWeight:700,fontSize:13,color:"#F0F0F0"}}>SKU 상세</span>
              <button onClick={()=>setSelectedSku(null)} style={{background:"none",border:"none",color:"#666",cursor:"pointer",fontSize:18,lineHeight:1}}>✕</button>
            </div>
            <div style={{fontWeight:600,fontSize:14,color:"#F0F0F0",marginBottom:3}}>{d.product_name}</div>
            {d.product_code&&<div style={{color:"#888",fontSize:11,fontFamily:"monospace",marginBottom:3}}>{d.product_code}</div>}
            {d.option_name&&<div style={{color:"#666",fontSize:12,marginBottom:12}}>{d.option_name}</div>}
            <div style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 10px",
              background:`${def?.color}20`,border:`1px solid ${def?.color}55`,borderRadius:14,marginBottom:16}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:def?.color,display:"inline-block"}}/>
              <span style={{fontSize:13,fontWeight:700,color:def?.color}}>{def?.label}</span>
            </div>
            {saleRec&&showSaleRec&&(
              <div style={{background:"rgba(200,123,123,0.1)",border:"1px solid rgba(200,123,123,0.35)",borderRadius:8,padding:"9px 12px",marginBottom:14}}>
                <div style={{fontSize:12,fontWeight:700,color:"#C87B7B",marginBottom:3}}>세일 추천</div>
                <div style={{fontSize:14,color:"#F0F0F0",fontWeight:700}}>{saleRec.recommendedDiscount}% 할인 권장</div>
              </div>
            )}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
              {[["판매가",`${(d.selling_price||0).toLocaleString()}원`],["현재고",`${(d.current_stock_qty||0).toLocaleString()}개`],
                ["재고 금액",`${(d.currentInventoryValue||0).toLocaleString()}원`],["미판매",fmtDays(d.noSalesDays)],
                ["SKU 기간",fmtDays(d.skuAge)],["입고 후",fmtDays(d.postRestockDays)]].map(([l,v])=>(
                <div key={l} style={{background:"#1e1e1e",borderRadius:7,padding:"9px 11px"}}>
                  <div style={{fontSize:12,color:"#555",marginBottom:3}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:600,color:"#F0F0F0"}}>{v}</div>
                </div>
              ))}
            </div>
            <div style={{borderTop:"1px solid #242424",paddingTop:13,marginBottom:12,fontSize:12,color:"#666"}}>
              <div style={{fontSize:12,color:"#555",marginBottom:6,fontWeight:600}}>스냅샷</div>
              데이터 날짜: <span style={{color:"#F0F0F0"}}>{d.snapshot_date}</span>
            </div>
            <div style={{borderTop:"1px solid #242424",paddingTop:13}}>
              <div style={{fontSize:12,color:"#555",marginBottom:8,fontWeight:600}}>판매/재고 요약</div>
              {[["처음입고일",d.first_inbound_date||"—"],["처음입고수량",`${(d.first_inbound_qty||0).toLocaleString()}개`],
                ["누적입고",`${(d.cumulative_inbound_qty||0).toLocaleString()}개`],["마지막입고일",d.latest_inbound_date||"—"],
                ["마지막배송일",d.last_delivery_date||"—"],["누적배송",`${(d.cumulative_delivery_qty||0).toLocaleString()}개`],
                ["판매효율 (STP)",String(d.sellThroughProxy)]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:5,color:"#666"}}>
                  <span>{l}</span><span style={{color:"#F0F0F0",fontWeight:500}}>{v}</span>
                </div>
              ))}
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ─────────────────────────────────────────────
// INV AGING TREND (STACKED AREA)
// ─────────────────────────────────────────────
function InvAgingTrend({DC,snapshotDates,refreshKey,onDateReady,stopRef}){
  const [rawByDate,setRawByDate]=useState({});
  const [loading,setLoading]=useState(false);
  const [aggUnit,setAggUnit]=useState("day");
  const [yMode,setYMode]=useState("qty");
  const [dateRange,setDateRange]=useState("14d"); // preset or "custom"
  useEffect(()=>{
    if(dateRange==="7d"||dateRange==="14d") setAggUnit("day");
    else if(dateRange==="30d"||dateRange==="90d") setAggUnit(p=>p==="day"?"week":p);
    else if(dateRange==="1y") setAggUnit(p=>(p==="day"||p==="week")?"month":p);
  },[dateRange]);
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [calPickMode,setCalPickMode]=useState("single"); // "single"|"range"
  const [showCal,setShowCal]=useState(false);
  const [drillPage,setDrillPage]=useState(0);
  const [drillTableVisible,setDrillTableVisible]=useState(false);
  const drillTableRef=useRef(null);
  const [clickedBar,setClickedBar]=useState(null); // {label, agingKey}
  const [nextSecVisible,setNextSecVisible]=useState(false);
  useEffect(()=>{
    if(!stopRef?.current) return;
    const obs=new IntersectionObserver(([e])=>setNextSecVisible(e.isIntersecting),{threshold:0});
    obs.observe(stopRef.current);
    return()=>obs.disconnect();
  },[stopRef]);
  const showDrillSticky=useStickyTable("drillTable",!!clickedBar&&drillTableVisible)&&!nextSecVisible;
  const [drillRows,setDrillRows]=useState([]);
  const [drillLoading,setDrillLoading]=useState(false);

  const rangeStart=useMemo(()=>{
    if(dateRange==="custom"&&customStart) return customStart;
    const d=dayjs();
    const map={["7d"]:7,["14d"]:14,["30d"]:30,["90d"]:90};
    if(map[dateRange]) return d.subtract(map[dateRange],"day").format("YYYY-MM-DD");
    if(dateRange==="1y") return d.subtract(1,"year").format("YYYY-MM-DD");
    return d.subtract(90,"day").format("YYYY-MM-DD");
  },[dateRange,customStart]);

  const rangeEnd=useMemo(()=>{
    if(dateRange==="custom"&&customEnd) return customEnd;
    return null;
  },[dateRange,customEnd]);

  useEffect(()=>{
    const inRange=snapshotDates.filter(d=>d>=rangeStart&&(!rangeEnd||d<=rangeEnd));
    if(!inRange.length){setRawByDate({});return;}
    setLoading(true);
    (async()=>{
      const db=await getSupabase();
      const map={};
      const PAGE=1000;let from=0;
      while(true){
        const{data:rows}=await db.from("inventory_snapshot")
          .select("snapshot_date,current_stock_qty,selling_price,last_delivery_date,first_inbound_date")
          .gte("snapshot_date",rangeStart)
          .order("snapshot_date",{ascending:true})
          .range(from,from+PAGE-1);
        if(!rows||rows.length===0) break;
        rows.forEach(r=>{
          if(rangeEnd&&r.snapshot_date>rangeEnd) return;
          const c=calcInvRow(r);
          if(!map[r.snapshot_date]) map[r.snapshot_date]={};
          if(!map[r.snapshot_date][c.agingKey]) map[r.snapshot_date][c.agingKey]={count:0,qty:0,value:0};
          map[r.snapshot_date][c.agingKey].count++;
          map[r.snapshot_date][c.agingKey].qty+=r.current_stock_qty||0;
          map[r.snapshot_date][c.agingKey].value+=c.currentInventoryValue||0;
        });
        if(rows.length<PAGE) break;
        from+=PAGE;
      }
      setRawByDate(map);
      setLoading(false);
    })();
  },[snapshotDates,rangeStart,rangeEnd,refreshKey]);

  const chartData=useMemo(()=>{
    const dates=Object.keys(rawByDate).sort();
    if(!dates.length) return[];
    const groups={};
    dates.forEach(d=>{
      let key,label;
      if(aggUnit==="day"){
        key=d;label=dayjs(d).format("M/D");
      } else if(aggUnit==="week"){
        const dd=dayjs(d);const sw=dd.subtract(dd.day(),"day");
        key=sw.format("YYYY-MM-DD");label=sw.format("M/D");
      } else if(aggUnit==="quarter"){
        const dd=dayjs(d);const q=Math.floor(dd.month()/3)+1;
        key=`${dd.year()}-Q${q}`;label=key;
      } else {
        key=d.slice(0,7);label=key;
      }
      if(!groups[key]){
        groups[key]={label,HEALTHY:0,SLOW:0,AGING:0,DEAD:0,_n:0};
        INV_AGING_KEYS.forEach(k=>{groups[key][`${k}_qty`]=0;groups[key][`${k}_val`]=0;groups[key][`${k}_count`]=0;});
      }
      INV_AGING_KEYS.forEach(k=>{
        const v=rawByDate[d][k];
        if(v){
          groups[key][k]+=yMode==="count"?v.count:v.qty;
          groups[key][`${k}_qty`]+=v.qty;
          groups[key][`${k}_val`]+=v.value;
          groups[key][`${k}_count`]+=v.count;
        }
      });
      groups[key]._n++;
    });
    return Object.values(groups).map(g=>{
      const r={label:g.label};
      INV_AGING_KEYS.forEach(k=>{
        r[k]=Math.round(g[k]/g._n);
        r[`${k}_qty`]=Math.round((g[`${k}_qty`]||0)/g._n);
        r[`${k}_val`]=Math.round((g[`${k}_val`]||0)/g._n);
        r[`${k}_count`]=Math.round((g[`${k}_count`]||0)/g._n);
      });
      return r;
    });
  },[rawByDate,aggUnit,yMode]);

  const latestDate=useMemo(()=>{const d=Object.keys(rawByDate).sort();return d[d.length-1]||null;},[rawByDate]);

  useEffect(()=>{if(onDateReady) onDateReady(latestDate);},[latestDate,onDateReady]);

  // label → snapshot dates mapping (for drill-down)
  const datesByLabel=useMemo(()=>{
    const map={};
    Object.keys(rawByDate).sort().forEach(d=>{
      let label;
      if(aggUnit==="day"){label=dayjs(d).format("M/D");}
      else if(aggUnit==="week"){const dd=dayjs(d);label=dd.subtract(dd.day(),"day").format("M/D");}
      else if(aggUnit==="quarter"){const dd=dayjs(d);label=`${dd.year()}-Q${Math.floor(dd.month()/3)+1}`;}
      else{label=d.slice(0,7);}
      if(!map[label]) map[label]=[];
      map[label].push(d);
    });
    return map;
  },[rawByDate,aggUnit]);

  const handleBarClick=useCallback((agingKey,data)=>{
    if(!data||!data.label) return;
    const label=data.label;
    // toggle off
    if(clickedBar&&clickedBar.label===label&&clickedBar.agingKey===agingKey){
      setClickedBar(null);setDrillRows([]);return;
    }
    const dates=datesByLabel[label]||[];
    const targetDate=dates[dates.length-1]||null;
    setClickedBar({label,agingKey,targetDate});
    setDrillLoading(true);
    if(!dates.length){setDrillLoading(false);return;}
    (async()=>{
      const db=await getSupabase();
      let all=[];let from=0;const PAGE=1000;
      while(true){
        const{data:rows,error}=await db.from("inventory_snapshot").select("*")
          .eq("snapshot_date",targetDate).range(from,from+PAGE-1);
        if(error||!rows||!rows.length) break;
        all=all.concat(rows);
        if(rows.length<PAGE) break;
        from+=PAGE;
      }
      setDrillRows(all.map(calcInvRow).filter(r=>r.agingKey===agingKey));
      setDrillLoading(false);
    })();
  },[datesByLabel,clickedBar]);

  useEffect(()=>{
    const el=drillTableRef.current;
    if(!el||!clickedBar){setDrillTableVisible(false);return;}
    const obs=new IntersectionObserver(([e])=>setDrillTableVisible(e.isIntersecting),{threshold:0.05});
    obs.observe(el);
    return()=>obs.disconnect();
  },[clickedBar]);

  useEffect(()=>{setDrillPage(0);},[clickedBar]);

  const kpi=useMemo(()=>{
    if(!latestDate||!rawByDate[latestDate]) return null;
    const d=rawByDate[latestDate];
    let total=0,totalQty=0,totalVal=0;
    INV_AGING_KEYS.forEach(k=>{total+=(d[k]?.count||0);totalQty+=(d[k]?.qty||0);totalVal+=(d[k]?.value||0);});
    const deadQty=d["DEAD"]?.qty||0;const healthyQty=d["HEALTHY"]?.qty||0;
    const qtyByKey={};
    INV_AGING_KEYS.forEach(k=>{
      const qty=d[k]?.qty||0;const val=d[k]?.value||0;const count=d[k]?.count||0;
      qtyByKey[k]={qty,pct:totalQty?(qty/totalQty*100).toFixed(1):"0.0",val,valPct:totalVal?(val/totalVal*100).toFixed(1):"0.0",count,skuPct:total?(count/total*100).toFixed(1):"0.0"};
    });
    const deadCount=d["DEAD"]?.count||0;const healthyCount=d["HEALTHY"]?.count||0;
    return{total,totalQty,totalVal,
      deadPct:totalQty?((deadQty/totalQty)*100).toFixed(1):"0.0",
      healthyPct:totalQty?((healthyQty/totalQty)*100).toFixed(1):"0.0",
      deadSkuPct:total?((deadCount/total)*100).toFixed(1):"0.0",
      healthySkuPct:total?((healthyCount/total)*100).toFixed(1):"0.0",
      qtyByKey};
  },[latestDate,rawByDate]);

  const fmtVal=v=>{
    if(v>=100000000) return `${(v/100000000).toFixed(1)}억`;
    if(v>=10000) return `${(v/10000).toFixed(0)}만`;
    return v.toLocaleString();
  };

  const AreaTooltip=({active,payload,label})=>{
    if(!active||!payload?.length) return null;
    const total=payload.reduce((s,p)=>s+(p.value||0),0);
    const unit="개";
    const dp=payload[0]?.payload||{};
    const totalVal=INV_AGING_KEYS.reduce((s,k)=>s+(dp[`${k}_val`]||0),0);
    return(
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:8,padding:"10px 14px",fontSize:12,minWidth:200}}>
        <div style={{color:DC.sub,marginBottom:7,fontWeight:600,fontSize:13}}>{label}</div>
        {[...payload].reverse().map(p=>{
          const val=dp[`${p.dataKey}_val`]||0;
          const qty=dp[`${p.dataKey}_qty`]||0;
          const cnt=dp[`${p.dataKey}_count`]||0;
          const totalQtyAll=INV_AGING_KEYS.reduce((s,k)=>s+(dp[`${k}_qty`]||0),0)||1;
          const totalCntAll=INV_AGING_KEYS.reduce((s,k)=>s+(dp[`${k}_count`]||0),0)||1;
          return(
            <div key={p.dataKey} style={{marginBottom:6,paddingBottom:6,borderBottom:`1px solid ${DC.border}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                <span style={{color:p.fill||p.stroke,fontWeight:600}}>{INV_AGING_DEFS[p.dataKey]?.label}</span>
                <span style={{fontSize:10,color:DC.sub}}>{cnt.toLocaleString()} SKU <span style={{fontWeight:400}}>({(cnt/totalCntAll*100).toFixed(1)}%)</span></span>
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <span style={{color:DC.text,fontWeight:600}}>{qty.toLocaleString()}개<span style={{color:DC.sub,fontWeight:400,marginLeft:3}}>({(qty/totalQtyAll*100).toFixed(1)}%)</span></span>
                <span style={{color:DC.sub}}>·</span>
                <span style={{color:DC.text}}>{fmtVal(val)}원{totalVal>0&&<span style={{color:DC.sub,marginLeft:3}}>({(val/totalVal*100).toFixed(1)}%)</span>}</span>
              </div>
            </div>
          );
        })}
        <div style={{display:"flex",justifyContent:"space-between",gap:14,paddingTop:2}}>
          <span style={{color:DC.sub}}>합계</span>
          <span style={{color:DC.text,fontWeight:600}}>{INV_AGING_KEYS.reduce((s,k)=>s+(dp[`${k}_qty`]||0),0).toLocaleString()}개 · {fmtVal(totalVal)}원</span>
        </div>
      </div>
    );
  };

  return(
    <div>
      {/* KPI cards */}
      {kpi&&(
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
          {[
            {label:"Dead Stock 비율",value:`${kpi.deadPct}%`,color:"#C87B7B"},
            {label:"Healthy 비율",value:`${kpi.healthyPct}%`,color:"#7EC8A4"},
            {label:"총 현재고",value:`${kpi.totalQty.toLocaleString()}개`,color:DC.text},
            {label:"총 재고 금액",value:fmtVal(kpi.totalVal)+"원",color:"#C8A87B"},
          ].map(c=>(
            <div key={c.label} style={{background:DC.bg,border:`1px solid ${DC.border}`,borderRadius:8,padding:"13px 15px"}}>
              <div style={{fontSize:12,color:DC.sub,marginBottom:5}}>{c.label}</div>
              <div style={{fontSize:18,fontWeight:700,color:c.color,letterSpacing:"-0.3px"}}>{c.value}</div>
            </div>
          ))}
        </div>
      )}

      {/* Controls */}
      <div style={{border:`1px solid ${DC.border}`,borderRadius:8,padding:"10px 14px",marginBottom:14}}>
        <div style={{fontSize:12,color:DC.sub,fontWeight:600,marginBottom:8}}>필터 <span style={{fontWeight:400,color:DC.dim}}>· 항목을 클릭해 기간·집계 단위·지표를 전환할 수 있습니다</span></div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{fontSize:12,color:DC.text,fontWeight:600,flexShrink:0,marginRight:2}}>기간</span>
          {[["7d","7일"],["14d","2주"],["30d","30일"],["90d","90일"],["1y","1년"]].map(([v,l])=>(
            <button key={v} data-hf onClick={()=>setDateRange(v)}
              style={{background:dateRange===v?DC.text:"transparent",color:dateRange===v?DC.card:DC.sub,
                border:`1px solid ${dateRange===v?DC.text:DC.border}`,borderRadius:5,padding:"4px 10px",fontSize:13,cursor:"pointer",fontWeight:dateRange===v?600:400}}>
              {l}
            </button>
          ))}
          <button data-hf onClick={()=>{setDateRange("custom");setShowCal(p=>!p);}}
            style={{background:dateRange==="custom"?DC.text:"transparent",
              color:dateRange==="custom"?DC.card:DC.sub,
              border:`1px solid ${dateRange==="custom"?DC.text:DC.border}`,
              borderRadius:5,padding:"4px 10px",fontSize:13,cursor:"pointer",
              fontWeight:dateRange==="custom"?600:400}}>
            직접 선택
          </button>
          <span style={{color:DC.border,margin:"0 3px",fontSize:14}}>|</span>
          <span style={{fontSize:12,color:DC.text,fontWeight:600,flexShrink:0,marginRight:2}}>집계</span>
          {[["day","일간"],["week","주간"],["month","월간"],["quarter","분기"]].map(([v,l])=>(
            <button key={v} data-hf onClick={()=>setAggUnit(v)}
              style={{background:aggUnit===v?DC.text:"transparent",color:aggUnit===v?DC.card:DC.sub,
                border:`1px solid ${aggUnit===v?DC.text:DC.border}`,borderRadius:5,padding:"4px 10px",fontSize:13,cursor:"pointer"}}>
              {l}
            </button>
          ))}
          <span style={{color:DC.border,margin:"0 3px",fontSize:14}}>|</span>
          <span style={{fontSize:12,color:DC.text,fontWeight:600,flexShrink:0,marginRight:2}}>단위</span>
          {[["qty","재고 수량"],["count","SKU 수"]].map(([v,l])=>(
            <button key={v} data-hf onClick={()=>setYMode(v)}
              style={{background:yMode===v?"rgba(126,200,164,0.15)":"transparent",color:yMode===v?"#7EC8A4":DC.sub,
                border:`1px solid ${yMode===v?"#7EC8A4":DC.border}`,borderRadius:5,padding:"4px 10px",fontSize:13,cursor:"pointer"}}>
              {l}
            </button>
          ))}
        </div>
      </div>
      {showCal&&(
        <div style={{marginBottom:10,padding:12,border:`1px solid ${DC.border}`,borderRadius:8,
          display:"inline-flex",flexDirection:"column",gap:8}}>
          <div style={{display:"flex",gap:4,marginBottom:4}}>
            {[["single","단일"],["range","기간"]].map(([k,l])=>(
              <button key={k} data-hf onClick={()=>setCalPickMode(k)}
                style={{flex:1,background:calPickMode===k?DC.text:"transparent",
                  color:calPickMode===k?DC.card:DC.sub,
                  border:`1px solid ${calPickMode===k?DC.text:DC.border}`,
                  borderRadius:5,padding:"4px 0",fontSize:12,cursor:"pointer"}}>
                {l}
              </button>
            ))}
          </div>
          <CalendarPicker
            mode={calPickMode}
            value={customStart}
            onChange={v=>{setCustomStart(v);setCustomEnd("");setShowCal(false);}}
            rangeStart={customStart}
            rangeEnd={customEnd}
            onRangeChange={({start,end})=>{
              setCustomStart(start);setCustomEnd(end||"");
              if(start&&end) setShowCal(false);
            }}
            availableDates={new Set(snapshotDates)}
            DC={{bg:"transparent",surface:"rgba(0,0,0,0.04)",border:DC.border,
              text:DC.text,sub:DC.sub,dim:DC.dim,green:"#7EC8A4",greenBg:"rgba(126,200,164,0.12)"}}
          />
        </div>
      )}

      <div style={{display:"flex",gap:16,alignItems:"flex-start"}}>
        {/* Stacked bar chart */}
        <div style={{flex:1,height:340}}>
          {loading
            ?<div style={{textAlign:"center",padding:"80px 0",color:DC.dim,fontSize:15}}>데이터 로딩 중...</div>
            :chartData.length===0
              ?<div style={{textAlign:"center",padding:"80px 0",color:DC.dim,fontSize:15}}>해당 기간 데이터 없음</div>
              :<ResponsiveContainer width="100%" height="100%">
                <AreaChart data={chartData} margin={{top:10,right:8,bottom:28,left:8}}>
                  <CartesianGrid strokeDasharray="2 5" stroke={DC.border} vertical={false}/>
                  <XAxis dataKey="label" tick={{fill:DC.text,fontSize:12}} axisLine={{stroke:DC.border}} tickLine={false}
                    angle={-20} textAnchor="end" interval="preserveStartEnd" dy={6}/>
                  <YAxis tick={{fill:DC.text,fontSize:12}} axisLine={{stroke:DC.border}} tickLine={false}
                    tickFormatter={v=>v>=1000?`${(v/1000).toFixed(1)}k`:String(v)}/>
                  <Tooltip content={<AreaTooltip/>} cursor={{stroke:DC.border,strokeDasharray:"3 3"}}/>
                  {INV_AGING_KEYS.map(k=>(
                    <Area key={k} type="monotone" dataKey={k} stackId="1"
                      stroke={INV_AGING_DEFS[k].color} strokeWidth={1.5}
                      fill={INV_AGING_DEFS[k].color}
                      fillOpacity={clickedBar&&clickedBar.agingKey!==k?0.25:0.7}
                      activeDot={{r:4,style:{cursor:"pointer"},onClick:(_,e)=>{const p=e?.payload;if(p)handleBarClick(k,p);}}}
                      style={{cursor:"pointer"}}
                      onClick={(data)=>handleBarClick(k,data?.payload||data)}/>
                  ))}
                </AreaChart>
              </ResponsiveContainer>
          }
        </div>

        {/* Legend panel */}
        <div style={{width:160,flexShrink:0,border:`1px solid ${DC.border}`,borderRadius:8,padding:"12px 14px"}}>
          <div style={{fontSize:12,fontWeight:700,color:DC.sub,marginBottom:12,letterSpacing:".06em"}}>AGING STATUS</div>
          {INV_AGING_KEYS.map(k=>{
            const def=INV_AGING_DEFS[k];
            const q=kpi?.qtyByKey[k];
            return(
              <div key={k} style={{marginBottom:13}}>
                <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:4}}>
                  <span style={{width:9,height:9,borderRadius:2,background:def.color,display:"inline-block",flexShrink:0}}/>
                  <span style={{fontSize:12,fontWeight:700,color:def.color}}>{def.label}</span>
                </div>
                <div style={{fontSize:10,color:DC.text,paddingLeft:14,lineHeight:1.4,marginBottom:3}}>{def.desc}</div>
                {q&&(
                  <div style={{paddingLeft:14,display:"flex",flexDirection:"column",gap:2}}>
                    <div style={{display:"flex",justifyContent:"space-between",gap:6}}>
                      <span style={{fontSize:10,color:DC.sub}}>수량</span>
                      <span style={{fontSize:10,color:DC.text,fontWeight:600}}>{q.qty.toLocaleString()}개 <span style={{color:DC.sub,fontWeight:400}}>({q.pct}%)</span></span>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",gap:6}}>
                      <span style={{fontSize:10,color:DC.sub}}>금액</span>
                      <span style={{fontSize:10,color:DC.text,fontWeight:600}}>{fmtVal(q.val)}원 <span style={{color:DC.sub,fontWeight:400}}>({q.valPct}%)</span></span>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",gap:6}}>
                      <span style={{fontSize:10,color:DC.sub,opacity:.7}}>SKU</span>
                      <span style={{fontSize:10,color:DC.sub}}>{q.count.toLocaleString()}개 <span style={{fontWeight:400}}>({q.skuPct}%)</span></span>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* Footnote */}
      <div style={{marginTop:16,paddingTop:12,borderTop:`1px solid ${DC.border}`,fontSize:12,color:DC.text,lineHeight:1.7}}>
        <span style={{color:DC.text,fontWeight:600,marginRight:6}}>계산 방식</span>
        마지막 판매일 기준 경과일수로 에이징을 분류합니다.&nbsp;
        <span style={{color:INV_AGING_DEFS.HEALTHY.color}}>Healthy</span> 0~30일 ·&nbsp;
        <span style={{color:INV_AGING_DEFS.SLOW.color}}>Slow-moving</span> 31~90일 ·&nbsp;
        <span style={{color:INV_AGING_DEFS.AGING.color}}>Aging</span> 91~180일 ·&nbsp;
        <span style={{color:INV_AGING_DEFS.DEAD.color}}>Dead Stock</span> 180일 초과.&nbsp;
        스택 높이는 선택된 지표(SKU 수 / 재고 수량 / 재고 금액)를 기간별로 집계한 값입니다.
      </div>

      {/* Drill-down table */}
      {clickedBar&&(
        <div ref={drillTableRef} style={{marginTop:16,borderTop:`1px solid ${DC.border}`,paddingTop:14}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexWrap:"wrap"}}>
            <span style={{fontSize:13,fontWeight:700,color:INV_AGING_DEFS[clickedBar.agingKey]?.color}}>
              {clickedBar.label} · {INV_AGING_DEFS[clickedBar.agingKey]?.label}
            </span>
            {!drillLoading&&<span style={{fontSize:12,color:DC.text}}>{drillRows.length.toLocaleString()}개 SKU</span>}
            {clickedBar.targetDate&&clickedBar.targetDate!==latestDate&&(
              <span style={{fontSize:11,color:"#e07b00",background:"#fff3e0",borderRadius:4,padding:"2px 7px",border:"1px solid #e07b00"}}>
                분석일 {clickedBar.targetDate} 기준 · 현재와 다를 수 있습니다
              </span>
            )}
            <button onClick={()=>{setClickedBar(null);setDrillRows([]);}}
              style={{marginLeft:"auto",background:"none",border:"none",color:DC.text,cursor:"pointer",fontSize:16,lineHeight:1}}>✕</button>
          </div>
          {drillLoading
            ?<div style={{textAlign:"center",padding:"30px 0",color:DC.text,fontSize:13}}>로딩 중...</div>
            :drillRows.length===0
              ?<div style={{textAlign:"center",padding:"20px 0",color:DC.text,fontSize:13}}>데이터 없음</div>
              :<div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{borderBottom:`1px solid ${DC.border}`}}>
                      {["상품코드","상품명","옵션","재고 수","미판매 일수","재고 금액"].map(h=>(
                        <th key={h} style={{padding:"6px 8px",textAlign:h==="상품명"||h==="옵션"||h==="상품코드"?"left":"center",
                          fontWeight:600,color:DC.text,fontSize:12,whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {[...drillRows].sort((a,b)=>(b.current_stock_qty||0)-(a.current_stock_qty||0)||b.noSalesDays-a.noSalesDays).slice(drillPage*50,(drillPage+1)*50).map((r,i)=>(
                      <tr key={r.id||i} style={{borderBottom:`1px solid ${DC.border}`,
                        background:i%2===0?"transparent":"rgba(0,0,0,0.02)"}}>
                        <td style={{padding:"6px 8px",color:DC.sub,fontFamily:"monospace",fontSize:11,maxWidth:90,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                          title={r.product_code}>{r.product_code||"—"}</td>
                        <td style={{padding:"6px 8px",color:DC.text,fontWeight:500,maxWidth:160,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                          title={r.product_name}>{r.product_name}</td>
                        <td style={{padding:"6px 8px",color:DC.text,maxWidth:100,
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}
                          title={r.option_name}>{r.option_name||"—"}</td>
                        <td style={{padding:"6px 8px",textAlign:"center",color:DC.text,fontWeight:600}}>
                          {(r.current_stock_qty||0).toLocaleString()}</td>
                        <td style={{padding:"6px 8px",textAlign:"center",fontWeight:700,
                          color:INV_AGING_DEFS[r.agingKey]?.color||DC.text}}>
                          {fmtDays(r.noSalesDays)}</td>
                        <td style={{padding:"6px 8px",textAlign:"center",color:DC.text}}>
                          {(r.currentInventoryValue||0)>=10000
                            ?`${Math.round((r.currentInventoryValue||0)/10000)}만`
                            :(r.currentInventoryValue||0).toLocaleString()}원</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {Math.ceil(drillRows.length/50)>1&&(
                  <div style={{display:"flex",justifyContent:"center",gap:4,marginTop:10,flexWrap:"wrap"}}>
                    {Array.from({length:Math.ceil(drillRows.length/50)}).map((_,i)=>(
                      <button key={i} onClick={()=>{setDrillPage(i);drillTableRef.current?.scrollIntoView({behavior:"smooth",block:"start"});}}
                        style={{background:drillPage===i?INV_AGING_DEFS[clickedBar.agingKey]?.color||DC.text:"transparent",
                          color:drillPage===i?"#fff":DC.sub,
                          border:`1px solid ${drillPage===i?(INV_AGING_DEFS[clickedBar.agingKey]?.color||DC.text):DC.border}`,
                          borderRadius:5,padding:"3px 10px",fontSize:12,cursor:"pointer"}}>
                        {i+1}
                      </button>
                    ))}
                  </div>
                )}
              </div>
          }
        </div>
      )}

      {/* Sticky close button — appears when drill table is in viewport (hides if another table also enters) */}
      {showDrillSticky&&(
        <div style={{position:"fixed",bottom:"20vh",left:"50%",transform:"translateX(-50%)",zIndex:800,pointerEvents:"none"}}>
          <button onClick={()=>{setClickedBar(null);setDrillRows([]);}}
            style={{pointerEvents:"auto",background:"rgba(30,30,30,0.92)",color:"#fff",
              border:"1px solid #444",borderRadius:20,padding:"8px 22px",
              fontSize:13,fontWeight:600,cursor:"pointer",backdropFilter:"blur(8px)",
              boxShadow:"0 4px 16px rgba(0,0,0,0.4)"}}>
            ✕ 표 닫기
          </button>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// REORDER CALCULATION
// SQL for new table (run once in Supabase):
// create table if not exists public.reorder_recommendations (
//   reorder_id uuid default gen_random_uuid() primary key,
//   reorder_data_date date not null,
//   reorder_product_code text default '',
//   reorder_product_name text not null,
//   reorder_option_name text default '',
//   reorder_available_stock integer default 0,
//   reorder_incoming_stock integer default 0,
//   reorder_effective_stock integer default 0,
//   reorder_weekly_sales integer default 0,
//   reorder_monthly_sales integer default 0,
//   reorder_expected_daily_sales numeric(10,4) default 0,
//   reorder_days_left numeric(10,2) default 9999,
//   reorder_trend_ratio numeric(10,4) default 0,
//   reorder_recommended_qty integer default 0,
//   reorder_created_at timestamptz default now()
// );
// ─────────────────────────────────────────────
async function computeAndSaveReorder(parsedRows,snapDate){
  const rows=parsedRows.filter(r=>r._r_weekly!=null);
  if(!rows.length) return;
  const computed=rows.map(r=>{
    const avail=r._r_avail||0;
    const incoming=r._r_incoming||0;
    const weekly=r._r_weekly||0;
    const monthly=r._r_monthly||0;
    const effective=avail+incoming;
    const daily7=weekly/7;
    const daily28=monthly>0?monthly/28:0;
    const expectedDaily=(daily7*0.7)+(daily28*0.3);
    if(expectedDaily<=0) return null;
    const daysLeft=effective/expectedDaily;
    const trendRatio=daily28>0?daily7/daily28:0;
    const recommended=Math.max(0,Math.round(expectedDaily*14-effective));
    return{
      reorder_data_date:snapDate,
      reorder_product_code:r.product_code||"",
      reorder_product_name:r.product_name,
      reorder_option_name:r.option_name||"",
      reorder_available_stock:avail,
      reorder_incoming_stock:incoming,
      reorder_effective_stock:effective,
      reorder_weekly_sales:weekly,
      reorder_monthly_sales:monthly,
      reorder_expected_daily_sales:Math.round(expectedDaily*10000)/10000,
      reorder_days_left:Math.round(daysLeft*100)/100,
      reorder_trend_ratio:Math.round(trendRatio*10000)/10000,
      reorder_recommended_qty:recommended,
    };
  }).filter(r=>r&&r.reorder_days_left<14);
  const db=await getSupabase();
  // Skip only if existing data is strictly newer (avoid overwriting newer upload with older)
  const{data:latest}=await db.from("reorder_recommendations")
    .select("reorder_data_date").order("reorder_data_date",{ascending:false}).limit(1);
  const latestDate=latest?.[0]?.reorder_data_date;
  if(latestDate&&snapDate<latestDate) return;
  // Always clear existing rows so stale dates don't linger when new upload has 0 reorder items
  await db.from("reorder_recommendations").delete().lte("reorder_created_at",new Date().toISOString());
  if(!computed.length) return;
  const CHUNK=200;
  for(let i=0;i<computed.length;i+=CHUNK){
    await db.from("reorder_recommendations").insert(computed.slice(i,i+CHUNK));
  }
}

// ─────────────────────────────────────────────
// REORDER CALCULATOR COMPONENT
// ─────────────────────────────────────────────
function ReorderCalculator({DC,refreshKey,onDateReady,latestSnapDate}){
  const [data,setData]=useState([]);
  const [loading,setLoading]=useState(false);
  const [search,setSearch]=useState("");
  const [sortKey,setSortKey]=useState("reorder_days_left");
  const [sortDir,setSortDir]=useState("asc");
  const [pg,setPg]=useState(0);
  const [selected,setSelected]=useState(()=>new Set());
  const PG=20;

  const load=useCallback(async()=>{
    setLoading(true);
    const db=await getSupabase();
    let all=[];let from=0;const PAGE=1000;
    while(true){
      const{data:rows,error}=await db.from("reorder_recommendations")
        .select("*").order("reorder_days_left",{ascending:true}).range(from,from+PAGE-1);
      if(error||!rows||rows.length===0) break;
      all=all.concat(rows);
      if(rows.length<PAGE) break;
      from+=PAGE;
    }
    setData(all);
    setLoading(false);
  },[]);

  useEffect(()=>{load();},[load,refreshKey]);

  const latestDataDate=useMemo(()=>{
    const dates=data.map(r=>r.reorder_data_date||"").filter(Boolean).sort();
    const fromData=dates[dates.length-1]||null;
    // Prefer the newer of the two so 기준일 reflects the most recent upload
    // even when the new upload had 0 SKUs needing reorder (table is empty)
    if(fromData&&latestSnapDate) return fromData>latestSnapDate?fromData:latestSnapDate;
    return fromData||latestSnapDate||null;
  },[data,latestSnapDate]);

  useEffect(()=>{if(onDateReady) onDateReady(latestDataDate);},[latestDataDate,onDateReady]);

  const kpi=useMemo(()=>{
    if(!data.length) return null;
    const n=data.length;
    const avgDays=data.reduce((s,r)=>s+(r.reorder_days_left||0),0)/n;
    const totalQty=data.reduce((s,r)=>s+(r.reorder_recommended_qty||0),0);
    const totalIncoming=data.reduce((s,r)=>s+(r.reorder_incoming_stock||0),0);
    return{n,avgDays,totalQty,totalIncoming};
  },[data]);

  const filtered=useMemo(()=>{
    let rows=data;
    if(search) rows=rows.filter(r=>
      (r.reorder_product_name||"").toLowerCase().includes(search.toLowerCase())||
      (r.reorder_product_code||"").toLowerCase().includes(search.toLowerCase())||
      (r.reorder_option_name||"").toLowerCase().includes(search.toLowerCase())
    );
    return[...rows].sort((a,b)=>{
      const va=a[sortKey]??0,vb=b[sortKey]??0;
      return sortDir==="asc"?(va>vb?1:-1):(va<vb?1:-1);
    });
  },[data,search,sortKey,sortDir]);

  const paged=filtered.slice(pg*PG,(pg+1)*PG);
  const totalPgs=Math.ceil(filtered.length/PG);

  const daysDistData=useMemo(()=>
    [{label:"0~3일",min:0,max:3},{label:"3~7일",min:3,max:7},{label:"7~10일",min:7,max:10},{label:"10~14일",min:10,max:14}]
      .map(b=>({label:b.label,count:data.filter(r=>(r.reorder_days_left||0)>=b.min&&(r.reorder_days_left||0)<b.max).length}))
  ,[data]);

  const topSales=useMemo(()=>
    [...data].sort((a,b)=>(b.reorder_expected_daily_sales||0)-(a.reorder_expected_daily_sales||0)).slice(0,5).map(r=>({
      name:`${r.reorder_product_name||""}${r.reorder_option_name?` / ${r.reorder_option_name}`:""}`.slice(0,18),
      value:r.reorder_expected_daily_sales||0,
    }))
  ,[data]);

  const topReorder=useMemo(()=>
    [...data].sort((a,b)=>(b.reorder_recommended_qty||0)-(a.reorder_recommended_qty||0)).slice(0,5).map(r=>({
      name:`${r.reorder_product_name||""}${r.reorder_option_name?` / ${r.reorder_option_name}`:""}`.slice(0,18),
      value:r.reorder_recommended_qty||0,
    }))
  ,[data]);

  const rising=useMemo(()=>
    [...data].filter(r=>(r.reorder_trend_ratio||0)>=1.2).sort((a,b)=>(b.reorder_trend_ratio||0)-(a.reorder_trend_ratio||0)).slice(0,5).map(r=>({
      name:`${r.reorder_product_name||""}${r.reorder_option_name?` / ${r.reorder_option_name}`:""}`.slice(0,18),
      value:Math.round((r.reorder_trend_ratio||0)*100)/100,
    }))
  ,[data]);

  const downloadCSV=(source)=>{
    const target=source||filtered;
    const cols=[
      {label:"상품코드",get:r=>r.reorder_product_code||"",align:"left"},
      {label:"상품명",get:r=>r.reorder_product_name,align:"left"},
      {label:"옵션",get:r=>r.reorder_option_name,align:"left"},
      {label:"가용재고",get:r=>(r.reorder_available_stock||0).toLocaleString(),align:"center"},
      {label:"입고대기",get:r=>(r.reorder_incoming_stock||0).toLocaleString(),align:"center"},
      {label:"실질 가용재고",get:r=>(r.reorder_effective_stock||0).toLocaleString(),align:"center",bold:true},
      {label:"1주 판매",get:r=>(r.reorder_weekly_sales||0).toLocaleString(),align:"center"},
      {label:"4주 판매",get:r=>(r.reorder_monthly_sales||0).toLocaleString(),align:"center"},
      {label:"예상 일판매",get:r=>(r.reorder_expected_daily_sales||0).toFixed(2),align:"center"},
      {label:"판매 추세",get:r=>trendTag(r).label,align:"center",colorBy:r=>trendTag(r).color},
      {label:"재고잔여일",get:r=>(r.reorder_days_left||0).toFixed(1)+"일",align:"center",bold:true,colorBy:r=>(r.reorder_days_left||0)<7?"#C87B7B":"#C8A87B"},
      {label:"추천 리오더",get:r=>(r.reorder_recommended_qty||0).toLocaleString(),align:"center",bold:true},
    ];
    const esc=s=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    const headerHtml=cols.map(c=>
      `<th style="background:#D4F0E1;color:#1a1a1a;border:1px solid #c0d6cb;padding:6px 10px;font-weight:600;text-align:${c.align};font-family:-apple-system,'Apple SD Gothic Neo',sans-serif;font-size:13px;">${esc(c.label)}</th>`
    ).join("");
    const bodyHtml=target.map(r=>"<tr>"+cols.map(c=>{
      const v=c.get(r);
      const color=c.colorBy?c.colorBy(r):"#1a1a1a";
      const fw=c.bold?700:400;
      return `<td style="border:1px solid #e0e0da;padding:5px 10px;text-align:${c.align};color:${color};font-weight:${fw};font-family:-apple-system,'Apple SD Gothic Neo',sans-serif;font-size:13px;">${esc(v)}</td>`;
    }).join("")+"</tr>").join("");
    const html=`<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"/></head><body><table style="border-collapse:collapse;"><thead><tr>${headerHtml}</tr></thead><tbody>${bodyHtml}</tbody></table></body></html>`;
    const blob=new Blob(["﻿"+html],{type:"application/vnd.ms-excel;charset=utf-8;"});
    const a=Object.assign(document.createElement("a"),{href:URL.createObjectURL(blob),download:`reorder_${new Date().toISOString().slice(0,10)}.xls`});
    a.click();URL.revokeObjectURL(a.href);
  };

  const rowKey=r=>r.reorder_id||`${r.reorder_product_name||""}__${r.reorder_option_name||""}`;
  const filteredKeys=useMemo(()=>filtered.map(rowKey),[filtered]);
  const allFilteredSelected=filteredKeys.length>0&&filteredKeys.every(k=>selected.has(k));
  const someFilteredSelected=filteredKeys.some(k=>selected.has(k));
  const toggleRow=k=>setSelected(p=>{const s=new Set(p);s.has(k)?s.delete(k):s.add(k);return s;});
  const toggleAllFiltered=()=>setSelected(p=>{
    const s=new Set(p);
    if(allFilteredSelected) filteredKeys.forEach(k=>s.delete(k));
    else filteredKeys.forEach(k=>s.add(k));
    return s;
  });
  const downloadSelected=()=>downloadCSV(filtered.filter(r=>selected.has(rowKey(r))));

  const trendTag=r=>{const t=r.reorder_trend_ratio||0;return t>=1.2?{label:"↑ 상승",color:"#7EC8A4"}:t>=0.8?{label:"→ 안정",color:"#7B9EC8"}:{label:"↓ 감소",color:"#C87B7B"};};

  const textCols=new Set(["reorder_product_code","reorder_product_name","reorder_option_name"]);
  const SortTh=({k,label})=>(
    <th onClick={()=>{if(sortKey===k)setSortDir(d=>d==="asc"?"desc":"asc");else{setSortKey(k);setSortDir("asc");setPg(0);}}}
      style={{padding:"4px 6px",textAlign:textCols.has(k)?"left":"center",fontWeight:600,color:DC.text,borderBottom:`1px solid ${DC.border}`,
        fontSize:12,whiteSpace:"nowrap",cursor:"pointer",userSelect:"none"}}>
      {label}{sortKey===k?(sortDir==="asc"?" ↑":" ↓"):""}
    </th>
  );

  const chartStyle={background:DC.bg,border:`1px solid ${DC.border}`,borderRadius:9,padding:"14px 12px"};
  const ttStyle={contentStyle:{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:7,fontSize:14},cursor:false};

  return(
    <div style={{marginTop:16,background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"20px 20px 28px"}}>
      <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:4,flexWrap:"wrap"}}>
        <span style={{fontWeight:600,fontSize:18,color:DC.text,letterSpacing:"-0.2px"}}>리오더 계산기</span>
        {latestDataDate&&<span style={{fontSize:12,color:DC.sub,marginLeft:4}}>· 기준일 {latestDataDate}</span>}
      </div>
      <div style={{fontSize:13,color:DC.sub,marginBottom:20}}>최근 판매량과 현재 재고를 기반으로 자동 리오더 필요 SKU를 분석합니다.</div>

      {/* Calculation flow card */}
      <div style={{marginBottom:20,background:DC.bg,border:`1px solid ${DC.border}`,borderRadius:10,padding:"14px 18px"}}>
        <div style={{fontSize:12,fontWeight:700,color:DC.text,letterSpacing:".06em",marginBottom:12}}>계산 기준 — 14일 재고 커버</div>
        <div style={{display:"flex",gap:0,alignItems:"center",flexWrap:"wrap"}}>
          {[
            {title:"판매속도",body:"(1주판매÷7)×70% + (4주판매÷28)×30%"},
            {title:"실질 가용재고",body:"가용재고 + 입고대기"},
            {title:"예상 재고잔여일",body:"실질가용재고 ÷ 예상일판매량"},
            {title:"14일 미만 → 리오더",body:"days_left < 14일"},
            {title:"추천 리오더 수량",body:"(일판매량×14) − 실질가용재고"},
          ].map((s,i,a)=>(
            <React.Fragment key={s.title}>
              <div style={{background:"rgba(0,0,0,0.03)",borderRadius:8,padding:"10px 14px",textAlign:"center",minWidth:120}}>
                <div style={{fontSize:12,color:DC.text,marginBottom:5,fontWeight:600}}>{s.title}</div>
                <div style={{fontSize:12,color:DC.sub,lineHeight:1.6}}>{s.body}</div>
              </div>
              {i<a.length-1&&<div style={{color:DC.text,fontSize:18,padding:"0 6px",flexShrink:0}}>→</div>}
            </React.Fragment>
          ))}
        </div>
      </div>

      {loading&&<div style={{textAlign:"center",padding:"40px 0",color:DC.text,fontSize:15}}>데이터 로딩 중...</div>}

      {!loading&&data.length===0&&(
        <div style={{textAlign:"center",padding:"40px 0",color:DC.text,fontSize:15,lineHeight:2}}>
          Inventory Trend 엑셀 업로드 완료 후 리오더 데이터가 자동 생성됩니다.<br/>
          <span style={{fontSize:13}}>엑셀에 <strong style={{color:DC.sub}}>가용재고 · 입고대기 · 1주발주합계 · 4주발주합계</strong> 컬럼 포함 필요</span>
        </div>
      )}

      {!loading&&data.length>0&&(
        <>
          {/* KPI */}
          {kpi&&(
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:20}}>
              {[
                {label:"리오더 추천 SKU",value:`${kpi.n}개`,color:"#C87B7B"},
                {label:"평균 재고잔여일",value:`${kpi.avgDays.toFixed(1)}일`,color:"#C8A87B"},
                {label:"총 추천 리오더 수량",value:`${kpi.totalQty.toLocaleString()}개`,color:DC.text},
                {label:"총 입고대기 수량",value:`${kpi.totalIncoming.toLocaleString()}개`,color:"#7B9EC8"},
              ].map(c=>(
                <div key={c.label} style={{background:DC.bg,border:`1px solid ${DC.border}`,borderRadius:8,padding:"13px 15px"}}>
                  <div style={{fontSize:12,color:DC.sub,marginBottom:5}}>{c.label}</div>
                  <div style={{fontSize:18,fontWeight:700,color:c.color,letterSpacing:"-0.3px"}}>{c.value}</div>
                </div>
              ))}
            </div>
          )}

          {/* Charts */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:20}}>
            <div style={chartStyle}>
              <div style={{fontSize:14,fontWeight:600,color:DC.text,marginBottom:12}}>판매 상승 SKU Top5 (추세비율)</div>
              <ResponsiveContainer width="100%" height={140}>
                <BarChart data={rising} layout="vertical" margin={{top:0,right:8,bottom:0,left:0}}>
                  <CartesianGrid strokeDasharray="2 4" stroke={DC.border} horizontal={false}/>
                  <XAxis type="number" tick={{fill:DC.text,fontSize:12}} axisLine={{stroke:DC.border}} tickLine={false}/>
                  <YAxis dataKey="name" type="category" tick={{fill:DC.text,fontSize:11}} axisLine={false} tickLine={false} width={140}/>
                  <Tooltip {...ttStyle}/>
                  <Bar dataKey="value" name="추세비율" fill="#7EC8A4" radius={[0,3,3,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={chartStyle}>
              <div style={{fontSize:14,fontWeight:600,color:DC.text,marginBottom:12}}>판매속도 Top5 (예상 일판매량)</div>
              <ResponsiveContainer width="100%" height={140}>
                <BarChart data={topSales} layout="vertical" margin={{top:0,right:8,bottom:0,left:0}}>
                  <CartesianGrid strokeDasharray="2 4" stroke={DC.border} horizontal={false}/>
                  <XAxis type="number" tick={{fill:DC.text,fontSize:12}} axisLine={{stroke:DC.border}} tickLine={false} tickFormatter={v=>v.toFixed(1)}/>
                  <YAxis dataKey="name" type="category" tick={{fill:DC.text,fontSize:11}} axisLine={false} tickLine={false} width={140}/>
                  <Tooltip {...ttStyle} formatter={v=>[v.toFixed(2),"일판매량"]}/>
                  <Bar dataKey="value" name="일판매량" fill="#7B9EC8" radius={[0,3,3,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={chartStyle}>
              <div style={{fontSize:14,fontWeight:600,color:DC.text,marginBottom:12}}>추천 리오더 수량 Top5</div>
              <ResponsiveContainer width="100%" height={140}>
                <BarChart data={topReorder} layout="vertical" margin={{top:0,right:8,bottom:0,left:0}}>
                  <CartesianGrid strokeDasharray="2 4" stroke={DC.border} horizontal={false}/>
                  <XAxis type="number" tick={{fill:DC.text,fontSize:12}} axisLine={{stroke:DC.border}} tickLine={false}/>
                  <YAxis dataKey="name" type="category" tick={{fill:DC.text,fontSize:11}} axisLine={false} tickLine={false} width={140}/>
                  <Tooltip {...ttStyle}/>
                  <Bar dataKey="value" name="추천리오더" fill="#C8A87B" radius={[0,3,3,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Table */}
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:8}}>
            <input placeholder="상품코드 / 상품명 / 옵션 검색" value={search} onChange={e=>{setSearch(e.target.value);setPg(0);}}
              style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,
                padding:"5px 10px",fontSize:13,color:DC.text,minWidth:180,outline:"none",fontFamily:"inherit"}}/>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <span style={{fontSize:13,color:DC.text}}>{filtered.length.toLocaleString()}개 SKU</span>
              {selected.size>0&&(
                <>
                  <span style={{fontSize:12,color:DC.sub}}>선택 {selected.size}</span>
                  <button onClick={()=>setSelected(new Set())}
                    style={{background:"transparent",color:DC.sub,border:`1px solid ${DC.border}`,borderRadius:5,
                      padding:"4px 10px",fontSize:13,cursor:"pointer"}}>선택 해제</button>
                  <button onClick={downloadSelected}
                    style={{background:"#7EC8A4",color:"#fff",border:"1px solid #7EC8A4",borderRadius:5,
                      padding:"4px 12px",fontSize:13,cursor:"pointer",fontWeight:600}}>↓ 선택 다운로드</button>
                </>
              )}
              <button onClick={()=>downloadCSV()}
                style={{background:"transparent",color:"#7EC8A4",border:"1px solid #7EC8A4",borderRadius:5,
                  padding:"4px 12px",fontSize:13,cursor:"pointer"}}>↓ 전체 다운로드</button>
            </div>
          </div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,tableLayout:"auto"}}>
              <thead style={{position:"sticky",top:0,background:DC.card,zIndex:2}}>
                <tr>
                  <th style={{padding:"4px 6px",borderBottom:`1px solid ${DC.border}`,width:28}}>
                    <input type="checkbox" checked={allFilteredSelected}
                      ref={el=>{if(el) el.indeterminate=!allFilteredSelected&&someFilteredSelected;}}
                      onChange={toggleAllFiltered} style={{cursor:"pointer"}}/>
                  </th>
                  <SortTh k="reorder_product_code" label="상품코드"/>
                  <SortTh k="reorder_product_name" label="상품명"/>
                  <SortTh k="reorder_option_name" label="옵션"/>
                  <SortTh k="reorder_available_stock" label="가용재고"/>
                  <SortTh k="reorder_incoming_stock" label="입고대기"/>
                  <SortTh k="reorder_effective_stock" label="실질 가용재고"/>
                  <SortTh k="reorder_weekly_sales" label="1주 판매"/>
                  <SortTh k="reorder_monthly_sales" label="4주 판매"/>
                  <SortTh k="reorder_expected_daily_sales" label="예상 일판매"/>
                  <SortTh k="reorder_trend_ratio" label="판매 추세"/>
                  <SortTh k="reorder_days_left" label="재고잔여일"/>
                  <SortTh k="reorder_recommended_qty" label="추천 리오더"/>
                </tr>
              </thead>
              <tbody>
                {paged.map((r,i)=>{
                  const trend=trendTag(r);
                  const urgent=(r.reorder_days_left||0)<7;
                  const k=rowKey(r);
                  const isSel=selected.has(k);
                  return(
                    <tr key={r.reorder_id||i} style={{borderBottom:`1px solid ${DC.border}`,background:isSel?"rgba(126,200,164,0.08)":"transparent"}}>
                      <td style={{padding:"4px 6px",textAlign:"center"}}>
                        <input type="checkbox" checked={isSel} onChange={()=>toggleRow(k)} style={{cursor:"pointer"}}/>
                      </td>
                      <td style={{padding:"4px 6px",color:DC.sub,fontFamily:"monospace",fontSize:12,maxWidth:90,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.reorder_product_code}>{r.reorder_product_code||"—"}</td>
                      <td style={{padding:"4px 6px",color:DC.text,fontWeight:500,maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.reorder_product_name}</td>
                      <td style={{padding:"4px 6px",color:DC.sub,maxWidth:90,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.reorder_option_name||"—"}</td>
                      <td style={{padding:"4px 6px",color:DC.sub,textAlign:"center"}}>{(r.reorder_available_stock||0).toLocaleString()}</td>
                      <td style={{padding:"4px 6px",color:DC.sub,textAlign:"center"}}>{(r.reorder_incoming_stock||0).toLocaleString()}</td>
                      <td style={{padding:"4px 6px",color:DC.text,textAlign:"center",fontWeight:500}}>{(r.reorder_effective_stock||0).toLocaleString()}</td>
                      <td style={{padding:"4px 6px",color:DC.sub,textAlign:"center"}}>{(r.reorder_weekly_sales||0).toLocaleString()}</td>
                      <td style={{padding:"4px 6px",color:DC.sub,textAlign:"center"}}>{(r.reorder_monthly_sales||0).toLocaleString()}</td>
                      <td style={{padding:"4px 6px",color:DC.text,textAlign:"center"}}>{(r.reorder_expected_daily_sales||0).toFixed(2)}</td>
                      <td style={{padding:"4px 6px",textAlign:"center"}}><span style={{fontSize:12,fontWeight:600,color:trend.color}}>{trend.label}</span></td>
                      <td style={{padding:"4px 6px",textAlign:"center",fontWeight:700,color:urgent?"#C87B7B":"#C8A87B"}}>{(r.reorder_days_left||0).toFixed(1)}일</td>
                      <td style={{padding:"4px 6px",textAlign:"center",fontWeight:700,color:DC.text}}>{(r.reorder_recommended_qty||0).toLocaleString()}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {totalPgs>1&&(
            <div style={{display:"flex",justifyContent:"center",gap:5,marginTop:14}}>
              {Array.from({length:totalPgs}).map((_,i)=>(
                <button key={i} onClick={()=>setPg(i)}
                  style={{background:pg===i?DC.text:"transparent",color:pg===i?DC.card:DC.sub,
                    border:`1px solid ${pg===i?DC.text:DC.border}`,borderRadius:5,padding:"3px 10px",fontSize:13,cursor:"pointer"}}>
                  {i+1}
                </button>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// INVENTORY TREND WRAPPER
// ─────────────────────────────────────────────
function InventoryTrend({DC,onReorderRefresh}){
  const [snapshotDates,setSnapshotDates]=useState([]);
  const [refreshKey,setRefreshKey]=useState(0);
  const [agingDate,setAgingDate]=useState(null);

  const loadDates=useCallback(async()=>{
    const db=await getSupabase();
    const{data}=await db.from("inventory_snapshot").select("snapshot_date").order("snapshot_date",{ascending:false});
    if(data) setSnapshotDates([...new Set(data.map(r=>r.snapshot_date))]);
  },[]);

  useEffect(()=>{loadDates();},[loadDates]);

  const onUploaded=useCallback(()=>{loadDates();setRefreshKey(k=>k+1);},[loadDates]);
  const onReorderDone=useCallback(()=>{if(onReorderRefresh) onReorderRefresh();},[onReorderRefresh]);
  const agingTrendSecRef=useRef(null);

  return(
    <div style={{marginTop:16,background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"20px 20px 28px"}}>
      <div style={{fontWeight:600,fontSize:14,color:DC.text,letterSpacing:"-0.2px",marginBottom:16}}>Inventory Trend</div>

      <InventoryUploader DC={DC} onUploaded={onUploaded} onReorderDone={onReorderDone}/>

      {/* SKU Risk Bubble */}
      <div style={{marginTop:32,paddingTop:24,borderTop:`1px solid ${DC.border}`}}>
        <div style={{fontWeight:600,fontSize:13,color:DC.text,marginBottom:16,letterSpacing:"-0.1px"}}>SKU Risk Bubble</div>
        <InvBubblePlot DC={DC} snapshotDates={snapshotDates} stopRef={agingTrendSecRef}/>
      </div>

      {/* Aging Trend */}
      <div ref={agingTrendSecRef} style={{marginTop:32,paddingTop:24,borderTop:`1px solid ${DC.border}`}}>
        <div style={{display:"flex",alignItems:"baseline",gap:12,marginBottom:16,flexWrap:"wrap"}}>
          <span style={{fontWeight:600,fontSize:13,color:DC.text,letterSpacing:"-0.1px"}}>Aging Trend</span>
          {agingDate&&<span style={{fontSize:13,color:DC.text}}>· 기준일 {agingDate}</span>}
          <span style={{fontSize:13,color:DC.text}}>재고 에이징은 마지막 판매일 이후 경과일을 기준으로 재고 건강도를 구간별로 추적하는 지표입니다.</span>
        </div>
        <InvAgingTrend DC={DC} snapshotDates={snapshotDates} refreshKey={refreshKey} onDateReady={setAgingDate}/>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA COMPARE
// ─────────────────────────────────────────────

const SKU_VOL_PASTEL={
  "자사몰":"#A8D8B8",
  "29CM":"#C87878",
  "무신사":"#C4B8E8",
  "오프라인 스토어":"#7AA8C8",
  common:"#E0CDB6",
  cross:"#C8AC8E",
};

function ActiveSkuVolume({orders=[],storeSales=[],DC}){
  const cardRef=useRef(null);
  const [aggUnit,setAggUnit]=useState("month");
  const [period,setPeriod]=useState("all");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [calOpenFor,setCalOpenFor]=useState(null);
  const [show4Jeolgi,setShow4Jeolgi]=useState(false);
  const [skuModal,setSkuModal]=useState(null);
  const [modalTab,setModalTab]=useState("common");
  const [selectedPK,setSelectedPK]=useState(null); // 요약 카드/인사이트 기준 월
  const [highlightKey,setHighlightKey]=useState(null); // 범례 클릭 하이라이트

  // 입절기 근사 날짜 (±1일 오차 허용)
  const SOLAR_TERMS=[
    {name:"입춘",mmdd:"02-04"},
    {name:"입하",mmdd:"05-06"},
    {name:"입추",mmdd:"08-07"},
    {name:"입동",mmdd:"11-07"},
  ];

  const {chartRows,channelOrder,jeolgiLines}=useMemo(()=>{
    const ONLINE_CHS=["자사몰","29CM","무신사"];
    const OFFLINE_CH="오프라인 스토어";
    const ALL_CHS=[...ONLINE_CHS,OFFLINE_CH];

    // Date range for current period setting
    let filterStart=null,filterEnd=null;
    if(period==="custom"&&customStart&&customEnd){
      filterStart=customStart;filterEnd=customEnd;
    } else if(period!=="all"){
      const d=new Date();
      if(period==="3m") d.setMonth(d.getMonth()-3);
      else if(period==="6m") d.setMonth(d.getMonth()-6);
      else if(period==="1y") d.setFullYear(d.getFullYear()-1);
      filterStart=[d.getFullYear(),String(d.getMonth()+1).padStart(2,"0"),String(d.getDate()).padStart(2,"0")].join("-");
      filterEnd=new Date().toISOString().slice(0,10);
    }

    const inRange=(dateStr)=>{
      if(!dateStr) return false;
      if(!filterStart) return true;
      return dateStr>=filterStart&&dateStr<=filterEnd;
    };

    const SEASONS=["봄","여름","가을","겨울"];

    const getPK=(dateStr)=>{
      if(!dateStr) return null;
      const d=new Date(dateStr);
      if(aggUnit==="week"){
        const jan1=new Date(d.getFullYear(),0,1);
        const dow=jan1.getDay();
        const weekNum=Math.ceil((Math.floor((d-jan1)/86400000)+dow+1)/7);
        return `${d.getFullYear()}-W${String(weekNum).padStart(2,"0")}`;
      }
      if(aggUnit==="quarter"){
        const q=Math.ceil((d.getMonth()+1)/3);
        return `${d.getFullYear()}-Q${q}`;
      }
      if(aggUnit==="jeolgi"){
        // 입춘 2/4 · 입하 5/6 · 입추 8/7 · 입동 11/7 기준 4계절 분기
        const m=d.getMonth()+1, day=d.getDate();
        let s, y=d.getFullYear();
        if((m===2&&day>=4)||m===3||m===4||(m===5&&day<=5)) s=1;
        else if((m===5&&day>=6)||m===6||m===7||(m===8&&day<=6)) s=2;
        else if((m===8&&day>=7)||m===9||m===10||(m===11&&day<=6)) s=3;
        else { s=4; if(m===1||(m===2&&day<=3)) y=y-1; }
        return `${y}-S${s}`;
      }
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
    };

    const getLbl=(pk)=>{
      if(aggUnit==="week"){const[y,w]=pk.split("-W");return `${y} W${w}`;}
      if(aggUnit==="quarter"){const[y,q]=pk.split("-Q");return `${y} Q${q}`;}
      if(aggUnit==="jeolgi"){const[y,s]=pk.split("-S");return `${y} ${SEASONS[parseInt(s)-1]}`;}
      const[y,m]=pk.split("-");return `${y}.${parseInt(m)}`;
    };

    const periodMap={};
    // 상품명 끝의 [컬러] 토큰은 제거하여 컬러 옵션은 1 SKU로 합산
    const normName=name=>String(name||"").replace(/\s*\[[^\]]*\]\s*$/,"").trim();

    const addSku=(pk,ch,name)=>{
      const n=normName(name);
      if(!pk||!n) return;
      if(!periodMap[pk]) periodMap[pk]={};
      if(!periodMap[pk][ch]) periodMap[pk][ch]=new Set();
      periodMap[pk][ch].add(n);
    };

    orders.forEach(o=>{
      if(!ONLINE_CHS.includes(o.channel)||o.status!=="배송") return;
      if(!inRange(o.order_date)) return;
      addSku(getPK(o.order_date),o.channel,o.product_name);
    });
    storeSales.forEach(s=>{
      if(!inRange(s.sale_date)) return;
      addSku(getPK(s.sale_date),OFFLINE_CH,s.product_name);
    });

    const sortedPKs=Object.keys(periodMap).sort();
    if(!sortedPKs.length) return{chartRows:[],channelOrder:ALL_CHS,solarTermRef:null};

    const gEx={};
    ALL_CHS.forEach(ch=>{gEx[ch]={only:0,total:0};});

    const processed=sortedPKs.map(pk=>{
      const chSets=periodMap[pk];
      const activeChs=ALL_CHS.filter(ch=>chSets[ch]?.size>0);
      const allSkus=new Set();
      activeChs.forEach(ch=>chSets[ch].forEach(s=>allSkus.add(s)));

      const commonSkus=new Set();  // 모든 활성 채널에 공통
      const crossSkus=new Set();   // 2~(n-1) 채널 교차
      const onlySkus={};
      activeChs.forEach(ch=>{onlySkus[ch]=new Set();});

      allSkus.forEach(sku=>{
        const inChs=activeChs.filter(ch=>chSets[ch].has(sku));
        if(inChs.length===activeChs.length&&activeChs.length>=2) commonSkus.add(sku);
        else if(inChs.length>=2) crossSkus.add(sku);
        else onlySkus[inChs[0]].add(sku);
      });

      activeChs.forEach(ch=>{
        gEx[ch].only+=(onlySkus[ch]?.size||0);
        gEx[ch].total+=(chSets[ch]?.size||0);
      });
      return{pk,label:getLbl(pk),chSets,commonSkus,crossSkus,onlySkus,allSkus,activeChs};
    });

    const channelOrder=[...ALL_CHS].sort((a,b)=>{
      const exA=gEx[a].total?gEx[a].only/gEx[a].total:0;
      const exB=gEx[b].total?gEx[b].only/gEx[b].total:0;
      return exA-exB;
    });

    const chartRows=processed.map(p=>({
      pk:p.pk,
      label:p.label,
      common:p.commonSkus.size,
      cross:p.crossSkus.size,
      _allCnt:p.allSkus.size,
      _commonSkus:[...p.commonSkus].sort(),
      _crossSkus:[...p.crossSkus].sort(),
      _chSets:Object.fromEntries(Object.entries(p.chSets).map(([k,v])=>[k,[...v].sort()])),
      _onlySkus:Object.fromEntries(Object.entries(p.onlySkus).map(([k,v])=>[k,[...v].sort()])),
      _activeChs:p.activeChs,
      ...Object.fromEntries(ALL_CHS.map(ch=>[`${ch}_only`,p.onlySkus[ch]?.size||0])),
    }));

    // 4절기 기준선 — 차트 데이터 연도 범위 내 모든 입절기 날짜
    let jeolgiLines=[];
    if(show4Jeolgi&&chartRows.length){
      const years=new Set();
      chartRows.forEach(r=>{
        const y=r.label.match(/^(\d{4})/)?.[1];
        if(y){years.add(parseInt(y));}
      });
      years.forEach(y=>{
        SOLAR_TERMS.forEach(({name,mmdd})=>{
          const date=`${y}-${mmdd}`;
          const pk=getPK(date);
          const lbl=getLbl(pk);
          if(chartRows.some(r=>r.label===lbl)){
            jeolgiLines.push({name,date,label:lbl});
          }
        });
      });
      jeolgiLines.sort((a,b)=>a.date.localeCompare(b.date));
    }

    return{chartRows,channelOrder,jeolgiLines};
  },[orders,storeSales,aggUnit,period,customStart,customEnd,show4Jeolgi]);

  const SkuTooltip=useCallback(({active,payload,label})=>{
    if(!active||!payload?.length) return null;
    const dp=payload[0]?.payload||{};
    const total=dp._allCnt||0;
    const pct=(n,d)=>d?(n/d*100).toFixed(1):"0.0";
    return(
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:8,padding:"10px 14px",minWidth:220,boxShadow:"0 4px 16px rgba(0,0,0,.08)",fontSize:12}}>
        <div style={{fontWeight:700,marginBottom:2,color:DC.text}}>{label}</div>
        <div style={{color:DC.sub,marginBottom:8,fontSize:11,borderBottom:`1px solid ${DC.border}`,paddingBottom:6}}>
          Active SKU: <b style={{color:DC.text}}>{total}개</b>
          <span style={{marginLeft:6,fontWeight:400}}>(공통 {dp.common||0} · 교차 {dp.cross||0} · only 합계 {total-(dp.common||0)-(dp.cross||0)})</span>
        </div>
        {dp.common>0&&(
          <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:5}}>
            <span style={{width:10,height:10,borderRadius:2,background:SKU_VOL_PASTEL.common,display:"inline-block",flexShrink:0}}/>
            <span style={{color:DC.text}}>공통 (전 채널)</span>
            <span style={{marginLeft:"auto",fontWeight:700,color:DC.text}}>{dp.common}개</span>
            <span style={{color:DC.sub,fontSize:11}}>({pct(dp.common,total)}%)</span>
          </div>
        )}
        {dp.cross>0&&(
          <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:5}}>
            <span style={{width:10,height:10,borderRadius:2,background:SKU_VOL_PASTEL.cross,display:"inline-block",flexShrink:0}}/>
            <span style={{color:DC.text}}>교차 (2–3채널)</span>
            <span style={{marginLeft:"auto",fontWeight:700,color:DC.text}}>{dp.cross}개</span>
            <span style={{color:DC.sub,fontSize:11}}>({pct(dp.cross,total)}%)</span>
          </div>
        )}
        {channelOrder.map(ch=>{
          const only=dp[`${ch}_only`]||0;
          const chActive=(dp._chSets?.[ch]?.length||0);
          if(!chActive) return null;
          return(
            <div key={ch} style={{marginBottom:5}}>
              <div style={{display:"flex",alignItems:"center",gap:5}}>
                <span style={{width:10,height:10,borderRadius:2,background:SKU_VOL_PASTEL[ch],display:"inline-block",flexShrink:0}}/>
                <span style={{color:DC.text}}>{ch} only</span>
                <span style={{marginLeft:"auto",fontWeight:700,color:DC.text}}>{only}개</span>
                <span style={{color:DC.sub,fontSize:11}}>({pct(only,chActive)}%)</span>
              </div>
              <div style={{marginLeft:15,fontSize:10,color:DC.dim,marginTop:1}}>
                {ch} 활성 {chActive}개 중 단독 판매
              </div>
            </div>
          );
        })}
        <div style={{borderTop:`1px solid ${DC.border}`,marginTop:4,paddingTop:5,fontSize:10,color:DC.dim}}>클릭하면 SKU 목록을 볼 수 있습니다</div>
      </div>
    );
  },[channelOrder,DC]);

  const stackKeys=["common","cross",...channelOrder.map(ch=>`${ch}_only`)];
  const stackColors={common:SKU_VOL_PASTEL.common,cross:SKU_VOL_PASTEL.cross,...Object.fromEntries(channelOrder.map(ch=>[`${ch}_only`,SKU_VOL_PASTEL[ch]]))};
  const stackLabels={common:"공통 (전 채널)",cross:"교차 (2–3채널)",...Object.fromEntries(channelOrder.map(ch=>[`${ch}_only`,`${ch} only`]))};

  const openModal=(dp)=>{
    if(!dp) return;
    const firstTab=dp.common>0?"common":dp.cross>0?"cross":((dp._activeChs||[])[0]||"common");
    setModalTab(firstTab);
    setSkuModal(dp);
  };

  const modalTabs=skuModal?[
    ...(skuModal.common>0?[{key:"common",label:`공통 (${skuModal.common})`}]:[]),
    ...(skuModal.cross>0?[{key:"cross",label:`교차 (${skuModal.cross})`}]:[]),
    ...channelOrder
      .filter(ch=>(skuModal._onlySkus?.[ch]?.length||0)>0)
      .map(ch=>({key:ch,label:`${ch} only (${skuModal._onlySkus?.[ch]?.length||0})`})),
  ]:[];

  const modalOnlySkus=skuModal&&!["common","cross"].includes(modalTab)?(skuModal._onlySkus?.[modalTab]||[]):[];

  const PERIOD_PRESETS=[["all","전체"],["3m","3개월"],["6m","6개월"],["1y","1년"]];

  // 선택된 월(없으면 마지막) — 요약 카드/인사이트 기준
  const activeRow=useMemo(()=>{
    if(!chartRows.length) return null;
    if(selectedPK){
      const f=chartRows.find(r=>r.pk===selectedPK);
      if(f) return f;
    }
    return chartRows[chartRows.length-1];
  },[chartRows,selectedPK]);

  // 직전 월 (MoM 비교)
  const prevRow=useMemo(()=>{
    if(!activeRow||!chartRows.length) return null;
    const idx=chartRows.findIndex(r=>r.pk===activeRow.pk);
    return idx>0?chartRows[idx-1]:null;
  },[chartRows,activeRow]);

  // 총 SKU 라인 차트용 데이터 (전월 대비 증감 포함)
  const totalLineData=useMemo(()=>chartRows.map((r,i)=>({
    pk:r.pk,label:r.label,total:r._allCnt,
    delta:i>0?r._allCnt-chartRows[i-1]._allCnt:null,
  })),[chartRows]);

  // 인사이트 자동 생성
  const insightText=useMemo(()=>{
    if(!activeRow) return "";
    if(!(activeRow._allCnt>0)) return "";
    const prevUnit=aggUnit==="week"?"전주":aggUnit==="quarter"?"전분기":aggUnit==="jeolgi"?"전 절기":"전월";
    // 채널별 활성 SKU 수 (해당 채널에서 판매된 모든 SKU)
    const chTotals=channelOrder.map(ch=>({ch,n:activeRow._chSets?.[ch]?.length||0}));
    const topCh=chTotals.filter(c=>c.n>0).sort((a,b)=>b.n-a.n)[0];
    let s="";
    if(topCh){
      s=`${activeRow.label} 기준, 활성 SKU가 가장 많은 채널은 ${topCh.ch} (${topCh.n.toLocaleString()}개)입니다.`;
    }
    if(prevRow){
      const drops=channelOrder.map(ch=>{
        const cur=activeRow._chSets?.[ch]?.length||0;
        const prev=prevRow._chSets?.[ch]?.length||0;
        return{ch,d:cur-prev,cur,prev};
      }).filter(c=>c.d<0).sort((a,b)=>a.d-b.d);
      if(drops.length){
        const c=drops[0];
        s+=` ${prevUnit} 대비 감소폭이 가장 큰 채널은 ${c.ch} (▼${Math.abs(c.d).toLocaleString()}개, ${c.prev.toLocaleString()} → ${c.cur.toLocaleString()})입니다.`;
      } else {
        s+=` ${prevUnit} 대비 모든 채널이 유지 또는 증가했습니다.`;
      }
    }
    return s;
  },[activeRow,prevRow,channelOrder,aggUnit]);

  // 채널 카드용 전체 채널(공통/교차 + 4채널 only)
  const summaryCards=useMemo(()=>{
    if(!activeRow) return [];
    const total=activeRow._allCnt||1;
    const pct=v=>(v/total*100).toFixed(1)+"%";
    return [
      {label:"전체",val:total,pct:"",color:"#1a1a1a",bg:"#fafaf7"},
      {label:"공통",val:activeRow.common||0,pct:pct(activeRow.common||0),color:SKU_VOL_PASTEL.common,bg:SKU_VOL_PASTEL.common+"22"},
      {label:"교차",val:activeRow.cross||0,pct:pct(activeRow.cross||0),color:SKU_VOL_PASTEL.cross,bg:SKU_VOL_PASTEL.cross+"22"},
      ...channelOrder.map(ch=>({
        label:`${ch} only`,
        val:activeRow[`${ch}_only`]||0,
        pct:pct(activeRow[`${ch}_only`]||0),
        color:SKU_VOL_PASTEL[ch],bg:SKU_VOL_PASTEL[ch]+"22",
      })),
    ];
  },[activeRow,channelOrder]);

  return(
    <div ref={cardRef} style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"20px 20px 24px",marginTop:16}}>
      {/* 헤더 */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"baseline",gap:10,flexWrap:"wrap",flex:1}}>
          <span style={{fontWeight:600,fontSize:16,color:DC.text}}>Active SKU 분석</span>
          <span style={{fontSize:12,color:DC.sub}}>채널별 실효 SKU 분포 · 비중·교집합·단독 구성</span>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
          {/* 분석 단위 (주/월/분기/절기) */}
          <div style={{display:"flex",gap:3}}>
            {[["week","주"],["month","월"],["quarter","분기"],["jeolgi","절기"]].map(([u,lbl])=>(
              <button key={u} data-hf onClick={()=>setAggUnit(u)}
                style={{background:aggUnit===u?DC.text:"transparent",color:aggUnit===u?"#fff":DC.sub,
                  border:`1px solid ${aggUnit===u?DC.text:DC.border}`,
                  borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontWeight:600,transition:"all .12s"}}>
                {lbl}
              </button>
            ))}
          </div>
          <span style={{color:DC.border}}>|</span>
          {/* 기간 필터 — CalDrop */}
          <CalDrop id="skuVol" period={period}
            setPeriod={v=>{setPeriod(v);if(v!=="custom"){setCustomStart("");setCustomEnd("");}}}
            presets={PERIOD_PRESETS}
            start={customStart} setStart={setCustomStart}
            end={customEnd} setEnd={setCustomEnd}
            calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}
            dark={false}/>
          <span style={{color:DC.border}}>|</span>
          <CaptureBtn cardRef={cardRef} filename="Active_SKU_볼륨" DC={DC}/>
        </div>
      </div>

      {/* 가이드 */}
      <div style={{background:"#f5f5f3",borderRadius:6,padding:"10px 14px",marginBottom:14,fontSize:11,color:DC.sub,lineHeight:2,display:"flex",flexDirection:"column",gap:1}}>
        <div><span style={{fontWeight:700,color:DC.text,marginRight:6}}>SKU 기준</span>상품명 단위로 집계합니다. 옵션(색상·사이즈)이 다르더라도 같은 상품명이면 1 SKU로 계산합니다.</div>
        <div><span style={{fontWeight:700,color:DC.text,marginRight:6}}>Active SKU 정의</span>온라인(자사몰·29CM·무신사)은 해당 기간 내 배송 완료 건이 1건 이상인 SKU, 오프라인 스토어는 매장 판매 데이터에 등록된 SKU(반품 포함)를 기준으로 합니다.</div>
        <div><span style={{fontWeight:700,color:DC.text,marginRight:6}}>공통 SKU</span>모든 활성 채널에서 동시에 판매된 SKU입니다. 2–3개 채널에만 걸친 SKU는 <b>교차</b>로 분리되며, 1개 채널에서만 판매된 SKU는 해당 채널의 <b>단독(only)</b>으로 분류됩니다. 스택 순서는 채널 간 단독 비율을 기준으로 자동 정렬됩니다.</div>
      </div>

      {/* 채널 요약 카드 (선택 월 기준) */}
      {activeRow&&(
        <div style={{marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
            <span style={{fontSize:11,fontWeight:600,color:DC.sub,letterSpacing:".06em"}}>요약 기준</span>
            <select value={activeRow.pk} onChange={e=>setSelectedPK(e.target.value)}
              style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:6,padding:"4px 10px",
                fontFamily:"ui-monospace,SFMono-Regular,Menlo,monospace",fontSize:13,fontWeight:700,color:DC.text,
                cursor:"pointer",outline:"none",minWidth:110}}>
              {chartRows.map(r=>(<option key={r.pk} value={r.pk}>{r.label}</option>))}
            </select>
            {selectedPK&&selectedPK!==chartRows[chartRows.length-1]?.pk&&
              <button data-hf onClick={()=>setSelectedPK(null)}
                style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,padding:"3px 9px",fontSize:11,cursor:"pointer",color:DC.sub}}>최신으로</button>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:`repeat(${summaryCards.length},minmax(0,1fr))`,gap:8}}>
            {summaryCards.map(c=>(
              <div key={c.label} style={{background:c.bg,border:`1px solid ${DC.border}`,borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:10,color:DC.sub,marginBottom:4,letterSpacing:".04em",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{c.label}</div>
                <div style={{display:"flex",alignItems:"baseline",gap:5}}>
                  <span style={{fontFamily:"ui-monospace,SFMono-Regular,Menlo,monospace",fontWeight:700,fontSize:18,color:c.color}}>{c.val.toLocaleString()}</span>
                  {c.pct&&<span style={{fontSize:10,color:DC.sub}}>{c.pct}</span>}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 범례 (클릭 시 하이라이트) */}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14}}>
        {[
          {key:"common",label:"공통 (전 채널)",color:SKU_VOL_PASTEL.common,icon:"●"},
          {key:"cross",label:"교차 (2–3채널)",color:SKU_VOL_PASTEL.cross,icon:"◆"},
          ...channelOrder.map((ch,i)=>({key:`${ch}_only`,label:`${ch} only`,color:SKU_VOL_PASTEL[ch],icon:["▲","■","▼","◀"][i%4]})),
        ].map(item=>{
          const dimmed=highlightKey&&highlightKey!==item.key;
          return(
            <button key={item.key} data-hf
              onClick={()=>setHighlightKey(highlightKey===item.key?null:item.key)}
              style={{display:"flex",alignItems:"center",gap:5,fontSize:12,color:DC.text,
                background:highlightKey===item.key?item.color+"33":"transparent",
                border:`1px solid ${highlightKey===item.key?item.color:DC.border}`,
                borderRadius:5,padding:"3px 8px",cursor:"pointer",
                opacity:dimmed?0.4:1,transition:"all .12s"}}>
              <span style={{color:item.color,fontSize:11,fontWeight:700}}>{item.icon}</span>
              <span style={{width:16,height:7,borderRadius:2,background:item.color,display:"inline-block"}}/>
              {item.label}
            </button>
          );
        })}
        {highlightKey&&<button data-hf onClick={()=>setHighlightKey(null)}
          style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,padding:"3px 8px",fontSize:11,cursor:"pointer",color:DC.sub}}>해제</button>}
      </div>

      {/* 메인 차트: 100% 스택 (비중) */}
      {!chartRows.length?(
        <div style={{textAlign:"center",padding:"60px 0",color:DC.sub,fontSize:14}}>
          해당 기간에 데이터가 없습니다
        </div>
      ):(
        <>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={chartRows} margin={{top:32,right:8,bottom:8,left:0}} barCategoryGap="22%" stackOffset="expand"
            onClick={({activePayload})=>{
              if(!activePayload?.length) return;
              const pk=activePayload[0].payload?.pk;
              if(pk) setSelectedPK(pk);
            }}>
            <CartesianGrid strokeDasharray="3 3" stroke={DC.border} vertical={false}/>
            <XAxis dataKey="label" tick={{fontSize:11,fill:DC.sub}} tickLine={false} axisLine={false}/>
            <YAxis tickFormatter={v=>`${Math.round(v*100)}%`} tick={{fontSize:11,fill:DC.sub}} tickLine={false} axisLine={false} width={42}/>
            <Tooltip content={<SkuTooltip/>} cursor={{fill:"rgba(0,0,0,.04)"}}/>
            {stackKeys.map(k=>(
              <Bar key={k} dataKey={k} stackId="1"
                fill={stackColors[k]} name={stackLabels[k]}
                fillOpacity={highlightKey&&highlightKey!==k?0.18:1}
                radius={k===stackKeys[stackKeys.length-1]?[3,3,0,0]:[0,0,0,0]}/>
            ))}
            {activeRow&&<ReferenceLine x={activeRow.label} stroke="#C8927B" strokeWidth={1.5} strokeDasharray="2 4"/>}
            {jeolgiLines.map(jl=>(
              <ReferenceLine key={jl.date} x={jl.label} stroke="#8899AA" strokeDasharray="4 3" strokeWidth={1.5}
                label={{value:`${jl.name} ${jl.date.slice(5).replace("-","/")}`,
                  position:"top",fontSize:9,fill:"#8899AA",fontWeight:700}}/>
            ))}
          </BarChart>
        </ResponsiveContainer>

        {/* 보조 라인 차트: 총 Active SKU 추이 */}
        <div style={{marginTop:18,paddingTop:14,borderTop:`1px solid ${DC.border}`}}>
          <div style={{display:"flex",alignItems:"baseline",gap:8,marginBottom:6}}>
            <span style={{fontSize:11,fontWeight:600,color:DC.sub,letterSpacing:".06em"}}>총 Active SKU 추이</span>
            {prevRow&&activeRow&&(
              <span style={{fontFamily:"ui-monospace,SFMono-Regular,Menlo,monospace",fontSize:11,color:DC.sub}}>
                {activeRow.label}: <b style={{color:DC.text}}>{activeRow._allCnt}</b>
                {(()=>{const d=activeRow._allCnt-prevRow._allCnt;return d===0?<span style={{marginLeft:5,color:DC.dim}}>(±0)</span>:
                  <span style={{marginLeft:5,color:d>0?"#7EC8A4":"#C87B7B"}}>({d>0?"▲":"▼"}{Math.abs(d)})</span>;})()}
              </span>
            )}
          </div>
          <ResponsiveContainer width="100%" height={120}>
            <LineChart data={totalLineData} margin={{top:8,right:8,bottom:0,left:0}}
              onClick={({activePayload})=>{
                if(!activePayload?.length) return;
                const pk=activePayload[0].payload?.pk;
                if(pk) setSelectedPK(pk);
              }}>
              <CartesianGrid strokeDasharray="3 3" stroke={DC.border} vertical={false}/>
              <XAxis dataKey="label" tick={{fontSize:10,fill:DC.sub}} tickLine={false} axisLine={false}/>
              <YAxis tick={{fontSize:10,fill:DC.sub}} tickLine={false} axisLine={false} width={36}/>
              <Tooltip contentStyle={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:6,fontSize:12}}
                formatter={(v,_n,p)=>{const d=p?.payload?.delta;return[`${v}개${d!=null?` (${d>=0?"▲":"▼"}${Math.abs(d)})`:""}`,"총 SKU"];}}/>
              <Line type="monotone" dataKey="total" stroke="#C8927B" strokeWidth={2}
                dot={{r:3,fill:"#fff",stroke:"#C8927B",strokeWidth:1.5}}
                activeDot={{r:5,style:{cursor:"pointer"}}}/>
              {activeRow&&<ReferenceLine x={activeRow.label} stroke="#C8927B" strokeWidth={1} strokeDasharray="2 4"/>}
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* 자동 인사이트 */}
        {insightText&&(
          <div style={{marginTop:14,padding:"12px 14px",background:"#fafaf7",border:`1px solid ${DC.border}`,borderRadius:8,fontSize:12,color:DC.text,lineHeight:1.7}}>
            <span style={{fontSize:10,fontWeight:700,color:"#C8927B",letterSpacing:".08em",marginRight:6}}>INSIGHT</span>
            {insightText}
          </div>
        )}
        </>
      )}

      {/* SKU 목록 모달 */}
      {skuModal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center"}}
          onClick={()=>setSkuModal(null)}>
          <div style={{background:"#fff",borderRadius:14,padding:"24px 28px",maxWidth:520,width:"92%",maxHeight:"75vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,.18)"}}
            onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:4,color:"#111"}}>{skuModal.label} — SKU 목록</div>
            <div style={{fontSize:11,color:"#888",marginBottom:14}}>전체 고유 SKU: {skuModal._allCnt}개 · 탭 선택 후 확인</div>
            <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:14}}>
              {modalTabs.map(t=>(
                <button key={t.key} onClick={()=>setModalTab(t.key)}
                  style={{background:modalTab===t.key?"#111":"transparent",color:modalTab===t.key?"#fff":"#555",
                    border:"1.5px solid",borderColor:modalTab===t.key?"#111":"#ddd",
                    borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontWeight:600,transition:"all .12s"}}>
                  {t.label}
                </button>
              ))}
            </div>
            {modalTab==="cross"&&(
              <div style={{fontSize:11,color:"#888",marginBottom:8}}>2~3개 채널에서 판매됐으나 전 채널 공통은 아닌 SKU</div>
            )}
            {!["common","cross"].includes(modalTab)&&(
              <div style={{fontSize:11,color:"#888",marginBottom:8}}>
                {modalTab}에서만 판매된 단독 SKU {modalOnlySkus.length}개
              </div>
            )}
            <div style={{overflowY:"auto",flex:1,borderTop:"1px solid #f0f0f0",paddingTop:8}}>
              {modalTab==="common"&&(skuModal._commonSkus||[]).map((s,i)=>(
                <div key={i} style={{padding:"5px 4px",borderBottom:"1px solid #f5f5f5",fontSize:13,color:"#222"}}>{s}</div>
              ))}
              {modalTab==="cross"&&(skuModal._crossSkus||[]).map((s,i)=>(
                <div key={i} style={{padding:"5px 4px",borderBottom:"1px solid #f5f5f5",fontSize:13,color:"#222"}}>{s}</div>
              ))}
              {!["common","cross"].includes(modalTab)&&(
                <>
                  {modalOnlySkus.length>0&&(
                    <>
                      <div style={{fontSize:11,fontWeight:600,color:"#555",marginBottom:4,marginTop:4}}>단독 SKU ({modalOnlySkus.length}개)</div>
                      {modalOnlySkus.map((s,i)=>(
                        <div key={i} style={{padding:"5px 4px",borderBottom:"1px solid #f5f5f5",fontSize:13,color:"#222"}}>{s}</div>
                      ))}
                    </>
                  )}
                </>
              )}
            </div>
            <button onClick={()=>setSkuModal(null)}
              style={{marginTop:16,padding:"8px 20px",background:"#111",color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontSize:13,fontWeight:600,alignSelf:"flex-end"}}>
              닫기
            </button>
          </div>
        </div>
      )}

      {/* 해석 가이드 + 각주 */}
      <div style={{marginTop:16,borderTop:`1px solid ${DC.border}`,paddingTop:14}}>
        <div style={{fontSize:11,fontWeight:700,color:DC.text,marginBottom:6}}>그래프 해석 방법</div>
        <div style={{fontSize:11,color:DC.sub,lineHeight:1.9,display:"flex",flexDirection:"column",gap:2}}>
          <div><span style={{display:"inline-block",width:8,height:8,borderRadius:2,background:SKU_VOL_PASTEL.common,marginRight:5,verticalAlign:"middle"}}/>
            <b style={{color:DC.text}}>공통 (전 채널) 바가 두껍다</b> — 브랜드 핵심 상품이 모든 채널에서 고르게 판매 중. 채널 간 SKU 전략이 일관적임.
          </div>
          <div><span style={{display:"inline-block",width:8,height:8,borderRadius:2,background:SKU_VOL_PASTEL.cross,marginRight:5,verticalAlign:"middle"}}/>
            <b style={{color:DC.text}}>교차 바가 크다</b> — 일부 채널에서만 공유되는 SKU가 많음. 채널별 기획 상품 운영 또는 단계적 채널 확장이 진행 중일 가능성.
          </div>
          <div><span style={{display:"inline-block",width:8,height:8,borderRadius:2,background:SKU_VOL_PASTEL["자사몰"],marginRight:5,verticalAlign:"middle"}}/>
            <b style={{color:DC.text}}>특정 채널 only 바가 크다</b> — 해당 채널 전용 기획이 활발하거나, 다른 채널에 미도입된 신상품이 집중된 상태.
          </div>
          <div style={{marginTop:2}}>
            <b style={{color:DC.text}}>바 전체 높이</b>가 기간 내 전체 고유 Active SKU 수입니다. 높이 증가는 라인 확대, 감소는 단종·시즌 오프를 의미합니다.
            시간 흐름에 따라 <b style={{color:DC.text}}>공통 비율이 높아지면</b> 채널 간 SKU 정합성이 향상되는 추세이며,
            <b style={{color:DC.text}}> only 비율이 커지면</b> 채널별 차별화 전략이 강화되고 있음을 나타냅니다.
          </div>
        </div>
        <div style={{marginTop:12,fontSize:11,color:DC.dim,lineHeight:1.8}}>
          시간의 흐름에 따라 판매되는 실제 SKU 수 파악을 통해 다음 시즌 상품 기획에 참고 데이터로 활용될 수 있으며, 각 판매처 간 실효 SKU량을 가늠할 수 있습니다.
        </div>
      </div>
    </div>
  );
}

const COMPARE_CH_COLOR={
  "자사몰":"#6EBF99",
  "29CM":"#F5C8A0",
  "오프라인 스토어":"#A8C8E0",
  "무신사":"#C4B8E8",
};
const COMPARE_CHANNELS=["자사몰","29CM","오프라인 스토어","무신사"];

function RevenueSankeyChart({periods,svgW}){
  const wrapRef=useRef(null);
  const [hoveredCh,setHoveredCh]=useState(null);
  const [selNodes,setSelNodes]=useState([]);   // max 2 [{key,pi,ch,amt,label}]
  const [modal,setModal]=useState(null);       // {x,y,a,b}
  const [orderWarn,setOrderWarn]=useState(false);

  const SVG_H=480,PAD_T=70,PAD_B=52,PAD_H=28,NODE_W=40,GAP=3,AVAIL_H=SVG_H-PAD_T-PAD_B;
  const CH_LABEL_W=52; // first-col label area

  const maxTotal=Math.max(...periods.map(p=>p.total),1);
  const heightScale=AVAIL_H/maxTotal;

  const cols=useMemo(()=>periods.map((p,pi)=>{
    const colX=periods.length===1?(svgW-NODE_W)/2
      :PAD_H+CH_LABEL_W+pi*(svgW-2*PAD_H-CH_LABEL_W-NODE_W)/(periods.length-1);
    let y=PAD_T;
    const nodes=COMPARE_CHANNELS.map(ch=>{
      const amt=p.byChannel[ch]||0;
      const h=Math.max(0,amt*heightScale);
      const n={ch,amt,x:colX,y:amt>0?y:0,h,color:COMPARE_CH_COLOR[ch]};
      if(amt>0) y+=h+GAP;
      return n;
    });
    return{...p,colX,nodes};
  }),[periods,svgW,heightScale]);

  // 채널별 첫 번째 데이터가 있는 컬럼 인덱스
  const firstVisPi=useMemo(()=>{
    const map={};
    COMPARE_CHANNELS.forEach(ch=>{
      const idx=cols.findIndex(col=>(col.nodes.find(n=>n.ch===ch)?.h||0)>=14);
      if(idx>=0) map[ch]=idx;
    });
    return map;
  },[cols]);

  // 베지어 곡선 리본
  const links=useMemo(()=>{
    const res=[];
    for(let pi=0;pi<cols.length-1;pi++){
      COMPARE_CHANNELS.forEach((ch,ci)=>{
        const ln=cols[pi].nodes.find(n=>n.ch===ch);
        const rn=cols[pi+1].nodes.find(n=>n.ch===ch);
        if(!ln||!rn||ln.h<1||rn.h<1) return;
        const x1=ln.x+NODE_W,x2=rn.x,mx=(x1+x2)/2;
        const path=[
          `M${x1} ${ln.y}C${mx} ${ln.y},${mx} ${rn.y},${x2} ${rn.y}`,
          `L${x2} ${rn.y+rn.h}C${mx} ${rn.y+rn.h},${mx} ${ln.y+ln.h},${x1} ${ln.y+ln.h}Z`,
        ].join(" ");
        res.push({ch,ci,path,color:COMPARE_CH_COLOR[ch]});
      });
    }
    return res;
  },[cols]);

  const fmtAmt=a=>{
    if(a>=1e8){
      const eok=Math.floor(a/1e8);
      const cheon=Math.floor((a%1e8)/1e7);
      return eok+"억"+(cheon>0?cheon+"천만":"");
    }
    if(a>=1e4) return Math.round(a/1e4)+"만";
    return a.toLocaleString();
  };

  const handleNodeClick=(e,pi,ch,amt,label)=>{
    e.stopPropagation();
    const rect=wrapRef.current?.getBoundingClientRect();
    const px=Math.round(e.clientX-(rect?.left||0));
    const py=Math.round(e.clientY-(rect?.top||0));
    const key=`${pi}__${ch}`;
    setSelNodes(prev=>{
      const already=prev.findIndex(n=>n.key===key);
      if(already>=0){
        const next=prev.filter(n=>n.key!==key);
        if(next.length<2) setModal(null);
        return next;
      }
      const node={key,pi,ch,amt,label};
      if(prev.length>=2){setModal(null);return[node];}
      const next=[...prev,node];
      if(next.length===2){
        if(next[0].pi>next[1].pi){
          // 미래→과거 순서: 경고 후 첫 번째 선택만 유지
          setOrderWarn(true);
          setTimeout(()=>setOrderWarn(false),2500);
          return[next[1]];
        }
        const mw=200,mh=140;
        const cx=Math.min(Math.max(px-mw/2,4),(svgW||600)-mw-4);
        const cy=Math.max(py-mh-12,4);
        setModal({x:cx,y:cy,a:next[0],b:next[1]});
      }
      return next;
    });
  };

  const isSel=(pi,ch)=>selNodes.some(n=>n.pi===pi&&n.ch===ch);
  const selIdx=(pi,ch)=>selNodes.findIndex(n=>n.pi===pi&&n.ch===ch);

  return(
    <div ref={wrapRef} style={{position:"relative"}}
      onClick={()=>{setSelNodes([]);setModal(null);setOrderWarn(false);}}>
      <svg width={svgW} height={SVG_H} style={{overflow:"visible",display:"block"}}
        onClick={()=>{setSelNodes([]);setModal(null);setOrderWarn(false);}}>
        <rect x={0} y={0} width={svgW} height={SVG_H} fill="transparent"/>
        <defs>
          {COMPARE_CHANNELS.map((ch,ci)=>{
            const id=`cg2_${ci}`;
            const c=COMPARE_CH_COLOR[ch];
            return(
              <linearGradient key={ch} id={id} x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stopColor={c} stopOpacity={0.75}/>
                <stop offset="50%" stopColor={c} stopOpacity={0.45}/>
                <stop offset="100%" stopColor={c} stopOpacity={0.75}/>
              </linearGradient>
            );
          })}
        </defs>

        {/* 직각 리본 */}
        {links.map((l,i)=>(
          <path key={i} d={l.path}
            fill={`url(#cg2_${l.ci})`}
            opacity={hoveredCh===null?0.65:hoveredCh===l.ch?0.9:0.04}
            style={{transition:"opacity .15s",cursor:"default"}}
            onMouseEnter={()=>setHoveredCh(l.ch)}
            onMouseLeave={()=>setHoveredCh(null)}
          />
        ))}

        {/* 노드 + 라벨 */}
        {cols.map((col,pi)=>(
          <g key={pi}>
            {col.nodes.map((n,ni)=>n.h>=1&&(
              <g key={ni} style={{cursor:"pointer"}}
                onClick={e=>handleNodeClick(e,pi,n.ch,n.amt,col.label)}
                onMouseEnter={()=>setHoveredCh(n.ch)}
                onMouseLeave={()=>setHoveredCh(null)}>
                {/* 하단 모서리만 둥근 노드 */}
                {(()=>{
                  const r=Math.min(3,n.h/2,NODE_W/2);
                  const{x:nx,y:ny,h:nh}=n;
                  const d=[
                    `M${nx} ${ny}`,
                    `H${nx+NODE_W}`,
                    `V${ny+nh-r}`,
                    `Q${nx+NODE_W} ${ny+nh} ${nx+NODE_W-r} ${ny+nh}`,
                    `H${nx+r}`,
                    `Q${nx} ${ny+nh} ${nx} ${ny+nh-r}`,
                    `Z`,
                  ].join(" ");
                  return(
                    <path d={d} fill={n.color}
                      opacity={hoveredCh===null?1:hoveredCh===n.ch?1:0.18}
                      stroke={isSel(pi,n.ch)?"#fff":"none"} strokeWidth={2}
                      style={{transition:"opacity .15s"}}/>
                  );
                })()}
                {/* 선택 번호 뱃지 */}
                {isSel(pi,n.ch)&&(
                  <text x={n.x+NODE_W/2} y={n.y+n.h/2} textAnchor="middle"
                    dominantBaseline="middle" fontSize={11} fontWeight={800} fill="#fff"
                    style={{pointerEvents:"none",userSelect:"none"}}>
                    {selIdx(pi,n.ch)+1}
                  </text>
                )}
                {/* 채널별 첫 번째 유효 노드 안에 채널명 삽입 */}
                {firstVisPi[n.ch]===pi&&n.h>=14&&(
                  <text x={n.x+NODE_W/2} y={n.y+n.h/2}
                    textAnchor="middle" dominantBaseline="middle"
                    fontSize={9} fontWeight={700} fill="#fff"
                    style={{pointerEvents:"none",userSelect:"none"}}>
                    {n.ch==="오프라인 스토어"?"오프라인":n.ch}
                  </text>
                )}
              </g>
            ))}
            {/* 기간 라벨 */}
            <text x={col.colX+NODE_W/2} y={SVG_H-PAD_B+18}
              textAnchor="middle" fontSize={10} fill="#111111" style={{pointerEvents:"none"}}>
              {col.label}
            </text>
            {/* 매출 합계 라벨 — 노드 스택 상단 */}
            {col.total>0&&(()=>{
              const vis=col.nodes.filter(n=>n.h>0);
              if(!vis.length) return null;
              const top=Math.min(...vis.map(n=>n.y));
              return(
                <text x={col.colX+NODE_W/2} y={top-7}
                  textAnchor="middle" dominantBaseline="auto"
                  fontSize={13} fontWeight={700} fill="#111111"
                  style={{pointerEvents:"none",userSelect:"none"}}>
                  {fmtAmt(col.total)}
                </text>
              );
            })()}
          </g>
        ))}

        {/* 컬럼 간 증감률 (가로 텍스트, 양수·음수 모두 표시) */}
        {cols.slice(0,-1).map((col,pi)=>{
          const next=cols[pi+1];
          if(!col.total||!next.total) return null;
          const pct=((next.total-col.total)/col.total*100);
          const up=pct>=0;
          const gapMidX=(col.colX+NODE_W+next.colX)/2;
          const y=PAD_T/2;
          return(
            <g key={`gr_${pi}`} style={{pointerEvents:"none"}}>
              <text x={gapMidX} y={y}
                textAnchor="middle" dominantBaseline="middle"
                fontSize={11} fontWeight={700} fill={up?"#7dbf9e":"#c97b7b"} style={{userSelect:"none"}}>
                {up?"▲":"▼"} {Math.abs(pct).toFixed(1)}%
              </text>
            </g>
          );
        })}
      </svg>

      {/* 비교 모달 */}
      {modal&&(()=>{
        const{x,y,a,b}=modal;
        const diff=b.amt-a.amt;
        const pct=a.amt>0?((diff/a.amt)*100):null;
        const up=diff>=0;
        return(
          <div onClick={e=>e.stopPropagation()}
            style={{position:"absolute",left:x,top:y,
              background:"#ffffff",border:"1px solid #d8d8d0",borderRadius:10,
              padding:"14px 16px 12px",minWidth:192,
              boxShadow:"0 4px 20px rgba(0,0,0,0.12)",zIndex:20,pointerEvents:"auto"}}>
            <button onClick={()=>{setSelNodes([]);setModal(null);}}
              style={{position:"absolute",top:7,right:9,background:"none",border:"none",
                color:"#111",cursor:"pointer",fontSize:15,lineHeight:1}}>✕</button>
            <div style={{fontSize:10,color:"#111",marginBottom:10,letterSpacing:"0.08em",textTransform:"uppercase"}}>매출 비교</div>
            {[a,b].map((nd,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <span style={{width:18,height:18,borderRadius:3,background:COMPARE_CH_COLOR[nd.ch]||"#444",
                  display:"inline-flex",alignItems:"center",justifyContent:"center",
                  fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>{i+1}</span>
                <div>
                  <div style={{fontSize:10,color:"#444"}}>{nd.label} · {nd.ch}</div>
                  <div style={{fontSize:13,color:"#111",fontWeight:700}}>{fmtWonShort(nd.amt)}</div>
                </div>
              </div>
            ))}
            <div style={{borderTop:"1px solid #e8e8e0",marginTop:8,paddingTop:8,display:"flex",alignItems:"baseline",gap:8}}>
              <span style={{fontSize:15,fontWeight:800,color:up?"#2a9a60":"#c0392b"}}>
                {up?"▲":"▼"} {pct!==null?`${Math.abs(pct).toFixed(1)}%`:"—"}
              </span>
              <span style={{fontSize:11,color:"#111"}}>
                ({up?"+":""}{fmtWonShort(Math.abs(diff))})
              </span>
            </div>
          </div>
        );
      })()}


      {/* 순서 경고 */}
      {orderWarn&&(
        <div style={{marginTop:10,textAlign:"center",fontSize:13,color:"#f87171",
          fontWeight:600,letterSpacing:"0.02em",userSelect:"none",
          animation:"fadeIn 0.2s ease"}}>
          매출 지점을 과거에서 미래로 선택해주세요
        </div>
      )}

      {/* 사용 안내 */}
      <div style={{marginTop:orderWarn?4:14,textAlign:"center",fontSize:15,color:"#111111",letterSpacing:"0.02em",userSelect:"none"}}>
        노드 매출 지점을 왼쪽에서 오른쪽 순으로 두번 클릭하면 해당 기간의 매출 증감률을 볼 수 있습니다
      </div>
    </div>
  );
}

function VolumeSlider({total,range,onChange,DC}){
  const trackRef=useRef(null);
  const drag=useRef(null);
  const startFrac=total<=1?0:range[0]/(total-1);
  const endFrac=total<=1?1:range[1]/(total-1);

  useEffect(()=>{
    const move=e=>{
      if(!drag.current||!trackRef.current) return;
      const clientX=e.touches?e.touches[0].clientX:e.clientX;
      const rect=trackRef.current.getBoundingClientRect();
      const dIdx=Math.round(((clientX-drag.current.startX)/rect.width)*(total-1));
      const {type,base}=drag.current;
      if(type==="left"){
        const s=Math.max(0,Math.min(base[1]-1,base[0]+dIdx));
        onChange([s,base[1]]);
      } else if(type==="right"){
        const e2=Math.max(base[0]+1,Math.min(total-1,base[1]+dIdx));
        onChange([base[0],e2]);
      } else {
        const span=base[1]-base[0];
        const s=Math.max(0,Math.min(total-1-span,base[0]+dIdx));
        onChange([s,s+span]);
      }
    };
    const up=()=>{drag.current=null;};
    document.addEventListener("mousemove",move);
    document.addEventListener("mouseup",up);
    document.addEventListener("touchmove",move,{passive:false});
    document.addEventListener("touchend",up);
    return()=>{
      document.removeEventListener("mousemove",move);
      document.removeEventListener("mouseup",up);
      document.removeEventListener("touchmove",move);
      document.removeEventListener("touchend",up);
    };
  },[total,onChange]);

  const startDrag=(type,e)=>{
    e.preventDefault();e.stopPropagation();
    const clientX=e.touches?e.touches[0].clientX:e.clientX;
    drag.current={type,startX:clientX,base:[...range]};
  };

  if(total<=1) return null;
  return(
    <div style={{padding:"10px 6px 2px"}}>
      <div ref={trackRef} style={{position:"relative",height:24,userSelect:"none"}}>
        {/* track */}
        <div style={{position:"absolute",top:10,left:0,right:0,height:4,background:DC.border,borderRadius:2}}/>
        {/* filled range — drag to scroll */}
        <div
          onMouseDown={e=>startDrag("body",e)}
          onTouchStart={e=>startDrag("body",e)}
          style={{position:"absolute",top:10,left:`${startFrac*100}%`,
            width:`${(endFrac-startFrac)*100}%`,height:4,
            background:"#7EC8A4",borderRadius:2,cursor:"grab"}}/>
        {/* left handle */}
        <div
          onMouseDown={e=>startDrag("left",e)}
          onTouchStart={e=>startDrag("left",e)}
          style={{position:"absolute",top:5,left:`${startFrac*100}%`,
            transform:"translateX(-50%)",width:14,height:14,
            background:"#fff",border:"2px solid #7EC8A4",borderRadius:"50%",
            cursor:"ew-resize",zIndex:3}}/>
        {/* right handle */}
        <div
          onMouseDown={e=>startDrag("right",e)}
          onTouchStart={e=>startDrag("right",e)}
          style={{position:"absolute",top:5,left:`${endFrac*100}%`,
            transform:"translateX(-50%)",width:14,height:14,
            background:"#fff",border:"2px solid #7EC8A4",borderRadius:"50%",
            cursor:"ew-resize",zIndex:3}}/>
      </div>
    </div>
  );
}

function CaptureBtn({cardRef,filename,DC}){
  const [busy,setBusy]=useState(false);
  const btnRef=useRef(null);
  const feedback=()=>{
    // Android: vibration API
    if(navigator.vibrate) { navigator.vibrate(80); return; }
    // iOS: short click via AudioContext + button pulse
    try{
      const ctx=new (window.AudioContext||window.webkitAudioContext)();
      const osc=ctx.createOscillator();
      const gain=ctx.createGain();
      osc.connect(gain); gain.connect(ctx.destination);
      osc.frequency.value=1000;
      gain.gain.setValueAtTime(0.08,ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+0.05);
      osc.start(); osc.stop(ctx.currentTime+0.05);
    }catch(_){}
    btnRef.current?.animate([{transform:"scale(1)"},{transform:"scale(1.18)"},{transform:"scale(1)"}],{duration:180,easing:"ease-out"});
  };
  const capture=async()=>{
    if(!cardRef?.current||busy) return;
    setBusy(true);
    // Hide all capture buttons inside the card before snapshot
    const btns=cardRef.current.querySelectorAll("[data-capture-hide]");
    btns.forEach(b=>{b._prevVis=b.style.visibility;b.style.visibility="hidden";});
    try{
      const {default:html2canvas}=await import("html2canvas");
      const canvas=await html2canvas(cardRef.current,{scale:2,useCORS:true,backgroundColor:null,logging:false});
      btns.forEach(b=>{b.style.visibility=b._prevVis||"";});
      const fname=`${filename}_${new Date().toISOString().slice(0,10)}.png`;
      const isIOS=/iPhone|iPad|iPod/i.test(navigator.userAgent);
      const isAndroid=/Android/i.test(navigator.userAgent);
      if(isIOS||isAndroid){
        canvas.toBlob(async blob=>{
          const file=new File([blob],fname,{type:"image/png"});
          const blobUrl=URL.createObjectURL(blob);
          // Web Share API: canShare 가드 없이 직접 시도 — iOS 공유 시트 "이미지 저장" → 사진앱
          if(navigator.share){
            try{
              await navigator.share({files:[file],title:fname});
              feedback(); setBusy(false); return;
            }catch(e){
              if(e.name==="AbortError"){ setBusy(false); return; }
              // files 미지원 시 url-only 로 재시도
              try{
                await navigator.share({title:fname,url:blobUrl});
                feedback(); setBusy(false); return;
              }catch(_){}
            }
          }
          // 최종 폴백: 다운로드
          const a=document.createElement("a");a.download=fname;a.href=blobUrl;a.click();
          feedback(); setBusy(false);
        },"image/png");
        return;
      }
      // PC: 로컬 다운로드
      const a=document.createElement("a");
      a.download=fname;a.href=canvas.toDataURL("image/png");a.click();
      feedback();
    }catch(e){btns.forEach(b=>{b.style.visibility=b._prevVis||"";});console.error(e);}
    setBusy(false);
  };
  return(
    <button ref={btnRef} data-capture-hide onClick={capture} disabled={busy} title="카드 이미지 저장"
      style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,
        padding:"3px 8px",fontSize:10,color:DC.sub,cursor:busy?"wait":"pointer",
        display:"flex",alignItems:"center",gap:3,opacity:busy?0.5:1,transition:"opacity .15s",flexShrink:0,fontWeight:400}}>
      <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
        <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
        <circle cx="12" cy="13" r="4"/>
      </svg>
      {busy?"저장중…":"저장"}
    </button>
  );
}

function DataCompare({revenues,storeSales=[],orders=[]}){
  const [volUnit,setVolUnit]=useState("month");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [volPeriod,setVolPeriod]=useState("all");
  const [volCalOpenFor,setVolCalOpenFor]=useState(null);
  const [sliderIdx,setSliderIdx]=useState([0,0]);
  const containerRef=useRef(null);
  const agingTrendSecRef=useRef(null);
  const reorderSecRef=useRef(null);
  const volCardRef=useRef(null);
  const bubbleCardRef=useRef(null);
  const agingCardRef=agingTrendSecRef; // reuse existing ref
  const reorderCardRef=reorderSecRef;  // reuse existing ref
  const [svgW,setSvgW]=useState(760);
  const [reorderKey,setReorderKey]=useState(0);
  const [snapshotDates,setSnapshotDates]=useState([]);
  const [invRefreshKey,setInvRefreshKey]=useState(0);
  const [agingDate,setAgingDate]=useState(null);
  const [reorderDate,setReorderDate]=useState(null);

  const loadSnapshotDates=useCallback(async()=>{
    const db=await getSupabase();
    const{data}=await db.from("inventory_snapshot").select("snapshot_date").order("snapshot_date",{ascending:false});
    if(data) setSnapshotDates([...new Set(data.map(r=>r.snapshot_date))]);
  },[]);
  useEffect(()=>{loadSnapshotDates();},[loadSnapshotDates]);

  useEffect(()=>{
    const obs=new ResizeObserver(es=>setSvgW(Math.max(380,es[0].contentRect.width-48)));
    if(containerRef.current) obs.observe(containerRef.current);
    return()=>obs.disconnect();
  },[]);

  // All available periods from data range to today
  const allVolPeriods=useMemo(()=>{
    const dates=[...revenues.map(r=>r.date),...storeSales.map(r=>r.sale_date)].filter(Boolean).sort();
    const today=new Date();
    const from=dates.length?new Date(dates[0]):new Date(today.getFullYear(),0,1);
    const res=[];
    if(volUnit==="year"){
      for(let y=from.getFullYear();y<=today.getFullYear();y++){
        res.push({label:String(y),start:`${y}-01-01`,end:`${y}-12-31`});
      }
    } else {
      let cur=new Date(from.getFullYear(),from.getMonth(),1);
      while(cur<=today&&res.length<60){
        const e=new Date(cur.getFullYear(),cur.getMonth()+1,0);
        res.push({label:`${cur.getFullYear()}.${cur.getMonth()+1}`,start:cur.toISOString().slice(0,10),end:e.toISOString().slice(0,10)});
        cur=new Date(cur.getFullYear(),cur.getMonth()+1,1);
      }
    }
    return res;
  },[volUnit,revenues,storeSales]);

  // Initialize slider to current year on first load; reset when unit changes
  useEffect(()=>{
    const n=allVolPeriods.length;
    if(!n){setSliderIdx([0,0]);return;}
    const curYear=new Date().getFullYear().toString();
    if(volUnit==="year"){
      const idx=allVolPeriods.findIndex(p=>p.label===curYear);
      if(idx>=0){setSliderIdx([0,n-1]);return;}
    } else {
      const inYear=allVolPeriods.map((p,i)=>({p,i})).filter(({p})=>p.label.startsWith(`${curYear}.`));
      if(inYear.length>0){setSliderIdx([inYear[0].i,inYear[inYear.length-1].i]);return;}
    }
    // fallback if no data for current year
    const def=volUnit==="year"?Math.min(n,5):Math.min(n,12);
    setSliderIdx([Math.max(0,n-def),n-1]);
  },[allVolPeriods]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSlider=useCallback(r=>setSliderIdx(r),[]);

  const volPeriods=useMemo(()=>{
    if(customStart&&customEnd){
      const rangeStart=new Date(customStart),rangeEnd=new Date(customEnd);
      if(rangeStart>rangeEnd) return [];
      const res=[];
      if(volUnit==="year"){
        for(let y=rangeStart.getFullYear();y<=rangeEnd.getFullYear();y++){
          res.push({label:String(y),start:`${y}-01-01`,end:`${y}-12-31`});
        }
      } else {
        let cur=new Date(rangeStart.getFullYear(),rangeStart.getMonth(),1);
        while(cur<=rangeEnd&&res.length<60){
          const e=new Date(cur.getFullYear(),cur.getMonth()+1,0);
          const eC=e>rangeEnd?rangeEnd:e;
          res.push({label:`${cur.getFullYear()}.${cur.getMonth()+1}`,start:cur.toISOString().slice(0,10),end:eC.toISOString().slice(0,10)});
          cur=new Date(cur.getFullYear(),cur.getMonth()+1,1);
        }
      }
      return res;
    }
    return allVolPeriods.slice(sliderIdx[0],sliderIdx[1]+1);
  },[volUnit,customStart,customEnd,allVolPeriods,sliderIdx]);

  const revenueData=useMemo(()=>volPeriods.map(p=>{
    const byChannel={};
    COMPARE_CHANNELS.forEach(ch=>{byChannel[ch]=0;});
    revenues.filter(r=>r.date>=p.start&&r.date<=p.end).forEach(r=>{
      if(COMPARE_CHANNELS.includes(r.channel)) byChannel[r.channel]+=(r.amount||0)-(r.refund_amount||0);
    });
    storeSales.filter(r=>{const d=r.sale_date||"";return d>=p.start&&d<=p.end;}).forEach(r=>{
      if(r.status==="배송") byChannel["오프라인 스토어"]+=(r.amount||0);
      else if(r.status==="반품") byChannel["오프라인 스토어"]-=(r.amount||0);
    });
    COMPARE_CHANNELS.forEach(ch=>{if(byChannel[ch]<0) byChannel[ch]=0;});
    const total=Object.values(byChannel).reduce((a,b)=>a+b,0);
    return{...p,byChannel,total};
  }),[revenues,storeSales,volPeriods]);

  const hasData=revenueData.some(p=>p.total>0);
  const showSlider=!customStart&&!customEnd&&allVolPeriods.length>1;

  const DC={bg:"#f8f8f6",card:"#ffffff",border:"#e0e0da",text:"#111111",sub:"#444444",dim:"#888888"};
  const LC=DC;
  const sectionCard={background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"20px 20px 24px",marginTop:16};

  return(
    <div style={{background:"#f8f8f6",minHeight:"100%",padding:"28px 28px 40px"}}>
      <div style={{fontWeight:700,fontSize:22,color:"#111111",letterSpacing:"-0.3px",marginBottom:24}}>데이터 컴페어</div>

      {/* ① 전체 매출 볼륨 — 다크 카드 */}
      <div ref={volCardRef} style={sectionCard}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
          <div style={{fontWeight:600,fontSize:16,color:DC.text}}>전체 매출 볼륨</div>
          <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
            {[["year","연"],["month","월"]].map(([u,lbl])=>(
              <button key={u} data-hf onClick={()=>{setVolUnit(u);setCustomStart("");setCustomEnd("");}}
                style={{background:volUnit===u?DC.text:"transparent",
                  color:volUnit===u?DC.card:DC.sub,
                  border:`1px solid ${volUnit===u?DC.text:DC.border}`,
                  borderRadius:6,padding:"4px 10px",fontSize:13,
                  cursor:"pointer",fontWeight:600,transition:"all .12s"}}>
                {lbl}
              </button>
            ))}
            <span style={{color:DC.border,fontSize:16,margin:"0 4px"}}>|</span>
            <CalDrop id="vol" period={volPeriod} setPeriod={v=>{setVolPeriod(v);if(v!=="custom"){setCustomStart("");setCustomEnd("");}}}
              presets={[]}
              start={customStart} setStart={setCustomStart}
              end={customEnd} setEnd={setCustomEnd}
              calOpenFor={volCalOpenFor} setCalOpenFor={setVolCalOpenFor}
              dark={true}/>
            <span style={{color:DC.border,margin:"0 2px"}}>|</span>
            <CaptureBtn cardRef={volCardRef} filename="매출볼륨" DC={DC}/>
          </div>
        </div>
        <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:12}}>
          {COMPARE_CHANNELS.map(ch=>(
            <span key={ch} style={{display:"flex",alignItems:"center",gap:5,fontSize:13,color:DC.text}}>
              <span style={{width:8,height:8,background:COMPARE_CH_COLOR[ch],display:"inline-block",flexShrink:0}}/>
              {ch}
            </span>
          ))}
        </div>
        <div ref={containerRef} style={{width:"100%",overflowX:"auto"}}>
          {hasData
            ?<RevenueSankeyChart periods={revenueData} svgW={svgW}/>
            :<div style={{textAlign:"center",padding:"80px 0",color:DC.text,fontSize:15}}>
              매출 데이터를 업로드하면 그래프가 표시됩니다
            </div>
          }
        </div>
        {showSlider&&(
          <div style={{paddingTop:8,borderTop:`1px solid ${DC.border}`,marginTop:8}}>
            <div style={{fontSize:12,color:DC.text,marginBottom:2,textAlign:"right"}}>
              {allVolPeriods[sliderIdx[0]]?.label} ~ {allVolPeriods[sliderIdx[1]]?.label}
              <span style={{marginLeft:8,color:DC.text}}>핸들 드래그로 기간 조정 · 가운데 드래그로 이동</span>
            </div>
            <VolumeSlider total={allVolPeriods.length} range={sliderIdx} onChange={handleSlider} DC={DC}/>
          </div>
        )}
      </div>

      {/* Inventory Uploader — 다크 카드 */}
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"20px 20px 24px",marginTop:16}}>
        <div style={{fontWeight:600,fontSize:16,color:DC.text,letterSpacing:"-0.2px",marginBottom:16}}>Inventory 업로더</div>
        <InventoryUploader DC={DC} onUploaded={()=>{loadSnapshotDates();setInvRefreshKey(k=>k+1);}} onReorderDone={()=>setReorderKey(k=>k+1)}/>
      </div>

      {/* ② SKU Risk Bubble — 다크 카드 */}
      <div ref={bubbleCardRef} style={sectionCard}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:600,fontSize:16,color:DC.text,letterSpacing:"-0.2px"}}>SKU Risk Bubble</div>
          <CaptureBtn cardRef={bubbleCardRef} filename="SKU_Risk_Bubble" DC={DC}/>
        </div>
        <InvBubblePlot DC={DC} snapshotDates={snapshotDates} stopRef={agingTrendSecRef}/>
      </div>

      {/* ③ Aging Trend — 섹션 카드 */}
      <div ref={agingTrendSecRef} style={sectionCard}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
          <div style={{display:"flex",alignItems:"baseline",gap:12,flexWrap:"wrap",flex:1,minWidth:0}}>
            <span style={{fontWeight:600,fontSize:18,color:DC.text,letterSpacing:"-0.2px"}}>Aging Trend</span>
            {agingDate&&<span style={{fontSize:12,color:DC.sub}}>· 기준일 {agingDate}</span>}
            <span style={{fontSize:13,color:DC.sub}}>재고 에이징은 마지막 판매일 이후 경과일을 기준으로 재고 건강도를 구간별로 추적하는 지표입니다.</span>
          </div>
          <CaptureBtn cardRef={agingCardRef} filename="Aging_Trend" DC={DC}/>
        </div>
        <InvAgingTrend DC={DC} snapshotDates={snapshotDates} refreshKey={invRefreshKey} onDateReady={setAgingDate} stopRef={reorderSecRef}/>
      </div>

      {/* ④ Active SKU 볼륨 */}
      <ActiveSkuVolume orders={orders} storeSales={storeSales} DC={DC}/>

      {/* ⑤ 리오더 계산기 (자체 스타일 포함) */}
      <div ref={reorderSecRef} style={{position:"relative"}}>
        <div style={{position:"absolute",top:20,right:20,zIndex:10}}>
          <CaptureBtn cardRef={reorderCardRef} filename="리오더_계산기" DC={DC}/>
        </div>
        <ReorderCalculator DC={DC} refreshKey={reorderKey} onDateReady={setReorderDate} latestSnapDate={snapshotDates.length?[...snapshotDates].sort()[snapshotDates.length-1]:null}/>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// CONTENT IMPACT — 콘텐츠(인스타그램 포스트) × 매출 attribution 캘린더
// ─────────────────────────────────────────────
// Phase 1: 월 셀렉터 + 빈 캘린더 그리드 (스캐폴드)
// 다음 단계: KPI 카드 → 셀별 매출/판매Top → IG 포스트 등록 → 리본
// 인스타그램 embed.js 1회 로드 + 새 임베드 process 호출
function useInstagramEmbedScript(deps){
  useEffect(()=>{
    const SRC="https://www.instagram.com/embed.js";
    let s=document.querySelector(`script[src="${SRC}"]`);
    if(!s){
      s=document.createElement("script");
      s.src=SRC; s.async=true;
      document.body.appendChild(s);
    }
    // 이미 로드된 경우 즉시 process, 아니면 onload 후 process
    const run=()=>{ try{ window.instgrm?.Embeds?.process(); }catch{} };
    if(window.instgrm) run();
    else s.addEventListener("load",run,{once:true});
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },deps);
}

// Microlink JSON 으로 og:image URL 받아옴 (DB 캐시용)
//   - 실패/CORS 에러는 null 반환 → 호출부에서 폴백 처리
async function fetchOgImage(postUrl){
  try{
    const res=await fetch(`https://api.microlink.io/?url=${encodeURIComponent(postUrl)}`);
    const json=await res.json();
    return json?.data?.image?.url||null;
  }catch{ return null; }
}

// 사용자가 업로드한 이미지를 캘린더 셀 픽셀 크기 + 비율(4:5)에 맞춰 center crop + JPEG 압축
// 셀 픽셀 ~268×336 에 가깝게 320×400 (1.2x) 로 저장하여 용량 절약 + 표시 sharp.
// 원본 2~5MB → 10~25KB (Storage 무료 1GB 안에 40,000+ 포스트 가능)
async function resizeImageForUpload(file, targetW=320, targetH=400, quality=0.75) {
  const dataUrl=await new Promise((res,rej)=>{
    const fr=new FileReader();
    fr.onload=()=>res(fr.result); fr.onerror=rej;
    fr.readAsDataURL(file);
  });
  const img=await new Promise((res,rej)=>{
    const i=new Image();
    i.onload=()=>res(i); i.onerror=rej;
    i.src=dataUrl;
  });
  const srcRatio=img.width/img.height;
  const targetRatio=targetW/targetH; // 4:5 = 0.8
  // 원본이 더 가로형이면 좌우 자르고 세로 가득, 더 세로형이면 상하 자르고 가로 가득
  let sw,sh,sx,sy;
  if(srcRatio>targetRatio){
    sh=img.height; sw=img.height*targetRatio;
    sx=(img.width-sw)/2; sy=0;
  } else {
    sw=img.width; sh=img.width/targetRatio;
    sx=0; sy=(img.height-sh)/2;
  }
  const canvas=document.createElement("canvas");
  canvas.width=targetW; canvas.height=targetH;
  canvas.getContext("2d").drawImage(img,sx,sy,sw,sh,0,0,targetW,targetH);
  const blob=await new Promise(res=>canvas.toBlob(res,"image/jpeg",quality));
  return blob;
}

// Supabase Storage 의 ig-thumbs 버킷에 업로드 → public URL 반환
async function uploadIgThumb(postId, blob) {
  const db=await getSupabase();
  const path=`post-${postId}-${Date.now()}.jpg`;
  const {error}=await db.storage.from("ig-thumbs").upload(path,blob,{
    contentType:"image/jpeg", upsert:true,
  });
  if(error) throw new Error(error.message);
  const {data}=db.storage.from("ig-thumbs").getPublicUrl(path);
  return data.publicUrl;
}

// Microlink 가 반환한 IG CDN 이미지 URL 을 받아 4:5 로 center crop 한 뒤 Storage 에 저장.
// CORS 차단(IG CDN 일부) 으로 캔버스 가공이 안 되면 원본 URL 그대로 반환 (폴백).
async function fetchAndCropOgImage(postUrl, postId) {
  const rawUrl=await fetchOgImage(postUrl);
  if(!rawUrl) return null;
  try{
    const img=await new Promise((res,rej)=>{
      const i=new Image();
      i.crossOrigin="anonymous";
      i.onload=()=>res(i); i.onerror=rej;
      i.src=rawUrl;
    });
    const targetW=320, targetH=400, q=0.75;
    const srcRatio=img.width/img.height;
    const targetRatio=targetW/targetH;
    let sw,sh,sx,sy;
    if(srcRatio>targetRatio){ sh=img.height; sw=img.height*targetRatio; sx=(img.width-sw)/2; sy=0; }
    else { sw=img.width; sh=img.width/targetRatio; sx=0; sy=(img.height-sh)/2; }
    const canvas=document.createElement("canvas");
    canvas.width=targetW; canvas.height=targetH;
    canvas.getContext("2d").drawImage(img,sx,sy,sw,sh,0,0,targetW,targetH);
    const blob=await new Promise((res,rej)=>{
      try{ canvas.toBlob(b=>b?res(b):rej(new Error("toBlob null")),"image/jpeg",q); }
      catch(e){ rej(e); }
    });
    return await uploadIgThumb(postId, blob);
  }catch{
    // CORS 차단 또는 캔버스 보안 에러 시 원본 URL 폴백 (셀에서 cover 로 잘림 있지만 표시는 됨)
    return rawUrl;
  }
}

// 인스타그램 포스트 썸네일 — DB에 캐시된 thumb_url 만 사용 (외부 호출 X)
//   thumb_url 누락 포스트는 빈 셀 — IGPostModal '썸네일 새로고침' 버튼으로 보충 가능
function InstagramThumb({ src, objectPosition="center" }) {
  const [err,setErr]=useState(false);
  if(err||!src) return null;
  return (
    <img src={src} alt="" loading="lazy" decoding="async"
      referrerPolicy="no-referrer"
      onError={()=>setErr(true)}
      style={{width:"100%",height:"100%",objectFit:"cover",objectPosition,display:"block"}}
    />
  );
}

// 인스타그램 임베드 — dangerouslySetInnerHTML 로 React reconcile 영역 밖에 두어
// embed.js 가 내부 DOM 을 iframe 으로 교체해도 NotFoundError 가 안 나도록 처리
function InstagramEmbed({ url, style }) {
  const safeUrl=String(url||"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const css=`margin:0;max-width:100%;min-width:100%;${style||""}`;
  return (
    <div dangerouslySetInnerHTML={{__html:
      `<blockquote class="instagram-media" data-instgrm-permalink="${safeUrl}" data-instgrm-version="14" style="${css}">`
      +`<a href="${safeUrl}" target="_blank" rel="noreferrer" style="color:#888;font-size:11px">포스트 보기</a>`
      +`</blockquote>`
    }}/>
  );
}

// URL 정규화: 트래킹 파라미터 제거 + permalink 형태로 표준화
function normalizeIgUrl(raw){
  const s=String(raw||"").trim();
  // 인스타그램 URL 패턴: /p/SHORTCODE/, /reel/SHORTCODE/, /tv/SHORTCODE/
  const m=s.match(/instagram\.com\/(p|reel|tv)\/([A-Za-z0-9_-]+)/i);
  if(!m) return null;
  return `https://www.instagram.com/${m[1]}/${m[2]}/`;
}

// 포스트별 소개 상품 태깅 (chip + 검색 추가)
function ProductTagger({ postId, tagged, allProducts, onChange }){
  const [q,setQ]=useState("");
  const [adding,setAdding]=useState(false);
  const [selected,setSelected]=useState(()=>new Set());

  const toggle=(name)=>{
    setSelected(prev=>{
      const next=new Set(prev);
      if(next.has(name)) next.delete(name); else next.add(name);
      return next;
    });
  };

  const handleBatchAdd=async()=>{
    if(selected.size===0) return;
    const db=await getSupabase();
    const rows=[...selected].map(name=>({post_id:postId,product_name:name}));
    const {error}=await db.from("instagram_post_products").insert(rows);
    if(!error){
      setSelected(new Set());
      setQ("");
      onChange();
      // adding mode 유지 → 검색창 그대로 두고 다른 상품 계속 추가 가능
    }
  };

  const handleRemove=async(name)=>{
    const db=await getSupabase();
    const {error}=await db.from("instagram_post_products").delete()
      .eq("post_id",postId).eq("product_name",name);
    if(!error) onChange();
  };

  const closeAdd=()=>{ setAdding(false); setQ(""); setSelected(new Set()); };

  const candidates=useMemo(()=>{
    if(!q.trim()) return [];
    const qLower=q.toLowerCase();
    return allProducts
      .filter(p=>!tagged.includes(p)&&p.toLowerCase().includes(qLower))
      .slice(0,20);
  },[q,allProducts,tagged]);

  return (
    <div style={{marginTop:10,paddingTop:10,borderTop:`1px dashed ${D.border}`}}>
      <div style={{display:"flex",flexWrap:"wrap",gap:5,alignItems:"center"}}>
        <span style={{fontSize:11,color:D.textMeta,marginRight:4,fontWeight:600}}>소개 상품:</span>
        {tagged.length===0&&!adding&&<span style={{fontSize:11,color:D.textMeta,fontStyle:"italic"}}>(없음)</span>}
        {tagged.map(t=>(
          <span key={t} style={{display:"inline-flex",alignItems:"center",gap:4,
            padding:"3px 9px",background:D.surfaceAlt,border:`1px solid ${D.border}`,
            borderRadius:12,fontSize:11,color:D.text}}>
            {t}
            <button onClick={()=>handleRemove(t)}
              style={{background:"none",border:"none",cursor:"pointer",
                color:D.textMeta,fontSize:11,padding:0,lineHeight:1}}>✕</button>
          </span>
        ))}
        {!adding&&(
          <button onClick={()=>setAdding(true)}
            style={{padding:"3px 9px",fontSize:11,background:"transparent",
              border:`1px dashed ${D.border}`,borderRadius:12,color:D.blue,
              cursor:"pointer"}}>＋ 상품 추가</button>
        )}
      </div>
      {adding&&(
        <div style={{marginTop:7,padding:"8px 10px",background:D.surfaceAlt,
          border:`1px solid ${D.border}`,borderRadius:8}}>
          <div style={{display:"flex",gap:6,alignItems:"center"}}>
            <input autoFocus value={q} onChange={e=>setQ(e.target.value)}
              onKeyDown={e=>{
                if(e.key==="Escape") closeAdd();
                if(e.key==="Enter"&&selected.size>0){ e.preventDefault(); handleBatchAdd(); }
              }}
              placeholder="상품명 검색 (예: 코튼, 데님). 여러 개 체크 후 한 번에 추가"
              style={{flex:1,padding:"6px 10px",fontSize:12,
                border:`1px solid ${D.border}`,borderRadius:6,boxSizing:"border-box",background:D.surface}}/>
            <button onClick={closeAdd}
              style={{padding:"5px 10px",fontSize:11,background:"transparent",
                border:`1px solid ${D.border}`,borderRadius:6,color:D.textMeta,cursor:"pointer"}}>닫기</button>
          </div>
          {candidates.length>0&&(
            <div style={{marginTop:6,background:D.surface,border:`1px solid ${D.border}`,borderRadius:6,
              maxHeight:240,overflowY:"auto"}}>
              {candidates.map(c=>{
                const sel=selected.has(c);
                return (
                  <div key={c} onClick={()=>toggle(c)}
                    style={{padding:"6px 10px",cursor:"pointer",fontSize:12,
                      color:sel?D.blue:D.text,fontWeight:sel?600:400,
                      background:sel?`${D.blue}12`:"transparent",
                      borderBottom:`1px solid ${D.border}`,
                      display:"flex",alignItems:"center",gap:8}}>
                    <input type="checkbox" checked={sel} readOnly
                      style={{cursor:"pointer",accentColor:D.blue,margin:0}}/>
                    <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c}</span>
                  </div>
                );
              })}
            </div>
          )}
          {q.trim()&&candidates.length===0&&(
            <div style={{marginTop:4,fontSize:10,color:D.textMeta}}>
              매칭되는 상품이 없습니다. 주문/매장 데이터에 등록된 상품명만 검색됩니다.
            </div>
          )}
          {selected.size>0&&(
            <div style={{marginTop:8,display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
              <span style={{fontSize:11,color:D.textMeta}}>선택 {selected.size}개</span>
              <button onClick={handleBatchAdd}
                style={{padding:"5px 14px",fontSize:11,background:D.blue,color:"#fff",
                  border:"none",borderRadius:6,cursor:"pointer",fontWeight:600}}>
                ＋ {selected.size}개 추가
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// 포스트 등록·관리 모달
// 사진 직접 업로드 — Microlink 차단/실패에 영향받지 않는 확실한 폴백
//   카메라 롤/갤러리/스크린샷 → 클라이언트 리사이즈 → Supabase Storage 업로드 → thumb_url 저장
function ThumbUploadButton({ post, onChange }) {
  const [state,setState]=useState("idle"); // idle | loading | ok | fail
  const [errMsg,setErrMsg]=useState("");
  const inputRef=useRef(null);
  const onPick=async(e)=>{
    const file=e.target.files?.[0];
    if(!file) return;
    setState("loading"); setErrMsg("");
    try{
      const blob=await resizeImageForUpload(file);
      const url=await uploadIgThumb(post.id, blob);
      const db=await getSupabase();
      const {error}=await db.from("instagram_posts").update({thumb_url:url}).eq("id",post.id);
      if(error) throw new Error(error.message);
      setState("ok");
      onChange();
      setTimeout(()=>setState("idle"),1500);
    }catch(ex){
      setState("fail");
      setErrMsg(String(ex?.message||ex));
    }finally{
      if(inputRef.current) inputRef.current.value="";
    }
  };
  const label=state==="loading"?"업로드 중..."
    :state==="ok"?"✓ 업로드됨"
    :state==="fail"?"실패 — 다시"
    :"📷 사진 업로드";
  return (
    <div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"flex-end"}}>
      <button onClick={()=>inputRef.current?.click()}
        disabled={state==="loading"}
        title="갤러리/카메라 롤/스크린샷에서 사진을 골라 캘린더 셀 배경으로 사용합니다 (자동 리사이즈)"
        style={{background:"transparent",border:`1px solid ${state==="fail"?D.red:D.border}`,borderRadius:5,
          padding:"3px 9px",fontSize:11,cursor:state==="loading"?"wait":"pointer",
          color:state==="fail"?D.red:state==="ok"?D.green:D.textSub,whiteSpace:"nowrap"}}>
        {label}
      </button>
      <input ref={inputRef} type="file" accept="image/*" onChange={onPick} style={{display:"none"}}/>
      {state==="fail"&&errMsg&&(
        <div style={{fontSize:10,color:D.red,maxWidth:220,textAlign:"right"}}>{errMsg}</div>
      )}
    </div>
  );
}

// 썸네일 가져오기/갱신 — Microlink API 호출이 자동 실행 시 실패하는 경우(차단·CORS·일일 한도)
// 수동 재시도용. 미리보기 + 직접 URL 입력 폴백 포함.
function ThumbRefreshButton({ post, onChange }) {
  const [state,setState]=useState("idle"); // idle | loading | ok | fail
  const [manual,setManual]=useState(false);
  const [manualUrl,setManualUrl]=useState(post.thumb_url||"");
  const tryFetch=async()=>{
    setState("loading");
    const thumb=await fetchAndCropOgImage(post.url, post.id);
    if(thumb){
      const db=await getSupabase();
      const {error}=await db.from("instagram_posts").update({thumb_url:thumb}).eq("id",post.id);
      if(!error){ setState("ok"); onChange(); setTimeout(()=>setState("idle"),1500); return; }
    }
    setState("fail");
  };
  const saveManual=async()=>{
    if(!manualUrl.trim()){ setManual(false); return; }
    setState("loading");
    const db=await getSupabase();
    const {error}=await db.from("instagram_posts").update({thumb_url:manualUrl.trim()}).eq("id",post.id);
    if(!error){ setState("ok"); onChange(); setManual(false); setTimeout(()=>setState("idle"),1500); }
    else setState("fail");
  };
  const hasThumb=!!post.thumb_url;
  const label=state==="loading"?"불러오는 중...":state==="ok"?"✓ 저장됨":state==="fail"?"실패 — 직접 입력":(hasThumb?"🔄 썸네일 갱신":"📷 썸네일 가져오기");
  return (
    <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end"}}>
      <button onClick={state==="fail"?()=>setManual(v=>!v):tryFetch}
        disabled={state==="loading"}
        title={hasThumb?`현재 썸네일: ${post.thumb_url}`:"캘린더 셀 배경 이미지를 자동으로 가져옵니다"}
        style={{background:"transparent",border:`1px solid ${state==="fail"?D.red:D.border}`,borderRadius:5,
          padding:"3px 9px",fontSize:11,cursor:state==="loading"?"wait":"pointer",
          color:state==="fail"?D.red:state==="ok"?D.green:D.textSub,whiteSpace:"nowrap"}}>
        {label}
      </button>
      {manual&&(
        <div style={{display:"flex",gap:4,alignItems:"center"}}>
          <input value={manualUrl} onChange={e=>setManualUrl(e.target.value)}
            placeholder="이미지 URL 직접 붙여넣기"
            style={{fontSize:11,padding:"3px 7px",border:`1px solid ${D.border}`,borderRadius:5,width:220}}/>
          <button onClick={saveManual}
            style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
              padding:"3px 9px",fontSize:11,cursor:"pointer"}}>저장</button>
        </div>
      )}
    </div>
  );
}

function IGPostModal({ date, posts, postProductsMap={}, allProducts=[], onClose, onChange }){
  // wizard step: 0=URL 입력, 1=URL 완료(임베드 확인) + 상품 매칭, 2=상품 매칭 완료
  const [step,setStep]=useState(0);
  const [url,setUrl]=useState("");
  const [memo,setMemo]=useState("");
  const [newPostId,setNewPostId]=useState(null);
  const [newPostUrl,setNewPostUrl]=useState("");
  const [saving,setSaving]=useState(false);
  const [err,setErr]=useState("");

  // 임베드 script 로드 + posts·step 변경 시마다 process
  useInstagramEmbedScript([posts.length,posts.map(p=>p.url).join("|"),step,newPostUrl]);

  // Step 0 → 1: URL 저장 (Run)
  const handleRun=async()=>{
    const norm=normalizeIgUrl(url);
    if(!norm){ setErr("올바른 인스타그램 포스트 URL이 아닙니다 (https://www.instagram.com/p/... 또는 /reel/...)"); return; }
    setSaving(true); setErr("");
    const db=await getSupabase();
    const {data,error}=await db.from("instagram_posts")
      .insert({post_date:date,url:norm,caption_memo:memo||null})
      .select().single();
    setSaving(false);
    if(error){ setErr("저장 실패: "+error.message); return; }
    setNewPostId(data.id);
    setNewPostUrl(norm);
    setStep(1);
    onChange(); // 부모 리스트 갱신
    // 백그라운드로 og:image 받아서 4:5 crop + Storage 업로드 후 thumb_url 캐시 (실패 무시)
    fetchAndCropOgImage(norm, data.id).then(thumb=>{
      if(thumb){ db.from("instagram_posts").update({thumb_url:thumb}).eq("id",data.id).then(()=>onChange()); }
    });
  };
  // Step 1 → 2: 상품 매칭 완료 (그냥 step 전환 — 상품은 chip 추가마다 즉시 저장됨)
  const handleComplete=()=>{ setStep(2); };
  // 다시 추가하기 (또는 자동 reset)
  const resetWizard=()=>{
    setStep(0); setUrl(""); setMemo(""); setNewPostId(null); setNewPostUrl(""); setErr("");
  };
  const handleDelete=async(id)=>{
    if(!confirm("이 포스트를 삭제하시겠습니까?")) return;
    const db=await getSupabase();
    const {error}=await db.from("instagram_posts").delete().eq("id",id);
    if(error){ setErr("삭제 실패: "+error.message); return; }
    onChange();
  };

  // wizard 진행률 표시용
  const stepLabel=["① URL 입력","② 임베드 확인 + 상품 매칭","✓ 완료"];

  return (
    <div onClick={onClose}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,
        display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()}
        style={{background:D.surface,borderRadius:14,padding:"24px 28px",
          width:"min(720px,95vw)",maxHeight:"85vh",overflowY:"auto",
          boxShadow:"0 8px 40px rgba(0,0,0,0.25)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:700,fontSize:16,color:D.black}}>{date} · 인스타그램 포스트</div>
          <button onClick={onClose}
            style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
              padding:"4px 12px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
        </div>

        {/* 새 포스트 추가 — wizard */}
        <div style={{background:`${D.blue}06`,borderRadius:8,padding:"14px 16px",marginBottom:18,
          border:`1.5px dashed ${D.blue}55`}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:8,flexWrap:"wrap"}}>
            <div style={{fontWeight:700,fontSize:14,color:D.blue}}>
              ＋ 이 날짜에 포스트 추가
            </div>
            {posts.length>0&&(
              <span style={{fontSize:11,color:D.textMeta}}>
                현재 {posts.length}개 등록됨 · 아래에서 확인/수정 가능
              </span>
            )}
          </div>
          {/* 화살표 형태 step indicator — 각 step 이 chevron 모양으로 흐름 직관화 */}
          <div style={{display:"flex",marginBottom:12,fontSize:11,gap:2}}>
            {stepLabel.map((s,i)=>{
              const isActive=step===i;
              const isDone=step>i;
              const isLast=i===stepLabel.length-1;
              const isFirst=i===0;
              const bg=isActive?D.blue:isDone?`${D.green}22`:D.surface;
              const fg=isActive?"#fff":isDone?D.green:D.textMeta;
              const bd=isActive?D.blue:isDone?`${D.green}55`:D.border;
              // chevron 화살표: 우측 끝에 뾰족한 모양 + 좌측은 들어간 모양 (첫/마지막 제외)
              const clip=isFirst&&isLast?"none":
                isFirst?"polygon(0 0, calc(100% - 12px) 0, 100% 50%, calc(100% - 12px) 100%, 0 100%)":
                isLast?"polygon(0 0, 100% 0, 100% 100%, 0 100%, 12px 50%)":
                "polygon(0 0, calc(100% - 12px) 0, 100% 50%, calc(100% - 12px) 100%, 0 100%, 12px 50%)";
              return (
                <div key={i} style={{flex:1,padding:isFirst?"7px 14px 7px 12px":"7px 14px 7px 22px",
                  textAlign:"center",fontWeight:600,
                  background:bg,color:fg,
                  border:`1px solid ${bd}`,
                  clipPath:clip,
                  display:"flex",alignItems:"center",justifyContent:"center",gap:5,
                  whiteSpace:"nowrap",overflow:"hidden"}}>
                  <span style={{fontSize:11,opacity:0.85}}>
                    {isDone?"✓":isActive?"●":i+1}
                  </span>
                  <span>{s.replace(/^[①②③✓]\s?/,"")}</span>
                </div>
              );
            })}
          </div>

          {step===0&&<>
            <input value={url} onChange={e=>setUrl(e.target.value)}
              placeholder="https://www.instagram.com/p/..."
              style={{width:"100%",padding:"8px 12px",border:`1px solid ${D.border}`,borderRadius:6,
                fontSize:13,marginBottom:8,boxSizing:"border-box"}}/>
            <input value={memo} onChange={e=>setMemo(e.target.value)}
              placeholder="메모 (선택)"
              style={{width:"100%",padding:"8px 12px",border:`1px solid ${D.border}`,borderRadius:6,
                fontSize:13,marginBottom:8,boxSizing:"border-box"}}/>
            {err&&<Alert type="error" msg={err}/>}
            <Btn onClick={handleRun} disabled={saving||!url} style={{width:"100%"}}>
              {saving?"저장 중…":"▶ Run (URL 등록)"}
            </Btn>
          </>}

          {step===1&&<>
            <Alert type="success" msg="✓ URL 등록 완료 — 아래 임베드를 확인하고 소개 상품을 매칭하세요"/>
            <div style={{fontSize:11,color:D.textSub,background:`${D.blue}10`,
              border:`1px solid ${D.blue}30`,borderRadius:6,padding:"8px 12px",margin:"8px 0",lineHeight:1.55}}>
              ℹ️ <b style={{color:D.text}}>사진이 캘린더에 추가되지 않을 경우</b>, 인스타그램 측의 차단 봇 활동으로 인해 자동 저장이 불가할 수 있습니다.
              아래 <b style={{color:D.text}}>등록된 포스트</b> 카드의 <b style={{color:D.text}}>📷 사진 업로드</b> 버튼을 눌러 사진 캡쳐본을 직접 업로드해주세요.
            </div>
            <div style={{minHeight:620,overflow:"hidden",margin:"10px 0"}}>
              <InstagramEmbed url={newPostUrl}/>
            </div>
            <ProductTagger postId={newPostId} tagged={postProductsMap[newPostId]||[]}
              allProducts={allProducts} onChange={onChange}/>
            {/* 우하단 정렬: 완료 버튼이 모달 우측 하단에 위치해 순서 종료를 안심하게 인지 */}
            <div style={{display:"flex",justifyContent:"flex-end",alignItems:"center",
              gap:8,marginTop:18,paddingTop:14,borderTop:`1px solid ${D.border}`}}>
              <button onClick={resetWizard}
                style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
                  padding:"7px 14px",fontSize:12,cursor:"pointer",color:D.textMeta}}>← 처음부터</button>
              <Btn onClick={handleComplete} style={{padding:"8px 22px",fontSize:13,fontWeight:700}}>
                ✓ 완료
              </Btn>
            </div>
          </>}

          {step===2&&<>
            <div style={{textAlign:"center",padding:"20px 0"}}>
              <div style={{fontSize:32,marginBottom:6}}>✓</div>
              <div style={{color:D.green,fontWeight:700,fontSize:14,marginBottom:4}}>등록 완료</div>
              <div style={{color:D.textMeta,fontSize:11,marginBottom:14}}>
                태그된 상품: {(postProductsMap[newPostId]||[]).length}개
              </div>
              <Btn onClick={resetWizard}>＋ 다른 포스트 추가</Btn>
            </div>
          </>}
        </div>

        {/* 등록된 포스트 리스트 */}
        <div style={{fontSize:12,fontWeight:600,color:D.text,marginBottom:10}}>
          등록된 포스트 ({posts.length})
        </div>
        {posts.length===0
          ? <div style={{color:D.textMeta,fontSize:12,padding:"30px 0",textAlign:"center"}}>아직 등록된 포스트가 없습니다.</div>
          : posts.map(p=>(
            // 카드/임베드 영역에 미리 공간 예약 → 임베드 iframe 로딩으로 인한 입력 위치 흔들림 방지
            <div key={p.id} style={{border:`1px solid ${D.border}`,borderRadius:8,padding:14,marginBottom:12,
              minHeight:760}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8,gap:8}}>
                <div style={{flex:1,minWidth:0}}>
                  <a href={p.url} target="_blank" rel="noreferrer"
                    style={{fontSize:11,color:D.blue,wordBreak:"break-all"}}>{p.url}</a>
                  {p.caption_memo&&<div style={{fontSize:12,color:D.textSub,marginTop:4}}>{p.caption_memo}</div>}
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0,alignItems:"flex-start"}}>
                  <ThumbUploadButton post={p} onChange={onChange}/>
                  <ThumbRefreshButton post={p} onChange={onChange}/>
                  <button onClick={()=>handleDelete(p.id)}
                    style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                      padding:"3px 9px",fontSize:11,cursor:"pointer",color:D.red}}>삭제</button>
                </div>
              </div>
              {/* 인스타그램 공식 임베드 — 임베드 iframe 로드 전후 높이 변동 차단 */}
              <div style={{minHeight:620,overflow:"hidden",marginBottom:8}}>
                <InstagramEmbed url={p.url}/>
              </div>
              {/* 소개 상품 태깅 */}
              <ProductTagger postId={p.id} tagged={postProductsMap[p.id]||[]}
                allProducts={allProducts} onChange={onChange}/>
            </div>
          ))}
      </div>
    </div>
  );
}

// 손그림 동그라미 + 별로 강조하는 매칭 상품 라벨
function MatchedProductBadge({ name, qty }) {
  // 손그림 동그라미 + 별 — 항상 다크 오버레이 위에 표시되므로 흰 텍스트 + drop-shadow
  return (
    <span style={{position:"relative",display:"inline-flex",alignItems:"center",gap:3,padding:"1px 4px"}}>
      {/* 별 스티커 */}
      <span style={{fontSize:11,color:"#F2B544",lineHeight:1,filter:"drop-shadow(0 1px 2px rgba(0,0,0,0.6))"}}>★</span>
      <span style={{fontWeight:700,color:"#fff",textShadow:"0 1px 2px rgba(0,0,0,0.6)",position:"relative",zIndex:1}}>{name}</span>
      <span style={{color:"#fff",opacity:0.85,textShadow:"0 1px 2px rgba(0,0,0,0.6)",position:"relative",zIndex:1}}>{qty}장</span>
      {/* 손그림 동그라미 — absolute fill, SVG inline */}
      <svg viewBox="0 0 100 30" preserveAspectRatio="none"
        style={{position:"absolute",inset:-3,width:"calc(100% + 6px)",height:"calc(100% + 6px)",pointerEvents:"none",zIndex:0,filter:"drop-shadow(0 1px 2px rgba(0,0,0,0.5))"}}>
        <path d="M 8 17 C 5 9, 25 4, 50 5 C 78 6, 96 12, 94 18 C 92 24, 70 27, 45 26 C 18 25, 6 21, 10 16"
          fill="none" stroke="#FF6B8A" strokeWidth="1.8" strokeLinecap="round" opacity="0.95"/>
        <path d="M 12 18 C 10 12, 28 7, 52 8 C 76 9, 92 14, 90 19"
          fill="none" stroke="#FF6B8A" strokeWidth="1.0" strokeLinecap="round" opacity="0.6"/>
      </svg>
    </span>
  );
}

function ContentImpact({ orders=[], revenues=[], storeSales=[] }) {
  const now=new Date();
  const [ym,setYm]=useState(()=>`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`);
  const [igPosts,setIgPosts]=useState([]);          // 전체 IG 포스트
  const [postProducts,setPostProducts]=useState([]); // 포스트별 소개 상품 (post_id, product_name)
  const [tableMissing,setTableMissing]=useState(false);
  const [postModalDate,setPostModalDate]=useState(null);  // 모달 열림 상태(iso date)
  const [postLoadTick,setPostLoadTick]=useState(0);

  // 캘린더 셀에 embed iframe 이 실제로 렌더되도록 process() 호출 트리거
  //   deps: igPosts 가 바뀌면 다시 process
  useInstagramEmbedScript([ym, igPosts.length, igPosts.map(p=>p.url).join("|")]);

  // IG 포스트 + 태깅 상품 동시 로딩
  useEffect(()=>{
    (async()=>{
      const db=await getSupabase();
      const [postsRes,prodsRes]=await Promise.all([
        db.from("instagram_posts").select("*").order("post_date",{ascending:false}).limit(1000),
        db.from("instagram_post_products").select("*").limit(5000),
      ]);
      if(postsRes.error){
        // 테이블 미생성 시 안내 (코드 42P01 = undefined_table)
        if(postsRes.error.code==="42P01"||/relation.*does not exist/i.test(postsRes.error.message||"")){
          setTableMissing(true);
        }
        return;
      }
      setIgPosts(postsRes.data||[]);
      setPostProducts(prodsRes.data||[]);
      setTableMissing(false);
      // 주의: thumb_url 자동 백필은 일부러 하지 않음 (Microlink 호출량 절감)
      //       - 저장 시점에만 1회 호출하여 DB 캐시
      //       - 캐시 누락 포스트는 IGPostModal 의 '썸네일 새로고침' 버튼으로 수동 갱신
    })();
  },[postLoadTick]);
  const refreshPosts=()=>setPostLoadTick(t=>t+1);

  // 일자별 포스트 인덱스
  const postsByDate=useMemo(()=>{
    const m={};
    igPosts.forEach(p=>{ if(!m[p.post_date]) m[p.post_date]=[]; m[p.post_date].push(p); });
    return m;
  },[igPosts]);

  // 포스트별 태깅 상품 인덱스 (post_id → [product_name])
  const postProductsMap=useMemo(()=>{
    const m={};
    postProducts.forEach(p=>{
      if(!m[p.post_id]) m[p.post_id]=[];
      m[p.post_id].push(p.product_name);
    });
    return m;
  },[postProducts]);

  // 검색 가능한 전체 상품 명단 (orders + store_sales의 distinct product_name)
  const allProducts=useMemo(()=>{
    const s=new Set();
    orders.forEach(o=>o.product_name&&s.add(o.product_name));
    storeSales.forEach(r=>r.product_name&&s.add(r.product_name));
    return [...s].sort();
  },[orders,storeSales]);

  // 포스트 ID → 색상 (Sankey 팔레트 순환) — 임팩트 카드 morph 셀 라인 컬러로 재사용
  const postColor=(postId)=>D.SANKEY[Math.abs(postId)%D.SANKEY.length];

  // 이름 매칭: 정확 일치 또는 한쪽이 다른 쪽의 부분 문자열 (대소문자 무시)
  const nameMatches=(a,b)=>{
    const al=String(a||"").toLowerCase();
    const bl=String(b||"").toLowerCase();
    if(!al||!bl) return false;
    return al===bl||al.includes(bl)||bl.includes(al);
  };

  // 월 이동
  const shiftMonth=(delta)=>{
    const [y,m]=ym.split("-").map(Number);
    const d=new Date(y,m-1+delta,1);
    setYm(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);
  };

  // ── 일별 데이터 집계 ────────────────────────────
  // 소스: revenues(온라인 매출) + storeSales(매장) + orders(상품 판매 수량)
  // 산출: 일자별 { channels:{ch:매출}, total:매출합, topProducts:[상위 3] }
  const dailyData=useMemo(()=>{
    const byDate={};
    const ensure=d=>{if(!byDate[d]) byDate[d]={channels:{},products:{},total:0,topProducts:[]};return byDate[d];};
    // 온라인 채널 매출 (revenues CSV)
    revenues.forEach(r=>{
      if(!r.date) return;
      const ch=r.channel||"미분류";
      const net=(r.amount||0)-(r.refund_amount||0);
      ensure(r.date).channels[ch]=(byDate[r.date].channels[ch]||0)+net;
    });
    // 매장 매출 (store_sales)
    storeSales.forEach(r=>{
      if(!r.sale_date) return;
      const amt=r.status==="배송"?(r.amount||0):r.status==="반품"?-(r.amount||0):0;
      ensure(r.sale_date).channels["오프라인 스토어"]=(byDate[r.sale_date].channels["오프라인 스토어"]||0)+amt;
    });
    // 판매 상품 수량 — 주문 기준(배송 여부 무관)
    //   반품·취소는 제외 (실제로 무효화된 주문이므로)
    orders.forEach(r=>{
      if(!r.order_date) return;
      if(r.status==="반품"||r.status==="취소") return;
      const name=r.product_name||"미분류";
      ensure(r.order_date).products[name]=(byDate[r.order_date].products[name]||0)+(r.qty||1);
    });
    // 파생 필드 — Top 5
    Object.values(byDate).forEach(d=>{
      d.total=Object.values(d.channels).reduce((s,v)=>s+v,0);
      d.topProducts=Object.entries(d.products)
        .sort((a,b)=>b[1]-a[1]).slice(0,5)
        .map(([name,qty])=>({name,qty}));
    });
    return byDate;
  },[orders,revenues,storeSales]);

  // 월 그리드 생성 — 일요일 시작, 월말일이 포함된 마지막 주까지만 표시
  const grid=useMemo(()=>{
    const [y,m]=ym.split("-").map(Number);
    const first=new Date(y,m-1,1);
    const last=new Date(y,m,0);  // 해당 월 마지막 날
    const startDay=first.getDay(); // 0=Sun
    // (월말일이 들어가는 주의 토요일까지) - (첫 주의 일요일) + 1 = 표시할 셀 수
    const totalDays=startDay+last.getDate();
    const weeks=Math.ceil(totalDays/7);
    const cellCount=weeks*7;
    const cells=[];
    for(let i=0;i<cellCount;i++){
      const d=new Date(y,m-1,i-startDay+1);
      cells.push({
        date:d,
        iso:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`,
        inMonth:d.getMonth()===m-1,
        isToday:d.toDateString()===now.toDateString(),
      });
    }
    return cells;
  },[ym]);

  const monthLabel=(()=>{
    const [y,m]=ym.split("-").map(Number);
    return `${y}년 ${m}월`;
  })();
  const WEEKDAYS=["일","월","화","수","목","금","토"];

  // ── 리본(attribution) 계산 ────────────────────────
  // 규칙: 포스트의 소개 상품이 그 후속 날짜의 판매 Top에 들었을 때만 연결
  // 방향: 포스트 날짜 → 미래 (역방향 X)
  const isoToGridIdx=useMemo(()=>{
    const m={};
    grid.forEach((c,i)=>{m[c.iso]=i;});
    return m;
  },[grid]);
  const lastInMonthIso=useMemo(()=>{
    const inMonths=grid.filter(c=>c.inMonth);
    return inMonths.length?inMonths[inMonths.length-1].iso:null;
  },[grid]);

  // ── 포스트 임팩트 점수 (0~5★) — Pre/Post Sales Lift ──
  // 산식: Lift% = (Post14_qty - Pre14_qty) / max(Pre14_qty, 1) × 100
  //   Pre14_qty  = 포스트 전 14일간 태깅 상품 총 판매 수량
  //   Post14_qty = 포스트 후 14일간 태깅 상품 총 판매 수량
  // 별 등급(임계치):
  //   ≥ +100% → 5★ (매출 2배+)
  //   +50~99% → 4★
  //   +20~49% → 3★
  //   +5~19%  → 2★
  //   -5~+5%  → 1★ (중립)
  //   ≤ -5%   → 0★
  const liftToStars=(lift)=>{
    if(lift>=100) return 5;
    if(lift>=50) return 4;
    if(lift>=20) return 3;
    if(lift>=5) return 2;
    if(lift>=-5) return 1;
    return 0;
  };
  const postScores=useMemo(()=>{
    const map={}; // postId → {stars, lift, preQty, postQty, preDays[], postDays[], tagged, WINDOW}
    const WINDOW=14;
    const dayMs=86400000;
    igPosts.forEach(post=>{
      const tagged=postProductsMap[post.id]||[];
      if(!tagged.length){
        map[post.id]={stars:0,lift:0,preQty:0,postQty:0,preDays:[],postDays:[],tagged:[],WINDOW,empty:true};
        return;
      }
      const baseMs=new Date(post.post_date).getTime();
      const collect=(offset)=>{ // offset: -14..-1 또는 +1..+14
        const iso=new Date(baseMs+offset*dayMs).toISOString().slice(0,10);
        const dd=dailyData[iso];
        let dQty=0;
        Object.entries(dd?.products||{}).forEach(([name,qty])=>{
          if(tagged.some(t=>nameMatches(t,name))) dQty+=qty;
        });
        return {iso,qty:dQty};
      };
      const preDays=[],postDays=[];
      for(let i=WINDOW;i>=1;i--) preDays.push(collect(-i));
      for(let i=1;i<=WINDOW;i++) postDays.push(collect(i));
      const preQty=preDays.reduce((s,d)=>s+d.qty,0);
      const postQty=postDays.reduce((s,d)=>s+d.qty,0);
      const lift=preQty>0?((postQty-preQty)/preQty*100):(postQty>0?100:0);
      const stars=liftToStars(lift);
      map[post.id]={stars,lift,preQty,postQty,preDays,postDays,tagged,WINDOW};
    });
    return map;
  },[igPosts,postProductsMap,dailyData]);

  // 임팩트 점수 산식 모달 상태
  const [scoreModalIso,setScoreModalIso]=useState(null);
  // 같은 날 다중 포스트 좌우 넘김 — iso → 현재 보여줄 인덱스
  const [cellPostIdx,setCellPostIdx]=useState({});
  // 포스트 임팩트 분석 모드 토글 — ON 시: 포스트가 있는 셀이 임팩트 카드로 변형 + 상단에 판매 속도 timeline
  const [impactMode,setImpactMode]=useState(false);
  // 임팩트 모드 헤더로 스크롤하기 위한 ref + 토글 ON 시 자동 스크롤
  const impactHeaderRef=useRef(null);
  useEffect(()=>{
    if(impactMode&&impactHeaderRef.current){
      // 약간 지연 후 스크롤 (DOM 렌더 완료 보장)
      requestAnimationFrame(()=>{
        impactHeaderRef.current?.scrollIntoView({behavior:"smooth",block:"start"});
      });
    }
  },[impactMode]);

  // 월 넘어가는 매칭 (이번 달 포스트 → 다음 달 이후 판매 Top)
  const crossMonthByPostDate=useMemo(()=>{
    if(!lastInMonthIso) return {};
    const out={}; // post_date(iso) → {nextYm, count}
    igPosts.forEach(post=>{
      if(isoToGridIdx[post.post_date]===undefined) return;
      const tagged=postProductsMap[post.id]||[];
      if(!tagged.length) return;
      // dailyData에서 lastInMonthIso 이후의 모든 날짜 검사
      Object.entries(dailyData).forEach(([iso,d])=>{
        if(iso<=lastInMonthIso) return;
        if(!d.topProducts?.length) return;
        for(const tagName of tagged){
          if(d.topProducts.find(tp=>nameMatches(tagName,tp.name))){
            if(!out[post.post_date]) out[post.post_date]={nextYm:iso.slice(0,7),count:0};
            if(iso.slice(0,7)<out[post.post_date].nextYm) out[post.post_date].nextYm=iso.slice(0,7);
            out[post.post_date].count++;
            break;
          }
        }
      });
    });
    return out;
  },[igPosts,postProductsMap,dailyData,lastInMonthIso,isoToGridIdx]);

  const weeksCount=grid.length/7;

  // 임팩트 모드 분석 요약 — KPI + sales velocity timeline + 인사이트 + 태그×매출 산점도 데이터
  const impactSummary=useMemo(()=>{
    if(!impactMode) return null;
    // ─ KPI ─
    let stars=0,liftSum=0,liftN=0,totalAttr=0;
    igPosts.forEach(p=>{
      const s=postScores[p.id]; if(!s) return;
      stars+=s.stars||0;
      if(typeof s.lift==="number"){ liftSum+=s.lift; liftN++; }
      totalAttr+=(s.postQty||0)-(s.preQty||0);
    });
    const avgStars=igPosts.length?stars/igPosts.length:0;
    const avgLift=liftN?liftSum/liftN:0;
    // ─ Sales velocity timeline (현재 달 daily qty) ─
    // 오늘 이후 또는 데이터 없는 trailing 날짜는 제외 (그래프가 0으로 떨어지지 않도록)
    const todayIso=dayjs().format("YYYY-MM-DD");
    const monthCells=grid.filter(c=>c.inMonth&&c.iso<=todayIso);
    const timeline=monthCells.map(c=>({
      iso:c.iso, day:c.date.getDate(),
      qty:(dailyData[c.iso]?.topProducts||[]).reduce((s,p)=>s+(p.qty||0),0),
      hasPost:(postsByDate[c.iso]||[]).length>0,
    }));
    // ─ 속도계: 후반 1/2 평균 vs 전반 1/2 평균 → ±% 가속/감속 ─
    const half=Math.floor(timeline.length/2);
    const firstAvg=half?timeline.slice(0,half).reduce((s,d)=>s+d.qty,0)/half:0;
    const secondAvg=(timeline.length-half)?timeline.slice(half).reduce((s,d)=>s+d.qty,0)/(timeline.length-half):0;
    const velocityChange=firstAvg?((secondAvg-firstAvg)/firstAvg)*100:0;
    // ─ 최고 임팩트 포스트 ─
    let bestPost=null,bestLift=-Infinity;
    igPosts.forEach(p=>{
      const s=postScores[p.id]; if(!s||typeof s.lift!=="number") return;
      if(s.lift>bestLift){ bestLift=s.lift; bestPost=p; }
    });
    // ─ 태그 상품 수 × 판매량 산점도 + 상관 r ─
    const tagScatter=igPosts.map(p=>{
      const s=postScores[p.id]||{};
      const tagCount=(postProductsMap[p.id]||[]).length;
      return {tagCount,postQty:s.postQty||0,iso:p.post_date,lift:s.lift||0};
    }).filter(x=>x.tagCount>0);
    let correlation=0;
    if(tagScatter.length>=3){
      const mx=tagScatter.reduce((s,x)=>s+x.tagCount,0)/tagScatter.length;
      const my=tagScatter.reduce((s,x)=>s+x.postQty,0)/tagScatter.length;
      let num=0,dx2=0,dy2=0;
      tagScatter.forEach(x=>{const dx=x.tagCount-mx,dy=x.postQty-my;num+=dx*dy;dx2+=dx*dx;dy2+=dy*dy;});
      correlation=(dx2&&dy2)?num/Math.sqrt(dx2*dy2):0;
    }
    const totalTags=tagScatter.reduce((s,x)=>s+x.tagCount,0);
    const totalQty=tagScatter.reduce((s,x)=>s+x.postQty,0);
    const avgQtyPerTag=totalTags?totalQty/totalTags:0;
    // ─ 평균 태그 상품 수 (전체 포스트 기준, 태그 없는 포스트 포함) ─
    const avgTagCount=igPosts.length?igPosts.reduce((s,p)=>s+(postProductsMap[p.id]||[]).length,0)/igPosts.length:0;
    // ─ LIFT가 가장 활발한 포스트의 태그 상품 수 ─
    const bestPostTagCount=bestPost?(postProductsMap[bestPost.id]||[]).length:0;
    // ─ 평균 어트리뷰션 속도 = 포스트 1회당 일평균 추가 판매 qty (장/일) ─
    //   각 포스트의 (postQty - preQty) / 14 의 평균
    const velList=igPosts.map(p=>{const s=postScores[p.id];return s?((s.postQty||0)-(s.preQty||0))/14:null;}).filter(v=>v!==null);
    const avgAttrVelocity=velList.length?velList.reduce((s,v)=>s+v,0)/velList.length:0;
    // ─ 포스트별 LIFT 랭킹 (best → worst) ─
    const postRanking=igPosts.map(p=>{
      const s=postScores[p.id]||{};
      const tags=postProductsMap[p.id]||[];
      // 베스트 매칭 상품
      const m=s.postPerProduct||{};
      const list=Object.entries(m).sort(([,a],[,b])=>b-a);
      const bestProduct=list.length&&list[0][1]>0?list[0][0]:null;
      return {iso:p.post_date,id:p.id,lift:s.lift||0,stars:s.stars||0,
        postQty:s.postQty||0,preQty:s.preQty||0,bestProduct,tagCount:tags.length};
    }).filter(x=>x.tagCount>0).sort((a,b)=>b.lift-a.lift);
    // ─ 포스트별 "전체 매출" 영향 랭킹 — 태깅되지 않은 상품 포함, 전체 매출(원) 기준 ─
    const dayMs=24*60*60*1000;
    const overallRanking=igPosts.map(p=>{
      const baseDate=new Date(p.post_date);
      let pre=0, post=0;
      for(let i=1;i<=14;i++){
        const preIso=new Date(baseDate.getTime()-i*dayMs).toISOString().slice(0,10);
        const postIso=new Date(baseDate.getTime()+i*dayMs).toISOString().slice(0,10);
        pre+=(dailyData[preIso]?.total||0);
        post+=(dailyData[postIso]?.total||0);
      }
      const lift=pre?((post-pre)/pre)*100:0;
      return {iso:p.post_date,id:p.id,lift,preRev:pre,postRev:post,delta:post-pre};
    }).sort((a,b)=>b.lift-a.lift);
    // ─ Cohort 평균 곡선 — 모든 포스트의 D-14 ~ D+14 평균 판매량 + 효과 정점 day ─
    const cohortDays=Array.from({length:29},(_,i)=>i-14); // -14..0..14
    const cohort=cohortDays.map(d=>{
      const vals=[];
      igPosts.forEach(p=>{
        const s=postScores[p.id]; if(!s) return;
        if(d<0){
          const idx=(s.preDays||[]).length+d;
          if(idx>=0&&idx<(s.preDays||[]).length) vals.push(s.preDays[idx].qty||0);
        } else if(d===0){
          // 게시일 — 보간을 위해 pre 마지막 + post 첫 평균
          const pl=s.preDays?.length?s.preDays[s.preDays.length-1].qty||0:0;
          const p0=s.postDays?.length?s.postDays[0].qty||0:0;
          vals.push((pl+p0)/2);
        } else {
          const idx=d-1;
          if(idx>=0&&idx<(s.postDays||[]).length) vals.push(s.postDays[idx].qty||0);
        }
      });
      return {d, avg:vals.length?vals.reduce((s,v)=>s+v,0)/vals.length:0, n:vals.length};
    });
    // 효과 정점: D>0 중 평균 최대인 day
    const postCohort=cohort.filter(c=>c.d>0&&c.n>0);
    const peakDay=postCohort.length?postCohort.reduce((a,b)=>a.avg>b.avg?a:b).d:0;
    const peakAvg=postCohort.length?Math.max(...postCohort.map(c=>c.avg)):0;
    // 전 평균 (D-14..D-1)
    const preCohort=cohort.filter(c=>c.d<0&&c.n>0);
    const preAvg=preCohort.length?preCohort.reduce((s,c)=>s+c.avg,0)/preCohort.length:0;
    // 후 평균 (D+1..D+14)
    const postCohortAvg=postCohort.length?postCohort.reduce((s,c)=>s+c.avg,0)/postCohort.length:0;
    // ─ 자동 인사이트 텍스트 (방향성 있는) ─
    const insights=[];
    if(avgLift>=15) insights.push({tone:"good",text:`평균 Lift +${avgLift.toFixed(0)}% — 콘텐츠 ROI가 매우 강력합니다. 유사한 포스팅 방식을 다음 달에도 유지하세요.`});
    else if(avgLift>=5) insights.push({tone:"good",text:`평균 Lift +${avgLift.toFixed(0)}% — 포스팅이 판매를 견인 중. 상위 임팩트 포스트를 재활용해 성과를 증폭하세요.`});
    else if(avgLift<=-5) insights.push({tone:"bad",text:`평균 Lift ${avgLift.toFixed(0)}% — 포스팅 후 오히려 판매가 줄었습니다. 태깅 상품·캡션·해시태그 전략을 재검토하세요.`});
    else insights.push({tone:"neutral",text:`평균 Lift ${avgLift>=0?"+":""}${avgLift.toFixed(0)}% — 영향이 크지 않습니다. 포스트 빈도·CTA를 강화해 lift를 끌어올리세요.`});
    if(velocityChange>=10) insights.push({tone:"good",text:`이번 달 후반 판매 속도가 전반 대비 +${velocityChange.toFixed(0)}% 가속 중 — 모멘텀이 살아있을 때 추가 포스트로 견인하세요.`});
    else if(velocityChange<=-10) insights.push({tone:"bad",text:`판매 속도 ${velocityChange.toFixed(0)}% 감속 — 신규 포스트·프로모션으로 흐름을 빠르게 회복하세요.`});
    if(bestPost){
      const bs=postScores[bestPost.id]||{};
      insights.push({tone:"good",text:`최고 임팩트: ${bestPost.post_date} (Lift ${bs.lift>=0?"+":""}${(bs.lift||0).toFixed(0)}%, ${bs.postQty||0}장) — 동일한 구도·태깅 패턴을 다시 시도하세요.`});
    }
    if(tagScatter.length>=3){
      if(correlation>=0.4) insights.push({tone:"good",text:`태그 상품 수↑ → 판매량↑ 강한 양의 상관 (r=${correlation.toFixed(2)}). 포스트당 태깅 상품을 더 적극적으로 늘려보세요.`});
      else if(correlation<=-0.2) insights.push({tone:"bad",text:`태그가 많을수록 오히려 판매↓ (r=${correlation.toFixed(2)}) — 핵심 1~2 상품에만 집중 노출하는 게 효율적입니다.`});
      else insights.push({tone:"neutral",text:`태그 수와 판매량 사이 뚜렷한 패턴 없음 (r=${correlation.toFixed(2)}) — 태깅보다 콘텐츠 품질이 결정적입니다.`});
    }
    return {postCount:igPosts.length,avgStars,avgLift,totalAttr,
      timeline,velocityChange,firstAvg,secondAvg,
      bestPost,bestPostTagCount,tagScatter,correlation,avgQtyPerTag,
      avgTagCount,avgAttrVelocity,
      postRanking,overallRanking,cohort,peakDay,peakAvg,preAvg,postCohortAvg,
      insights};
  },[impactMode,igPosts,postScores,postProductsMap,grid,dailyData,postsByDate]);

  return (
    <div style={{padding:"20px 24px",maxWidth:1960,margin:"0 auto"}}>
      {/* 임팩트 모드 상단 분석 패널 — KPI + 속도계 + Sales Velocity Timeline + 태그×매출 산점도 */}
      {impactMode&&impactSummary&&(
        <div ref={impactHeaderRef} style={{scrollMarginTop:16}}>
          <ImpactAnalysisHeader summary={impactSummary} monthLabel={monthLabel}/>
        </div>
      )}

      {/* 월 셀렉터 — 중앙 정렬 */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:14,marginBottom:24}}>
        <button onClick={()=>shiftMonth(-1)} data-hf
          style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:8,
            padding:"8px 16px",fontSize:16,cursor:"pointer",color:D.text,fontWeight:600}}>◂</button>
        <div style={{fontSize:22,fontWeight:700,color:D.black,minWidth:170,textAlign:"center"}}>{monthLabel}</div>
        <button onClick={()=>shiftMonth(1)} data-hf
          style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:8,
            padding:"8px 16px",fontSize:16,cursor:"pointer",color:D.text,fontWeight:600}}>▸</button>
        <button onClick={()=>setYm(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`)} data-hf
          style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:8,
            padding:"7px 14px",fontSize:12,cursor:"pointer",color:D.textMeta,marginLeft:8}}>이번 달</button>
      </div>

      {/* 캘린더 그리드 — iPad 사이즈 (셀 크게) */}
      <Card>
        <div style={{position:"relative"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:1,background:D.border,border:`1px solid ${D.border}`,borderRadius:8,overflow:"hidden"}}>
          {WEEKDAYS.map((w,i)=>(
            <div key={w} style={{background:D.surfaceAlt,padding:"10px 12px",fontSize:13,fontWeight:600,
              color:i===0?D.red:i===6?D.blue:D.textMeta,textAlign:"center"}}>{w}</div>
          ))}
          {grid.map((c,i)=>{
            const d=dailyData[c.iso];
            // 채널 매출: 매출 큰 순 정렬, 채널당 1줄 mini-bar
            const channels=d?Object.entries(d.channels).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]):[];
            const maxAmt=channels.length?channels[0][1]:1;
            // 같은 날 포스트는 별점 desc 정렬 — 임팩트 큰 포스트 먼저 보임
            const postsRaw=postsByDate[c.iso]||[];
            const posts=[...postsRaw].sort((a,b)=>(postScores[b.id]?.stars||0)-(postScores[a.id]?.stars||0));
            const curIdx=Math.min(cellPostIdx[c.iso]||0,Math.max(0,posts.length-1));
            const curPost=posts[curIdx];
            const cm=crossMonthByPostDate[c.iso];
            // 당일 임베드 포스트의 태깅 상품 ∩ 당일 판매 Top → 강조 매칭 set
            const samedayTagged=new Set();
            posts.forEach(p=>(postProductsMap[p.id]||[]).forEach(n=>samedayTagged.add(n)));
            const isHighlighted=name=>{
              for(const t of samedayTagged) if(nameMatches(t,name)) return true;
              return false;
            };
            // 포스트 있는 셀은 검정 오버레이 위 → 흰 텍스트, 없으면 기본
            const hasBg=c.inMonth&&posts.length>0;
            const fg=hasBg?"#fff":D.text;
            const fgMeta=hasBg?"rgba(255,255,255,0.75)":D.textMeta;
            const ts=hasBg?"0 1px 2px rgba(0,0,0,0.6)":"none"; // text-shadow for legibility
            return (
            <div key={i} onClick={c.inMonth?()=>setPostModalDate(c.iso):undefined}
              style={{
              background:c.inMonth?D.surface:D.bg,
              aspectRatio:"4 / 5",minHeight:252,padding:0,
              opacity:c.inMonth?1:0.35,
              position:"relative",overflow:"hidden",
              cursor:c.inMonth?"pointer":"default",
              display:"flex",flexDirection:"column"}}>
              {/* IG 포스트 썸네일 배경 — 임팩트 모드에서는 셀 morph가 자체 썸네일을 가지므로 표시 안 함 */}
              {!impactMode&&c.inMonth&&posts.length>0&&curPost?.thumb_url&&(
                <div style={{position:"absolute",inset:0,zIndex:0,overflow:"hidden",pointerEvents:"none"}}>
                  <InstagramThumb src={curPost.thumb_url}/>
                </div>
              )}
              {/* 반투명 검정 레이어 — 흰 텍스트 가독성 확보 */}
              {!impactMode&&c.inMonth&&posts.length>0&&(
                <div style={{position:"absolute",inset:0,zIndex:1,pointerEvents:"none",
                  background:"rgba(0,0,0,0.42)"}}/>
              )}
              {/* 같은 날 멀티 포스트 좌우 넘김 화살표 (별점 desc 순) */}
              {c.inMonth&&posts.length>1&&(
                <>
                  <button onClick={e=>{e.stopPropagation();setCellPostIdx(p=>({...p,[c.iso]:(curIdx-1+posts.length)%posts.length}));}}
                    title={`이전 포스트 (${curIdx+1}/${posts.length})`}
                    style={{position:"absolute",left:4,top:"50%",transform:"translateY(-50%)",zIndex:3,
                      background:"rgba(255,255,255,0.9)",border:`1px solid ${D.border}`,borderRadius:14,
                      width:24,height:24,padding:0,fontSize:13,cursor:"pointer",color:D.text,fontWeight:700,
                      lineHeight:1,boxShadow:"0 1px 4px rgba(0,0,0,0.15)"}}>‹</button>
                  <button onClick={e=>{e.stopPropagation();setCellPostIdx(p=>({...p,[c.iso]:(curIdx+1)%posts.length}));}}
                    title={`다음 포스트 (${curIdx+1}/${posts.length})`}
                    style={{position:"absolute",right:4,top:"50%",transform:"translateY(-50%)",zIndex:3,
                      background:"rgba(255,255,255,0.9)",border:`1px solid ${D.border}`,borderRadius:14,
                      width:24,height:24,padding:0,fontSize:13,cursor:"pointer",color:D.text,fontWeight:700,
                      lineHeight:1,boxShadow:"0 1px 4px rgba(0,0,0,0.15)"}}>›</button>
                  <div style={{position:"absolute",bottom:3,left:"50%",transform:"translateX(-50%)",zIndex:3,
                    fontSize:9,color:"#000",fontWeight:700,background:"rgba(255,255,255,0.85)",
                    padding:"1px 7px",borderRadius:8,whiteSpace:"nowrap"}}>
                    {curIdx+1} / {posts.length}
                  </div>
                </>
              )}
              {/* 임팩트 모드 + 포스트 있는 셀: PR1 카드 스타일로 morph */}
              {impactMode&&posts.length>0&&curPost?(
                <ImpactCellMorph post={curPost} score={postScores[curPost.id]}
                  tags={postProductsMap[curPost.id]||[]}
                  dateNum={c.date.getDate()} isToday={c.isToday}
                  postsCount={posts.length}
                  onStarClick={()=>setScoreModalIso(c.iso)}
                  onEditClick={()=>setPostModalDate(c.iso)}
                  color={postColor(curPost.id)}
                  cm={cm} setYm={setYm}/>
              ):impactMode?(
                /* 임팩트 모드 + 포스트 없는 셀: 빈 셀 — 날짜만 옅게 표시 */
                <div style={{position:"relative",zIndex:2,padding:"10px 12px",
                  display:"flex",alignItems:"flex-start",flex:1}}>
                  <span style={{fontSize:14,fontWeight:c.isToday?700:500,
                    color:c.isToday?D.blue:i%7===0?`${D.red}66`:i%7===6?`${D.blue}66`:D.textMeta,
                    opacity:c.inMonth?0.6:0.25}}>
                    {c.date.getDate()}
                  </span>
                </div>
              ):(
              <div style={{position:"relative",zIndex:2,padding:"10px 12px",
                display:"flex",flexDirection:"column",flex:1,minHeight:0}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                  <div style={{display:"flex",alignItems:"center",gap:4,flexWrap:"wrap"}}>
                    <span style={{fontSize:14,fontWeight:c.isToday?700:600,
                      color:hasBg?"#fff":(c.isToday?D.blue:i%7===0?D.red:i%7===6?D.blue:D.text),
                      textShadow:ts}}>
                      {c.date.getDate()}
                    </span>
                    {posts.length>0&&(()=>{
                      // 다수 포스트면 max 점수 사용 (사용자 입장: 최고 임팩트 포스트로 평가)
                      const stars=posts.reduce((m,p)=>Math.max(m,postScores[p.id]?.stars||0),0);
                      return (
                        <button onClick={e=>{e.stopPropagation();setScoreModalIso(c.iso);}}
                          data-star-iso={c.iso}
                          title="임팩트 점수 산식 보기"
                          style={{background:"none",border:"none",padding:0,cursor:"pointer",
                            fontSize:11,letterSpacing:0.5,lineHeight:1,whiteSpace:"nowrap",
                            filter:"drop-shadow(0 1px 2px rgba(0,0,0,0.55))"}}>
                          <span style={{color:"#fff"}}>{"★".repeat(stars)}</span>
                          <span style={{color:"rgba(255,255,255,0.40)"}}>{"★".repeat(5-stars)}</span>
                        </button>
                      );
                    })()}
                  </div>
                  <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap",justifyContent:"flex-end"}}>
                    {posts.length>0&&(
                      <span title={`${posts.length}개 포스트`}
                        style={{display:"inline-flex",alignItems:"center",gap:3}}>
                        {/* 흰색 인스타그램 로고 */}
                        <svg width="13" height="13" viewBox="0 0 24 24"
                          style={{filter:"drop-shadow(0 1px 2px rgba(0,0,0,0.55))"}}>
                          <rect x="3" y="3" width="18" height="18" rx="5" stroke="#fff" strokeWidth="2.2" fill="none"/>
                          <circle cx="12" cy="12" r="4" stroke="#fff" strokeWidth="2.2" fill="none"/>
                          <circle cx="17.5" cy="6.5" r="1.3" fill="#fff"/>
                        </svg>
                        {/* 흰색 숫자 뱃지 */}
                        <span style={{fontSize:10,fontWeight:700,color:"#111",
                          background:"#fff",padding:"0 6px",borderRadius:8,
                          minWidth:12,textAlign:"center",lineHeight:"16px",
                          boxShadow:"0 1px 3px rgba(0,0,0,0.25)"}}>{posts.length}</span>
                      </span>
                    )}
                    {cm&&<button onClick={e=>{e.stopPropagation();setYm(cm.nextYm);}}
                      title={`다음: ${cm.nextYm} (${cm.count}건)`}
                      style={{fontSize:9,color:D.blue,fontWeight:700,background:`${D.blue}10`,
                        border:`1px solid ${D.blue}30`,borderRadius:10,padding:"2px 6px",cursor:"pointer"}}>
                      ▸{parseInt(cm.nextYm.split("-")[1])}월
                    </button>}
                    {c.isToday&&<span style={{fontSize:10,color:D.blue,fontWeight:700,
                      background:`${D.blue}15`,padding:"2px 6px",borderRadius:10}}>오늘</span>}
                    {/* 임베드 수정/삭제 진입 버튼 — 모달 안에서 삭제·수정 가능 */}
                    {c.inMonth&&posts.length>0&&(
                      <button onClick={e=>{e.stopPropagation();setPostModalDate(c.iso);}}
                        title="임베드 포스트 수정/삭제"
                        style={{fontSize:11,color:D.textSub,background:"rgba(255,255,255,0.85)",
                          border:`1px solid ${D.border}`,borderRadius:5,padding:"1px 6px",cursor:"pointer",lineHeight:1}}>
                        ✎
                      </button>
                    )}
                  </div>
                </div>
                {c.inMonth&&d&&d.total>0&&(
                  <>
                    <div style={{fontSize:11,fontWeight:700,color:fg,marginBottom:6,textShadow:ts}}>
                      {fmtWonShort(d.total)}
                    </div>
                    {/* 채널별 1줄 미니 바 — 매출 큰 순 */}
                    <div style={{marginBottom:8}}>
                      {channels.slice(0,4).map(([ch,amt])=>(
                        <div key={ch} title={`${ch} ${fmtWonShort(amt)}`}
                          style={{display:"flex",alignItems:"center",gap:5,marginBottom:2}}>
                          <span style={{fontSize:9,color:hasBg?"#fff":chColor(ch),fontWeight:600,
                            minWidth:34,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",textShadow:ts}}>
                            {ch.length>3?ch.slice(0,3):ch}
                          </span>
                          <div style={{flex:1,height:4,background:hasBg?"rgba(255,255,255,0.25)":`${D.border}80`,borderRadius:2,overflow:"hidden"}}>
                            <div style={{width:`${Math.max(4,amt/maxAmt*100)}%`,height:"100%",background:chColor(ch),borderRadius:2}}/>
                          </div>
                          <span style={{fontSize:9,color:fgMeta,minWidth:36,textAlign:"right",fontVariantNumeric:"tabular-nums",textShadow:ts}}>
                            {fmtWonShort(amt)}
                          </span>
                        </div>
                      ))}
                    </div>
                  </>
                )}
                {c.inMonth&&(samedayTagged.size>0||d?.topProducts?.length>0)&&(
                  <div style={{marginTop:"auto",display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,fontSize:9,color:fg,lineHeight:1.6,textShadow:ts}}>
                    {/* 좌측: 소개 상품 (당일 포스트 태깅 상품) */}
                    <div style={{minWidth:0}}>
                      <div style={{fontSize:9,fontWeight:700,color:fg,letterSpacing:"0.05em",
                        textTransform:"uppercase",marginBottom:2}}>소개 상품</div>
                      {samedayTagged.size>0
                        ?[...samedayTagged].slice(0,5).map((name,j)=>(
                          <div key={j} data-tag-iso={c.iso} data-tag-name={name}
                            style={{whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",color:fg}}>
                            · {name}
                          </div>
                        ))
                        :<div style={{color:fgMeta,opacity:0.7}}>—</div>}
                    </div>
                    {/* 우측: 판매 베스트 (Top 5) — 매칭은 손그림 강조 */}
                    <div style={{minWidth:0}}>
                      <div style={{fontSize:9,fontWeight:700,color:fg,letterSpacing:"0.05em",
                        textTransform:"uppercase",marginBottom:2}}>판매 top</div>
                      {d?.topProducts?.length>0
                        ?d.topProducts.map((p,j)=>{
                          const hi=isHighlighted(p.name);
                          return (
                            <div key={j} data-iso={c.iso} data-pidx={j} className="impact-prod-row"
                              style={{display:"flex",alignItems:"center",gap:3,color:fg,
                              whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",
                              position:"relative"}}>
                              {hi
                                ?<MatchedProductBadge name={p.name} qty={p.qty}/>
                                :<><span style={{overflow:"hidden",textOverflow:"ellipsis",minWidth:0,color:fg}}>• {p.name}</span><span style={{color:fg,flexShrink:0}}>{p.qty}장</span></>}
                            </div>
                          );
                        })
                        :<div style={{color:fgMeta,opacity:0.7}}>—</div>}
                    </div>
                  </div>
                )}
              </div>
              )}
            </div>
          );})}
        </div>
        </div>
      </Card>

      {tableMissing&&<Alert type="warn" msg={`Supabase에 instagram_posts 테이블이 없습니다. 아래 SQL을 Supabase Editor에서 실행해주세요:\n\nCREATE TABLE IF NOT EXISTS instagram_posts (\n  id            serial PRIMARY KEY,\n  post_date     date NOT NULL,\n  url           text NOT NULL,\n  caption_memo  text,\n  created_at    timestamptz DEFAULT now()\n);\nCREATE INDEX IF NOT EXISTS instagram_posts_date_idx ON instagram_posts (post_date);`}/>}

      {/* 임팩트 모드 하단 자동 인사이트 — 방향성 있는 추천 문구 */}
      {impactMode&&impactSummary&&impactSummary.insights.length>0&&(
        <ImpactInsightsFooter insights={impactSummary.insights}/>
      )}

      <div style={{marginTop:16,fontSize:11,color:D.textMeta,lineHeight:1.7,background:D.bg,
        padding:"10px 14px",borderRadius:6,border:`1px dashed ${D.border}`}}>
        <b>사용법</b><br/>
        ① 셀 클릭 → 인스타그램 URL 등록 → 모달에서 소개 상품 검색·태깅<br/>
        ② 하단 중앙 토글로 <b style={{color:D.text}}>포스트 임팩트 분석</b> ON → 상단에 KPI + 판매 속도 timeline + 속도계 + 태그×매출 산점도 + 자동 인사이트 표시. 포스트가 있는 셀은 임팩트 카드로 변형됩니다.<br/>
        ③ 매칭이 다음 달까지 이어지면 셀에 <span style={{color:D.blue,fontWeight:600}}>▸N월</span> 버튼 표시 — 클릭하면 그 달로 이동<br/>
        <span style={{color:D.text}}>★ 임팩트 점수</span>: 포스팅 후 14일 vs 전 14일 태깅 상품 판매 변화율을 5★ 등급으로 변환 — ★를 클릭하면 산식·일별 표 모달
      </div>

      {/* 포스트 등록·관리 모달 */}
      {postModalDate&&<IGPostModal date={postModalDate} posts={postsByDate[postModalDate]||[]}
        postProductsMap={postProductsMap} allProducts={allProducts}
        onClose={()=>setPostModalDate(null)} onChange={refreshPosts}/>}

      {/* 임팩트 점수 산식 모달 */}
      {scoreModalIso&&(
        <ImpactScoreModal iso={scoreModalIso}
          posts={postsByDate[scoreModalIso]||[]}
          postScores={postScores}
          onClose={()=>setScoreModalIso(null)}/>
      )}

      {/* Sticky 포스트 임팩트 분석 토글 + 토글 하단 sticky 가이드 (모드 ON 시에만) */}
      {igPosts.length>0&&(
        <div style={{position:"fixed",left:"50%",bottom:"4vh",transform:"translateX(-50%)",
          zIndex:1500,display:"flex",flexDirection:"column",alignItems:"center",gap:6,
          maxWidth:"min(720px, 92vw)"}}>
          <button onClick={()=>setImpactMode(v=>!v)}
            title={impactMode?"포스트 임팩트 분석 모드 끄기 (캘린더 일반 보기)":"포스트 임팩트 분석 모드 켜기"}
            style={{padding:"10px 22px",fontSize:12,fontWeight:600,letterSpacing:"0.02em",
              background:D.black,color:"#fff",
              border:`1px solid ${impactMode?"#fff":D.black}`,borderRadius:6,cursor:"pointer",
              boxShadow:"0 4px 14px rgba(0,0,0,0.25)",
              display:"inline-flex",alignItems:"center",gap:8}}>
            <span style={{width:7,height:7,borderRadius:"50%",
              background:impactMode?"#22c55e":"rgba(255,255,255,0.35)",
              boxShadow:impactMode?"0 0 6px #22c55e":"none"}}/>
            포스트 임팩트 분석 {impactMode?`ON · ${igPosts.length}개`:"OFF"}
          </button>
          {impactMode&&<ImpactGuideSticky monthLabel={monthLabel}/>}
        </div>
      )}
    </div>
  );
}

// 토글 버튼 하단 sticky 가이드 — 분석 모드의 해석 방법을 간결하게 안내
function ImpactGuideSticky({ monthLabel }) {
  return (
    <div style={{
      background:D.surface,color:D.text,
      border:`1px solid ${D.border}`,
      borderRadius:10,padding:"14px 18px",
      boxShadow:"0 6px 20px rgba(0,0,0,0.18)",
      width:"min(1280px, 94vw)",maxHeight:"58vh",overflowY:"auto"
    }}>
      <div style={{fontSize:13,fontWeight:700,letterSpacing:"0.03em",marginBottom:8,
        display:"flex",alignItems:"center",justifyContent:"center",gap:8,flexWrap:"wrap",color:D.black,textAlign:"center"}}>
        포스트 임팩트 분석 모드 · {monthLabel}
        <span style={{fontSize:11,fontWeight:500,color:D.textMeta,marginLeft:2}}>
          업계 표준 Pre/Post Sales Lift 방식
        </span>
      </div>
      <div style={{fontSize:12,color:D.textSub,lineHeight:1.65,marginBottom:10,textAlign:"center"}}>
        <b style={{color:D.text}}>이 페이지는 무엇을 보여주나요?</b><br/>
        인스타그램에 어떤 상품을 소개(태깅)한 포스트를 올렸을 때, 그 상품의 판매량이 정말 늘었는지 — 그리고 얼마나 늘었는지를 한눈에 볼 수 있게 해줍니다.
        포스트가 효과가 있었는지 <b style={{color:D.text}}>"숫자"</b>로 판단하세요.
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(260px, 1fr))",gap:8,marginTop:6}}>
        <div style={{background:D.surfaceAlt,borderRadius:8,padding:"10px 12px",border:`1px solid ${D.border}`}}>
          <div style={{fontSize:11,fontWeight:700,color:"#15803d",letterSpacing:"0.04em",marginBottom:5}}>
            ① LIFT (가장 중요) — 포스트 효과 측정
          </div>
          <div style={{fontSize:11,color:D.textSub,lineHeight:1.6}}>
            포스트 게시 <b style={{color:D.text}}>전 14일</b>의 태깅 상품 평균 판매량과 <b style={{color:D.text}}>후 14일</b>의 평균 판매량의 변화율(%).
            <br/>
            <span style={{color:"#15803d"}}>+30% 이상</span> 매우 좋음 / <span style={{color:"#15803d"}}>+10~30%</span> 양호 / <span style={{color:D.textMeta}}>−10~+10%</span> 영향 미미 / <span style={{color:"#b91c1c"}}>−10% 이하</span> 역효과
          </div>
        </div>
        <div style={{background:D.surfaceAlt,borderRadius:8,padding:"10px 12px",border:`1px solid ${D.border}`}}>
          <div style={{fontSize:11,fontWeight:700,color:"#b45309",letterSpacing:"0.04em",marginBottom:5}}>
            ② ★ 점수 — LIFT를 5단계 등급으로
          </div>
          <div style={{fontSize:11,color:D.textSub,lineHeight:1.6}}>
            LIFT% 절댓값 기준 ★1~5 변환. ★5 = +30% 이상, ★4 = +10~30%, ★3 = ±10% 이내, ★2 = −10~−20%, ★1 = −20% 이하.
            <br/>
            ★를 클릭하면 일별 판매표 + 산식 모달을 볼 수 있습니다.
          </div>
        </div>
        <div style={{background:D.surfaceAlt,borderRadius:8,padding:"10px 12px",border:`1px solid ${D.border}`}}>
          <div style={{fontSize:11,fontWeight:700,color:"#1d4ed8",letterSpacing:"0.04em",marginBottom:5}}>
            ③ 셀 안 차트 (날짜별 카드) — 흐름 읽기
          </div>
          <div style={{fontSize:11,color:D.textSub,lineHeight:1.6}}>
            X축 가운데 <b style={{color:D.text}}>점선(=0)</b>이 포스트 게시일. 왼쪽 회색 = 포스트 <b>전 14일</b> 판매, 오른쪽 컬러 = 포스트 <b>후 14일</b> 판매. 오른쪽이 왼쪽보다 위로 올라가면 효과 있음.
            아래 <b style={{color:D.text}}>속도</b> = 일평균 추가 판매 장수 / <b style={{color:D.text}}>인사이트</b> = 다음 액션 추천.
          </div>
        </div>
        <div style={{background:D.surfaceAlt,borderRadius:8,padding:"10px 12px",border:`1px solid ${D.border}`}}>
          <div style={{fontSize:11,fontWeight:700,color:"#6d28d9",letterSpacing:"0.04em",marginBottom:5}}>
            ④ 상단 차트 — 월 전체 흐름
          </div>
          <div style={{fontSize:11,color:D.textSub,lineHeight:1.6}}>
            <b style={{color:D.text}}>모든 포스트 평균 효과 곡선</b>: D-14 ~ D+14 일별 평균 판매량. 효과 정점일(D+N)이 표시됩니다.
            <br/>
            <b style={{color:D.text}}>속도계</b>: 월 전반 14일 평균 vs 후반 14일 평균 — 판매가 가속 / 감속 중인지.
          </div>
        </div>
      </div>
    </div>
  );
}

// Pre/Post 14일 일별 qty mini Lift 차트
function LiftMiniChart({ preDays, postDays, color, height=110 }) {
  const data=useMemo(()=>{
    const out=[];
    const pre=Array.isArray(preDays)?preDays:[];
    const post=Array.isArray(postDays)?postDays:[];
    pre.forEach((d,i)=>out.push({d:-pre.length+i,iso:d.iso,pre:d.qty||0,post:null}));
    // 경계점에서 라인 연결을 위해 D=0 가상 분기 (pre 끝 값을 post 시작에도)
    const preLast=pre.length?pre[pre.length-1].qty||0:0;
    out.push({d:0,iso:"D",pre:preLast,post:preLast});
    post.forEach((d,i)=>out.push({d:i+1,iso:d.iso,pre:null,post:d.qty||0}));
    return out;
  },[preDays,postDays]);
  const total=(preDays||[]).reduce((s,d)=>s+(d.qty||0),0)+(postDays||[]).reduce((s,d)=>s+(d.qty||0),0);
  if(total===0){
    return (
      <div style={{height,display:"flex",alignItems:"center",justifyContent:"center",
        background:D.surfaceAlt,borderRadius:6,fontSize:11,color:D.textMeta,
        border:`1px solid ${D.border}`}}>
        판매 데이터 없음
      </div>
    );
  }
  return (
    <div style={{height}}>
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={data} margin={{top:18,right:8,bottom:4,left:4}}>
          <CartesianGrid strokeDasharray="3 3" stroke={D.border} vertical={false}/>
          <XAxis dataKey="d" tick={{fontSize:9,fill:D.textMeta}}
            ticks={[-14,-7,0,7,14]} type="number" domain={[-14,14]}/>
          <YAxis hide={true}/>
          <Tooltip
            contentStyle={{fontSize:11,padding:"4px 8px",borderRadius:5}}
            formatter={(v)=>v===null?"":`${v}장`}
            labelFormatter={(d)=>{const row=data.find(x=>x.d===d);return row?row.iso:`D${d>=0?"+":""}${d}`;}}/>
          <ReferenceLine x={0} stroke={D.black} strokeDasharray="3 3"
            label={{value:"포스트",fontSize:9,fill:D.text,position:"top",offset:6,fontWeight:600}}/>
          <Line type="monotone" dataKey="pre" stroke={D.textMeta} strokeWidth={1.5} dot={false} connectNulls={false}/>
          <Line type="monotone" dataKey="post" stroke="#7BB7E5" strokeWidth={2} dot={false} connectNulls={false}/>
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

// ─────────────────────────────────────────────
// 셀 morph — 캘린더 날짜 셀을 PR1 PostImpactCard 스타일로 치환
//   썸네일 + 날짜 + ★ + Lift% + LiftMiniChart + mini 속도 + 인사이트 한 줄 + 태그 칩
// ─────────────────────────────────────────────
function ImpactCellMorph({ post, score, tags, dateNum, isToday, postsCount, onStarClick, onEditClick, color, cm, setYm }) {
  const stars=score?.stars||0;
  const lift=score?.lift||0;
  const liftCol=lift>=0?"#10b981":"#ef4444";
  // 카드별 속도계: 일평균 attribution qty (post 14일 - pre 14일) / 14
  const dailyAttr=((score?.postQty||0)-(score?.preQty||0))/14;
  // 베스트 매칭 상품 — postPerProduct 가장 높은 항목
  const bestProduct=(()=>{
    const m=score?.postPerProduct||{};
    const list=Object.entries(m).sort(([,a],[,b])=>b-a);
    return list.length&&list[0][1]>0?list[0][0]:null;
  })();
  // 카드별 자동 인사이트 한 줄 (Lift 기반 방향성)
  const insight=(()=>{
    if(lift>=30) return {text:"강한 견인 - 동일 패턴 재구현 추천",tone:"good"};
    if(lift>=10) return {text:"양호한 견인 - 후속 콘텐츠 유지",tone:"good"};
    if(lift>=0)  return {text:"약한 양의 효과 - 콘텐츠 보강 필요",tone:"neutral"};
    if(lift>=-10) return {text:"효과 미미 - 태깅·캡션 재구성",tone:"neutral"};
    return {text:"역효과 - 전략 재검토",tone:"bad"};
  })();
  const toneCol={good:"#15803d",bad:"#b91c1c",neutral:D.textSub}[insight.tone];
  return (
    <div style={{position:"relative",zIndex:2,padding:10,
      display:"flex",flexDirection:"column",flex:1,minHeight:0,gap:5,background:D.surface}}>
      {/* 헤더: 썸네일 + 날짜+★ / Lift% */}
      <div style={{display:"flex",alignItems:"flex-start",gap:8}}>
        <div style={{width:36,height:45,borderRadius:6,overflow:"hidden",background:D.surfaceAlt,flexShrink:0,
          border:`1px solid ${D.border}`}}>
          {post?.thumb_url
            ?<InstagramThumb src={post.thumb_url}/>
            :<div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,color:D.textMeta}}>—</div>}
        </div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:11,color:D.textMeta,marginBottom:2,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
            <span style={{width:6,height:6,background:color,borderRadius:"50%",display:"inline-block"}}/>
            <b style={{color:D.text,fontWeight:700,fontSize:12}}>{dateNum}일</b>
            {isToday&&<span style={{fontSize:9,color:D.blue,fontWeight:700,background:`${D.blue}15`,padding:"1px 5px",borderRadius:8}}>오늘</span>}
            {postsCount>1&&<span style={{fontSize:9,color:D.textMeta}}>· {postsCount}개</span>}
          </div>
          <button onClick={e=>{e.stopPropagation();onStarClick();}}
            title="임팩트 점수 산식 보기"
            style={{background:"none",border:"none",padding:0,cursor:"pointer",
              fontSize:11,letterSpacing:0.4,lineHeight:1,color:"#F2B544",whiteSpace:"nowrap"}}>
            <span>{"★".repeat(stars)}</span>
            <span style={{color:"#cfcfcf"}}>{"★".repeat(5-stars)}</span>
          </button>
        </div>
        <div style={{textAlign:"right",flexShrink:0}}>
          <div style={{fontSize:9,color:D.textMeta,fontWeight:700,letterSpacing:"0.05em"}}>LIFT</div>
          <div style={{fontSize:16,fontWeight:800,color:liftCol,lineHeight:1.1}}>
            {lift>=0?"+":""}{lift.toFixed(1)}%
          </div>
        </div>
      </div>
      {/* 차트 또는 CTA */}
      {tags.length>0&&score
        ?<div style={{flex:1,minHeight:80}}>
            <LiftMiniChart preDays={score.preDays||[]} postDays={score.postDays||[]} color={color} height={"100%"}/>
          </div>
        :<div onClick={e=>{e.stopPropagation();onEditClick();}}
            style={{flex:1,minHeight:80,display:"flex",alignItems:"center",justifyContent:"center",
              background:D.surfaceAlt,borderRadius:6,fontSize:10,color:D.textSub,
              border:`1px dashed ${D.border}`,cursor:"pointer"}}>
            태깅 상품 없음 — 클릭하여 연결
          </div>}
      {/* mini 속도계 + 인사이트 한 줄 (속도 hover popover 안내) */}
      <div style={{display:"flex",gap:6,alignItems:"center",fontSize:10,
        background:D.surfaceAlt,borderRadius:5,padding:"4px 8px"}}>
        <SpeedHoverInfo value={dailyAttr} score={score}>
          <MiniSpeedGauge value={dailyAttr}/>
          <span style={{fontWeight:600,color:D.textMeta}}>속도</span>
          <b style={{color:dailyAttr>=0?"#15803d":"#b91c1c",fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap"}}>
            {dailyAttr>=0?"+":""}{dailyAttr.toFixed(1)}장/일
          </b>
        </SpeedHoverInfo>
        <span style={{flex:1,color:toneCol,textAlign:"right",
          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
          {insight.text}
        </span>
      </div>
      {/* 푸터 — 베스트/태그 칩 + cm 버튼 + ✎ */}
      <div style={{display:"flex",flexWrap:"wrap",gap:3,alignItems:"center",minHeight:20}}>
        {bestProduct&&(
          <span style={{fontSize:9,fontWeight:700,color:"#10b981",
            background:"#10b98112",border:"1px solid #10b98133",
            padding:"2px 6px",borderRadius:10,maxWidth:120,
            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
            🏆 {bestProduct}
          </span>
        )}
        {tags.filter(t=>t!==bestProduct).slice(0,2).map(t=>(
          <span key={t} style={{fontSize:9,color:D.textSub,background:D.surfaceAlt,
            border:`1px solid ${D.border}`,padding:"2px 6px",borderRadius:10,
            maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t}</span>
        ))}
        {tags.length>3&&<span style={{fontSize:9,color:D.textMeta}}>+{tags.length-3}</span>}
        <span style={{flex:1}}/>
        {cm&&<button onClick={e=>{e.stopPropagation();setYm(cm.nextYm);}}
          title={`다음: ${cm.nextYm} (${cm.count}건)`}
          style={{fontSize:9,color:D.blue,fontWeight:700,background:`${D.blue}10`,
            border:`1px solid ${D.blue}30`,borderRadius:8,padding:"1px 5px",cursor:"pointer"}}>
          ▸{parseInt(cm.nextYm.split("-")[1])}월
        </button>}
        <button onClick={e=>{e.stopPropagation();onEditClick();}}
          title="포스트 수정"
          style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
            padding:"1px 6px",fontSize:11,cursor:"pointer",color:D.textMeta,lineHeight:1}}>
          ✎
        </button>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 포스트 임팩트 분석 모드 상단 패널 — KPI · 속도계 · Sales Velocity Timeline · 태그×매출 산점도
// ─────────────────────────────────────────────
function ImpactAnalysisHeader({ summary, monthLabel }) {
  const {avgStars,avgLift,velocityChange,firstAvg,secondAvg,
    bestPost,bestPostTagCount,avgTagCount,avgAttrVelocity,
    cohort,peakDay,peakAvg,preAvg,postCohortAvg}=summary;
  const liftColor=avgLift>=0?"#10b981":"#ef4444";
  const attrVelColor=avgAttrVelocity>=0?"#10b981":"#ef4444";
  const cohortLiftPct=preAvg?((postCohortAvg-preAvg)/preAvg)*100:0;
  return (
    <div style={{marginBottom:16,display:"flex",flexDirection:"column",gap:10}}>
      {/* KPI Row — 6 타일 + 속도계 */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(140px, 1fr))",gap:8}}>
        <KpiTile label="평균 임팩트 점수 ★" value={avgStars.toFixed(1)} unit="/ 5" badge={
          <span style={{fontSize:10,color:"#F2B544",letterSpacing:0.4}}>
            {"★".repeat(Math.round(avgStars))}{"★".repeat(5-Math.round(avgStars)).split("").map((_,i)=><span key={i} style={{color:"#cfcfcf"}}>★</span>)}
          </span>
        }/>
        <KpiTile label="평균 LIFT" value={`${avgLift>=0?"+":""}${avgLift.toFixed(1)}%`} valueColor={liftColor}
          hint="포스트 전 14일 대비 후 14일 태깅 상품 판매 변화율의 평균"/>
        <KpiTile label="평균 태그 상품 수" value={avgTagCount.toFixed(1)} unit="개 / 포스트"
          hint="포스트 1개당 평균 몇 개의 상품을 소개했는지"/>
        <KpiTile label="최고 LIFT 포스트의 상품 태그 수" value={`${bestPostTagCount}`} unit={bestPost?`개 · ${bestPost.post_date}`:"개"}
          hint="가장 효과 좋았던 포스트가 몇 개 상품을 소개했는지"/>
        <KpiTile label="포스트당 태깅 상품 일평균 추가 판매" value={`${avgAttrVelocity>=0?"+":""}${avgAttrVelocity.toFixed(1)}`} unit="장 / 일" valueColor={attrVelColor}
          hint="포스트 1회가 태깅한 상품을 하루 평균 몇 장 더 팔게 만들었는지 (포스트 후 14일 일평균 − 전 14일 일평균)"/>
        <VelocityGauge change={velocityChange} firstAvg={firstAvg} secondAvg={secondAvg}/>
      </div>

      {/* Cohort 평균 효과 곡선 — full width */}
      <div>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:6,gap:10,flexWrap:"wrap"}}>
            <div>
              <div style={{fontSize:13,fontWeight:700,color:D.black}}>모든 포스트 평균 효과 곡선</div>
              <div style={{fontSize:10,color:D.textMeta,marginTop:2,lineHeight:1.45}}>
                D-14 ~ D+14 일별 평균 판매량을 모든 포스트 기준으로 합친 곡선. 효과가 보통 며칠 만에 정점에 도달하는지 보세요.
              </div>
            </div>
          </div>
          {cohort.some(c=>c.n>0)?(
            <>
              <div style={{height:140}}>
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={cohort} margin={{top:14,right:10,left:0,bottom:4}}>
                    <CartesianGrid strokeDasharray="3 3" stroke={D.border} vertical={false}/>
                    <XAxis dataKey="d" tick={{fontSize:10,fill:D.textMeta}} ticks={[-14,-7,0,7,14]}
                      type="number" domain={[-14,14]}/>
                    <YAxis tick={{fontSize:10,fill:D.textMeta}} width={36}
                      label={{value:"평균 (장)",fontSize:9,fill:D.textMeta,angle:-90,position:"insideLeft",offset:10,style:{textAnchor:"middle"}}}/>
                    <Tooltip
                      contentStyle={{fontSize:11,padding:"6px 10px",borderRadius:5}}
                      formatter={(v)=>[`${v.toFixed(1)}장`,"평균 판매"]}
                      labelFormatter={(d)=>`D${d>0?"+":""}${d}`}/>
                    <ReferenceLine x={0} stroke={D.black} strokeDasharray="3 3"
                      label={{value:"포스트일",fontSize:9,fill:D.text,position:"top",fontWeight:600,offset:4}}/>
                    {peakDay>0&&<ReferenceLine x={peakDay} stroke="#15803d" strokeDasharray="2 2"
                      label={{value:`정점 D+${peakDay}`,fontSize:9,fill:"#15803d",position:"top",offset:4}}/>}
                    <Line type="monotone" dataKey="avg" stroke="#7BB7E5" strokeWidth={2.5} dot={{r:2,fill:"#7BB7E5"}}/>
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div style={{fontSize:11,color:D.textSub,marginTop:6,lineHeight:1.55,
                background:D.surfaceAlt,borderRadius:5,padding:"6px 10px"}}>
                <b style={{color:D.text}}>효과 정점</b>: D+{peakDay}일 ({peakAvg.toFixed(1)}장)
                · <b style={{color:D.text}}>전 평균</b> {preAvg.toFixed(1)}장
                → <b style={{color:D.text}}>후 평균</b> {postCohortAvg.toFixed(1)}장
                ({cohortLiftPct>=0?"+":""}{cohortLiftPct.toFixed(0)}%)
              </div>
            </>
          ):(
            <div style={{height:140,display:"flex",alignItems:"center",justifyContent:"center",
              fontSize:11,color:D.textMeta,background:D.surfaceAlt,borderRadius:6}}>
              평균 곡선 표시를 위해 태깅된 포스트가 필요합니다
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

// KPI Tile — 라벨 + 값 + 단위
function KpiTile({ label, value, unit, valueColor, badge, hint }) {
  return (
    <div title={hint||undefined}
      style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:8,padding:"12px 14px",
        display:"flex",flexDirection:"column",gap:4}}>
      <div style={{fontSize:11,color:D.textSub,fontWeight:600,lineHeight:1.35}}>{label}</div>
      <div style={{display:"flex",alignItems:"baseline",gap:3}}>
        <span style={{fontSize:20,fontWeight:700,color:valueColor||D.black,lineHeight:1,fontVariantNumeric:"tabular-nums"}}>{value}</span>
        {unit&&<span style={{fontSize:10,color:D.textMeta}}>{unit}</span>}
      </div>
      {badge&&<div style={{marginTop:2}}>{badge}</div>}
      {hint&&<div style={{fontSize:10,color:D.textMeta,lineHeight:1.4,marginTop:2}}>{hint}</div>}
    </div>
  );
}

// 속도 hover 시 깔끔한 설명 popover (portal로 viewport에 떠 있게)
function SpeedHoverInfo({ value, score, children }) {
  const [pos,setPos]=useState(null); // {x,y}
  const ref=useRef(null);
  const onEnter=()=>{
    const r=ref.current?.getBoundingClientRect();
    if(r) setPos({x:r.left+r.width/2, y:r.top});
  };
  const onLeave=()=>setPos(null);
  const pre=score?.preQty||0;
  const post=score?.postQty||0;
  const delta=post-pre;
  return (
    <>
      <div ref={ref} onMouseEnter={onEnter} onMouseLeave={onLeave}
        style={{display:"flex",alignItems:"center",gap:6,cursor:"help"}}>
        {children}
      </div>
      {pos&&createPortal(
        <div style={{position:"fixed",left:pos.x,top:pos.y-8,transform:"translate(-50%, -100%)",
          zIndex:3500,pointerEvents:"none",
          background:D.black,color:"#fff",padding:"10px 14px",borderRadius:8,
          width:260,fontSize:11,lineHeight:1.6,
          boxShadow:"0 8px 24px rgba(0,0,0,0.22)"}}>
          <div style={{fontWeight:700,fontSize:12,marginBottom:6,letterSpacing:"0.02em"}}>
            속도란?
          </div>
          <div style={{color:"rgba(255,255,255,0.85)"}}>
            태깅한 상품이 <b style={{color:"#fff"}}>포스트 이후</b> 하루 평균 몇 장 더 팔렸는지.
            <br/>
            <span style={{color:"rgba(255,255,255,0.6)"}}>= (후 14일 일평균 − 전 14일 일평균)</span>
          </div>
          <div style={{marginTop:8,paddingTop:8,borderTop:"1px solid rgba(255,255,255,0.15)",
            display:"grid",gridTemplateColumns:"1fr 1fr",gap:4,fontSize:11}}>
            <span style={{color:"rgba(255,255,255,0.55)"}}>전 14일 합</span>
            <b style={{textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{pre.toLocaleString()}장</b>
            <span style={{color:"rgba(255,255,255,0.55)"}}>후 14일 합</span>
            <b style={{textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{post.toLocaleString()}장</b>
            <span style={{color:"rgba(255,255,255,0.55)"}}>증감</span>
            <b style={{textAlign:"right",color:delta>=0?"#7BB7E5":"#fca5a5",fontVariantNumeric:"tabular-nums"}}>
              {delta>=0?"+":""}{delta.toLocaleString()}장
            </b>
            <span style={{color:"rgba(255,255,255,0.55)"}}>일평균</span>
            <b style={{textAlign:"right",color:value>=0?"#7BB7E5":"#fca5a5",fontVariantNumeric:"tabular-nums"}}>
              {value>=0?"+":""}{value.toFixed(2)}장/일
            </b>
          </div>
        </div>,
        document.body
      )}
    </>
  );
}

// 카드용 미니 속도계 — 단일 포스트의 일평균 증분 qty 를 작은 반원 게이지로 (반으로 줄임)
function MiniSpeedGauge({ value }) {
  // -5 ~ +5 장/일 범위로 정규화. 좌끝 -90°, 우끝 +90°
  const clamped=Math.max(-5,Math.min(5,value));
  const angle=(clamped/5)*90;
  const color=value>=1?"#15803d":value>=0?"#3b82f6":value>=-1?"#b45309":"#b91c1c";
  const r=6,cx=7,cy=7;
  return (
    <svg viewBox="0 0 14 9" width={14} height={9} style={{flexShrink:0}}>
      <path d={`M ${cx-r},${cy} A ${r},${r} 0 0 1 ${cx+r},${cy}`}
        stroke={D.border} strokeWidth={1.2} fill="none" strokeLinecap="round"/>
      {(()=>{
        const t=(angle+90)/180;
        const endA=-90+t*180;
        const rad=endA*Math.PI/180;
        const ex=cx+r*Math.cos(rad),ey=cy+r*Math.sin(rad);
        const large=t>0.5?1:0;
        return <path d={`M ${cx-r},${cy} A ${r},${r} 0 ${large} 1 ${ex},${ey}`}
          stroke={color} strokeWidth={1.2} fill="none" strokeLinecap="round"/>;
      })()}
      <line x1={cx} y1={cy} x2={cx+(r-1)*Math.cos((angle-90)*Math.PI/180)} y2={cy+(r-1)*Math.sin((angle-90)*Math.PI/180)}
        stroke={D.black} strokeWidth={0.9} strokeLinecap="round"/>
      <circle cx={cx} cy={cy} r={0.9} fill={D.black}/>
    </svg>
  );
}

// 속도계 게이지 — 가속/감속 % 시각화 (반원 게이지)
function VelocityGauge({ change, firstAvg, secondAvg }) {
  // -50% ~ +50% 범위로 정규화. needle angle = -90deg (좌끝) ~ +90deg (우끝)
  const clamped=Math.max(-50,Math.min(50,change));
  const angle=(clamped/50)*90; // deg
  const color=change>=10?"#10b981":change>=0?"#3b82f6":change>=-10?"#f59e0b":"#ef4444";
  const label=change>=20?"매우 빠름":change>=5?"가속":change>=-5?"보합":change>=-20?"감속":"매우 느림";
  // SVG arc: 반원 (180°). center (50,50), radius 38
  const r=38, cx=50, cy=50;
  return (
    <div style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:8,padding:"8px 10px",
      display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",position:"relative"}}>
      <div style={{fontSize:10,color:D.textMeta,fontWeight:700,letterSpacing:"0.06em",
        textTransform:"uppercase",marginBottom:2,alignSelf:"flex-start"}}>판매 속도</div>
      <svg viewBox="0 0 100 62" width="100%" style={{maxWidth:120}}>
        {/* 배경 호 */}
        <path d={`M ${cx-r},${cy} A ${r},${r} 0 0 1 ${cx+r},${cy}`}
          stroke={D.border} strokeWidth={6} fill="none" strokeLinecap="round"/>
        {/* 컬러 호 — 현재까지 */}
        {(()=>{
          const t=(angle+90)/180; // 0..1
          const endAngle=-90+t*180;
          const rad=endAngle*Math.PI/180;
          const ex=cx+r*Math.cos(rad), ey=cy+r*Math.sin(rad);
          const largeArc=t>0.5?1:0;
          return <path d={`M ${cx-r},${cy} A ${r},${r} 0 ${largeArc} 1 ${ex},${ey}`}
            stroke={color} strokeWidth={6} fill="none" strokeLinecap="round"/>;
        })()}
        {/* needle */}
        <line x1={cx} y1={cy} x2={cx+(r-4)*Math.cos((angle-90)*Math.PI/180)} y2={cy+(r-4)*Math.sin((angle-90)*Math.PI/180)}
          stroke={D.black} strokeWidth={2} strokeLinecap="round"/>
        <circle cx={cx} cy={cy} r={3} fill={D.black}/>
      </svg>
      <div style={{display:"flex",alignItems:"baseline",gap:4,marginTop:-4}}>
        <span style={{fontSize:16,fontWeight:800,color}}>{change>=0?"+":""}{change.toFixed(0)}%</span>
        <span style={{fontSize:9,color:D.textMeta,fontWeight:600}}>{label}</span>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 페이지 하단 자동 인사이트 — 방향성 있는 추천 문구
// ─────────────────────────────────────────────
function ImpactInsightsFooter({ insights }) {
  const toneCol={good:D.text,bad:D.text,neutral:D.text};
  const toneAccent={good:"#15803d",bad:"#b91c1c",neutral:D.textMeta};
  const toneIcon={good:"▲",bad:"▼",neutral:"·"};
  return (
    <Card style={{marginTop:16}}>
      <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:10,letterSpacing:"0.02em"}}>
        분석 인사이트
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {insights.map((ins,i)=>(
          <div key={i} style={{display:"flex",gap:10,alignItems:"flex-start",
            background:D.surface,borderLeft:`2px solid ${toneAccent[ins.tone]}`,
            padding:"7px 12px",fontSize:12,lineHeight:1.55,color:toneCol[ins.tone]}}>
            <span style={{color:toneAccent[ins.tone],fontSize:10,fontWeight:700,lineHeight:1.4,marginTop:1}}>
              {toneIcon[ins.tone]||"·"}
            </span>
            <span>{ins.text}</span>
          </div>
        ))}
      </div>
    </Card>
  );
}

// ─────────────────────────────────────────────
// 임팩트 점수 산식 모달 — 산식 + 업계 근거 + 실제 계산 과정
// ─────────────────────────────────────────────
function ImpactScoreModal({ iso, posts, postScores, onClose }) {
  return (
    <div onClick={onClose}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,
        display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()}
        style={{background:D.surface,borderRadius:14,padding:"24px 28px",
          width:"min(820px,95vw)",maxHeight:"90vh",overflowY:"auto",
          boxShadow:"0 8px 40px rgba(0,0,0,0.22)"}}>
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:14,gap:10}}>
          <div>
            <div style={{fontWeight:700,fontSize:16,color:D.black}}>임팩트 점수 — Pre/Post Sales Lift
              <span style={{fontSize:12,color:D.textMeta,fontWeight:500,marginLeft:6}}>· {iso}</span>
            </div>
            <div style={{fontSize:11,color:D.textMeta,marginTop:3}}>
              포스트 게시 <b>전 14일</b> vs <b>후 14일</b> 태깅 상품 판매 수량 변화율(%)로 5★ 점수 산출 — 인플루언서·콘텐츠 마케팅 측정의 가장 통용되는 방식
            </div>
          </div>
          <button onClick={onClose}
            style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
              padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta,flexShrink:0}}>✕ 닫기</button>
        </div>

        {/* 산식 */}
        <div style={{background:D.surfaceAlt,borderRadius:8,padding:"12px 14px",marginBottom:14,fontSize:12,color:D.text,lineHeight:1.7}}>
          <div style={{fontWeight:700,marginBottom:6}}>산식</div>
          <div style={{fontFamily:"'JetBrains Mono','Courier New',monospace",fontSize:12,marginBottom:8,color:D.textSub}}>
            Lift% = (Post14_qty − Pre14_qty) / max(Pre14_qty, 1) × 100<br/>
            Pre14_qty&nbsp; = 포스트 전 14일간 태깅 상품 총 판매 수량<br/>
            Post14_qty = 포스트 후 14일간 태깅 상품 총 판매 수량
          </div>
          <div style={{fontWeight:700,marginTop:10,marginBottom:6,color:D.text}}>별 등급 (업계 통용 임계치)</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11.5}}>
            <thead><tr style={{color:D.textMeta}}>
              <th style={{textAlign:"left",padding:"4px 6px",fontWeight:600}}>Lift</th>
              <th style={{textAlign:"left",padding:"4px 6px",fontWeight:600}}>의미</th>
              <th style={{textAlign:"center",padding:"4px 6px",fontWeight:600}}>★</th>
            </tr></thead>
            <tbody>
              {[
                ["≥ +100%","매출 2배 이상 — 강한 임팩트",5],
                ["+50 ~ 99%","매우 유의미",4],
                ["+20 ~ 49%","유의미",3],
                ["+5 ~ 19%","약한 양의 효과",2],
                ["−5 ~ +5%","중립 / 변동 범위",1],
                ["≤ −5%","부정 효과",0],
              ].map(([range,desc,n])=>(
                <tr key={range} style={{borderTop:`1px solid ${D.border}`}}>
                  <td style={{padding:"4px 6px",fontFamily:"'JetBrains Mono','Courier New',monospace",color:D.text}}>{range}</td>
                  <td style={{padding:"4px 6px",color:D.textSub}}>{desc}</td>
                  <td style={{padding:"4px 6px",textAlign:"center",letterSpacing:0.5}}>
                    <span style={{color:"#F2B544"}}>{"★".repeat(n)}</span>
                    <span style={{color:"#cfcfcf"}}>{"★".repeat(5-n)}</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{marginTop:10,padding:"6px 10px",background:`${D.amber}10`,border:`1px solid ${D.amber}40`,borderRadius:4,color:D.text,fontSize:11.5}}>
            <b>한계</b>: 통제군 없는 단순 Pre/Post 비교라 계절성·외부 이벤트 영향을 분리하지 못합니다. 정확한 ROI 를 보려면 A/B(Lift study) 또는 MMM(Marketing Mix Modeling) 도구를 함께 사용해야 합니다.
          </div>
        </div>

        {/* 포스트별 실제 계산 */}
        {posts.map((post,pi)=>{
          const s=postScores[post.id];
          if(!s) return null;
          return (
            <div key={post.id} style={{border:`1px solid ${D.border}`,borderRadius:8,padding:"14px 16px",marginBottom:12}}>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:10,marginBottom:10}}>
                <div>
                  <div style={{fontSize:12,color:D.textMeta}}>포스트 {pi+1} · {post.post_date}</div>
                  <a href={post.url} target="_blank" rel="noreferrer"
                    style={{fontSize:11,color:D.blue,wordBreak:"break-all"}}>{post.url}</a>
                </div>
                <div style={{fontSize:18,letterSpacing:1,lineHeight:1,whiteSpace:"nowrap",flexShrink:0}}>
                  <span style={{color:"#F2B544"}}>{"★".repeat(s.stars)}</span>
                  <span style={{color:"#cfcfcf"}}>{"★".repeat(5-s.stars)}</span>
                  <span style={{fontSize:11,color:D.textMeta,marginLeft:6}}>{s.stars}/5</span>
                </div>
              </div>

              {/* 태깅 상품 */}
              <div style={{marginBottom:10}}>
                <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>태깅 상품 ({s.tagged?.length||0}개)</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                  {(s.tagged||[]).map(t=>(
                    <span key={t} style={{padding:"2px 8px",fontSize:11,background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:10}}>{t}</span>
                  ))}
                  {(!s.tagged||s.tagged.length===0)&&<span style={{fontSize:11,color:D.textMeta,fontStyle:"italic"}}>(없음 — 점수 0)</span>}
                </div>
              </div>

              {/* 요약 지표 — Pre / Post / Lift */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:12}}>
                <div style={{background:D.surfaceAlt,borderRadius:6,padding:"8px 10px"}}>
                  <div style={{fontSize:10,color:D.textMeta}}>Pre14_qty <span style={{color:D.textMeta}}>(D−14 ~ D−1)</span></div>
                  <div style={{fontSize:14,fontWeight:700,color:D.text}}>{(s.preQty||0).toLocaleString()}장</div>
                </div>
                <div style={{background:D.surfaceAlt,borderRadius:6,padding:"8px 10px"}}>
                  <div style={{fontSize:10,color:D.textMeta}}>Post14_qty <span style={{color:D.textMeta}}>(D+1 ~ D+14)</span></div>
                  <div style={{fontSize:14,fontWeight:700,color:D.text}}>{(s.postQty||0).toLocaleString()}장</div>
                </div>
                <div style={{background:`${(s.lift||0)>=0?D.green:D.red}10`,
                  border:`1px solid ${(s.lift||0)>=0?D.green:D.red}30`,
                  borderRadius:6,padding:"8px 10px"}}>
                  <div style={{fontSize:10,color:D.textMeta}}>Sales Lift</div>
                  <div style={{fontSize:14,fontWeight:700,color:(s.lift||0)>=0?D.green:D.red}}>
                    {(s.lift||0)>=0?"+":""}{(s.lift||0).toFixed(1)}%
                  </div>
                </div>
              </div>

              {/* 최종 계산 step */}
              <div style={{background:`${D.blue}08`,border:`1px solid ${D.blue}25`,borderRadius:6,padding:"8px 12px",marginBottom:12,
                fontFamily:"'JetBrains Mono','Courier New',monospace",fontSize:11.5,color:D.text,lineHeight:1.7}}>
                Lift% = ({(s.postQty||0).toLocaleString()} − {(s.preQty||0).toLocaleString()}) / max({(s.preQty||0).toLocaleString()}, 1) × 100 = <b>{(s.lift||0).toFixed(1)}%</b><br/>
                ★ = liftToStars({(s.lift||0).toFixed(1)}%) = <b>{s.stars}/5</b>
              </div>

              {/* Pre / Post 일별 breakdown */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                {[
                  {label:"Pre 14일 (D−14 ~ D−1)",rows:s.preDays||[],sign:-1,total:s.preQty||0},
                  {label:"Post 14일 (D+1 ~ D+14)",rows:s.postDays||[],sign:1,total:s.postQty||0},
                ].map((blk,bi)=>(
                  <div key={bi}>
                    <div style={{fontSize:11,color:D.textMeta,marginBottom:4,fontWeight:600}}>{blk.label}</div>
                    <div style={{maxHeight:240,overflowY:"auto",border:`1px solid ${D.border}`,borderRadius:6}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                        <thead style={{background:D.surfaceAlt,position:"sticky",top:0}}>
                          <tr>
                            <th style={{padding:"5px 8px",textAlign:"left",fontWeight:600,color:D.textMeta}}>D{blk.sign>0?"+":"−"}i</th>
                            <th style={{padding:"5px 8px",textAlign:"left",fontWeight:600,color:D.textMeta}}>일자</th>
                            <th style={{padding:"5px 8px",textAlign:"right",fontWeight:600,color:D.textMeta}}>태깅 qty</th>
                          </tr>
                        </thead>
                        <tbody>
                          {blk.rows.map((d,i)=>{
                            const idx=blk.sign<0?(blk.rows.length-i):(i+1);
                            return (
                              <tr key={d.iso} style={{borderTop:`1px solid ${D.border}`}}>
                                <td style={{padding:"4px 8px",color:D.textMeta}}>D{blk.sign>0?"+":"−"}{idx}</td>
                                <td style={{padding:"4px 8px",color:D.text}}>{d.iso}</td>
                                <td style={{padding:"4px 8px",textAlign:"right",
                                  color:d.qty>0?(blk.sign>0?D.blue:D.textSub):D.textMeta,
                                  fontWeight:d.qty>0?600:400}}>{d.qty.toLocaleString()}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr style={{borderTop:`2px solid ${D.border}`,background:D.surfaceAlt,fontWeight:700}}>
                            <td colSpan={2} style={{padding:"5px 8px",color:D.text}}>합계</td>
                            <td style={{padding:"5px 8px",textAlign:"right",color:blk.sign>0?D.blue:D.text}}>{blk.total.toLocaleString()}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// APP ROOT
// ─────────────────────────────────────────────
export default function App() {
  const validPages=["dashboard","flow","promo","input","compare","impact"];
  const hashPage=()=>{const h=window.location.hash.replace("#","");return validPages.includes(h)?h:"dashboard";};
  const [page,setPageState]=useState(hashPage);
  const setPage=useCallback(p=>{window.location.hash=p;setPageState(p);},[]);
  useEffect(()=>{
    const onHash=()=>{const h=window.location.hash.replace("#","");if(validPages.includes(h))setPageState(h);};
    window.addEventListener("hashchange",onHash);
    return()=>window.removeEventListener("hashchange",onHash);
  },[]);
  const [orders,setOrders]=useState([]);
  const [stocks,setStocks]=useState([]);
  const [revenues,setRevenues]=useState([]);
  const [storeSales,setStoreSales]=useState([]);
  const [appLoading,setAppLoading]=useState(true);
  const firstLoad=useRef(true);
  const [ts,setTs]=useState(()=>{
    try{return JSON.parse(localStorage.getItem("merryon_ts")||"null")||{orders:null,stock:null,revenue:null,store:null};}
    catch{return{orders:null,stock:null,revenue:null,store:null};}
  });

  const loadData=useCallback(async()=>{
    const db=await getSupabase();
    const PAGE=1000;

    async function fetchAll(table,orderCol,asc=true){
      let rows=[],offset=0;
      while(true){
        const {data,error}=await db.from(table).select("*").order(orderCol,{ascending:asc}).range(offset,offset+PAGE-1);
        if(error||!data||data.length===0) break;
        rows=rows.concat(data);
        if(data.length<PAGE) break;
        offset+=PAGE;
      }
      return rows;
    }

    const [allHeaders,allItems,allStocks,allRevRaw,allStoreSales,tsRes]=await Promise.all([
      fetchAll("order_headers","order_date",true),
      fetchAll("order_items","item_id",true),
      fetchAll("stock_uploads","upload_date",false),
      fetchAll("revenues","date",false),
      fetchAll("store_sales","sale_date",true),
      db.from("upload_ts").select("*").order("id",{ascending:true}).limit(1),
    ]);

    // 중복 제거: 같은 date+channel은 id 가장 큰 것(최신)만 유지
    const revMap={};
    allRevRaw.forEach(r=>{const k=`${r.date}__${r.channel}`;if(!revMap[k]||r.id>revMap[k].id)revMap[k]=r;});
    const allRevenues=Object.values(revMap);

    // order_headers ⨝ order_items → 기존 orders 단일 행 모양으로 머지
    // (Dashboard/analyze/LogisticsFlow 등 소비처가 같은 필드를 그대로 참조)
    // 스키마: headers={order_no,order_date,channel,payment_amount}, items={item_id,order_no,product_name,option_name,qty,sale_price,status,delivery_date,raw_status}
    const headerMap={};
    allHeaders.forEach(h=>{headerMap[h.order_no]=h;});
    const allOrders=allItems.map(it=>{
      const h=headerMap[it.order_no]||{};
      return {
        order_no:it.order_no,
        order_date:h.order_date||null,
        delivery_date:it.delivery_date||null,   // items에 위치
        channel:h.channel||"",
        payment_amount:h.payment_amount||0,
        product_name:it.product_name,
        option_name:it.option_name,
        qty:it.qty,
        sale_price:it.sale_price,                // amount 컬럼 없음 — analyze()의 r.amount는 자동 폴백 처리됨
        status:it.status,
        raw_status:it.raw_status,
        // 하위 호환 합성 키 (기존 코드의 r.order_id 폴백용)
        order_id:`${h.order_date||""}||${it.order_no}||${it.product_name||""}||${it.option_name||""}`,
      };
    });

    // 매장 반품은 DB에 저장만 두고 집계 로직에서는 제외 (요청사항)
    //   - DataHistoryPanel 은 Supabase 직접 조회라 반품 행 그대로 보임
    //   - 여기서 한 번 필터 → analyze/Dashboard/PromoFlow/ContentImpact 모두 영향 0
    const activeStoreSales=allStoreSales.filter(r=>r.status!=="반품");

    // store_sales → 주문 호환 rows (채널은 "오프라인 스토어"로 정규화)
    const storeOrderRows=activeStoreSales.map(r=>({
      order_date:r.sale_date,
      channel:"오프라인 스토어",
      product_name:r.product_name,
      option_name:r.option_name,
      qty:r.qty,
      status:r.status,
      order_id:r.order_id,
    }));
    // 예약거래는 매장 CSV(store_sales)로 별도 집계 — store_sales가 있으면 orders의 예약거래 행 제외 (이중 합산 방지)
    const hasStoreSales=activeStoreSales.length>0;
    const baseOrders=hasStoreSales
      ?allOrders.filter(o=>String(o.channel||"").trim()!=="예약거래")
      :allOrders;
    setOrders([...baseOrders.map(o=>({...o,channel:normChannel(o.channel)})),...storeOrderRows]);
    setStocks(allStocks);
    setRevenues(allRevenues);
    setStoreSales(activeStoreSales);
    const tsData=tsRes?.data;
    if(tsData&&tsData.length>0){
      const t=tsData[0];
      const next={orders:t.orders||null,stock:t.stock||null,revenue:t.revenue||null,store:t.store||null};
      setTs(next);try{localStorage.setItem("merryon_ts",JSON.stringify(next));}catch{}
    }
    if(firstLoad.current){
      setAppLoading(false);
      firstLoad.current=false;
    }
  },[]);

  useEffect(()=>{ loadData(); },[loadData]);

  const updateTs=useCallback(async(key,val)=>{
    setTs(prev=>{
      const next={...prev,[key]:val};
      try{localStorage.setItem("merryon_ts",JSON.stringify(next));}catch{}
      return next;
    });
    const db=await getSupabase();
    await db.from("upload_ts").upsert({id:1,[key]:val},{onConflict:"id"});
  },[]);

  const nav=[
    {key:"dashboard",label:"대시보드"},
    {key:"compare",label:"데이터 컴페어"},
    {key:"promo",label:"프로모션 플로우"},
    {key:"flow",label:"물류 플로우"},
    {key:"impact",label:"콘텐츠 임팩트"},
    {key:"input",label:"데이터 입력"},
  ];

  const [visible,setVisible]=useState(false);
  useEffect(()=>{if(!appLoading){const t=setTimeout(()=>setVisible(true),30);return()=>clearTimeout(t);}},[appLoading]);

  if(appLoading) return <LoadingScreen/>;

  const isDark=page==="compare";
  const DK={bg:"#0A0A0A",surface:"#141414",border:"#2a2a2a",text:"#F0F0F0",sub:"#888",active:"#242424"};

  return (
    <div style={{ display:"flex", flexDirection:"column", minHeight:"100vh",
      background:isDark?DK.bg:D.bg,
      fontFamily:"'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      color:isDark?DK.text:D.text, fontSize:14,
      opacity:visible?1:0, transition:"opacity 0.35s ease, background 0.2s ease" }}>
      <style>{`input,textarea,select{font-family:'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif;}input::placeholder,textarea::placeholder{font-family:inherit;font-size:inherit;}`}</style>

      {/* top bar — always light */}
      <div style={{ background:D.surface,
        borderBottom:`1px solid ${D.border}`,
        padding:"0 24px", display:"flex", alignItems:"center", gap:24, height:48, flexShrink:0 }}>
        <div style={{ display:"flex", flexDirection:"column", lineHeight:1.1, marginRight:8 }}>
          <span style={{ fontWeight:800, fontSize:13, letterSpacing:"0.08em", color:D.black }}>MERRYON</span>
          <span style={{ fontSize:10, color:D.textMeta, letterSpacing:"0.06em" }}>COMMERCE · Made by Jihoon</span>
        </div>
        <nav style={{ display:"flex", gap:2, flex:1 }}>
          {nav.map(n=>(
            <button key={n.key} onClick={()=>setPage(n.key)}
              style={{ background:page===n.key?D.surfaceAlt:"transparent",
                color:page===n.key?D.black:D.textSub,
                border:"none", borderRadius:6, padding:"6px 14px",
                cursor:"pointer", fontSize:12,
                fontWeight:page===n.key?600:400 }}>
              {n.label}
            </button>
          ))}
        </nav>
        <div style={{ color:D.textMeta, fontSize:11, flexShrink:0, display:"flex", flexDirection:"column", alignItems:"flex-end", gap:1 }}>
          <span>{new Date().toLocaleDateString("ko-KR",{year:"numeric",month:"long",day:"numeric"})}</span>
          <span style={{fontSize:9,opacity:0.5}}>build {__BUILD_TIME__}</span>
        </div>
      </div>

      {/* main content */}
      <div style={{ flex:1, overflowY:"auto" }}>
        {page==="dashboard"&&(
          <Dashboard orders={orders} stocks={stocks} revenues={revenues} storeSales={storeSales} ts={ts}
            onRefresh={loadData}/>
        )}
        {page==="flow"&&<LogisticsFlow orders={orders} stocks={stocks} ts={ts}/>}
        {page==="promo"&&<PromoFlow revenues={revenues} storeSales={storeSales} orders={orders}/>}
        {page==="compare"&&<DataCompare revenues={revenues} storeSales={storeSales} orders={orders}/>}
        {page==="impact"&&<ContentImpact orders={orders} revenues={revenues} storeSales={storeSales}/>}
        {page==="input"&&(
          <DataInput
            onUpdate={updateTs}
            onDataChange={loadData}
            orders={orders}
            stocks={stocks}
            revenues={revenues}
            storeSales={storeSales}
          />
        )}
      </div>
    </div>
  );
}
