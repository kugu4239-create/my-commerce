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
const MUTE_BLUE="#5E81AC"; // 이익률(베타) 뱃지·모달 뮤트 블루 테마
const MUTE_GREEN="#7CA989"; // 상승 뱃지 등 뮤트 그린 테마

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

// 총(gross) 출고 판정 — 실제 출고된 행: 순배송 + 배송후 취소(반품) + 배송후 교환.
// 배송전 취소('취소')·미배송('주문')은 제외. 반품/교환은 별도 status 로 계속 집계되므로
// '배송'(총 출고)은 반품·교환을 부분집합으로 포함한다. (반품률 분모 = 총 출고)
const wasShipped = s => s === "배송" || s === "반품" || s === "교환";

// 상품명 정규화 — SaleCalcModal 내부 정의(normProdName)와 동일 동작을 유지해야 한다
// (calc_supply_override 의 norm_name 키가 양쪽에서 일치해야 매칭됨). 인벤토리 "가격 DB" 업로드와
// 이익률 계산 훅(useInventoryPricing)에서 상품 매칭에 공통 사용.
const normProdName = s => String(s||"").trim().toLowerCase()
  .replace(/[​‌‍ ﻿]/g,"")
  .replace(/[\s_·•~\-]*\d+\s*colou?rs?\b/gi,"")
  .replace(/\[[^\]]*\]/g,"")
  .replace(/\([^)]*\)/g,"")
  .replace(/[_·•~]+/g,"")
  .replace(/\s+/g,"").trim();

// 카페24 상품코드 매칭 전용 정규화 — 공용 normProdName 과 달리 색상 [BLACK]/[BROWN] 등은 보존하고
// 배송 배지(*오늘 출발, *MM.DD 예약 발송, *당일발송 …)만 제거한다. 띄어쓰기는 무시(유연 매칭).
// 저장(cafe24_product_codes.norm_name)과 SKU Risk 매칭에 동일하게 사용해 양쪽 키가 일치해야 한다.
const normCafe24Name = s => String(s||"").split("*")[0].toLowerCase()
  .replace(/[​‌‍ ﻿]/g,"")
  .replace(/\s+/g,"").trim();
// 색상([BLACK]/[WHITE] …) 제거 키 — 한쪽에만 색상이 있을 때를 위한 폴백 매칭용.
const cafe24BaseKey = s => normCafe24Name(s).replace(/\[[^\]]*\]/g,"");
// 색상 토큰 → 대표 hex (KR↔EN 동의어 통합: WHITE≡화이트). 색상으로 인식 안 되면 null.
function colorToHex(tok){
  const t=String(tok||"").trim();
  if(!t) return null;
  return COLOR_HEX[t]||COLOR_HEX[t.toUpperCase()]||COLOR_HEX[t.toLowerCase()]||null;
}
// 문자열에서 색상 hex 추출 — 대괄호 [WHITE] 안이나 옵션값(화이트, 화이트-L 등)에서.
function extractColorHex(...parts){
  for(const p of parts){
    const s=String(p||"");
    for(const b of (s.match(/\[([^\]]+)\]/g)||[])){ const h=colorToHex(b.slice(1,-1)); if(h) return h; }
    for(const seg of s.split(/[\s/\-_·•,]+/)){ const h=colorToHex(seg); if(h) return h; }
  }
  return null;
}

// 이익률 집계 대상 판정 — 취소/교환/반품(환불) 주문만 제외하고
// 배송·주문·접수·발주 등 그 외 상태는 모두 포함한다.
// 내부 status(정규화) + raw_status(원본 CS처리 텍스트) 양쪽을 확인.
const PROFIT_EXCLUDE_STATUS = new Set(["취소","반품","교환"]);
const isProfitCountable = r => {
  if(PROFIT_EXCLUDE_STATUS.has(r.status)) return false;
  const raw=String(r.raw_status||"").toLowerCase().replace(/\s/g,"");
  if(/취소|환불|반품|교환|cancel|refund|return|exchange/.test(raw)) return false;
  return true;
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
  // 한국 표준시(KST, UTC+9) 기준 YYYY-MM-DD — 브라우저/서버 타임존과 무관하게 동일
  return new Date(Date.now()+offsetDays*86400000+32400000).toISOString().slice(0,10);
}
// Date 객체를 로컬 구성요소로 YYYY-MM-DD 포맷 (toISOString의 UTC 밀림 방지 — 로컬=KST 환경 기준)
function ymd(d){ return [d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-'); }

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
    return rows.filter(r => r[dateField] >= localDate(-7) && r[dateField] <= ceiling);
  }
  if (period === "14d") {
    return rows.filter(r => r[dateField] >= localDate(-14) && r[dateField] <= ceiling);
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
  const today = localDate(0);
  const initMonth = () => {
    // YYYY-MM-DD 문자열을 직접 파싱(타임존 무관) — new Date(str)는 UTC 파싱이라 월이 밀릴 수 있음
    const base = (mode==="range" ? (rangeStart||value||today) : (value||today)) || today;
    const [y,m] = String(base).slice(0,10).split("-").map(Number);
    return { y: y||Number(today.slice(0,4)), m: (m||1)-1 };
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
          const isToday = dateStr === today;
          const bg = sel ? C.green : rng ? C.greenBg : "transparent";
          const col = sel ? "#fff" : isToday ? C.green : avail ? C.text : C.dim;
          return (
            <button key={day} onClick={()=>handleClick(dateStr)} disabled={!avail}
              title={isToday?"오늘":undefined}
              style={{ ...btnBase, background:bg, color:col,
                fontWeight:(sel||isToday)?700:avail?500:400,
                cursor:avail?"pointer":"default",
                outline: (isStart(dateStr)||isEnd(dateStr)) ? `2px solid ${C.green}` : "none",
                outlineOffset: -1,
                boxShadow: (!sel && isToday) ? `inset 0 0 0 1.5px ${C.green}` : undefined,
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
  const today = localDate(0);
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
      if (wasShipped(r.status)) { prodMap[key].shipped++;   prodMap[key].byChannel[ch].shipped++; }
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
      if (wasShipped(r.status)) { chanMap[ch].shipped++; chanMap[ch].byProd[r.product_name||"미분류"].shipped++; }
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
// orderRows = 주문일(order_date) 필터, shipRows = 배송일(delivery_date) 필터.
//   주문·매출·객단가 = orderRows(주문일) / 배송·반품·반품률 = shipRows(배송일)
function analyze(orderRows, stockRows, revenueRows, storeRows=[], shipRows=orderRows) {
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

  // ── [배송·반품 KPI] (온라인만, 매장 제외 · 배송일 기준) ─────────────
  // 소스: 이지어드민 orders CSV(배송일 필터) — 오프라인 채널(매장 판매 머지 행 포함)은 제외
  // 계산: ('배송'=총 출고 = 배송+배송후 취소(반품)+배송후 교환. 배송전 취소·미배송 제외)
  //   - 배송 수    = COUNT(DISTINCT order_no||order_id) where 총 출고
  //   - 반품 수    = COUNT(DISTINCT order_no||order_id) where status="반품"
  //   - 배송 장수  = SUM(qty) where 총 출고
  //   - 반품 장수  = SUM(qty) where 반품
  //   - 반품률     = 반품 장수 / 총 출고 장수 * 100 (장수 기준 — 분모는 반품·교환 포함)
  const onlineRows       = orderRows.filter(r=>!isOffline(r));       // 주문일 기준 (주문 KPI)
  const onlineShipRows   = shipRows.filter(r=>!isOffline(r));        // 배송일 기준 (배송·반품 KPI)
  const shippedRows      = onlineShipRows.filter(r=>wasShipped(r.status));
  const returnedRows     = onlineShipRows.filter(r=>r.status==="반품");
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
  // 반품률 = 온라인 반품 qty ÷ 온라인 총 출고 qty × 100 (분모는 반품·교환 포함)
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
  const chOrderIds={};      // 채널별 Set(oid) where 총 출고(배송+반품+교환) — 배송 카운트
  const chReturnedIds={};   // 채널별 Set(oid) where 반품 — 반품 카운트
  const chAllOrderIds={};   // 채널별 Set(oid) 모든 상태 — 판매처 상세의 '주문 수' 컬럼용
  const chOrderedQty={};    // 채널별 SUM(qty) 모든 상태 — '주문 장수'
  const chShippedQty={};    // 채널별 SUM(qty) 총 출고 — '배송 장수' + 반품률 분모(반품·교환 포함)
  const chReturnedQty={};   // 채널별 SUM(qty) 반품   — '반품 장수' + 반품률 분자
  const PAYMENT_CH=new Set(["자사몰"]); // payment_amount(MAX) 사용 채널
  // 주문(모든 상태) + 주문 장수 + 객단가(AOV) — 주문일(order_date) 기준
  orderRows.forEach(r=>{
    if(isExcl(r)) return; // MERRYON OVERSEA 제외
    const ch=r.channel||"미분류";
    // 주문번호 키: 신규 order_no 필드 우선, 없으면 order_id 전체(이전 데이터 호환)
    const oid=r.order_no||r.order_id||"";
    const status=(r.status==="배송"&&/^CORD/i.test(oid))?"반품":r.status;
    const qty=(r.qty||1);
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    // 전체 주문 Set: 모든 상태(배송/반품/교환 등) 포함
    if(oid){
      if(!chAllOrderIds[ch]) chAllOrderIds[ch]=new Set();
      chAllOrderIds[ch].add(oid);
    }
    chOrderedQty[ch]=(chOrderedQty[ch]||0)+qty;
    // 객단가(AOV) 금액 맵 — 순배송 기준
    if(status!=="배송") return;
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
  // 배송/반품 고유 주문번호 + 장수 — 배송일(delivery_date) 기준
  shipRows.forEach(r=>{
    if(isExcl(r)) return;
    const ch=r.channel||"미분류";
    const oid=r.order_no||r.order_id||"";
    const status=(r.status==="배송"&&/^CORD/i.test(oid))?"반품":r.status;
    const qty=(r.qty||1);
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    if(status==="반품"){
      if(oid){
        if(!chReturnedIds[ch]) chReturnedIds[ch]=new Set();
        chReturnedIds[ch].add(oid);
      }
      chReturnedQty[ch]=(chReturnedQty[ch]||0)+qty;
    }
    if(!wasShipped(status)) return;   // 취소·주문 제외 (배송+반품+교환 = 총 출고)
    if(!chOrderIds[ch]) chOrderIds[ch]=new Set();
    chOrderIds[ch].add(oid);
    chShippedQty[ch]=(chShippedQty[ch]||0)+qty;
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
  // 계산: delivery_date(배송일)의 YYYY-MM 단위 그룹 → 배송/반품 라인 카운트, 반품률
  // (라인 카운트 유지: 차트는 추세 시각화 목적이므로 행 단위로 충분)
  const byMonth={};
  shipRows.forEach(r=>{
    const ym=r.delivery_date?r.delivery_date.slice(0,7):null;   // 배송일 기준
    if(!ym) return;
    if(!byMonth[ym]) byMonth[ym]={month:ym,shipped:0,returned:0};
    if(wasShipped(r.status)) byMonth[ym].shipped++;   // 총 출고(배송+반품+교환)
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
  // 로컬(KST) 구성요소로 포맷 — toISOString(UTC)는 KST 새벽에 하루 밀림
  const fmt=d=>[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
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
  // 배송·반품 KPI 전용 — 배송일(delivery_date) 기준 기간 필터 (배송일 없는 행은 기간 집계서 제외)
  const deliveryFilteredOrders=useMemo(()=>filterByDate(orders,"delivery_date",period,customStart,customEnd),[orders,period,customStart,customEnd]);
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
  const stats=useMemo(()=>analyze(filteredOrders,filteredStocks,filteredRevenues,filteredStoreSales,deliveryFilteredOrders),[filteredOrders,filteredStocks,filteredRevenues,filteredStoreSales,deliveryFilteredOrders]);

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

  // 월별 배송량 차트 데이터 — 배송일(delivery_date) 기준 (배송일 없는 행 제외)
  const shippingChartData=useMemo(()=>{
    const today=localDate(-1);
    if(shippingPeriod==="yd"){
      const yStr=localDate(-1);
      const byDay={};
      orders.filter(r=>r.delivery_date===yStr).forEach(r=>{
        if(!byDay[yStr]) byDay[yStr]={date:yStr.slice(5),shipped:0};
        if(wasShipped(r.status)) byDay[yStr].shipped++;
      });
      return Object.values(byDay);
    }
    if(shippingPeriod==="7d"||shippingPeriod==="1m"){
      const c=new Date();
      if(shippingPeriod==="7d") c.setDate(c.getDate()-7);
      else c.setMonth(c.getMonth()-1);
      const cut=ymd(c);
      const byDay={};
      orders.filter(r=>r.delivery_date>=cut&&r.delivery_date<=today).forEach(r=>{
        const d=r.delivery_date;
        if(!byDay[d]) byDay[d]={date:d.slice(5),shipped:0};
        if(wasShipped(r.status)) byDay[d].shipped++;
      });
      return Object.values(byDay).sort((a,b)=>a.date>b.date?1:-1);
    }
    if(shippingPeriod==="custom"&&shippingCustomStart&&shippingCustomEnd){
      const diff=(new Date(shippingCustomEnd)-new Date(shippingCustomStart))/86400000;
      const src=orders.filter(r=>r.delivery_date>=shippingCustomStart&&r.delivery_date<=shippingCustomEnd);
      if(diff<=60){
        const byDay={};
        src.forEach(r=>{const d=r.delivery_date;if(!byDay[d])byDay[d]={date:d.slice(5),shipped:0};if(wasShipped(r.status))byDay[d].shipped++;});
        return Object.values(byDay).sort((a,b)=>a.date>b.date?1:-1);
      }
      const byMonth={};
      src.forEach(r=>{const ym=r.delivery_date?.slice(0,7);if(!ym)return;if(!byMonth[ym])byMonth[ym]={date:ym,shipped:0};if(wasShipped(r.status))byMonth[ym].shipped++;});
      return Object.values(byMonth).sort((a,b)=>a.date>b.date?1:-1);
    }
    const c=new Date(); c.setMonth(c.getMonth()-3);
    const cut=ymd(c);
    const byMonth={};
    orders.filter(r=>r.delivery_date>=cut).forEach(r=>{
      const ym=r.delivery_date?.slice(0,7); if(!ym) return;
      if(!byMonth[ym]) byMonth[ym]={date:ym,shipped:0};
      if(wasShipped(r.status)) byMonth[ym].shipped++;
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
      start=ymd(d);
    } else {
      const d=new Date(); d.setMonth(d.getMonth()-(returnPeriod==="1m"?1:3));
      start=ymd(d);
    }
    const filteredRet=orders.filter(r=>r.delivery_date>=start&&r.delivery_date<=end&&r.channel!=="오프라인 스토어");
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
      const d=r.delivery_date; const ch=r.channel||"미분류";
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
          <div style={{fontSize:10,color:D.textMeta,marginTop:6}}>배송일 기준{getPeriodStr(shippingPeriod,shippingCustomStart,shippingCustomEnd)?` · ${getPeriodStr(shippingPeriod,shippingCustomStart,shippingCustomEnd)}`:""}</div>
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
          // 배송 건: 이지어드민 orders 중 총 출고(배송+배송후 취소/교환), 매장 제외 · 배송일 기준
          const OFFL2=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
          const isOff=r=>OFFL2.has(r.channel||"");
          const shipped=deliveryFilteredOrders.filter(r=>wasShipped(r.status)&&!isOff(r));
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
            const d=r.delivery_date||"—";
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
                소스: <b>주문·배송 업로드 데이터</b> (orders, 실제 출고 = 배송 + 배송후 취소(반품) + 배송후 교환) · <b>배송일 기준</b><br/>
                <b>배송 건</b> = COUNT(DISTINCT 주문번호) where 실제 출고 (배송일 기준)<br/>
                <b>배송 수량(장)</b> = SUM(qty) where 실제 출고 · 반품/교환은 별도 KPI 로도 집계
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
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>배송일별 (최근 30일)</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>배송일</th>
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
          // 반품률 = 반품 수량(장) / 총 출고 수량(장) * 100 — 동기간 내(배송일 기준), 매장 제외
          //   총 출고 = 배송 + 배송후 취소(반품) + 배송후 교환 (반품은 분모에도 포함)
          const OFFL3=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
          const isOff=r=>OFFL3.has(r.channel||"");
          const byCh={};
          deliveryFilteredOrders.filter(r=>!isOff(r)&&wasShipped(r.status)).forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={shippedQty:0,returnedQty:0};
            const q=r.qty||1;
            byCh[ch].shippedQty+=q;                      // 총 출고
            if(r.status==="반품") byCh[ch].returnedQty+=q; // 반품(분자) — 분모에도 포함됨
          });
          const chRows=Object.entries(byCh).sort((a,b)=>b[1].returnedQty-a[1].returnedQty);
          // 매장 반품률 — 매장 판매 데이터에서 별도 계산 (배송 카운트와 무관)
          const storeShippedQty =filteredStoreSales.filter(r=>wasShipped(r.status)).reduce((s,r)=>s+(r.qty||1),0);
          const storeReturnedQty=filteredStoreSales.filter(r=>r.status==="반품").reduce((s,r)=>s+(r.qty||1),0);
          const storeRate=storeShippedQty>0?(storeReturnedQty/storeShippedQty*100):0;
          const hasStore=storeShippedQty>0||storeReturnedQty>0;
          // top return products — 반품 수량(장) 기준 (온라인 + 매장)
          const byProd={};
          deliveryFilteredOrders.filter(r=>!isOff(r)&&r.status==="반품").forEach(r=>{
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
                소스: <b>주문·배송 업로드 데이터</b> (온라인 채널, 총 출고·반품) · <b>매장 판매 데이터</b> (오프라인 스토어)<br/>
                <b>반품률</b> = <b>반품 수량(장) ÷ 총 출고 수량(장)</b> × 100 (총 출고 = 배송+반품+교환 · 온라인과 매장 각각 별도 계산)<br/>
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
  const today=localDate(0);
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
    try{return localStorage.getItem("merryon_aging_date")||localDate(0);}
    catch{return localDate(0);}
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
  const todayStr=localDate(0);
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

// ─────────────────────────────────────────────
// 프로모션 할인율 그리드 — Editor + 표시 컴포넌트
// 저장 구조 (신):
//   { products: { period:{start,end}, rows:[{group,rate}] },
//     coupons:  [{name,rate,start,end,stack}] }   // stack=true → 중복 적용(곱셈 누적)
// 구버전 호환 (products 가 배열인 경우 첫 행의 start/end 를 공통 period 로 마이그레이트)
// ─────────────────────────────────────────────
function emptyProductRow(){return{group:"",rate:"",markup:"",cpn:0,products:[]};}
// 쿠폰 타입 모델 — same-type 끼리 중복 불가, share 는 누구와도 중복 불가, 그 외 cross-type 만 누적
const COUPON_TYPES=[
  {key:"product",label:"상품 쿠폰",  short:"상품",   color:D.blue,  bg:`${D.blue}10`,  border:`${D.blue}55`},
  {key:"cart",   label:"장바구니 쿠폰",short:"장바구니",color:D.green, bg:`${D.green}10`, border:`${D.green}55`},
  {key:"share",  label:"분담 쿠폰",  short:"분담",   color:D.amber, bg:`${D.amber}10`, border:`${D.amber}55`},
];
const COUPON_TYPE_BY_KEY=Object.fromEntries(COUPON_TYPES.map(t=>[t.key,t]));
function couponTypeOf(c){const t=c?.type;return t&&COUPON_TYPE_BY_KEY[t]?t:"product";}
function couponUnitOf(c){return c?.unit==="won"?"won":"pct";}
function canStack(a,b){const ta=couponTypeOf(a),tb=couponTypeOf(b);if(ta==="share"||tb==="share") return false;return ta!==tb;}
function emptyCouponRow(type="product"){return{name:"",rate:"",start:"",end:"",type,unit:"pct",stack:false,excludeGroups:[],stacksWith:[]};}
// 쿠폰 표시 이름 (이름 비어있으면 `쿠폰N`)
function couponDisplayName(c,i){const n=(c?.name||"").trim();return n||`쿠폰${i+1}`;}
function normalizePlan(p){
  const coupons=(Array.isArray(p?.coupons)?p.coupons:[]).map(c=>({
    name:c.name||"",
    rate:c.rate||"",
    start:c.start||"",
    end:c.end||"",
    type:(c?.type&&COUPON_TYPE_BY_KEY[c.type])?c.type:"product",
    unit:c?.unit==="won"?"won":"pct",
    stack:!!c.stack,
    excludeGroups:Array.isArray(c.excludeGroups)?c.excludeGroups:[],
    stacksWith:Array.isArray(c.stacksWith)?c.stacksWith:[],
  }));
  // 신 포맷
  if(p?.products&&!Array.isArray(p.products)&&Array.isArray(p.products.rows)){
    return{
      products:{
        period:{start:p.products.period?.start||"",end:p.products.period?.end||""},
        rows:p.products.rows.map(r=>({group:r.group||"",rate:r.rate||"",markup:r.markup||"",cpn:r.cpn||0,products:Array.isArray(r.products)?r.products:[]})),
      },
      coupons,
    };
  }
  // 구 포맷: 행마다 start/end 있던 경우 → 첫 비어있지 않은 행의 기간을 공통 period 로
  if(Array.isArray(p?.products)){
    const first=p.products.find(r=>r.start||r.end);
    return{
      products:{
        period:{start:first?.start||"",end:first?.end||""},
        rows:p.products.map(r=>({group:r.group||"",rate:r.rate||"",markup:r.markup||"",cpn:r.cpn||0,products:Array.isArray(r.products)?r.products:[]})),
      },
      coupons,
    };
  }
  return{products:{period:{start:"",end:""},rows:[]},coupons};
}

// 할인율 매트릭스 — 곱연산(가격 기준): 최종 = 1-(1-d_p)*factor
//   열: 상품할인 + 쿠폰별 개별 시나리오 (중복 쿠폰도 겹치지 않고 한 장씩 따로 표시)
//   행: 상품군
function computeDiscountMatrix(plan){
  const p=normalizePlan(plan);
  const groups=p.products.rows.filter(r=>(r.group||"").trim()||(+r.rate||0)>0)
    .map(r=>({group:(r.group||"").trim()||"전체",rate:+r.rate||0,markup:r.markup?parseFloat(r.markup):null,products:Array.isArray(r.products)?r.products:[],cpn:r.cpn||0}));
  const coupons=p.coupons.filter(c=>(+c.rate||0)>0||(c.name||"").trim());
  const fin=(dp,factor)=>Math.round((1-(1-dp/100)*factor)*1000)/10;
  const exOf=c=>Array.isArray(c.excludeGroups)?c.excludeGroups:[];

  // 컬럼 구성:
  //   1) 프런트 할인
  //   2) 단독 (combo=false): product / cart 단독 컬럼 (다른 쿠폰과 누적 가능하므로 비-final)
  //   3) 조합 (combo=true): 분담 단독 (어차피 중복 불가 → final 비교용) + cart×product 쌍
  const cols=[{key:"prod",label:"프런트 할인"}];
  coupons.forEach((c,i)=>{
    if(couponTypeOf(c)==="share") return; // share 는 조합 그룹으로 이동
    const nm=couponDisplayName(c,i);
    const rate=+c.rate||0;
    const tInfo=COUPON_TYPE_BY_KEY[couponTypeOf(c)];
    cols.push({
      key:"c"+i,coupon:true,combo:false,indexes:[i],
      name:nm,sub:`(${tInfo.short}·${rate}%)`,label:`${nm} (${tInfo.short}·${rate}%)`,
    });
  });
  // 분담 쿠폰 단독 (어차피 누구와도 누적 불가 → final 그 자체)
  coupons.forEach((c,i)=>{
    if(couponTypeOf(c)!=="share") return;
    const nm=couponDisplayName(c,i);
    const rate=+c.rate||0;
    const tInfo=COUPON_TYPE_BY_KEY[couponTypeOf(c)];
    cols.push({
      key:"c"+i,coupon:true,combo:true,indexes:[i],
      name:nm,sub:`(${tInfo.short}·${rate}%)`,label:`${nm} (${tInfo.short}·${rate}%)`,
    });
  });
  // 누적(pair) — 모든 canStack 쌍을 cartesian product 로 펼침
  //  · 상품 쿠폰을 먼저, 장바구니 쿠폰을 나중으로 배치 (실제 적용 순서와 동일하게 표시)
  for(let i=0;i<coupons.length;i++){
    for(let j=i+1;j<coupons.length;j++){
      if(!canStack(coupons[i],coupons[j])) continue;
      let ii=i,jj=j;
      if(couponTypeOf(coupons[i])==="cart" && couponTypeOf(coupons[j])==="product"){
        ii=j; jj=i; // 상품 쿠폰을 먼저 표시
      }
      const ci=coupons[ii],cj=coupons[jj];
      const ti=COUPON_TYPE_BY_KEY[couponTypeOf(ci)],tj=COUPON_TYPE_BY_KEY[couponTypeOf(cj)];
      const ni=couponDisplayName(ci,ii),nj=couponDisplayName(cj,jj);
      const ri=+ci.rate||0,rj=+cj.rate||0;
      cols.push({
        key:"cp"+ii+"_"+jj,coupon:true,combo:true,indexes:[ii,jj],
        name:`${ni} × ${nj}`,
        sub:`(${ti.short} ${ri}% + ${tj.short} ${rj}%)`,
        label:`${ni} × ${nj} (${ti.short}+${tj.short})`,
      });
    }
  }
  // 조합(combo=true) 컬럼들에 케이스 번호 부여 — case 1, case 2, ...
  {let n=0; cols.forEach(c=>{ if(c.combo) c.caseNum=++n; });}

  // 각 행: 컬럼별로 indexes 의 쿠폰 중 하나라도 제외되면 미적용,
  //         아니면 factor = ∏(1 - rate/100) 누적 → fin(g.rate, factor)
  const rows=groups.map(g=>{
    const cells={prod:fin(g.rate,1)};
    cols.forEach(col=>{
      if(col.key==="prod") return;
      const idxs=col.indexes||[];
      const excluded=idxs.some(idx=>exOf(coupons[idx]).includes(g.group));
      if(excluded){
        cells[col.key]=null;
        return;
      }
      const factor=idxs.reduce((f,idx)=>f*(1-(+coupons[idx].rate||0)/100),1);
      cells[col.key]=fin(g.rate,factor);
    });
    return {group:g.group,rate:g.rate,markup:g.markup,products:g.products,cpn:g.cpn,cells};
  });

  return {groups,coupons,cols,rows,hasGroup:groups.length>0,hasCoupon:coupons.length>0};
}

// 상품군×시나리오 매트릭스 표 (에디터·등록 카드 공용)
function DiscountMatrix({ plan, compact=false, circledKeys, onToggleCircle }){
  const m=computeDiscountMatrix(plan);
  const [localCircled,setLocalCircled]=useState(()=>new Set());
  // 묶음 상품 보기 — 저장된 매트릭스에서도 클릭 시 인라인 펼침
  const [bundleOpenIdx,setBundleOpenIdx]=useState(null);
  if(!m.hasGroup) return null;
  // 값 클릭 시 파란 원 강조 토글. onToggleCircle 있으면 제어형(저장·공유), 없으면 로컬
  const controlled=!!onToggleCircle;
  const circled=controlled?new Set(circledKeys||[]):localCircled;
  const toggleCircle=k=>controlled?onToggleCircle(k):setLocalCircled(prev=>{const s=new Set(prev);s.has(k)?s.delete(k):s.add(k);return s;});
  const cell={padding:compact?"2px 6px":"4px 8px",fontSize:compact?10:11,textAlign:"center",whiteSpace:"nowrap"};
  const th={...cell,color:D.textSub,fontWeight:600,borderBottom:`1px solid ${D.border}`,verticalAlign:"bottom"};
  // 구분선:
  //   - 프런트 → 쿠폰 단독: 첫 쿠폰 열 좌측
  //   - 쿠폰 단독 → 쿠폰 누적: 첫 combo 열 좌측
  const divAt=(c,ci)=>{
    const prev=m.cols[ci-1];
    if(c.coupon&&!prev?.coupon) return {borderLeft:`2px solid ${D.borderMid}`};
    if(c.combo&&!prev?.combo) return {borderLeft:`2px solid ${D.borderMid}`};
    return null;
  };
  return (
    <div style={{overflowX:"auto",marginTop:6,display:"flex",justifyContent:"center"}}>
      <style>{`
        .mat-tip:hover::after{
          content:attr(data-tip);
          position:absolute;
          bottom:calc(100% + 6px);
          left:50%;
          transform:translateX(-50%);
          background:${D.black};
          color:#fff;
          padding:8px 12px;
          border-radius:6px;
          font-size:11px;
          font-family:'Noto Sans KR','Pretendard',sans-serif;
          font-weight:400;
          white-space:pre-line;
          z-index:1000;
          pointer-events:none;
          box-shadow:0 4px 16px rgba(0,0,0,0.22);
          max-width:340px;
          width:max-content;
          line-height:1.55;
          text-align:left;
          letter-spacing:-0.01em;
        }
        .mat-tip:hover::before{
          content:'';
          position:absolute;
          bottom:calc(100% + 1px);
          left:50%;
          transform:translateX(-50%);
          border:5px solid transparent;
          border-top-color:${D.black};
          z-index:1001;
          pointer-events:none;
        }
      `}</style>
      <table style={{borderCollapse:"collapse",fontSize:compact?10:11}}>
        <thead>
          {/* Row 1: Case N 라벨 — 별도 행으로 분리해 컬럼 좌측 디바이더가 라벨 위로 올라오지 않게 함 */}
          {m.cols.some(c=>c.combo&&c.caseNum)&&(
            <tr>
              <th style={{padding:"0 0 4px",border:"none"}}/>
              <th style={{padding:"0 0 4px",border:"none"}}/>
              {m.cols.map(c=>(
                <th key={"case"+c.key} style={{padding:"0 8px 4px",border:"none",textAlign:"center"}}>
                  {c.combo&&c.caseNum&&(
                    <span style={{fontSize:compact?9:10,fontWeight:700,color:D.blue,letterSpacing:"0.04em"}}>
                      Case {c.caseNum}
                    </span>
                  )}
                </th>
              ))}
            </tr>
          )}
          {/* Row 2: 기존 컬럼 헤더 (디바이더는 이 행부터 시작) */}
          <tr>
            <th style={{...th,textAlign:"left"}}>상품군</th>
            <th style={{...th,textAlign:"center",color:D.green}} title="시나리오(프런트 할인×쿠폰) 적용 후 실수령 ÷ 원가 · ×3 이하 적색">시나리오 적용 마크업</th>
            {m.cols.map((c,ci)=>(
              <th key={c.key} style={{...th,...divAt(c,ci)}} title={c.label}>
                {c.name?(
                  <div style={{display:"flex",flexDirection:"column",alignItems:"center"}}>
                    <span style={{maxWidth:170,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</span>
                    <span style={{color:c.combo?D.blue:D.textMeta}}>{c.sub}</span>
                  </div>
                ):c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {m.rows.map((r,i)=>{
            const products=Array.isArray(r.products)?r.products:[];
            const hasBundle=products.length>0;
            const isOpen=bundleOpenIdx===i;
            return (<React.Fragment key={i}>
            <tr>
              <td style={{...cell,textAlign:"left",color:D.textSub,maxWidth:280,overflow:"hidden",textOverflow:"ellipsis"}} title={r.markup!=null?`${r.group} · 시나리오 적용 마크업 ×${r.markup.toFixed(2)}`:r.group}>
                {r.group}
                {hasBundle&&(
                  <button onClick={()=>setBundleOpenIdx(isOpen?null:i)}
                    title={`묶음 상품 ${products.length}개 보기`}
                    style={{background:isOpen?"#fff":D.black,color:isOpen?D.black:"#fff",
                      border:`1px solid ${D.black}`,borderRadius:4,
                      padding:"1px 6px",fontSize:compact?9:10,cursor:"pointer",fontWeight:700,marginLeft:6}}>
                    {isOpen?"▾":"▸"} 묶음 {products.length}
                  </button>
                )}
              </td>
              <td style={{...cell,textAlign:"center",fontWeight:700}}>
                {r.markup!=null&&!isNaN(r.markup)
                  ?<span style={{color:r.markup<=3?D.red:D.green}}>×{r.markup.toFixed(2)}</span>
                  :<span style={{color:D.textMeta,fontWeight:400}}>—</span>}
              </td>
              {m.cols.map((c,ci)=>{
                const v=r.cells[c.key];
                // 누적(combo) 셀은 강조, 단독 쿠폰 셀은 보조 색, 프런트 셀은 기본 색
                const isCombo=!!c.combo;
                const k=(r.group||"전체")+"|"+c.key;
                // 마우스 오버 툴팁 — 계산 방식 안내
                const tip=(()=>{
                  if(c.key==="prod") return `${r.group} 프런트 할인 ${r.rate}%`;
                  if(v==null){
                    const ex=(c.indexes||[]).filter(idx=>{
                      const ec=m.coupons[idx]; const eg=Array.isArray(ec.excludeGroups)?ec.excludeGroups:[]; return eg.includes(r.group);
                    }).map(idx=>couponDisplayName(m.coupons[idx],idx));
                    return ex.length?`연관 없음 — ${ex.join(", ")} 는 ${r.group}와 연관 없음`:"연관 없음";
                  }
                  const parts=[`${r.group} 프런트 할인 ${r.rate}%`];
                  let hasShare=false;
                  (c.indexes||[]).forEach(idx=>{
                    const cp=m.coupons[idx];
                    const t=COUPON_TYPE_BY_KEY[couponTypeOf(cp)];
                    if(t.key==="share") hasShare=true;
                    parts.push(`${couponDisplayName(cp,idx)}(${t.short}) ${+cp.rate||0}%`);
                  });
                  let s=parts.join(" × ")+` = ${v}%`;
                  if(hasShare) s+="\n* 분담 쿠폰은 다른 쿠폰과 조합이 불가합니다.";
                  return s;
                })();
                return <td key={c.key} onClick={v==null?undefined:()=>toggleCircle(k)}
                  className="mat-tip" data-tip={tip}
                  style={{...cell,...divAt(c,ci),cursor:v==null?"default":"pointer",position:"relative",
                  fontWeight:v==null?500:(isCombo?700:500),
                  color:v==null?D.textMeta:(isCombo?D.blue:D.textSub)}}>
                  {v==null?"연관 없음":(circled.has(k)
                    ?<span style={{display:"inline-block",border:`2px solid ${D.blue}`,borderRadius:"50%",padding:compact?"1px 6px":"2px 9px",lineHeight:1}}>{v}%</span>
                    :v+"%")}
                </td>;
              })}
            </tr>
            {isOpen&&hasBundle&&(
              <tr>
                <td colSpan={2+m.cols.length} style={{padding:"0 6px 8px",background:D.surfaceAlt}}>
                  <div style={{margin:"4px 0",border:`1px solid ${D.borderMid}`,borderRadius:6,overflow:"hidden",background:D.surface}}>
                    <div style={{padding:"6px 10px",background:D.surfaceAlt,fontSize:10,fontWeight:700,color:D.black,
                      display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <span>{r.group} · 묶음 상품 {products.length}개{(r.cpn||0)>0?` · 쿠폰율 ${r.cpn}%`:""}</span>
                      <button onClick={()=>setBundleOpenIdx(null)}
                        style={{background:"none",border:"none",cursor:"pointer",color:D.textMeta,fontSize:11}}>✕</button>
                    </div>
                    <div style={{maxHeight:280,overflow:"auto"}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
                        <thead><tr style={{background:D.surfaceAlt,color:D.textMeta}}>
                          {["상품명","정가","쿠폰율","기본 할인율","프런트 판매가","최종 노출가","최종 할인율","자사부담","수수료","채널보전","자사 정산","공급가","마진","마크업"].map((h,k)=>(
                            <th key={k} style={{padding:"4px 8px",textAlign:k===0?"left":"right",fontWeight:600,position:"sticky",top:0,background:D.surfaceAlt,whiteSpace:"nowrap"}}>{h}</th>
                          ))}
                        </tr></thead>
                        <tbody>
                          {products.map((p,j)=>{
                            const sv=p.supplyIncVat||Math.round((p.supply||0)*1.1);
                            const won=n=>"₩"+(Math.round(n||0)).toLocaleString();
                            return (
                              <tr key={j} style={{borderTop:`1px solid ${D.border}`}}>
                                <td title={p.name} style={{padding:"3px 8px",maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap"}}>{won(p.list)}</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap"}}>{r.cpn||0}%</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{p.baseDisc||0}%</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.blue,background:"#eef3ff",fontWeight:600}}>{won(p.basePrice)}</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.textSub}}>{won(p.finalPrice)}</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.blue,background:"#eef3ff",fontWeight:700}}>{p.finalDisc||0}%</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.selfBurden||0)>0?D.red:D.textMeta}}>
                                  {(p.selfBurden||0)>0?`−${won(p.selfBurden)}`:"—"}
                                </td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.red}}>−{won(p.fee||0)} <span style={{fontSize:9,color:D.textMeta}}>({p.feeRate||0}%)</span></td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.channelBurden||0)>0?D.blue:D.textMeta}}>
                                  {(p.channelBurden||0)>0?`+${won(p.channelBurden)}`:"—"}
                                </td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{won(p.net||0)}</td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.supply||0)>0?D.text:D.textMeta}}>
                                  {(p.supply||0)>0?won(sv):"—"}
                                </td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600,
                                  color:(p.supply||0)>0?((p.margin||0)>=0?D.text:D.red):D.textMeta}}>
                                  {(p.supply||0)>0?won(p.margin||0):"—"}
                                </td>
                                <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:700,
                                  color:(p.supply||0)>0?((p.markup||0)>3?D.green:D.red):D.textMeta}}>
                                  {(p.supply||0)>0?`×${(p.markup||0).toFixed(2)}`:"—"}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </td>
              </tr>
            )}
            </React.Fragment>);
          })}
        </tbody>
      </table>
    </div>
  );
}

function DiscountPlanEditor({ value, onChange, calOpenFor, setCalOpenFor, idPrefix="dp", platform }) {
  const plan=normalizePlan(value);
  // 표시용 행: 비어 있으면 기본 3행/1행. 입력 중에는 빈 행도 그대로 유지 (입력 끊김·삭제 버그 방지)
  const productRows=plan.products.rows.length?plan.products.rows:[emptyProductRow(),emptyProductRow(),emptyProductRow()];
  const coupons    =plan.coupons.length?plan.coupons:[emptyCouponRow()];
  const [calcOpen,setCalcOpen]=useState(false);
  const [dragIdx,setDragIdx]=useState(null); // 쿠폰 드래그 중 인덱스
  const [prodDragIdx,setProdDragIdx]=useState(null); // 상품군 드래그 중 인덱스
  const [bundleViewIdx,setBundleViewIdx]=useState(null); // 묶음 상품 보기 펼친 행 인덱스
  // 인라인 계산기 — null | { platform: '29CM'|'자사몰', targetRowIdx: number|null, initialCoupon?: number, initialCouponName?: string }
  //   targetRowIdx 가 null 이면 결과는 새 상품군 행들로 append
  //   숫자면 그 행의 묶음만 채움 (단일 행)
  const [inlineCalc,setInlineCalc]=useState(null);
  // 묶음 추가 직전 — 기존 쿠폰 사용/새 쿠폰 선택 프롬프트
  //   null | { platform, targetRowIdx, step: 'choose'|'pick', selectedIdxs: Set<number> }
  const [couponPrompt,setCouponPrompt]=useState(null);

  // 빈 행 필터링은 저장 시점이 아닌 곳에선 하지 않음 — UI 행 상태 유지
  const setProductPeriod=(field,v)=>onChange({
    products:{period:{...plan.products.period,[field]:v},rows:productRows},
    coupons,
  });
  const setProductRows=(arr)=>onChange({
    products:{period:plan.products.period,rows:arr},
    coupons,
  });
  const setCoupons=(arr)=>onChange({
    products:plan.products,
    coupons:arr,
  });
  // 묶음 추가 흐름 진입 — 기존 쿠폰이 있으면 '기존 쿠폰 사용 / 새 쿠폰' 프롬프트 노출.
  //   없으면 곧바로 계산기 모달 오픈.
  const requestInlineCalc=(platformPick,targetRowIdx)=>{
    const usableCoupons=coupons.filter(c=>(+c.rate||0)>0);
    if(usableCoupons.length>0){
      setCouponPrompt({platform:platformPick,targetRowIdx,step:"choose",selectedIdxs:new Set()});
    }else{
      setInlineCalc({platform:platformPick,targetRowIdx});
    }
  };
  // 인라인 계산기 결과 머지
  //   payload: { platform, discount_plan:{ products:{period,rows}, coupons } }
  //   - targetRowIdx 가 null  → 새 상품군 행을 append, 새 쿠폰은 기존 상품군과 연관 없음
  //   - targetRowIdx 가 숫자   → 해당 행 묶음만 채움 (rate/group 은 사용자가 정한 값 유지),
  //                              새 쿠폰은 그 행에만 적용 (다른 기존 행과는 연관 없음)
  //   동률 쿠폰 머지: 새 쿠폰이 기존 쿠폰과 같은 (rate, type) 이면 새 쿠폰 추가 없이
  //                  기존 쿠폰이 새 상품군에도 적용되도록 excludeGroups 만 조정.
  const attachInlineCalc=(payload,targetRowIdx)=>{
    const incRows=Array.isArray(payload?.discount_plan?.products?.rows)?payload.discount_plan.products.rows:[];
    const incCouponsRaw=Array.isArray(payload?.discount_plan?.coupons)?payload.discount_plan.coupons:[];
    const prevRows=productRows;
    const prevCoupons=coupons;
    // 의미있는 새 쿠폰만 (rate>0 or name)
    const incCoupons=incCouponsRaw.filter(c=>(+c.rate||0)>0||(c.name||"").trim());
    // 새 쿠폰 → 기존 쿠폰 매칭 (rate + type 동일하면 동률로 간주)
    const matchIdxOf=(nc)=>{
      const nr=+nc.rate||0, nt=nc.type||"product";
      if(nr===0) return -1;
      return prevCoupons.findIndex(c=>(+c.rate||0)===nr&&(c.type||"product")===nt);
    };
    const matchMap=incCoupons.map(matchIdxOf); // -1 = no match
    const unmatchedNewCoupons=incCoupons.filter((_,i)=>matchMap[i]===-1);
    const matchedPrevIdxs=new Set(matchMap.filter(i=>i>=0));
    if(targetRowIdx==null){
      // 새 행 append + 새 쿠폰 append. 기존 ↔ 새 교차는 excludeGroups 로 차단
      const preExistingGroupNames=prevRows
        .filter(r=>(r.group||"").trim()||(+r.rate||0)>0)
        .map(r=>(r.group||"").trim()||"전체");
      const newGroupNames=incRows
        .filter(r=>(r.group||"").trim()||(+r.rate||0)>0)
        .map(r=>(r.group||"").trim()||"전체");
      // 비매칭 새 쿠폰: 기존 상품군에 미적용
      const taggedNewCoupons=unmatchedNewCoupons.map(c=>({
        ...emptyCouponRow(c.type||"product"),
        ...c,
        excludeGroups:[...new Set([...(Array.isArray(c.excludeGroups)?c.excludeGroups:[]),...preExistingGroupNames])],
      }));
      // 기존 쿠폰: 매칭된 것은 새 상품군에도 적용, 아니면 새 상품군 제외
      const updatedPrevCoupons=prevCoupons.map((c,i)=>{
        const ex=Array.isArray(c.excludeGroups)?c.excludeGroups:[];
        if(matchedPrevIdxs.has(i)){
          // 동률 매칭 — 새 상품군은 적용 (excludeGroups 에서 제거)
          return {...c,excludeGroups:ex.filter(g=>!newGroupNames.includes(g))};
        }
        return {...c,excludeGroups:[...new Set([...ex,...newGroupNames])]};
      });
      const cleanedPrevRows=prevRows.filter(r=>(r.group||"").trim()||(+r.rate||0)>0||(Array.isArray(r.products)&&r.products.length>0));
      onChange({
        products:{period:plan.products.period,rows:[...cleanedPrevRows,...incRows]},
        coupons:[...updatedPrevCoupons.filter(c=>(+c.rate||0)>0||(c.name||"").trim()),...taggedNewCoupons],
      });
    }else{
      // 단일 행 묶음 채움 — incRows 의 모든 product 를 한 묶음으로 병합
      const allProducts=incRows.flatMap(r=>Array.isArray(r.products)?r.products:[]);
      const matchedProducts=allProducts.filter(p=>(p.supply||0)>0||(p.supplyIncVat||0)>0);
      const avgMarkup=matchedProducts.length>0
        ?Math.round(matchedProducts.reduce((s,p)=>s+(+p.markup||0),0)/matchedProducts.length*100)/100
        :null;
      const cpnFromRows=incRows.find(r=>(+r.cpn||0)>0)?.cpn;
      const cpnFromCoupons=incCoupons.find(c=>(+c.rate||0)>0)?.rate;
      const cpn=Number(cpnFromRows||cpnFromCoupons||0)||0;
      const targetGroupName=(prevRows[targetRowIdx]?.group||"").trim()||"전체";
      const otherPreExistingGroupNames=prevRows
        .filter((r,i)=>i!==targetRowIdx&&((r.group||"").trim()||(+r.rate||0)>0))
        .map(r=>(r.group||"").trim()||"전체");
      const nextRows=prevRows.map((r,i)=>i===targetRowIdx?{
        ...r,
        markup:avgMarkup!=null?String(avgMarkup):r.markup,
        cpn,
        products:allProducts,
      }:r);
      // 비매칭 새 쿠폰: 이 행 제외한 기존 상품군에 미적용
      const taggedNewCoupons=unmatchedNewCoupons.map(c=>({
        ...emptyCouponRow(c.type||"product"),
        ...c,
        excludeGroups:[...new Set([...(Array.isArray(c.excludeGroups)?c.excludeGroups:[]),...otherPreExistingGroupNames])],
      }));
      // 동률 매칭된 기존 쿠폰 — 이 행은 반드시 적용 (excludeGroups 에서 이 행 제거)
      const updatedPrevCoupons=prevCoupons.map((c,i)=>{
        if(!matchedPrevIdxs.has(i)) return c;
        const ex=Array.isArray(c.excludeGroups)?c.excludeGroups:[];
        return {...c,excludeGroups:ex.filter(g=>g!==targetGroupName)};
      });
      onChange({
        products:{period:plan.products.period,rows:nextRows},
        coupons:[...updatedPrevCoupons.filter(c=>(+c.rate||0)>0||(c.name||"").trim()),...taggedNewCoupons],
      });
    }
    setInlineCalc(null);
  };

  // 29CM 계산기에서 선택한 조합을 상품군·쿠폰 입력에 반영
  // {tier, baseDisc, primaryCoupon, stackRates:[중복쿠폰율 ...]}
  const applyCalc=({tier,baseDisc,primaryCoupon,stackRates=[]})=>{
    const groupLabel=`${tier.name} ${tier.range}`;
    const filledProducts=productRows.filter(r=>r.group||r.rate);
    const idx=filledProducts.findIndex(r=>r.group===groupLabel);
    const nextProductRows=idx>=0
      ?filledProducts.map((r,i)=>i===idx?{...r,rate:String(baseDisc)}:r)
      :[...filledProducts,{group:groupLabel,rate:String(baseDisc)}];
    // 쿠폰 자동 입력: 기본 쿠폰 → cart 타입, 중복 쿠폰들 → product 타입 (사용자가 사후 조정)
    const cleanStacks=stackRates.filter(r=>r>0);
    const nextCoupons=[
      {...emptyCouponRow("cart"),name:"29CM 쿠폰",rate:String(primaryCoupon)},
      ...cleanStacks.map((r,i)=>({...emptyCouponRow("product"),name:`29CM Case ${i+1}`,rate:String(r)})),
    ];
    onChange({
      products:{period:plan.products.period,rows:nextProductRows},
      coupons:nextCoupons,
    });
  };
  const firstCouponRate=Number((coupons.find(c=>couponTypeOf(c)==="cart"&&c.rate)||coupons.find(c=>c.rate)||{}).rate)||10;
  // 매트릭스에 쓰이는 상품군 목록 (쿠폰별 적용 여부 토글용)
  const matrixGroups=[...new Set(productRows
    .filter(r=>(r.group||"").trim()||(+r.rate||0)>0)
    .map(r=>(r.group||"").trim()||"전체"))];
  const toggleCouponGroup=(i,g)=>{
    const c=coupons[i];
    const ex=Array.isArray(c.excludeGroups)?c.excludeGroups:[];
    const next=ex.includes(g)?ex.filter(x=>x!==g):[...ex,g];
    const n=[...coupons];n[i]={...c,excludeGroups:next};setCoupons(n);
  };
  // 쿠폰 타입 변경
  const setCouponType=(i,nextType)=>{
    if(!COUPON_TYPE_BY_KEY[nextType]) return;
    const n=[...coupons];n[i]={...n[i],type:nextType};setCoupons(n);
  };

  const cellInp={background:D.surface,border:`1px solid ${D.border}`,borderRadius:5,
    padding:"4px 7px",fontSize:11,color:D.text,width:"100%",boxSizing:"border-box",
    fontFamily:"'Noto Sans KR','Pretendard',sans-serif"};
  const lbl={fontSize:11,color:D.textMeta,marginBottom:4,fontWeight:600};
  const head={fontSize:11,color:D.textMeta,fontWeight:600,padding:"4px 6px",textAlign:"left"};

  return (
    <div style={{border:`1px solid ${D.border}`,borderRadius:6,padding:"10px 12px",background:D.surface}}>
      <div style={{fontWeight:700,fontSize:13,color:D.black,marginBottom:8}}>할인율</div>

      <div style={{display:"flex",gap:20,flexWrap:"wrap",alignItems:"flex-start",marginBottom:12}}>
      {/* 상품 할인 */}
      <div style={{flex:"0 1 640px",minWidth:340}}>
        <div style={{...lbl,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
          <span>프런트 할인 <span style={{color:D.textMeta,fontWeight:400}}>· 상품군별 할인율·평균 마크업 (전체 동일 기간)</span></span>
          {platform==="29CM"&&(
            <button onClick={()=>setCalcOpen(true)}
              style={{background:D.blue,color:"#fff",border:"none",borderRadius:5,
                padding:"4px 10px",fontSize:10,cursor:"pointer",fontWeight:700,whiteSpace:"nowrap"}}>
              계산기 이용
            </button>
          )}
        </div>
        {/* 공통 기간 */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:6,maxWidth:360}}>
          <div>
            <div style={{fontSize:10,color:D.textMeta,marginBottom:3}}>시작</div>
            <DateDrop id={`${idPrefix}_prodStart`} value={plan.products.period.start}
              onChange={v=>setProductPeriod("start",v)} calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}
              placeholder="날짜 선택"/>
          </div>
          <div>
            <div style={{fontSize:10,color:D.textMeta,marginBottom:3}}>종료</div>
            <DateDrop id={`${idPrefix}_prodEnd`} value={plan.products.period.end}
              onChange={v=>setProductPeriod("end",v)} calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}
              placeholder="날짜 선택"/>
          </div>
        </div>
        <table style={{width:"100%",maxWidth:520,borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>
            <th style={{...head,width:22}}/>
            <th style={{...head,width:"42%"}}>상품군</th>
            <th style={{...head,width:"20%"}}>할인율(%)</th>
            <th style={{...head,width:"26%",color:D.green}} title="시나리오(프런트 할인×쿠폰) 적용 후 실수령 ÷ 원가 · ×3 이하 적색">시나리오 적용 마크업</th>
            <th style={{...head,width:"10%"}}/>
          </tr></thead>
          <tbody>
            {productRows.map((row,i)=>{
              const muVal=row.markup===""||row.markup==null?null:parseFloat(row.markup);
              const muLow=muVal!=null&&!isNaN(muVal)&&muVal<=3;
              return (
              <tr key={i}
                onDragOver={e=>{
                  if(prodDragIdx===null||prodDragIdx===i) return;
                  e.preventDefault();
                  e.dataTransfer.dropEffect="move";
                  const arr=[...productRows];
                  const [m]=arr.splice(prodDragIdx,1);
                  arr.splice(i,0,m);
                  setProductRows(arr);
                  setProdDragIdx(i);
                }}
                style={{opacity:prodDragIdx===i?0.4:1,transition:"opacity 0.12s",
                  background:prodDragIdx!==null&&prodDragIdx!==i?`${D.blue}06`:"transparent"}}>
                <td style={{padding:"3px 0",textAlign:"center"}}>
                  <span draggable="true"
                    onDragStart={e=>{e.dataTransfer.effectAllowed="move";setProdDragIdx(i);}}
                    onDragEnd={()=>setProdDragIdx(null)}
                    title="드래그하여 상품군 순서 변경"
                    style={{cursor:"grab",color:D.textMeta,fontSize:11,userSelect:"none",
                      display:"inline-block",padding:"0 2px",lineHeight:1}}>
                    ⋮⋮
                  </span>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input value={row.group} onChange={e=>{const n=[...productRows];n[i]={...row,group:e.target.value};setProductRows(n);}}
                    style={cellInp} placeholder="예: 신상품, 전체"/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <input type="number" onWheel={e=>e.currentTarget.blur()} value={row.rate} onChange={e=>{const n=[...productRows];n[i]={...row,rate:e.target.value};setProductRows(n);}}
                    style={cellInp} placeholder="0" min="0" max="100"/>
                </td>
                <td style={{padding:"3px 4px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:3}}>
                    <span style={{fontSize:11,color:muLow?D.red:D.textMeta,fontWeight:muVal!=null?700:400}}>×</span>
                    <input type="number" step="0.01" min="0" onWheel={e=>e.currentTarget.blur()}
                      value={row.markup||""}
                      onChange={e=>{const n=[...productRows];n[i]={...row,markup:e.target.value};setProductRows(n);}}
                      style={{...cellInp,color:muLow?D.red:(muVal!=null?D.green:D.text),fontWeight:muVal!=null?700:400}}
                      placeholder="0.00"
                      title="실수령 ÷ 원가 (예: ×3.50)"/>
                  </div>
                </td>
                <td style={{padding:"3px 4px",textAlign:"right",whiteSpace:"nowrap"}}>
                  {Array.isArray(row.products)&&row.products.length>0?(
                    <span style={{display:"inline-flex",alignItems:"center",gap:3,marginRight:4}}>
                      <button onClick={()=>setBundleViewIdx(bundleViewIdx===i?null:i)}
                        title={`묶음 상품 ${row.products.length}개 보기`}
                        style={{background:bundleViewIdx===i?D.black:"transparent",
                          color:bundleViewIdx===i?"#fff":D.text,
                          border:`1px solid ${D.borderMid}`,borderRadius:4,
                          padding:"2px 6px",fontSize:10,cursor:"pointer",fontWeight:600}}>
                        묶음 {row.products.length}
                      </button>
                      <button onClick={()=>{
                        if(!window.confirm(`'${row.group||`행 ${i+1}`}' 의 묶음 상품 ${row.products.length}개를 삭제할까요?`)) return;
                        if(bundleViewIdx===i) setBundleViewIdx(null);
                        const n=[...productRows];n[i]={...row,products:[],markup:""};setProductRows(n);
                      }}
                        title="묶음 삭제 — 묶음 비우고 다시 추가 가능"
                        style={{background:"transparent",color:D.red,
                          border:`1px solid ${D.border}`,borderRadius:4,
                          padding:"2px 5px",fontSize:10,cursor:"pointer",fontWeight:700}}>
                        ✕
                      </button>
                    </span>
                  ):(
                    ((row.group||"").trim()||(+row.rate||0)>0)&&(platform==="29CM"||platform==="자사몰")&&(
                      <span style={{display:"inline-flex",alignItems:"center",gap:3,marginRight:4}}>
                        {platform==="29CM"&&(
                          <button onClick={()=>requestInlineCalc("29CM",i)}
                            title="29CM 계산기로 이 행의 묶음 채우기"
                            style={{background:"transparent",color:D.text,
                              border:`1px dashed ${D.borderMid}`,borderRadius:4,
                              padding:"2px 6px",fontSize:10,cursor:"pointer",fontWeight:600}}>
                            + 29CM 묶음
                          </button>
                        )}
                        {platform==="자사몰"&&(
                          <button onClick={()=>requestInlineCalc("자사몰",i)}
                            title="자사몰 계산기로 이 행의 묶음 채우기"
                            style={{background:"transparent",color:D.text,
                              border:`1px dashed ${D.borderMid}`,borderRadius:4,
                              padding:"2px 6px",fontSize:10,cursor:"pointer",fontWeight:600}}>
                            + 자사몰 묶음
                          </button>
                        )}
                      </span>
                    )
                  )}
                  <button onClick={()=>{const n=productRows.filter((_,j)=>j!==i);setProductRows(n.length?n:[emptyProductRow()]);}}
                    style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:14,lineHeight:1}}>✕</button>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
        {bundleViewIdx!=null&&Array.isArray(productRows[bundleViewIdx]?.products)&&productRows[bundleViewIdx].products.length>0&&(()=>{
          const r=productRows[bundleViewIdx];
          const cpnRow=Number(r.cpn||0)||0;
          const won=n=>"₩"+(Math.round(n||0)).toLocaleString();
          return (
          <div style={{marginTop:8,border:`1px solid ${D.borderMid}`,borderRadius:6,overflow:"hidden"}}>
            <div style={{padding:"6px 10px",background:D.surfaceAlt,fontSize:10,fontWeight:700,color:D.black,
              display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <span>{r.group||`행 ${bundleViewIdx+1}`} · 묶음 상품 {r.products.length}개{cpnRow>0?` · 쿠폰율 ${cpnRow}%`:""}</span>
              <button onClick={()=>setBundleViewIdx(null)}
                style={{background:"none",border:"none",cursor:"pointer",color:D.textMeta,fontSize:11}}>✕</button>
            </div>
            <div style={{maxHeight:280,overflow:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
                <thead><tr style={{background:D.surfaceAlt,color:D.textMeta}}>
                  {["상품명","정가","쿠폰율","기본 할인율","프런트 판매가","최종 노출가","최종 할인율","자사부담","수수료","채널보전","자사 정산","공급가","마진","마크업"].map((h,k)=>(
                    <th key={k} style={{padding:"4px 8px",textAlign:k===0?"left":"right",fontWeight:600,position:"sticky",top:0,background:D.surfaceAlt,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {r.products.map((p,j)=>{
                    const sv=p.supplyIncVat||Math.round((p.supply||0)*1.1);
                    return (
                      <tr key={j} style={{borderTop:`1px solid ${D.border}`}}>
                        <td title={p.name} style={{padding:"3px 8px",maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap"}}>{won(p.list)}</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap"}}>{cpnRow}%</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{p.baseDisc||0}%</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.blue,background:"#eef3ff",fontWeight:600}}>{won(p.basePrice)}</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.textSub}}>{won(p.finalPrice)}</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.blue,background:"#eef3ff",fontWeight:700}}>{p.finalDisc||0}%</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.selfBurden||0)>0?D.red:D.textMeta}}>
                          {(p.selfBurden||0)>0?`−${won(p.selfBurden)}`:"—"}
                        </td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.red}}>−{won(p.fee||0)} <span style={{fontSize:9,color:D.textMeta}}>({p.feeRate||0}%)</span></td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.channelBurden||0)>0?D.blue:D.textMeta}}>
                          {(p.channelBurden||0)>0?`+${won(p.channelBurden)}`:"—"}
                        </td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{won(p.net||0)}</td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.supply||0)>0?D.text:D.textMeta}}>
                          {(p.supply||0)>0?won(sv):"—"}
                        </td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600,
                          color:(p.supply||0)>0?((p.margin||0)>=0?D.text:D.red):D.textMeta}}>
                          {(p.supply||0)>0?won(p.margin||0):"—"}
                        </td>
                        <td style={{padding:"3px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:700,
                          color:(p.supply||0)>0?((p.markup||0)>3?D.green:D.red):D.textMeta}}>
                          {(p.supply||0)>0?`×${(p.markup||0).toFixed(2)}`:"—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
          );
        })()}
        <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",marginTop:6}}>
          <button onClick={()=>setProductRows([...productRows,emptyProductRow()])}
            style={{background:"transparent",border:`1px dashed ${D.border}`,borderRadius:5,
              padding:"4px 12px",fontSize:11,color:D.textMeta,cursor:"pointer"}}>+ 행 추가</button>
          {platform==="29CM"&&(
            <button onClick={()=>requestInlineCalc("29CM",null)}
              title="29CM 계산기로 새 상품군·쿠폰 묶음 추가"
              style={{background:"transparent",border:`1px dashed ${D.borderMid}`,borderRadius:5,
                padding:"4px 12px",fontSize:11,color:D.text,fontWeight:600,cursor:"pointer"}}>
              + 29CM 묶음 추가
            </button>
          )}
          {platform==="자사몰"&&(
            <button onClick={()=>requestInlineCalc("자사몰",null)}
              title="자사몰 계산기로 새 상품군·쿠폰 묶음 추가"
              style={{background:"transparent",border:`1px dashed ${D.borderMid}`,borderRadius:5,
                padding:"4px 12px",fontSize:11,color:D.text,fontWeight:600,cursor:"pointer"}}>
              + 자사몰 묶음 추가
            </button>
          )}
        </div>
      </div>

      {/* 쿠폰 */}
      <div style={{flex:"1 1 540px",minWidth:340}}>
        <div style={lbl}>쿠폰 <span style={{color:D.textMeta,fontWeight:400}}>· 할인율 + 기간 (프런트 할인 적용 후 추가 적용) · 타입별 누적 규칙 자동 적용 · 칩으로 적용 상품군 선택</span></div>
        {/* 쿠폰 타입별 누적 규칙 안내 */}
        <div style={{display:"flex",gap:8,marginBottom:8,flexWrap:"wrap",fontSize:10,lineHeight:1.4}}>
          {[
            {key:"product",desc:"상품 쿠폰끼리 중복 불가 · 장바구니와 누적 가능 · 분담과는 누적 불가"},
            {key:"cart",desc:"장바구니 쿠폰끼리 중복 불가 · 상품과 누적 가능 · 분담과는 누적 불가"},
            {key:"share",desc:"분담 쿠폰은 다른 어떤 쿠폰과도 중복 불가 (단독 적용)"},
          ].map(({key,desc})=>{
            const t=COUPON_TYPE_BY_KEY[key];
            return (
              <div key={key} style={{display:"flex",alignItems:"center",gap:5,
                background:t.bg,border:`1px solid ${t.border}`,borderRadius:6,padding:"3px 7px",
                flex:"1 1 200px",minWidth:200}}>
                <span style={{background:t.color,color:"#fff",fontWeight:700,fontSize:9,
                  padding:"2px 6px",borderRadius:3,whiteSpace:"nowrap",flexShrink:0}}>{t.short}</span>
                <span style={{color:t.color,fontSize:10}}>{desc}</span>
              </div>
            );
          })}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:8,maxWidth:760}}>
          {coupons.map((row,i)=>(
            <div key={i}
              onDragOver={e=>{
                if(dragIdx===null||dragIdx===i) return;
                e.preventDefault();
                e.dataTransfer.dropEffect="move";
                // 실시간 재배치 — 드래그 중인 쿠폰을 즉시 i 위치로 이동
                const arr=[...coupons];
                const [moved]=arr.splice(dragIdx,1);
                arr.splice(i,0,moved);
                setCoupons(arr);
                setDragIdx(i);
              }}
              onDrop={e=>{e.preventDefault();setDragIdx(null);}}
              style={(()=>{const tInfo=COUPON_TYPE_BY_KEY[couponTypeOf(row)];return{
                border:`1px solid ${dragIdx!==null&&dragIdx!==i?`${D.blue}80`:tInfo.border}`,borderRadius:8,
                padding:"10px 12px",background:tInfo.bg,
                display:"flex",flexDirection:"column",gap:8,
                opacity:dragIdx===i?0.45:1,
                transition:"opacity 0.12s, border-color 0.12s"};})()}>
              {/* 드래그 핸들 · 타입 세그먼트 · 쿠폰명 · 할인율 · 삭제 */}
              <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                <span draggable="true"
                  onDragStart={e=>{setDragIdx(i);e.dataTransfer.effectAllowed="move";}}
                  onDragEnd={()=>setDragIdx(null)}
                  title="드래그하여 쿠폰 순서 변경"
                  style={{cursor:"grab",color:D.textMeta,fontSize:13,padding:"0 4px",
                    userSelect:"none",flexShrink:0,lineHeight:1}}>
                  ⋮⋮
                </span>
                {/* 타입 세그먼트 — 상품 · 장바구니 · 분담 (같은 타입 끼리 중복 불가, 분담은 누구와도 중복 불가) */}
                <div style={{display:"flex",flexShrink:0,borderRadius:6,overflow:"hidden",border:`1px solid ${D.border}`}}>
                  {COUPON_TYPES.map(t=>{
                    const active=couponTypeOf(row)===t.key;
                    return <button key={t.key} type="button" onClick={()=>setCouponType(i,t.key)}
                      title={t.key==="share"?`${t.label} — 누구와도 중복 적용 안 됨`:`${t.label} — 같은 타입끼리만 중복 불가`}
                      style={{padding:"4px 8px",fontSize:10,fontWeight:700,cursor:"pointer",
                        border:"none",
                        background:active?t.color:"transparent",
                        color:active?"#fff":D.textMeta,
                        fontFamily:"inherit",whiteSpace:"nowrap"}}>
                      {t.short}
                    </button>;
                  })}
                </div>
                <input value={row.name} onChange={e=>{const n=[...coupons];n[i]={...row,name:e.target.value};setCoupons(n);}}
                  style={{...cellInp,width:280,maxWidth:"100%",flex:"0 0 auto"}} placeholder="쿠폰명 (예: 신규가입 쿠폰)"/>
                <div style={{display:"flex",alignItems:"center",gap:4,flexShrink:0}}>
                  <input type="number" onWheel={e=>e.currentTarget.blur()} value={row.rate} onChange={e=>{const n=[...coupons];n[i]={...row,rate:e.target.value};setCoupons(n);}}
                    style={{...cellInp,width:couponUnitOf(row)==="won"?86:62,textAlign:"right"}}
                    placeholder="0" min="0" max={couponUnitOf(row)==="won"?undefined:"100"}/>
                  {/* % / 원 토글 — 작은 세그먼트 */}
                  <div style={{display:"flex",border:`1px solid ${D.border}`,borderRadius:4,overflow:"hidden"}}>
                    {[{k:"pct",l:"%"},{k:"won",l:"원"}].map(u=>{
                      const active=couponUnitOf(row)===u.k;
                      return <button key={u.k} type="button"
                        onClick={()=>{const n=[...coupons];n[i]={...row,unit:u.k};setCoupons(n);}}
                        title={u.k==="won"?"정액 차감(원 단위) — 매트릭스 누적에서는 제외":"퍼센트 할인"}
                        style={{background:active?D.black:"transparent",color:active?"#fff":D.textMeta,
                          border:"none",padding:"3px 6px",fontSize:10,fontWeight:700,cursor:"pointer",
                          fontFamily:"inherit",lineHeight:1}}>{u.l}</button>;
                    })}
                  </div>
                </div>
                <button onClick={()=>{const n=coupons.filter((_,j)=>j!==i);setCoupons(n.length?n:[emptyCouponRow()]);}}
                  title="쿠폰 삭제"
                  style={{flexShrink:0,background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:15,lineHeight:1,padding:"0 2px"}}>✕</button>
              </div>
              {/* 기간 */}
              <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                <span style={{fontSize:10,color:D.textMeta,fontWeight:600,width:60,flexShrink:0}}>기간</span>
                <div style={{minWidth:130}}>
                  <DateDrop id={`${idPrefix}_coupon${i}Start`} value={row.start}
                    onChange={v=>{const n=[...coupons];n[i]={...row,start:v};setCoupons(n);}}
                    calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="시작"/>
                </div>
                <span style={{color:D.textMeta,fontSize:11}}>~</span>
                <div style={{minWidth:130}}>
                  <DateDrop id={`${idPrefix}_coupon${i}End`} value={row.end}
                    onChange={v=>{const n=[...coupons];n[i]={...row,end:v};setCoupons(n);}}
                    calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="종료"/>
                </div>
              </div>
              {/* 적용 상품군 */}
              {matrixGroups.length>0&&(
                <div style={{display:"flex",alignItems:"flex-start",gap:6,flexWrap:"wrap"}}>
                  <span style={{fontSize:10,color:D.textMeta,fontWeight:600,width:60,flexShrink:0,paddingTop:3}}>적용 상품군</span>
                  <div style={{display:"flex",flexWrap:"wrap",gap:4,flex:1}}>
                    {matrixGroups.map(g=>{
                      const off=(row.excludeGroups||[]).includes(g);
                      return <button key={g} type="button" onClick={()=>toggleCouponGroup(i,g)}
                        title={off?`${g} 적용 안 함 → 클릭 시 적용`:`${g} 적용 중 → 클릭 시 제외`}
                        style={{fontSize:10,padding:"2px 9px",borderRadius:12,cursor:"pointer",lineHeight:1.5,
                          border:`1px solid ${off?D.border:D.blue}`,background:off?D.surfaceAlt:`${D.blue}14`,
                          color:off?D.textMeta:D.blue,textDecoration:off?"line-through":"none",fontWeight:600}}>
                        {off?"":"✓ "}{g}
                      </button>;
                    })}
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>
        <div style={{display:"flex",gap:6,marginTop:6}}>
          <button onClick={()=>setCoupons([...coupons,emptyCouponRow("product")])}
            style={{background:"transparent",border:`1px dashed ${D.border}`,borderRadius:5,
              padding:"4px 12px",fontSize:11,color:D.textMeta,cursor:"pointer"}}>+ 쿠폰 추가</button>
        </div>
      </div>
      </div>

      {/* 실시간 최종 할인율 매트릭스 (상품군 × 시나리오) */}
      {computeDiscountMatrix(plan).hasGroup&&(
        <div style={{marginTop:10,paddingTop:10,borderTop:`1px solid ${D.border}`}}>
          <div style={{...lbl,marginBottom:2}}>예상 최종 할인율 <span style={{color:D.textMeta,fontWeight:400}}>· 곱연산(프런트할인×쿠폰) · 파랑=예상 최종 · 값 클릭 시 원 강조</span></div>
          <DiscountMatrix plan={plan}/>
        </div>
      )}

      {calcOpen&&(
        <Promo29CMCalcModal
          initialCoupon={firstCouponRate}
          onApply={applyCalc}
          onClose={()=>setCalcOpen(false)}/>
      )}
      {inlineCalc?.platform==="29CM"&&(
        <SaleCalcModal
          onClose={()=>setInlineCalc(null)}
          onAttachInlineCalc={(payload)=>attachInlineCalc(payload,inlineCalc.targetRowIdx)}
          attachMode={inlineCalc.targetRowIdx==null?"new":"fill"}
          initialCoupon={inlineCalc.initialCoupon}
          initialPrimaryType={inlineCalc.initialPrimaryType}
          initialCoupons={inlineCalc.initialCoupons}/>
      )}
      {inlineCalc?.platform==="자사몰"&&(
        <OwnMallSaleCalcModal
          onClose={()=>setInlineCalc(null)}
          onAttachInlineCalc={(payload)=>attachInlineCalc(payload,inlineCalc.targetRowIdx)}
          attachMode={inlineCalc.targetRowIdx==null?"new":"fill"}
          initialCoupon={inlineCalc.initialCoupon}
          initialCouponName={inlineCalc.initialCouponName}/>
      )}
      {couponPrompt&&(
        <div onClick={()=>setCouponPrompt(null)}
          style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:2200,
            display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div onClick={e=>e.stopPropagation()}
            style={{background:D.surface,borderRadius:10,border:`1px solid ${D.black}`,
              padding:"18px 20px",minWidth:340,maxWidth:480,fontFamily:"'Noto Sans KR','Pretendard',sans-serif"}}>
            {couponPrompt.step==="choose"&&(
              <>
                <div style={{fontSize:13,fontWeight:700,color:D.black,marginBottom:4}}>
                  {couponPrompt.platform} 묶음 추가
                </div>
                <div style={{fontSize:11,color:D.textMeta,marginBottom:14,lineHeight:1.55}}>
                  이 묶음에 사용할 쿠폰을 선택하세요.<br/>
                  · <b>기존 쿠폰</b>: 매트릭스 컬럼 확장 없이 기존 쿠폰 시나리오에 묶음을 추가<br/>
                  · <b>새 쿠폰</b>: 별도 Case 로 매트릭스 컬럼을 확장 (기존 상품군과는 '연관 없음')
                </div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <button onClick={()=>setCouponPrompt({...couponPrompt,step:"pick"})}
                    style={{flex:"1 1 140px",background:D.surface,color:D.black,
                      border:`1px solid ${D.black}`,borderRadius:6,padding:"9px 12px",
                      fontSize:12,fontWeight:700,cursor:"pointer"}}>
                    기존 쿠폰 사용
                  </button>
                  <button onClick={()=>{
                    setInlineCalc({platform:couponPrompt.platform,targetRowIdx:couponPrompt.targetRowIdx});
                    setCouponPrompt(null);
                  }}
                    style={{flex:"1 1 140px",background:D.black,color:"#fff",
                      border:`1px solid ${D.black}`,borderRadius:6,padding:"9px 12px",
                      fontSize:12,fontWeight:700,cursor:"pointer"}}>
                    새 쿠폰 (별도 Case)
                  </button>
                </div>
                <div style={{marginTop:10,textAlign:"right"}}>
                  <button onClick={()=>setCouponPrompt(null)}
                    style={{background:"transparent",border:"none",color:D.textMeta,
                      fontSize:11,cursor:"pointer"}}>취소</button>
                </div>
              </>
            )}
            {couponPrompt.step==="pick"&&(()=>{
              const validIdxs=coupons.map((c,i)=>(+c.rate||0)>0?i:-1).filter(i=>i>=0);
              const selectedIdxs=couponPrompt.selectedIdxs||new Set();
              const selCount=selectedIdxs.size;
              const toggleIdx=(idx)=>{
                const next=new Set(selectedIdxs);
                next.has(idx)?next.delete(idx):next.add(idx);
                setCouponPrompt({...couponPrompt,selectedIdxs:next});
              };
              const confirmPick=()=>{
                if(selCount===0) return;
                const ordered=validIdxs.filter(i=>selectedIdxs.has(i));
                const picked=ordered.map(i=>{
                  const c=coupons[i];
                  return {
                    rate:+c.rate||0,
                    type:couponTypeOf(c),
                    name:(c.name||"").trim()||`쿠폰 ${i+1}`,
                  };
                });
                const first=picked[0];
                setInlineCalc({
                  platform:couponPrompt.platform,
                  targetRowIdx:couponPrompt.targetRowIdx,
                  initialCoupons:picked,
                  // 호환 — 기존 prop 도 1번째 쿠폰으로 채워둠 (자사몰 등)
                  initialCoupon:first.rate,
                  initialCouponName:first.name,
                  initialPrimaryType:first.type,
                });
                setCouponPrompt(null);
              };
              return (
                <>
                  <div style={{fontSize:13,fontWeight:700,color:D.black,marginBottom:4}}>
                    기존 쿠폰 선택 <span style={{color:D.textMeta,fontWeight:400,fontSize:11}}>(여러 개 선택 가능)</span>
                  </div>
                  <div style={{fontSize:11,color:D.textMeta,marginBottom:14,lineHeight:1.55}}>
                    선택한 쿠폰의 시나리오가 계산기에 그대로 복제됩니다.<br/>
                    {couponPrompt.platform==="29CM"
                      ?"· 첫번째 선택 = 기본 쿠폰, 나머지 = 누적(추가) 쿠폰"
                      :"· 자사몰 계산기는 단일 쿠폰만 사용 — 첫번째 선택만 적용"}
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:6,maxHeight:280,overflow:"auto"}}>
                    {validIdxs.map((i,order)=>{
                      const c=coupons[i];
                      const r=+c.rate||0;
                      const nm=(c.name||"").trim()||`쿠폰 ${i+1}`;
                      const tInfo=COUPON_TYPE_BY_KEY[couponTypeOf(c)];
                      const checked=selectedIdxs.has(i);
                      const selOrder=checked?[...validIdxs].filter(x=>selectedIdxs.has(x)).indexOf(i):-1;
                      return (
                        <label key={i}
                          style={{display:"flex",alignItems:"center",gap:8,textAlign:"left",
                            background:checked?"#eef3ff":D.surface,color:D.text,
                            border:`1px solid ${checked?D.blue:D.border}`,borderRadius:6,padding:"8px 10px",
                            fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>
                          <input type="checkbox" checked={checked} onChange={()=>toggleIdx(i)}
                            style={{cursor:"pointer"}}/>
                          {checked&&(
                            <span style={{display:"inline-flex",alignItems:"center",justifyContent:"center",
                              width:18,height:18,borderRadius:"50%",background:D.blue,color:"#fff",
                              fontSize:10,fontWeight:700}}>{selOrder+1}</span>
                          )}
                          <span style={{color:D.black,fontWeight:700}}>{nm}</span>
                          <span style={{color:D.blue,fontWeight:700}}>{r}%</span>
                          <span style={{color:D.textMeta,fontWeight:400}}>· {tInfo.short}</span>
                        </label>
                      );
                    })}
                  </div>
                  <div style={{marginTop:14,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,flexWrap:"wrap"}}>
                    <button onClick={()=>setCouponPrompt({...couponPrompt,step:"choose"})}
                      style={{background:"transparent",border:"none",color:D.textMeta,
                        fontSize:11,cursor:"pointer"}}>← 뒤로</button>
                    <div style={{display:"inline-flex",gap:6,alignItems:"center"}}>
                      <button onClick={()=>setCouponPrompt(null)}
                        style={{background:"transparent",border:"none",color:D.textMeta,
                          fontSize:11,cursor:"pointer"}}>취소</button>
                      <button onClick={confirmPick} disabled={selCount===0}
                        style={{background:selCount>0?D.black:D.surfaceAlt,
                          color:selCount>0?"#fff":D.textMeta,
                          border:`1px solid ${selCount>0?D.black:D.border}`,borderRadius:6,
                          padding:"7px 14px",fontSize:11,fontWeight:700,
                          cursor:selCount>0?"pointer":"default"}}>
                        선택 완료 ({selCount})
                      </button>
                    </div>
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}
    </div>
  );
}

// 저장 직전 빈 행 정리 — addPromo/patchPromo 직전에 호출
function cleanDiscountPlan(plan){
  const p=normalizePlan(plan);
  return{
    products:{
      period:p.products.period,
      rows:p.products.rows.filter(r=>r.group||r.rate),
    },
    coupons:p.coupons.filter(r=>r.rate||r.start||r.end||r.name),
  };
}

// 표 셀에 표시되는 컴팩트 보기 — 상품 할인 + 쿠폰 기간을 작은 가로 막대 두 줄로 시각화
function DiscountPlanView({ plan, marks={}, onToggleGroup, onToggleCircle, compact=true, showBadges=true }) {
  const p=normalizePlan(plan);
  const cleanedProducts=p.products.rows.filter(r=>r.group||r.rate);
  const cleanedCoupons=p.coupons.filter(r=>r.rate||r.start||r.end);
  const hasAny=cleanedProducts.length||cleanedCoupons.length||p.products.period.start;
  if(!hasAny) return <span style={{color:D.textMeta,fontSize:11}}>—</span>;

  // 뱃지 크기: compact 11/2x7, non-compact 도 -30% 스케일로 같게 작게 유지
  const badgeFs=compact?11:9;
  const badgePad=compact?"2px 7px":"2px 7px";
  return (
    <div style={{fontSize:compact?11:12,lineHeight:1.75,minWidth:compact?140:0,fontFamily:"'Noto Sans KR','Pretendard',sans-serif"}}>
      {/* 상품군 할인율 — 뱃지 형태. 클릭 시 흑백 전환 마킹(저장·공유). 비-compact 모드에서는 뱃지와 표 사이 간격을 넓힘.
          showBadges=false 면 뱃지 행 자체 미노출 (예: 임팩트 모달) */}
      {showBadges&&cleanedProducts.length>0&&(
        <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:compact?5:24,justifyContent:compact?"flex-start":"center"}}>
          {cleanedProducts.map((r,i)=>{
            const g=r.group||"전체";
            const on=(marks.groups||[]).includes(g);
            return (
            <span key={"p"+i} onClick={onToggleGroup?()=>onToggleGroup(g):undefined}
              style={{display:"inline-flex",alignItems:"center",gap:6,
              padding:badgePad,background:on?D.black:"#fff",border:`1px solid ${D.black}`,
              color:on?"#fff":D.black,borderRadius:10,fontSize:badgeFs,fontWeight:600,whiteSpace:"nowrap",
              cursor:onToggleGroup?"pointer":"default"}}>
              <span>{g}</span>
              <span style={{width:1,height:10,background:on?"#fff":D.black,display:"inline-block"}}/>
              <b style={{fontWeight:700}}>{r.rate||0}%</b>
            </span>
            );
          })}
        </div>
      )}
      {/* 상품군 × 시나리오 최종 할인율 매트릭스 (쿠폰은 매트릭스 열로 표시) */}
      {cleanedCoupons.length>0&&<DiscountMatrix plan={plan} compact={compact} circledKeys={marks.circles} onToggleCircle={onToggleCircle}/>}
    </div>
  );
}

// 핀셋 상품 선택기 — 배송(orders) 데이터 기반 상품 검색(상품 단위) + 상품별 메모
//   value: [{name, memo}]
function PinnedProductPicker({ value=[], onChange, orders=[] }) {
  const [q,setQ]=useState("");
  const [checked,setChecked]=useState(()=>new Set());
  const allProducts=useMemo(()=>{
    const s=new Set();
    orders.forEach(r=>{const n=(r.product_name||"").trim(); if(n) s.add(n);});
    return [...s].sort((a,b)=>a.localeCompare(b,"ko"));
  },[orders]);
  const selectedNames=useMemo(()=>new Set(value.map(v=>v.name)),[value]);
  const matches=useMemo(()=>{
    const kw=q.trim().toLowerCase();
    if(!kw) return [];
    return allProducts.filter(n=>n.toLowerCase().includes(kw)&&!selectedNames.has(n)).slice(0,30);
  },[q,allProducts,selectedNames]);
  const pillInp={background:D.surface,border:`1px solid ${D.border}`,borderRadius:6,padding:"6px 10px",fontSize:11,color:D.text,width:"100%",boxSizing:"border-box",fontFamily:"'Noto Sans KR','Pretendard',sans-serif"};
  const addNames=names=>{
    const add=names.filter(n=>!selectedNames.has(n));
    if(!add.length) return;
    onChange([...value,...add.map(n=>({name:n,memo:""}))]);
    setChecked(new Set()); setQ("");
  };
  const toggle=n=>setChecked(prev=>{const s=new Set(prev);s.has(n)?s.delete(n):s.add(n);return s;});
  const toggleAll=()=>setChecked(prev=>{
    const s=new Set(prev);
    if(matches.every(n=>s.has(n))) matches.forEach(n=>s.delete(n));
    else matches.forEach(n=>s.add(n));
    return s;
  });
  const checkedCount=matches.filter(n=>checked.has(n)).length;
  const allChecked=matches.length>0&&matches.every(n=>checked.has(n));
  const remove=i=>onChange(value.filter((_,j)=>j!==i));
  const setMemo=(i,memo)=>onChange(value.map((v,j)=>j===i?{...v,memo}:v));
  return (
    <div>
      <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>핀셋 상품 <span style={{opacity:.6}}>(배송 데이터 기반 · 상품 단위 · 체크 후 한번에 추가 · 임팩트 분석에서 전/후 판매량 비교)</span></div>
      <div style={{position:"relative"}}>
        <input value={q} onChange={e=>setQ(e.target.value)} style={{...pillInp,width:"30%",minWidth:200}} placeholder="상품명 검색 (배송 데이터)"/>
        {matches.length>0&&(
          <div style={{position:"absolute",top:"100%",left:0,right:0,zIndex:50,background:D.surface,
            border:`1px solid ${D.border}`,borderRadius:6,marginTop:2,maxHeight:260,overflowY:"auto",
            boxShadow:"0 4px 16px rgba(0,0,0,0.12)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,
              padding:"6px 10px",borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt,
              position:"sticky",top:0}}>
              <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:D.textSub,cursor:"pointer"}}>
                <input type="checkbox" checked={allChecked} onChange={toggleAll} style={{cursor:"pointer"}}/>
                모두 선택 ({matches.length})
              </label>
              <button onClick={()=>addNames(matches.filter(n=>checked.has(n)))} disabled={checkedCount===0}
                style={{background:checkedCount?D.black:D.surfaceAlt,color:checkedCount?"#fff":D.textMeta,
                  border:`1px solid ${checkedCount?D.black:D.border}`,borderRadius:5,padding:"4px 10px",
                  fontSize:12,fontWeight:600,cursor:checkedCount?"pointer":"default",whiteSpace:"nowrap"}}>
                선택 {checkedCount}개 추가
              </button>
            </div>
            {matches.map(n=>(
              <label key={n} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",fontSize:13,
                cursor:"pointer",color:D.text,borderBottom:`1px solid ${D.border}`}}>
                <input type="checkbox" checked={checked.has(n)} onChange={()=>toggle(n)} style={{cursor:"pointer",flexShrink:0}}/>
                <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={n}>{n}</span>
              </label>
            ))}
          </div>
        )}
      </div>
      {value.length>0&&(
        <div style={{marginTop:8,maxWidth:560}}>
          {value.length>10&&(
            <div style={{fontSize:10,color:D.textMeta,marginBottom:4}}>
              총 {value.length}개 · 위 영역 스크롤
            </div>
          )}
          {/* 10개 행 고정 높이 + 인너 스크롤. 행 높이(약 28px) × 10 + 약간의 padding */}
          <div style={{display:"flex",flexDirection:"column",gap:3,
            maxHeight:value.length>10?300:"none",overflowY:value.length>10?"auto":"visible",
            border:value.length>10?`1px solid ${D.border}`:"none",
            borderRadius:value.length>10?6:0,padding:value.length>10?4:0}}>
            {value.map((v,i)=>(
              <div key={v.name+i} style={{display:"flex",alignItems:"center",gap:4,
                background:D.surfaceAlt,borderRadius:4,padding:"3px 6px",flexShrink:0}}>
                <span style={{flex:"0 0 55%",fontSize:11,color:D.text,
                  overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={v.name}>📌 {v.name}</span>
                <input value={v.memo||""} onChange={e=>setMemo(i,e.target.value)}
                  style={{...pillInp,flex:1,padding:"3px 6px",fontSize:11,background:D.surface}} placeholder="메모"/>
                <button onClick={()=>remove(i)}
                  style={{background:"none",border:"none",color:D.textMeta,
                    cursor:"pointer",padding:"0 4px",fontSize:13,lineHeight:1,flexShrink:0}}>✕</button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// 프로모션 상세 내용 — 하이라이트 가능한 contentEditable 에디터
function HighlightEditor({ value, onChange, placeholder, minHeight, inputStyle }){
  const ref=useRef(null);
  const lastSyncRef=useRef("");
  useEffect(()=>{
    if(ref.current&&(value||"")!==lastSyncRef.current){
      ref.current.innerHTML=value||"";
      lastSyncRef.current=value||"";
    }
  },[value]);
  const updateOnChange=()=>{
    if(!ref.current) return;
    const html=ref.current.innerHTML;
    lastSyncRef.current=html;
    onChange(html);
  };
  const applyHighlight=(color)=>{
    const sel=window.getSelection();
    if(!sel||sel.rangeCount===0||sel.isCollapsed) return;
    try{
      document.execCommand("styleWithCSS",false,true);
      document.execCommand("hiliteColor",false,color);
    }catch{
      try{ document.execCommand("backColor",false,color); }catch{}
    }
    updateOnChange();
  };
  const clearHighlight=()=>{
    const sel=window.getSelection();
    if(!sel||sel.rangeCount===0||sel.isCollapsed) return;
    try{
      document.execCommand("styleWithCSS",false,true);
      document.execCommand("hiliteColor",false,"transparent");
    }catch{}
    updateOnChange();
  };
  const handlePaste=e=>{
    e.preventDefault();
    const text=(e.clipboardData||window.clipboardData).getData("text/plain");
    document.execCommand("insertText",false,text);
    updateOnChange();
  };
  return (
    <div>
      <div style={{display:"flex",gap:4,marginBottom:4,flexWrap:"wrap"}}>
        <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>applyHighlight("#cdd9e6")}
          title="선택한 텍스트 뮤트 블루 하이라이트"
          style={{background:"#cdd9e6",border:`1px solid ${D.borderMid}`,borderRadius:4,
            padding:"3px 10px",fontSize:11,cursor:"pointer",fontWeight:700,color:"#2d3e54"}}>
          블루
        </button>
        <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>applyHighlight("#e6cdcd")}
          title="선택한 텍스트 뮤트 레드 하이라이트"
          style={{background:"#e6cdcd",border:`1px solid ${D.borderMid}`,borderRadius:4,
            padding:"3px 10px",fontSize:11,cursor:"pointer",fontWeight:700,color:"#5b3030"}}>
          레드
        </button>
        <button type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>applyHighlight("#cfcec8")}
          title="선택한 텍스트 연한 차콜 하이라이트"
          style={{background:"#cfcec8",border:`1px solid ${D.borderMid}`,borderRadius:4,
            padding:"3px 10px",fontSize:11,cursor:"pointer",fontWeight:700,color:"#3d3d36"}}>
          차콜
        </button>
        <button type="button" onMouseDown={e=>e.preventDefault()} onClick={clearHighlight}
          title="선택 영역 하이라이트 해제"
          style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
            padding:"3px 10px",fontSize:11,cursor:"pointer",color:D.textSub}}>
          해제
        </button>
      </div>
      <div ref={ref} contentEditable suppressContentEditableWarning
        onInput={updateOnChange} onBlur={updateOnChange} onPaste={handlePaste}
        data-placeholder={placeholder||""}
        style={{minHeight:minHeight||144,padding:"8px 10px",fontSize:13,
          border:`1px solid ${D.border}`,borderRadius:5,background:D.surface,
          lineHeight:1.5,outline:"none",color:D.text,whiteSpace:"pre-wrap",
          fontFamily:"'Noto Sans KR','Pretendard',sans-serif",overflowWrap:"break-word",
          ...inputStyle}}/>
    </div>
  );
}

function PromoFlow({ revenues, storeSales=[], orders=[] }) {
  const [promos,setPromos]=useState(getPromosCache);
  const [showForm,setShowForm]=useState(false);
  const [addTimePickerFor,setAddTimePickerFor]=useState(null); // "start_date" | "end_date" | null
  const [editTimePickerFor,setEditTimePickerFor]=useState(null); // 동일 — 수정 모드
  const [form,setForm]=useState({name:"",platform:"자사몰",start_date:"",end_date:"",memo:"",content:"",files:[],discount_plan:{products:[],coupons:[]},pinned_products:[],submit_date:""});
  const today=localDate(0); // 로컬(KST) 날짜 — UTC 변환 시 새벽에 하루 밀려 오늘 시작 프로모션의 임팩트 버튼이 사라지던 문제 방지
  const [impactModal,setImpactModal]=useState(null);
  const [profitModal,setProfitModal]=useState(null); // 이익률 계산(베타) 모달 — 자사몰 한정
  const [filePreview,setFilePreview]=useState(null);
  const [viewStart,setViewStart]=useState(()=>localDate(-30));
  const [viewEnd,setViewEnd]=useState(()=>localDate(30));
  const [viewPeriod,setViewPeriod]=useState("2m");
  const [calOpenFor,setCalOpenFor]=useState(null);
  const handleViewPeriod=v=>{
    setViewPeriod(v);
    if(v==="custom") return;
    const days=v==="1m"?15:v==="2m"?30:45;
    setViewStart(localDate(-days));
    setViewEnd(localDate(days));
  };

  const [hoveredPromo,setHoveredPromo]=useState(null);
  const [fileAddTarget,setFileAddTarget]=useState(null);
  const fileInputRef=useRef(null);
  const [isDragging,setIsDragging]=useState(false);
  const dragRef=useRef(null);
  const formFileRef=useRef(null);
  const promoCardRefs=useRef({}); // 등록 프로모션 카드별 DOM ref (이미지 다운로드용)
  const [formFileDragOver,setFormFileDragOver]=useState(false);
  const [tableFileDragOver,setTableFileDragOver]=useState(null);
  const [pinnedModal,setPinnedModal]=useState(null); // 핀셋 상품 전체 보기 모달 — { promo } | null
  // Hidden promo log (localStorage only — no schema change needed)
  const getHiddenLog=()=>{try{return JSON.parse(localStorage.getItem("hidden_promo_log")||"[]");}catch{return[];}};
  const saveHiddenLogLocal=d=>localStorage.setItem("hidden_promo_log",JSON.stringify(d));
  const [hiddenLog,setHiddenLog]=useState(getHiddenLog);
  const hiddenIds=useMemo(()=>new Set(hiddenLog.map(h=>h.id)),[hiddenLog]);
  const [selHiddenIds,setSelHiddenIds]=useState(new Set());

  // 가려진 종료 프로모션 — Supabase 동기화 (기기 간 공유). 미연결 시 localStorage 폴백
  useEffect(()=>{
    (async()=>{
      const local=getHiddenLog();
      try{
        const db=await getSupabase();
        const{data,error}=await db.from("hidden_promo_log").select("*");
        if(error||!data) return; // 오프라인/미연결 → 로컬 초기값 유지
        const rows=data.map(r=>({...(r.data||{}),id:r.id,hidden_at:r.hidden_at}));
        if(rows.length>0){
          setHiddenLog(rows);saveHiddenLogLocal(rows);
        } else if(local.length>0){
          // Supabase 비어있고 로컬에 있으면 1회 마이그레이션
          const payload=local.map(h=>({id:h.id,hidden_at:h.hidden_at||new Date().toISOString(),data:h}));
          await db.from("hidden_promo_log").insert(payload);
        }
      }catch{/* 로컬 초기값 유지 */}
    })();
  },[]);
  // Promo search
  const [searchOpen,setSearchOpen]=useState(false);
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
  const startEditPromo=p=>{setEditingPromoId(p.id);setEditPromoForm({name:p.name,platform:p.platform,start_date:p.start_date,end_date:p.end_date,content:p.content||p.memo||"",discount_plan:p.discount_plan||{products:[],coupons:[]},pinned_products:p.pinned_products||[],submit_date:p.submit_date||""});};
  const savePromoEdit=()=>{
    const savedId=editingPromoId;
    patchPromo(editingPromoId,{...editPromoForm,memo:editPromoForm.content,discount_plan:cleanDiscountPlan(editPromoForm.discount_plan)});
    setEditingPromoId(null);
    // 수정 완료 후 해당 카드로 화면 스크롤
    setTimeout(()=>{
      const el=promoCardRefs.current[savedId];
      if(el) el.scrollIntoView({behavior:"smooth",block:"center"});
    },80);
  };
  // 한국시(KST) 현재 시각 YYYY-MM-DDTHH:MM — end_date(KST 기준)와 비교해 종료 판정 (UTC면 9시간 늦게 종료 처리됨)
  const nowStr=new Date(Date.now()+32400000).toISOString().slice(0,16);
  const isEnded=p=>p.end_date&&String(p.end_date)<nowStr;
  const readFileData=(file,cb)=>{const r=new FileReader();r.onload=e=>cb({name:file.name,type:file.type,data:e.target.result});r.readAsDataURL(file);};

  const patchPromo=useCallback(async(id,updates)=>{
    setPromos(prev=>{const next=prev.map(p=>p.id===id?{...p,...updates}:p);setPromosCache(next);return next;});
    const db=await getSupabase();
    await db.from("promotions").update(updates).eq("id",id);
  },[]);
  // 핀셋 상품 뱃지 하이라이트(흑백) 토글 — pinned_products(jsonb)에 저장되어 기기 간 공유됨
  const togglePinHighlight=(p,idx)=>patchPromo(p.id,{pinned_products:(p.pinned_products||[]).map((pp,i)=>i===idx?{...pp,highlight:!pp.highlight}:pp)});
  // 할인율 매트릭스 원형 강조 / 상품군 뱃지 마킹 — discount_marks(jsonb)에 저장되어 기기 간 공유됨
  const toggleMark=(p,kind,key)=>{const m=p.discount_marks||{};const set=new Set(m[kind]||[]);set.has(key)?set.delete(key):set.add(key);patchPromo(p.id,{discount_marks:{...m,[kind]:[...set]}});};

  // Load from Supabase — localStorage는 Supabase에 실제 데이터 있을 때만 덮어씀
  useEffect(()=>{
    (async()=>{
      const local=getPromosCache(); // 먼저 읽어둠 (덮어쓰기 방지)
      const db=await getSupabase();
      const{data,error}=await db.from("promotions").select("*").order("start_date",{ascending:true});
      if(!error&&data){
        const rows=data.map(p=>({...p,files:p.files||(p.file?[p.file]:[]),discount_plan:p.discount_plan||{products:[],coupons:[]},pinned_products:p.pinned_products||[],discount_marks:p.discount_marks||{},submit_date:p.submit_date||""}));
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
    const newP={...form,discount_plan:cleanDiscountPlan(form.discount_plan),id:Date.now()};
    setPromos(prev=>{const next=[...prev,newP];setPromosCache(next);return next;});
    const db=await getSupabase();
    await db.from("promotions").insert(newP);
    setForm({name:"",platform:"자사몰",start_date:"",end_date:"",memo:"",content:"",files:[],discount_plan:{products:[],coupons:[]},pinned_products:[],submit_date:""});
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
  const hidePromo=async p=>{
    const entry={...p,hidden_at:new Date().toISOString()};
    const next=[...hiddenLog.filter(h=>h.id!==p.id),entry];
    setHiddenLog(next);saveHiddenLogLocal(next);
    try{
      const db=await getSupabase();
      await db.from("hidden_promo_log").upsert({id:p.id,hidden_at:entry.hidden_at,data:entry},{onConflict:"id"});
    }catch{/* 로컬 저장은 완료됨 */}
  };
  const delFromHiddenLog=async ids=>{
    const next=hiddenLog.filter(h=>!ids.has(h.id));
    setHiddenLog(next);setSelHiddenIds(new Set());
    saveHiddenLogLocal(next);
    try{
      const db=await getSupabase();
      await db.from("hidden_promo_log").delete().in("id",[...ids]);
    }catch{/* 로컬 저장은 완료됨 */}
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
      const db=await getSupabase();
      const{data,error}=await db.from("submit_promotions").select("*").order("id",{ascending:true});
      if(!error&&Array.isArray(data)){
        // DB가 단일 진실원천 — 빈 배열이면 그대로 반영(삭제 항목 부활 방지)
        setSubmitPromos(data);saveSubmitPromosLocal(data);
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
    const{error}=await db.from("submit_promotions").delete().eq("id",id);
    if(error){
      // 삭제 실패 시 롤백 + 알림 (재등장 방지)
      alert("제출 완료 처리 실패: "+error.message);
      setSubmitPromos(submitPromos);saveSubmitPromosLocal(submitPromos);
    }
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
    // 실제 매출 데이터 채우기 — 온라인 채널은 순매출(매출 − 환불)
    revenues.filter(r=>r.date>=viewStart&&r.date<=viewEnd).forEach(r=>{
      if(!byDate[r.date]) return;
      byDate[r.date][r.channel]=(byDate[r.date][r.channel]||0)+((r.amount||0)-(r.refund_amount||0));
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

  const inp={background:D.surface,border:`1px solid ${D.border}`,borderRadius:6,
    padding:"6px 10px",fontSize:11,color:D.text,width:"100%",boxSizing:"border-box",
    fontFamily:"'Noto Sans KR','Pretendard',sans-serif"};

  // ── 프로모션 전략 메모 (채널별, 좌측 책갈피 드로어) ──
  const [strategyOpen,setStrategyOpen]=useState(false);
  const [strategy,setStrategy]=useState(()=>{try{return JSON.parse(localStorage.getItem("promo_strategy")||"{}");}catch{return{};}});
  useEffect(()=>{(async()=>{
    const db=await getSupabase();
    const{data,error}=await db.from("promo_strategy").select("*");
    if(!error&&Array.isArray(data)&&data.length){
      const m={}; data.forEach(r=>{m[r.channel]=r.memo||"";});
      setStrategy(m); localStorage.setItem("promo_strategy",JSON.stringify(m));
    }
  })();},[]);
  const setStrategyMemo=(ch,memo)=>setStrategy(prev=>{const n={...prev,[ch]:memo};localStorage.setItem("promo_strategy",JSON.stringify(n));return n;});
  const saveStrategy=async ch=>{const db=await getSupabase();await db.from("promo_strategy").upsert({channel:ch,memo:strategy[ch]||""},{onConflict:"channel"});};
  const [strategySaved,setStrategySaved]=useState(false);
  const saveAllStrategy=async()=>{const db=await getSupabase();await db.from("promo_strategy").upsert(PROMO_PLATFORMS.map(ch=>({channel:ch,memo:strategy[ch]||""})),{onConflict:"channel"});setStrategySaved(true);setTimeout(()=>setStrategySaved(false),1500);};

  // ── 공백 알림 (채널별 등록 프로모션 사이 빈 기간) ──
  const [gapOpen,setGapOpen]=useState(false);
  const [calcOpen,setCalcOpen]=useState(false);
  const [mallCalcOpen,setMallCalcOpen]=useState(false); // 자사몰 세일율 계산기(베타) 모달
  const promoGaps=useMemo(()=>{
    // UTC 기준 날짜 연산 — 로컬(KST) 파싱 후 toISOString 변환 시 하루 밀리는 버그 방지
    const addDays=(d,n)=>{const dt=new Date(d+"T00:00:00Z");dt.setUTCDate(dt.getUTCDate()+n);return dt.toISOString().slice(0,10);};
    // channels=null → 전체 채널, 배열이면 해당 채널만 통합해 빈 기간 계산
    const gapsFor=(channels)=>{
      const ivs=promos.filter(p=>!hiddenIds.has(p.id)&&p.start_date&&p.end_date&&(!channels||channels.includes(p.platform)))
        .map(p=>({s:String(p.start_date).slice(0,10),e:String(p.end_date).slice(0,10)}))
        .filter(iv=>iv.s&&iv.e&&iv.s<=iv.e)
        .sort((a,b)=>a.s>b.s?1:-1);
      const merged=[];
      ivs.forEach(iv=>{const last=merged[merged.length-1]; if(last&&iv.s<=addDays(last.e,1)){if(iv.e>last.e)last.e=iv.e;} else merged.push({...iv});});
      const gaps=[];
      for(let i=1;i<merged.length;i++){const gs=addDays(merged[i-1].e,1),ge=addDays(merged[i].s,-1); if(gs<=ge) gaps.push({start:gs,end:ge,days:Math.round((new Date(ge)-new Date(gs))/86400000)+1});}
      return {count:ivs.length,gaps};
    };
    return {all:gapsFor(null),core:gapsFor(["자사몰","29CM","오프라인 스토어"])};
  },[promos,hiddenIds]);

  // 등록/가려진 카드 공용 본문
  //   상단 행: 기간 · 상세 · 핀셋(최대 10) · 첨부
  //   하단 영역: 할인율 매트릭스 (카드 전체 폭, 비-컴팩트)
  const PIN_LIMIT=10;
  const promoDetailBody=(p)=>{
    const pins=p.pinned_products||[];
    const visiblePins=pins.slice(0,PIN_LIMIT);
    const hiddenCount=Math.max(0,pins.length-PIN_LIMIT);
    return (
    <div style={{marginTop:30}}>
      {/* 상단 정보 행 — 프로모션명 아래 3배 여백 (10 → 30) */}
      <div style={{display:"flex",gap:16,alignItems:"flex-start",flexWrap:"wrap"}}>
        <div style={{flex:"0 0 auto",minWidth:120}}>
          <div style={{fontSize:11,color:D.black,fontWeight:700,marginBottom:2}}>기간</div>
          {[p.start_date,p.end_date].map((dt,i)=>{
            const [d,t]=(dt||"").split("T");
            const wd=d?["일","월","화","수","목","금","토"][new Date(d+"T00:00:00").getDay()]:"";
            return <div key={i}><span style={{fontWeight:700,fontSize:13,color:D.text}}>{d}</span>{wd&&<span style={{fontSize:12,color:D.textSub,marginLeft:3}}>({wd})</span>}{t&&<span style={{fontSize:12,color:D.textSub,marginLeft:4}}>{t}</span>}</div>;
          })}
        </div>
        <div style={{flex:pins.length>0?"1 1 240px":"3 1 480px",minWidth:200,fontSize:12,color:D.textSub,whiteSpace:"pre-wrap"}}>
          <div style={{fontSize:11,color:D.black,fontWeight:700,marginBottom:2}}>상세 내용</div>
          {(p.content||p.memo)
            ? <div dangerouslySetInnerHTML={{__html:(p.content||p.memo)}}/>
            : "—"}
        </div>
        {pins.length>0&&(
          <div style={{flex:"2 1 280px",minWidth:200}}>
            <div style={{fontSize:11,color:D.black,fontWeight:700,marginBottom:2}}>
              핀셋 상품{pins.length>PIN_LIMIT&&<span style={{opacity:.6,fontWeight:500,marginLeft:6}}>총 {pins.length}개 중 {PIN_LIMIT}개 표시</span>}
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
              {visiblePins.map((pp,i)=>(
                <span key={i} onClick={()=>togglePinHighlight(p,i)}
                  title={pp.memo?`${pp.name} · ${pp.memo}`:pp.name}
                  style={{cursor:"pointer",borderRadius:8,padding:"1px 7px",fontSize:10,maxWidth:160,
                    overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                    background:pp.highlight?D.black:D.surfaceAlt,
                    color:pp.highlight?"#fff":D.textSub,
                    border:`1px solid ${pp.highlight?D.black:D.border}`}}>{pp.name}</span>
              ))}
              {hiddenCount>0&&(
                <button data-capture-hide onClick={()=>setPinnedModal({promo:p})}
                  style={{cursor:"pointer",borderRadius:8,padding:"1px 8px",fontSize:11,fontWeight:500,
                    background:"transparent",color:D.black,border:`1px dashed ${D.black}`,whiteSpace:"nowrap",
                    fontFamily:"-apple-system, BlinkMacSystemFont, 'SF Pro Text', 'SF Pro', 'Helvetica Neue', Arial, sans-serif",
                    letterSpacing:"-0.01em"}}>
                  + {hiddenCount}개 더보기
                </button>
              )}
            </div>
          </div>
        )}
        <div style={{flex:"0 0 auto",marginLeft:"auto",minWidth:150}}>
          <div style={{fontSize:10,color:D.textMeta,marginBottom:2}}>첨부 파일{p.submit_date?` · 제출일 ${p.submit_date}`:""}</div>
          <div
            onDragOver={e=>{e.preventDefault();setTableFileDragOver(p.id);}}
            onDragLeave={e=>{if(!e.currentTarget.contains(e.relatedTarget))setTableFileDragOver(null);}}
            onDrop={e=>{e.preventDefault();setTableFileDragOver(null);
              addFilesFromList(e.dataTransfer.files,(p.files||[]).length,f=>addFileToPromo(p.id,f));
            }}
            style={{display:"flex",flexDirection:"column",gap:3,
              border:`1px dashed ${tableFileDragOver===p.id?D.blue:"transparent"}`,
              borderRadius:4,padding:tableFileDragOver===p.id?4:0,
              background:tableFileDragOver===p.id?"#eef3ff":"transparent",minHeight:22,transition:"all 0.15s"}}>
            {(p.files||[]).map((f,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:4}}>
                <span style={{fontSize:11,color:D.textSub,wordBreak:"break-all",flex:1,textAlign:"left"}}
                  title={f.name}>📎 {f.name}</span>
                <button data-capture-hide onClick={()=>setFilePreview(f)} title="미리보기"
                  style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,color:D.textMeta,
                    cursor:"pointer",padding:"1px 6px",fontSize:10,whiteSpace:"nowrap",flexShrink:0,lineHeight:1}}>
                  미리보기
                </button>
                <button data-capture-hide onClick={()=>removeFileFromPromo(p.id,i)} title="첨부파일 삭제"
                  style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,color:D.textMeta,
                    cursor:"pointer",padding:"1px 6px",fontSize:10,whiteSpace:"nowrap",flexShrink:0}}>첨부파일 삭제</button>
              </div>
            ))}
            {(p.files||[]).length<3&&(
              tableFileDragOver===p.id
                ?<span style={{fontSize:11,color:D.blue,textAlign:"center",padding:"2px 0"}}>여기에 놓기 ↓</span>
                :<button data-capture-hide onClick={()=>{setFileAddTarget(p.id);fileInputRef.current.value="";fileInputRef.current.click();}}
                  style={{background:"transparent",border:`1px dashed ${D.border}`,borderRadius:3,
                    padding:"2px 6px",fontSize:11,color:D.textMeta,cursor:"pointer",
                    whiteSpace:"nowrap",alignSelf:"flex-start"}}>+ 파일 추가</button>
            )}
            {!(p.files||[]).length&&tableFileDragOver!==p.id&&<span style={{color:D.textMeta,fontSize:11}}>—</span>}
          </div>
        </div>
      </div>
      {/* 하단: 할인율 매트릭스 — 카드 전체 폭으로 펼치고 표 자체는 좌우 중앙 정렬 */}
      {computeDiscountMatrix(p.discount_plan||{}).hasGroup&&(
        <div style={{marginTop:14,paddingTop:12,borderTop:`1px dashed ${D.border}`}}>
          <div style={{fontSize:11,color:D.black,fontWeight:700,marginBottom:6}}>할인율 매트릭스</div>
          <DiscountPlanView plan={p.discount_plan} marks={p.discount_marks||{}}
            onToggleGroup={g=>toggleMark(p,"groups",g)} onToggleCircle={k=>toggleMark(p,"circles",k)}
            compact={false}/>
        </div>
      )}
    </div>
  );};

  return (
    <div style={{padding:"20px 24px",maxWidth:1600,margin:"0 auto"}}>
      {/* 좌측 책갈피 — 채널별 프로모션 전략 메모 */}
      <div style={{position:"fixed",top:140,left:0,zIndex:1500}}>
        {!strategyOpen?(gapOpen?null:(
          <button onClick={()=>{setStrategyOpen(true);setGapOpen(false);}}
            style={{writingMode:"vertical-rl",background:D.surface,color:D.text,border:`1px solid ${D.borderMid}`,
              borderRadius:"0 8px 8px 0",padding:"14px 7px",fontSize:12,fontWeight:700,cursor:"pointer",
              letterSpacing:"0.12em",boxShadow:"2px 2px 8px rgba(0,0,0,0.18)"}}>
            프로모션 전략 메모
          </button>
        )):(
          <div style={{width:"70vw",maxHeight:"72vh",overflowY:"auto",background:D.black,
            border:`1px solid ${D.border}`,borderRadius:"0 10px 10px 0",
            boxShadow:"4px 4px 24px rgba(0,0,0,0.18)",padding:"14px 16px",
            fontFamily:"'Noto Sans KR','Pretendard',sans-serif"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <b style={{fontSize:11,color:"#fff"}}>프로모션 전략 메모</b>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <button onClick={saveAllStrategy}
                  style={{background:strategySaved?D.green:"#fff",border:"none",borderRadius:5,
                    color:strategySaved?"#fff":D.black,cursor:"pointer",fontSize:11,fontWeight:700,padding:"4px 12px"}}>
                  {strategySaved?"저장됨 ✓":"저장"}
                </button>
                <button onClick={()=>setStrategyOpen(false)}
                  style={{background:"none",border:"none",color:"#fff",cursor:"pointer",fontSize:15,lineHeight:1}}>✕</button>
              </div>
            </div>
            {["자사몰","29CM","오프라인 스토어","무신사"].map(ch=>(
              <div key={ch} style={{marginBottom:10}}>
                <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:3}}>
                  <span style={{width:7,height:7,borderRadius:"50%",background:chColor(ch),display:"inline-block"}}/>
                  <span style={{fontSize:11,fontWeight:600,color:"#fff"}}>{ch}</span>
                </div>
                <textarea value={strategy[ch]||""} onChange={e=>setStrategyMemo(ch,e.target.value)} onBlur={()=>saveStrategy(ch)}
                  style={{width:"100%",boxSizing:"border-box",minHeight:ch==="무신사"?54:ch==="29CM"?324:108,resize:"vertical",
                    background:"#2a2a2a",border:`1px solid #444`,borderRadius:6,padding:"6px 8px",fontSize:11,color:"#fff",
                    fontFamily:"'Noto Sans KR','Pretendard',sans-serif",lineHeight:1.6}}/>
              </div>
            ))}
          </div>
        )}
      </div>
      {/* 좌측 책갈피 2 — 공백 알림 (등록 프로모션 사이 빈 기간) */}
      <div style={{position:"fixed",top:332,left:0,zIndex:1500}}>
        {!gapOpen?(strategyOpen?null:(
          <button onClick={()=>{setGapOpen(true);setStrategyOpen(false);}}
            style={{writingMode:"vertical-rl",background:D.surface,color:D.text,border:`1px solid ${D.borderMid}`,
              borderRadius:"0 8px 8px 0",padding:"14px 7px",fontSize:12,fontWeight:700,cursor:"pointer",
              letterSpacing:"0.12em",boxShadow:"2px 2px 8px rgba(0,0,0,0.18)"}}>
            공백 알림
          </button>
        )):(
          <div style={{width:"min(440px,90vw)",maxHeight:"72vh",overflowY:"auto",background:D.surface,
            border:`1px solid ${D.border}`,borderRadius:"0 10px 10px 0",
            boxShadow:"4px 4px 24px rgba(0,0,0,0.18)",padding:"14px 16px",
            fontFamily:"'Noto Sans KR','Pretendard',sans-serif"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
              <b style={{fontSize:12,color:D.black}}>공백 알림</b>
              <button onClick={()=>setGapOpen(false)}
                style={{background:"none",border:"none",color:D.textMeta,cursor:"pointer",fontSize:15,lineHeight:1}}>✕</button>
            </div>
            <div style={{fontSize:10,color:D.textMeta,marginBottom:10}}>프로모션이 하나도 없는 빈 기간</div>
            {[
              {key:"all",title:"전체 채널 공백",info:promoGaps.all},
              {key:"core",title:"자사몰 · 29CM · 오프라인 스토어 공백",info:promoGaps.core},
            ].map(sec=>(
              <div key={sec.key} style={{marginBottom:14}}>
                <div style={{fontSize:11,fontWeight:700,color:D.black,marginBottom:5,paddingBottom:3,borderBottom:`1px solid ${D.border}`}}>{sec.title}</div>
                {sec.info.count===0?(
                  <div style={{fontSize:11,color:D.textMeta}}>등록된 프로모션이 없습니다.</div>
                ):sec.info.gaps.length===0?(
                  <div style={{fontSize:11,color:D.green}}>공백 없음 — 전 기간 진행 중</div>
                ):sec.info.gaps.map((g,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:D.text,lineHeight:1.9}}>
                    <span style={{color:D.red,fontWeight:700}}>•</span>
                    <span style={{fontWeight:600}}>{g.start} ~ {g.end}</span>
                    <span style={{color:D.textMeta}}>({g.days}일 공백)</span>
                  </div>
                ))}
              </div>
            ))}
          </div>
        )}
      </div>
      {/* 좌측 책갈피 3 — 29CM 세일율 계산기 (모달) */}
      {!strategyOpen&&!gapOpen&&!calcOpen&&!mallCalcOpen&&(
        <button onClick={()=>setCalcOpen(true)}
          style={{position:"fixed",top:470,left:0,zIndex:1500,writingMode:"vertical-rl",
            background:D.surface,color:D.text,border:`1px solid ${D.borderMid}`,borderRadius:"0 8px 8px 0",
            padding:"14px 7px",fontSize:12,fontWeight:700,cursor:"pointer",letterSpacing:"0.12em",
            boxShadow:"2px 2px 8px rgba(0,0,0,0.18)"}}>
          29CM 세일율 계산기
        </button>
      )}
      {/* 좌측 책갈피 4 — 자사몰 세일율 계산기 (29CM 계산기 아래) */}
      {!strategyOpen&&!gapOpen&&!calcOpen&&!mallCalcOpen&&(
        <button onClick={()=>setMallCalcOpen(true)}
          style={{position:"fixed",top:655,left:0,zIndex:1500,writingMode:"vertical-rl",
            background:D.surface,color:D.text,border:`1px solid ${D.borderMid}`,borderRadius:"0 8px 8px 0",
            padding:"14px 7px",fontSize:12,fontWeight:700,cursor:"pointer",letterSpacing:"0.12em",
            boxShadow:"2px 2px 8px rgba(0,0,0,0.18)"}}>
          자사몰 세일율 계산기
        </button>
      )}
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
        </div>
      </div>

      {/* 추가 버튼 — 타임라인 위 중앙 (폼 열리면 숨김) */}
      {!showForm&&(
        <div style={{display:"flex",justifyContent:"center",marginBottom:16}}>
          <button onClick={()=>setShowForm(true)}
            style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
              padding:"8px 22px",fontSize:13,cursor:"pointer",fontWeight:600}}>
            + 프로모션 추가
          </button>
        </div>
      )}

      {showForm&&(
        <Card style={{marginBottom:20}}>
          <div style={{fontWeight:600,fontSize:14,marginBottom:16}}>프로모션 추가</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,alignItems:"start"}}>
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>프로모션명</div>
              <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={inp} placeholder="예: 오픈 기념 할인"/>
            </div>
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>플랫폼</div>
              <div style={{display:"flex",gap:3}}>
                {PROMO_PLATFORMS.map(p=>(
                  <button key={p} onClick={()=>setForm(f=>({...f,platform:p}))}
                    style={{flex:1,background:form.platform===p?chColor(p):"transparent",
                      color:form.platform===p?"#fff":D.textSub,
                      border:`1px solid ${form.platform===p?chColor(p):D.border}`,
                      borderRadius:4,padding:"4px 3px",fontSize:9,cursor:"pointer"}}>
                    {p}
                  </button>
                ))}
              </div>
            </div>
            {[["시작일시","start_date"],["종료일시","end_date"]].map(([label,field])=>(
              <div key={field}>
                <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>{label}</div>
                <DateDrop id={`promo_${field}`}
                  value={form[field]?.slice(0,10)||""}
                  onChange={v=>{const time=form[field]?.slice(10)||"";setForm(f=>({...f,[field]:v+time}));}}
                  calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor}
                  placeholder="날짜 선택"/>
                <div style={{display:"flex",gap:3,marginTop:4}}>
                  {[["T10:00","오전 10시"],["T11:00","오전 11시"],["T23:59","오후 23:59"]].map(([time,tl])=>(
                    <button key={time} onClick={()=>{
                      const base=form[field]?form[field].slice(0,10):localDate(0);
                      setForm(f=>({...f,[field]:`${base}${time}`}));
                    }} style={{flex:1,fontSize:11,padding:"3px 2px",
                      background:form[field]&&form[field].includes(time)?D.black:D.surfaceAlt,
                      color:form[field]&&form[field].includes(time)?"#fff":D.textSub,
                      border:`1px solid ${form[field]&&form[field].includes(time)?D.black:D.border}`,
                      borderRadius:3,cursor:"pointer",whiteSpace:"nowrap"}}>
                      {tl}
                    </button>
                  ))}
                  <button onClick={()=>setAddTimePickerFor(addTimePickerFor===field?null:field)}
                    style={{fontSize:11,padding:"3px 8px",
                    background:addTimePickerFor===field?D.black:"transparent",
                    color:addTimePickerFor===field?"#fff":D.textSub,
                    border:`1px solid ${addTimePickerFor===field?D.black:D.border}`,borderRadius:3,cursor:"pointer",whiteSpace:"nowrap"}}>
                    그외
                  </button>
                  <button onClick={()=>{
                    const base=form[field]?form[field].slice(0,10):"";
                    setForm(f=>({...f,[field]:base}));
                  }} style={{fontSize:11,padding:"3px 6px",background:"transparent",
                    border:`1px solid ${D.border}`,borderRadius:3,cursor:"pointer",color:D.textMeta,whiteSpace:"nowrap"}}>
                    시간 삭제
                  </button>
                </div>
                {addTimePickerFor===field&&(
                  <div style={{marginTop:6,padding:"6px 8px",background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:5,display:"flex",alignItems:"center",gap:6}}>
                    <span style={{fontSize:11,color:D.textSub}}>시간 직접 입력:</span>
                    <input type="time" value={form[field]?form[field].slice(11,16)||"":""}
                      onChange={e=>{
                        const t=e.target.value;
                        const base=form[field]?form[field].slice(0,10):localDate(0);
                        setForm(f=>({...f,[field]:t?`${base}T${t}`:base}));
                      }}
                      style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:4,
                        padding:"3px 6px",fontSize:12,color:D.text,fontFamily:"inherit"}}/>
                  </div>
                )}
              </div>
            ))}
          </div>
          <div style={{marginTop:10,display:"grid",gridTemplateColumns:"3fr 1fr",gap:8}}>
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>프로모션 내용 <span style={{opacity:.6,fontWeight:400}}>(드래그 후 하이라이트 가능)</span></div>
              <HighlightEditor value={form.content||form.memo||""}
                onChange={v=>setForm(f=>({...f,content:v,memo:v}))}
                placeholder="할인율, 대상 상품, 조건 등 (선택)" minHeight={144}/>
            </div>
            <div>
              <div style={{fontSize:11,color:D.textMeta,marginBottom:4}}>첨부 파일 <span style={{opacity:.6}}>(최대 3개)</span></div>
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
            <PinnedProductPicker value={form.pinned_products||[]}
              onChange={v=>setForm(f=>({...f,pinned_products:v}))} orders={orders}/>
          </div>
          <div style={{marginTop:14}}>
            <DiscountPlanEditor value={form.discount_plan}
              onChange={v=>setForm(f=>({...f,discount_plan:v}))}
              calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} idPrefix="adddp"
              platform={form.platform}/>
          </div>
          {/* 제출일 + 저장/취소 — 폼 하단 중앙 */}
          <div style={{marginTop:18,display:"flex",justifyContent:"center",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <span style={{fontSize:11,color:D.textMeta,fontWeight:600}}>제출일</span>
              <DateDrop id="promo_submit_date" value={form.submit_date||""}
                onChange={v=>setForm(f=>({...f,submit_date:v}))}
                calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="제출일 선택"/>
            </div>
            <button onClick={addPromo}
              style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
                padding:"8px 26px",fontSize:13,cursor:"pointer",fontWeight:600}}>저장</button>
            <button onClick={()=>setShowForm(false)}
              style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"8px 18px",fontSize:13,cursor:"pointer",color:D.textSub}}>취소</button>
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
            <span style={{position:"absolute",left:`${tp}%`,top:-14,transform:"translateX(-50%)",
              fontSize:11,color:D.primary,fontWeight:700,background:D.surface,
              padding:"0 4px",borderRadius:3,whiteSpace:"nowrap"}}>오늘</span>
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
                      onClick={()=>{
                        const el=promoCardRefs.current[promo.id];
                        if(el) el.scrollIntoView({behavior:"smooth",block:"start"});
                      }}
                      title={`${promo.name} — 카드로 이동`}
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
        <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>기간별 플랫폼 순매출</div>
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

      {/* 오른쪽 도트 네비 — 카드 영역(maxWidth:1600 centered)을 침범하지 않도록 우측 여백으로 push */}
      {(()=>{
        const navPromos=[...promos].filter(p=>!hiddenIds.has(p.id)).sort((a,b)=>a.start_date>b.start_date?1:-1);
        if(navPromos.length===0) return null;
        // 카드 폭 1600 가정. 가능한 경우 오른쪽 마진 안쪽으로 배치, 아닐 경우 viewport 오른쪽 끝.
        return (
          <div data-capture-hide
            style={{position:"fixed",
              right:"max(4px, calc((100vw - 1600px) / 2 - 156px))",
              top:"22vh",zIndex:80,
              display:"flex",flexDirection:"column",gap:3,maxHeight:"70vh",overflowY:"auto",
              padding:"6px 7px",width:150,boxSizing:"border-box",
              background:`${D.surface}f2`,border:`1px solid ${D.border}`,borderRadius:6,
              boxShadow:"0 2px 8px rgba(0,0,0,0.08)",
              fontFamily:"'Noto Sans KR','Pretendard',sans-serif"}}>
            <div style={{fontSize:8,color:D.textMeta,fontWeight:700,letterSpacing:"0.06em",marginBottom:2}}>프로모션 바로가기</div>
            {navPromos.map(p=>(
              <button key={p.id} onClick={()=>{
                const el=promoCardRefs.current[p.id];
                if(el) el.scrollIntoView({behavior:"smooth",block:"start"});
              }} title={`${p.platform} ${p.name}`}
                style={{display:"flex",alignItems:"center",gap:5,
                  background:"transparent",border:"none",padding:"1px 2px",borderRadius:3,
                  cursor:"pointer",textAlign:"left",fontFamily:"inherit",
                  width:"100%",overflow:"hidden"}}>
                <span style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),flexShrink:0,
                  border:isEnded(p)?`1px solid ${D.border}`:"none",opacity:isEnded(p)?0.55:1}}/>
                <span style={{fontSize:9,color:D.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                  textDecoration:isEnded(p)?"line-through":"none",opacity:isEnded(p)?0.6:1}}>
                  {p.platform} {p.name}
                </span>
              </button>
            ))}
          </div>
        );
      })()}

      {/* 등록된 프로모션 — 카드 목록 */}
      {promos.filter(p=>!hiddenIds.has(p.id)).length>0&&(
        <Card>
          <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>등록된 프로모션</div>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {[...promos].filter(p=>!hiddenIds.has(p.id)).sort((a,b)=>a.start_date>b.start_date?1:-1).map(p=>{
              const ended=isEnded(p);
              const isEditing=editingPromoId===p.id;
              const inp3={background:D.surface,border:`1px solid ${D.border}`,borderRadius:5,
                padding:"6px 10px",fontSize:11,color:D.text,width:"100%",boxSizing:"border-box",
                fontFamily:"'Noto Sans KR','Pretendard',sans-serif"};
              return (
              <div key={p.id} ref={el=>{promoCardRefs.current[p.id]=el;}}
                style={{position:"relative",border:`1px solid ${ended&&!isEditing?D.borderMid:D.black}`,borderRadius:10,
                  padding:"14px 16px",background:D.surface,
                  fontFamily:"'Noto Sans KR','Pretendard',sans-serif",lineHeight:1.7}}>
                {ended&&!isEditing&&(
                  <div data-capture-hide aria-hidden="true"
                    style={{position:"absolute",inset:0,borderRadius:10,
                      background:"rgba(255,255,255,0.62)",pointerEvents:"none",zIndex:1}}/>
                )}
                {isEditing?(
                  <div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,marginBottom:8}}>
                      <div>
                        <div style={{fontSize:11,color:D.textMeta,marginBottom:3}}>프로모션명</div>
                        <input value={editPromoForm.name} onChange={e=>setEditPromoForm(f=>({...f,name:e.target.value}))} style={inp3}/>
                      </div>
                      <div>
                        <div style={{fontSize:11,color:D.textMeta,marginBottom:3}}>플랫폼</div>
                        <div style={{display:"flex",gap:3}}>
                          {PROMO_PLATFORMS.map(pl=>(
                            <button key={pl} onClick={()=>setEditPromoForm(f=>({...f,platform:pl}))}
                              style={{flex:1,background:editPromoForm.platform===pl?chColor(pl):"transparent",
                                color:editPromoForm.platform===pl?"#fff":D.textSub,
                                border:`1px solid ${editPromoForm.platform===pl?chColor(pl):D.border}`,
                                borderRadius:4,padding:"3px 2px",fontSize:8,cursor:"pointer"}}>{pl}</button>
                          ))}
                        </div>
                      </div>
                      {[["시작일시","start_date"],["종료일시","end_date"]].map(([label,field])=>(
                        <div key={field}>
                          <div style={{fontSize:11,color:D.textMeta,marginBottom:3}}>{label}</div>
                          <DateDrop id={`editdate_${field}_${p.id}`} value={editPromoForm[field]?.slice(0,10)||""}
                            onChange={v=>{const time=editPromoForm[field]?.slice(10)||"";setEditPromoForm(f=>({...f,[field]:v+time}));}}
                            calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="날짜 선택"/>
                          <div style={{display:"flex",gap:3,marginTop:3,flexWrap:"wrap"}}>
                            {[["T10:00","10시"],["T11:00","11시"],["T23:59","23:59"]].map(([time,tl])=>(
                              <button key={time} onClick={()=>{
                                const base=(editPromoForm[field]||localDate(0)).slice(0,10);
                                setEditPromoForm(f=>({...f,[field]:`${base}${time}`}));
                              }} style={{flex:1,fontSize:10,padding:"2px",background:D.surfaceAlt,
                                border:`1px solid ${D.border}`,borderRadius:3,cursor:"pointer",color:D.textSub}}>{tl}</button>
                            ))}
                            <button onClick={()=>{
                              const fkey=`${p.id}|${field}`;
                              setEditTimePickerFor(editTimePickerFor===fkey?null:fkey);
                            }} style={{fontSize:10,padding:"2px 6px",
                              background:editTimePickerFor===`${p.id}|${field}`?D.black:"transparent",
                              color:editTimePickerFor===`${p.id}|${field}`?"#fff":D.textSub,
                              border:`1px solid ${editTimePickerFor===`${p.id}|${field}`?D.black:D.border}`,
                              borderRadius:3,cursor:"pointer",whiteSpace:"nowrap"}}>그외</button>
                            <button onClick={()=>{
                              const base=(editPromoForm[field]||"").slice(0,10);
                              setEditPromoForm(f=>({...f,[field]:base}));
                            }} style={{fontSize:10,padding:"2px 6px",background:"transparent",
                              border:`1px solid ${D.border}`,borderRadius:3,cursor:"pointer",color:D.textMeta,whiteSpace:"nowrap"}}>시간 삭제</button>
                          </div>
                          {editTimePickerFor===`${p.id}|${field}`&&(
                            <div style={{marginTop:5,padding:"5px 8px",background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:5,display:"flex",alignItems:"center",gap:6}}>
                              <span style={{fontSize:10,color:D.textSub}}>시간:</span>
                              <input type="time" value={editPromoForm[field]?editPromoForm[field].slice(11,16)||"":""}
                                onChange={e=>{
                                  const t=e.target.value;
                                  const base=(editPromoForm[field]||localDate(0)).slice(0,10);
                                  setEditPromoForm(f=>({...f,[field]:t?`${base}T${t}`:base}));
                                }}
                                style={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:4,
                                  padding:"2px 5px",fontSize:11,color:D.text,fontFamily:"inherit"}}/>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                    <div style={{marginBottom:8,display:"grid",gridTemplateColumns:"3fr 1fr",gap:8}}>
                      <div>
                        <div style={{fontSize:11,color:D.textMeta,marginBottom:3}}>프로모션 내용 <span style={{opacity:.6,fontWeight:400}}>(드래그 후 하이라이트 가능)</span></div>
                        <HighlightEditor value={editPromoForm.content||""}
                          onChange={v=>setEditPromoForm(f=>({...f,content:v}))}
                          placeholder="할인율, 대상 상품, 조건 등" minHeight={80}/>
                      </div>
                      <div>
                        <div style={{fontSize:11,color:D.textMeta,marginBottom:3}}>제출일</div>
                        <DateDrop id={`editsubmit_${p.id}`} value={editPromoForm.submit_date||""}
                          onChange={v=>setEditPromoForm(f=>({...f,submit_date:v}))}
                          calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} placeholder="제출일 선택"/>
                      </div>
                    </div>
                    <div style={{marginBottom:8}}>
                      <PinnedProductPicker value={editPromoForm.pinned_products||[]}
                        onChange={v=>setEditPromoForm(f=>({...f,pinned_products:v}))} orders={orders}/>
                    </div>
                    <div style={{marginBottom:8}}>
                      <DiscountPlanEditor value={editPromoForm.discount_plan}
                        onChange={v=>setEditPromoForm(f=>({...f,discount_plan:v}))}
                        calOpenFor={calOpenFor} setCalOpenFor={setCalOpenFor} idPrefix={`editdp${editingPromoId}`}
                        platform={editPromoForm.platform}/>
                    </div>
                    <div style={{display:"flex",gap:6,justifyContent:"center"}}>
                      <button onClick={savePromoEdit}
                        style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                          padding:"7px 22px",fontSize:13,cursor:"pointer",fontWeight:600}}>저장</button>
                      <button onClick={()=>setEditingPromoId(null)}
                        style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                          padding:"7px 16px",fontSize:13,cursor:"pointer",color:D.textSub}}>취소</button>
                    </div>
                  </div>
                ):(
                <>
                  {/* 우측 상단 액션 — 이미지 다운로드 + 수정 (캡처 시 숨김) */}
                  <div data-capture-hide style={{position:"absolute",top:10,right:12,display:"flex",gap:6,alignItems:"center"}}>
                    <CaptureBtn cardRef={{get current(){return promoCardRefs.current[p.id];}}}
                      filename={`프로모션_${p.name||p.id}`} DC={{border:D.border,sub:D.textMeta}}/>
                    <button onClick={()=>startEditPromo(p)} title="수정"
                      style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                        color:D.textSub,cursor:"pointer",padding:"3px 9px",fontSize:11,fontWeight:600}}>✎ 수정</button>
                  </div>
                  {/* 채널 + 이름 + 임팩트 분석 (이름 오른쪽) */}
                  {(()=>{
                    const ichg=(p.start_date&&p.start_date.slice(0,10)<=today)?promoRevenueChg(p,revenues,storeSales).chg:null;
                    const canImpact=!!(p.start_date&&p.start_date.slice(0,10)<=today);
                    return (
                    <div style={{display:"flex",alignItems:"center",gap:6,paddingRight:130,flexWrap:"wrap",marginTop:14}}>
                      <span style={{width:7,height:7,borderRadius:"50%",background:chColor(p.platform),flexShrink:0}}/>
                      <span style={{fontSize:11,color:D.textMeta}}>{p.platform}</span>
                      <b style={{fontSize:28,color:D.black,lineHeight:1.2}}>{p.name}</b>
                      {canImpact&&(
                        <button onClick={()=>setImpactModal(p)} data-capture-hide
                          style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                            padding:"3px 11px",fontSize:11,cursor:"pointer",fontWeight:600,marginLeft:4,
                            position:"relative",zIndex:2}}>임팩트 분석</button>
                      )}
                      {ichg!=null&&ichg>=20&&(
                        <span data-capture-hide title={`직전 동일기간 대비 매출 +${ichg.toFixed(1)}%`}
                          style={{display:"inline-flex",alignItems:"center",justifyContent:"center",
                            width:18,height:18,borderRadius:"50%",background:MUTE_GREEN,color:"#fff",
                            fontSize:13,fontWeight:800,lineHeight:1,flexShrink:0}}>⤴</span>
                      )}
                      {canImpact&&p.platform==="자사몰"&&(
                        <button onClick={()=>setProfitModal(p)} data-capture-hide title="마진율 계산 (베타 테스트 중)"
                          style={{background:"transparent",color:MUTE_BLUE,border:`1px solid ${MUTE_BLUE}`,borderRadius:5,
                            padding:"3px 9px",fontSize:11,cursor:"pointer",fontWeight:600,marginLeft:4,
                            position:"relative",zIndex:2,display:"inline-flex",alignItems:"center",gap:4}}>
                          마진율 계산
                          <span style={{fontSize:8,fontWeight:800,color:"#fff",background:MUTE_BLUE,borderRadius:3,padding:"1px 4px"}}>베타</span>
                        </button>
                      )}
                      {ended&&<span style={{fontSize:11,fontWeight:500,color:D.red,marginLeft:4}}>· 종료된 프로모션</span>}
                    </div>
                    );
                  })()}
                  {promoDetailBody(p)}
                  {/* 하단 액션 — 가리기 / 프로모션 삭제 */}
                  <div data-capture-hide style={{display:"flex",gap:6,marginTop:10,justifyContent:"flex-end"}}>
                    {ended&&(
                      <button onClick={()=>hidePromo(p)} title="가리기 (종료 프로모션 로그)"
                        style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,
                          color:D.textMeta,cursor:"pointer",padding:"3px 10px",fontSize:11}}>가리기</button>
                    )}
                    <button onClick={()=>{
                      if(window.confirm(`"${p.name}" 프로모션을 삭제하시겠습니까?\n삭제 후 되돌릴 수 없습니다.`)){
                        delPromo(p.id);
                      }
                    }}
                      style={{background:"transparent",border:`1px solid ${D.red}55`,borderRadius:4,
                        color:D.red,cursor:"pointer",padding:"3px 10px",fontSize:11}}>프로모션 삭제</button>
                  </div>
                </>
                )}
              </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* 프로모션 검색 — 기본 닫힘, 헤더 클릭으로 토글 */}
      <Card style={{marginTop:12}}>
        <button onClick={()=>setSearchOpen(o=>!o)} type="button"
          style={{background:"transparent",border:"none",padding:0,cursor:"pointer",fontFamily:"inherit",
            display:"flex",alignItems:"center",gap:6,marginBottom:searchOpen?12:0,
            fontWeight:600,fontSize:14,color:D.black}}>
          <span style={{display:"inline-block",transform:searchOpen?"rotate(90deg)":"rotate(0deg)",
            transition:"transform 0.15s",color:D.textMeta,fontSize:11}}>▶</span>
          프로모션 검색
        </button>
        {searchOpen&&(<>
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
        </>)}
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
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {[...hiddenLog].sort((a,b)=>b.hidden_at>a.hidden_at?1:-1).map(h=>(
              <div key={h.id} style={{position:"relative",border:`1px solid ${D.black}`,borderRadius:10,
                padding:"14px 16px",background:D.surface,opacity:0.92,
                fontFamily:"'Noto Sans KR','Pretendard',sans-serif",lineHeight:1.7}}>
                {/* 우측 상단 — 일괄 액션을 위한 체크박스 */}
                <div data-capture-hide style={{position:"absolute",top:10,right:12,display:"flex",gap:6,alignItems:"center"}}>
                  <label style={{display:"inline-flex",alignItems:"center",gap:4,cursor:"pointer",fontSize:11,color:D.textMeta}}>
                    <input type="checkbox" checked={selHiddenIds.has(h.id)}
                      onChange={ev=>{const s=new Set(selHiddenIds);ev.target.checked?s.add(h.id):s.delete(h.id);setSelHiddenIds(s);}}
                      style={{cursor:"pointer"}}/>
                    선택
                  </label>
                </div>
                {/* 채널 + 이름 + 임팩트 분석 — 진행 중 카드와 동일 형태 */}
                <div style={{display:"flex",alignItems:"center",gap:6,paddingRight:130,flexWrap:"wrap",marginTop:14}}>
                  <span style={{width:7,height:7,borderRadius:"50%",background:chColor(h.platform),flexShrink:0}}/>
                  <span style={{fontSize:11,color:D.textMeta}}>{h.platform}</span>
                  <b style={{fontSize:28,color:D.black,lineHeight:1.2}}>{h.name}</b>
                  <button onClick={()=>setImpactModal(h)} data-capture-hide
                    style={{background:D.black,color:"#fff",border:"none",borderRadius:5,
                      padding:"3px 11px",fontSize:11,cursor:"pointer",fontWeight:600,marginLeft:4}}>임팩트 분석</button>
                  <span style={{fontSize:11,fontWeight:500,color:D.red,marginLeft:4}}>· 가려진 종료 프로모션</span>
                </div>
                {promoDetailBody(h)}
              </div>
            ))}
          </div>
        </Card>
      )}

      {impactModal&&<PromoImpactModal promo={impactModal} onClose={()=>setImpactModal(null)} revenues={revenues} storeSales={storeSales} orders={orders}/>}
      {profitModal&&<ProfitCalcModal promo={profitModal} orders={orders} onClose={()=>setProfitModal(null)}/>}
      {filePreview&&<FilePreviewModal file={filePreview} onClose={()=>setFilePreview(null)}/>}
      {calcOpen&&<SaleCalcModal onClose={()=>setCalcOpen(false)}
        onCreatePromo={(prefill)=>{
          setForm({name:"",platform:prefill.platform||"29CM",start_date:"",end_date:"",
            memo:prefill.content||"",content:prefill.content||"",files:[],
            discount_plan:prefill.discount_plan||{products:{period:{start:"",end:""},rows:[]},coupons:[]},
            pinned_products:[],submit_date:""});
          setShowForm(true);
          setCalcOpen(false);
        }}/>}
      {mallCalcOpen&&<OwnMallSaleCalcModal onClose={()=>setMallCalcOpen(false)}
        onCreatePromo={(prefill)=>{
          setForm({name:"",platform:prefill.platform||"자사몰",start_date:"",end_date:"",
            memo:prefill.content||"",content:prefill.content||"",files:[],
            discount_plan:prefill.discount_plan||{products:{period:{start:"",end:""},rows:[]},coupons:[]},
            pinned_products:[],submit_date:""});
          setShowForm(true);
          setMallCalcOpen(false);
        }}/>}
      {pinnedModal&&<PinnedListModal promo={pinnedModal.promo} onToggleHighlight={idx=>togglePinHighlight(pinnedModal.promo,idx)} onClose={()=>setPinnedModal(null)}/>}
    </div>
  );
}

// ─────────────────────────────────────────────
// 프로모션 임팩트 분석 모달 — 종료된 프로모션 임팩트 분석
//   - 일별 매출: 직전 동일 기간(점선 전) → 프로모션 기간(점선 후)
//   - Top 20: 프로모션 기간(주문일 기준) + 해당 채널 · status="배송" 상품 판매 랭킹
// ─────────────────────────────────────────────
// 프로모션 매출 증감률 — 직전 동일기간 대비 (임팩트 분석 요약과 동일 기준)
function promoRevenueChg(promo, revenues=[], storeSales=[]){
  const ch=promo.platform;
  const dayMs=86400000;
  const todayStr=localDate(0);
  const yesterdayStr=localDate(-1);
  const promoStart=String(promo.start_date||"").slice(0,10);
  const promoEndRaw=String(promo.end_date||"").slice(0,10);
  if(!promoStart||!promoEndRaw) return {prevTotal:0,promoTotal:0,chg:null};
  const isOngoing=promoEndRaw>=todayStr; // 종료일이 미래거나 당일(오늘)이면 오늘 데이터 미완성 → 전일까지만 집계
  const promoEnd=isOngoing?(yesterdayStr>=promoStart?yesterdayStr:promoStart):promoEndRaw;
  const lenDays=Math.max(0,(new Date(promoEnd)-new Date(promoStart))/dayMs)+1;
  const prevStart=new Date(new Date(promoStart).getTime()-lenDays*dayMs).toISOString().slice(0,10);
  let prevTotal=0,promoTotal=0;
  if(ch==="오프라인 스토어"){
    storeSales.forEach(r=>{
      const d=r.sale_date; if(!d||d<prevStart||d>promoEnd) return;
      const amt=r.status==="배송"?(r.amount||0):r.status==="반품"?-(r.amount||0):0;
      if(d<promoStart) prevTotal+=amt; else promoTotal+=amt;
    });
  }else{
    revenues.forEach(r=>{
      if(r.channel!==ch) return;
      const d=r.date; if(!d||d<prevStart||d>promoEnd) return;
      const amt=(r.amount||0)-(r.refund_amount||0);
      if(d<promoStart) prevTotal+=amt; else promoTotal+=amt;
    });
  }
  const chg=prevTotal>0?((promoTotal-prevTotal)/prevTotal*100):null;
  return {prevTotal,promoTotal,chg};
}

// 29CM 할인율 계산기 — 가격대별 P75 목표 할인율을 쿠폰율로 역산 (단일 시뮬 + 일괄 엑셀)
const CALC_SLOTS=[
  {id:"S1",name:"구간 1",min:0,     max:100000,   disc:32,color:"#854F0B",bg:"#FAEEDA",n:1597,range:"< ₩100,000"},
  {id:"S2",name:"구간 2",min:100000,max:150000,   disc:23,color:"#534AB7",bg:"#EEEDFE",n:422, range:"₩100,000 ~ ₩149,999"},
  {id:"S3",name:"구간 3",min:150000,max:200000,   disc:30,color:"#0F6E56",bg:"#E1F5EE",n:131, range:"₩150,000 ~ ₩199,999"},
  {id:"S4",name:"구간 4",min:200000,max:250000,   disc:26,color:"#993C1D",bg:"#FAECE7",n:18,  range:"₩200,000 ~ ₩249,999"},
  {id:"S5",name:"구간 5",min:250000,max:Infinity, disc:40,color:"#3C3489",bg:"#EEEDFE",n:16,  range:"≥ ₩250,000"},
];
const CALC_CONDS={S1:"정가 < ₩100,000",S2:"₩100,000 ≤ 정가 < ₩150,000",S3:"₩150,000 ≤ 정가 < ₩200,000",S4:"₩200,000 ≤ 정가 < ₩250,000",S5:"정가 ≥ ₩250,000"};
const calcClassify=list=>{for(const s of CALC_SLOTS) if(list>=s.min&&list<s.max) return s; return CALC_SLOTS[CALC_SLOTS.length-1];};
const wonFmt=n=>new Intl.NumberFormat("ko-KR").format(Math.round(n));
// 기본 할인율 일의 자리가 6~9면 다음 10단위로 올림 (예: 7→10, 16~19→20, 26~29→30)
const roundUpBaseDisc=d=>d%10>=6?Math.ceil(d/10)*10:d;
const calcReverse=(list,p75,coupon)=>{
  const factorFinal=1-p75/100, factorCoupon=1-coupon/100;
  if(factorCoupon<=0) return {baseDisc:0,basePrice:list,finalPrice:list,finalDisc:0};
  let baseFactor=factorFinal/factorCoupon; if(baseFactor>1) baseFactor=1;
  // 올림 규칙 적용 후, 올림된 할인율로 가격 재계산
  const baseDisc=roundUpBaseDisc(Math.round((1-baseFactor)*1000)/10);
  baseFactor=1-baseDisc/100;
  const basePrice=Math.round(list*baseFactor/10)*10;
  const finalPrice=Math.round(basePrice*factorCoupon/10)*10;
  // 결과값으로 나온 실제 최종 할인율 (정가 대비)
  const finalDisc=list>0?Math.round((1-finalPrice/list)*1000)/10:0;
  return {baseDisc, basePrice, finalPrice, finalDisc};
};
function SaleCalcModal({ onClose, onCreatePromo, onAttachInlineCalc, attachMode, initialCoupon, initialPrimaryType, initialCoupons }){
  // initialCoupons: [{rate, type, name}] — 여러 쿠폰을 받아서 첫번째 → 기본, 나머지 → 누적 쿠폰
  const initCoupons=Array.isArray(initialCoupons)&&initialCoupons.length>0?initialCoupons:null;
  const initPrimary=initCoupons?initCoupons[0]:null;
  const initStacks=initCoupons?initCoupons.slice(1):[];
  // 디폴트 — 자주 쓰는 시나리오 첫번째(29CM 지원 쿠폰 15% · 장바구니 · 채널부담)
  const [coupon,setCoupon]=useState(initPrimary?initPrimary.rate:(initialCoupon!=null?initialCoupon:15));
  const [primaryType,setPrimaryType]=useState(
    initPrimary&&COUPON_TYPE_BY_KEY[initPrimary.type]?initPrimary.type
    :(initialPrimaryType&&COUPON_TYPE_BY_KEY[initialPrimaryType]?initialPrimaryType:"cart")
  ); // 기본 쿠폰 타입
  const [primaryBurden,setPrimaryBurden]=useState("channel"); // self=자사부담, channel=채널부담
  const [primaryShareRate,setPrimaryShareRate]=useState(50); // 분담 type일 때 채널부담률 %
  const [stackCoupons,setStackCoupons]=useState(
    initStacks.map(c=>({
      rate:+c.rate||0,
      type:COUPON_TYPE_BY_KEY[c.type]?c.type:"product",
      burden:"channel",
      shareRate:50,
    }))
  ); // [{rate, type, burden, shareRate}] — 추가 쿠폰 목록
  const [scenarioIdx,setScenarioIdx]=useState(0); // 선택한 시나리오 인덱스
  const [listPrice,setListPrice]=useState(129000);
  const [processed,setProcessed]=useState(null);
  const [summary,setSummary]=useState("");
  const [showResults,setShowResults]=useState(false);
  const [dragOver,setDragOver]=useState(false);
  // 일괄 표의 임시 제거 / 체크 — 다운로드/재업로드/리로드 시 복원
  const [removedRows,setRemovedRows]=useState(()=>new Set());
  const [checkedRows,setCheckedRows]=useState(()=>new Set());
  // 드래그 다중 선택 — null | "add" | "remove"
  const dragSelectModeRef=useRef(null);
  const fileRef=useRef(null);
  const wbRef=useRef(null), fnameRef=useRef(""), sheetRef=useRef(""), rawRef=useRef(null);
  const cpnPrimary=(()=>{const v=Number(coupon); return isNaN(v)||v<0?0:Math.min(v,60);})();
  // 현재 상태가 프리셋과 일치하는지 — 일치하면 시나리오 영역을 페이드 처리
  const presetActiveIdx=(()=>{
    const s0=stackCoupons[0];
    if(cpnPrimary===15&&primaryType==="cart"&&stackCoupons.length===0) return 0;
    if(cpnPrimary===15&&primaryType==="cart"&&stackCoupons.length===1&&s0?.type==="product"&&Number(s0?.rate)===10) return 1;
    if(cpnPrimary===29&&primaryType==="share"&&Number(primaryShareRate)===40&&stackCoupons.length===0) return 2;
    return -1;
  })();
  const resetPreset=()=>{
    setCoupon(0);setPrimaryType("product");setPrimaryBurden("self");setPrimaryShareRate(50);
    setStackCoupons([]);setScenarioIdx(0);
  };
  // 입력된 모든 쿠폰 (이름·타입·부담 주체·분담률 포함)
  const allCoupons=useMemo(()=>{
    // 장바구니 쿠폰은 자사가 발행할 수 없으므로 burden 을 항상 channel 로 고정
    const burdenFor=(type,b)=>type==="cart"?"channel":(b||"self");
    const list=[{rate:cpnPrimary,type:primaryType,burden:burdenFor(primaryType,primaryBurden),shareRate:primaryShareRate,label:"기본"}];
    stackCoupons.forEach((s,i)=>{
      const r=Number(s.rate);
      const rate=isNaN(r)||r<0?0:Math.min(r,60);
      const sr=Number(s.shareRate);
      const type=s.type||"product";
      list.push({
        rate,
        type,
        burden:burdenFor(type,s.burden),
        shareRate:isNaN(sr)?50:Math.max(0,Math.min(100,sr)),
        label:`Case ${i+1}`,
      });
    });
    return list.filter(c=>c.rate>0);
  },[cpnPrimary,primaryType,primaryBurden,primaryShareRate,stackCoupons]);
  // 타입 규칙 기반 가능한 시나리오 (단독 + cart×product 쌍 + 분담 단독)
  const scenarios=useMemo(()=>{
    const by={product:[],cart:[],share:[]};
    allCoupons.forEach(c=>by[c.type]?.push(c));
    const sc=[];
    // 단독 (모든 쿠폰) — "장바구니 15%" 형식
    allCoupons.forEach(c=>{
      const tInfo=COUPON_TYPE_BY_KEY[c.type];
      sc.push({label:`${tInfo.short} ${c.rate}%`,items:[c]});
    });
    // 쌍: 상품 × 장바구니 — 상품 쿠폰을 먼저 적용한 뒤 장바구니 쿠폰 적용
    by.cart.forEach(cc=>{
      by.product.forEach(pp=>{
        sc.push({label:`상품 ${pp.rate}% + 장바구니 ${cc.rate}%`,items:[pp,cc]});
      });
    });
    sc.forEach(s=>{
      s.factor=s.items.reduce((f,c)=>f*(1-c.rate/100),1);
      s.eff=Math.round((1-s.factor)*1000)/10;
    });
    // 효과 큰 순 정렬 후 Case 번호 부여
    sc.sort((a,b)=>b.eff-a.eff).forEach((s,i)=>{s.caseNum=i+1;});
    return sc;
  },[allCoupons]);
  const selectedScenario=scenarios[scenarioIdx]||scenarios[0]||{factor:1-cpnPrimary/100,eff:cpnPrimary,items:[]};
  // 유효 쿠폰율 = 선택된 시나리오의 결과
  const cpn=selectedScenario.eff;
  // 단일 시뮬 — 기본 할인율 사용자 수동 오버라이드 (디폴트 0%)
  const [singleManualBase,setSingleManualBase]=useState(null);
  const single=useMemo(()=>{
    if(!listPrice) return null;
    const bd=singleManualBase!=null
      ? Math.max(0,Math.min(100,parseFloat(singleManualBase)||0))
      : 0; // 디폴트 0% (구간/P75 제거)
    const baseFactor=1-bd/100;
    const basePrice=Math.round(listPrice*baseFactor/10)*10;
    const finalPrice=Math.round(basePrice*(1-cpn/100)/10)*10;
    const finalDisc=listPrice>0?Math.round((1-finalPrice/listPrice)*1000)/10:0;
    return {baseDisc:bd,basePrice,finalPrice,finalDisc};
  },[listPrice,cpn,singleManualBase]);
  // 부담 주체별 차감액 계산 → 자사 매출 → 수수료 → 마진
  // 프런트 할인(기본 할인율)은 항상 자사부담. 각 쿠폰은 burden(self/channel) 혹은 share 의 shareRate(채널부담률) 에 따라 분배.
  // 정산 모델 (해석 B + 수수료 프런트 판매가 기준): 수수료는 프런트 판매가 기준 부과,
  //                      채널 보전(channelBurden) 은 수수료 없이 그대로 정산에 가산.
  //                      net = finalPrice − fee + channelBurden
  // 부가세 처리 (옵션 B · 양쪽 VAT 포함): 인벤토리 공급가는 세전(공급가액) 이므로
  //                      마진 계산 시 ×1.1 적용. 결제액은 이미 소비자가(VAT 포함).
  const computeMargin=useCallback((list,baseDisc,items,supply)=>{
    if(!list) return {finalPrice:0,selfBurden:0,channelBurden:0,fee:0,feeRate:28,net:0,margin:0,markup:0,supplyIncVat:0};
    let priceAfter=list*(1-baseDisc/100);
    let selfBurden=list*(baseDisc/100); // 프런트는 자사
    let channelBurden=0;
    (items||[]).forEach(c=>{
      const cut=priceAfter*((c.rate||0)/100);
      if(c.type==="share"){
        const channelPart=cut*((c.shareRate||0)/100);
        channelBurden+=channelPart;
        selfBurden+=cut-channelPart;
      } else if(c.burden==="channel") {
        channelBurden+=cut;
      } else {
        selfBurden+=cut;
      }
      priceAfter-=cut;
    });
    // 채널 수수료율 = 28% − 기본 세일율(baseDisc) 10% 단위마다 -1%p (최소 0)
    const fr=Math.max(0,28-Math.floor(baseDisc/10));
    const basePriceR=Math.round(list*(1-baseDisc/100)/10)*10; // 프런트 판매가 (쿠폰 적용 전)
    const finalPriceR=Math.round(priceAfter/10)*10;       // 고객 결제액 (10원 단위)
    const channelBurdenR=Math.round(channelBurden/10)*10; // 채널 보전 (10원 단위)
    const fee=Math.round(basePriceR*(fr/100)/10)*10;      // 수수료는 프런트 판매가 기준
    const net=finalPriceR-fee+channelBurdenR;             // 자사 정산
    const supplyIncVat=Math.round((supply||0)*1.1);       // 공급가 (세포) = 인벤토리 공급가액 × 1.1
    const margin=net-supplyIncVat;
    // 실수령 마크업 = 실수령액(정산액) ÷ 원가
    const markup=supplyIncVat>0?Math.round(net/supplyIncVat*100)/100:0;
    return {finalPrice:finalPriceR,selfBurden:Math.round(selfBurden),channelBurden:channelBurdenR,fee,feeRate:fr,net,margin,markup,supplyIncVat};
  },[]);
  // 인벤토리 공급가 맵 + 상품 목록 — 최근 스냅샷 기준
  const [invMap,setInvMap]=useState({});
  const [invProducts,setInvProducts]=useState([]); // [{name,code,selling,supply}]
  // 사용자 업로드 공급가 오버라이드 (calc_supply_override 테이블) — 인벤토리보다 우선
  const [overrideMap,setOverrideMap]=useState({}); // {normName: supplyPrice}
  const [overrideStatus,setOverrideStatus]=useState("");
  const [overrideDrag,setOverrideDrag]=useState(false);
  const overrideFileRef=useRef(null);
  // 상품명 정규화 — 인벤토리 / 오버라이드 / supplyOf 모두 공통 사용
  const normProdName=useCallback(s=>String(s||"").trim().toLowerCase()
    .replace(/[​‌‍ ﻿]/g,"")
    .replace(/[\s_·•~\-]*\d+\s*colou?rs?\b/gi,"")
    .replace(/\[[^\]]*\]/g,"")
    .replace(/\([^)]*\)/g,"")
    .replace(/[_·•~]+/g,"")
    .replace(/\s+/g,"").trim(),[]);
  // 단일 시뮬 검색·선택 상태
  const [singleQuery,setSingleQuery]=useState("");
  const [singleSelected,setSingleSelected]=useState(null); // {name,code,selling,supply}
  // 마지막 단일 분석 상품 Supabase 영구 저장 (sale_calc_last_single id=1)
  const lastSingleLoaded=useRef(false);
  useEffect(()=>{
    let alive=true;
    (async()=>{
      try{
        const db=await getSupabase();
        const {data,error}=await db.from("sale_calc_last_single").select("*").eq("id",1).maybeSingle();
        if(error||!data||!alive){ lastSingleLoaded.current=true; return; }
        if(data.product_name){
          setSingleSelected({
            name:data.product_name,
            code:data.product_code||"",
            selling:data.product_selling||0,
            supply:data.product_supply||0,
          });
        }
        if(data.list_price) setListPrice(data.list_price);
        if(data.manual_base!=null&&data.manual_base!=="") setSingleManualBase(data.manual_base);
      }catch{}
      lastSingleLoaded.current=true;
    })();
    return()=>{alive=false;};
  },[]);
  useEffect(()=>{
    if(!lastSingleLoaded.current) return;
    const t=setTimeout(async()=>{
      try{
        const db=await getSupabase();
        await db.from("sale_calc_last_single").upsert({
          id:1,
          product_name:singleSelected?.name||null,
          product_code:singleSelected?.code||null,
          product_selling:singleSelected?.selling||null,
          product_supply:singleSelected?.supply||null,
          list_price:listPrice||null,
          manual_base:singleManualBase==null?null:String(singleManualBase),
          updated_at:new Date().toISOString(),
        },{onConflict:"id"});
      }catch{}
    },600);
    return()=>clearTimeout(t);
  },[singleSelected,listPrice,singleManualBase]);
  // 오버라이드 로드 + 업로드 / 삭제 핸들러
  useEffect(()=>{
    let alive=true;
    (async()=>{
      try{
        const db=await getSupabase();
        let all=[],from=0;const PAGE=1000;
        while(true){
          const {data,error}=await db.from("calc_supply_override").select("norm_name,supply_price").range(from,from+PAGE-1);
          if(error||!data||data.length===0) break;
          all=all.concat(data);
          if(data.length<PAGE) break;
          from+=PAGE;
        }
        if(!alive) return;
        const m={};
        all.forEach(r=>{if(r.norm_name) m[r.norm_name]=r.supply_price||0;});
        setOverrideMap(m);
        if(Object.keys(m).length>0) setOverrideStatus(`업로드 공급가 ${Object.keys(m).length}건 로드 완료 (인벤토리보다 우선 적용)`);
      }catch{}
    })();
    return()=>{alive=false;};
  },[]);
  const handleSupplyUpload=async file=>{
    setOverrideStatus("파일 분석 중…");
    try{
      const XLSX=await getXLSX();
      const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,dateNF:"YYYY-MM-DD"});
      if(!raw||raw.length<2){ setOverrideStatus("데이터 행이 없습니다"); return; }
      // 헤더 행 자동 탐색 (첫 5행 내) — 인벤토리 업로더와 동일 휴리스틱
      let headerRow=0;
      for(let i=0;i<Math.min(raw.length,5);i++){
        const text=(raw[i]||[]).join(" ").toLowerCase();
        if(/상품명|product_name|product|name|품명/.test(text)
           &&/공급가|supply_price|원가|cost|단가/.test(text)){
          headerRow=i; break;
        }
      }
      const headers=(raw[headerRow]||[]).map(h=>String(h||"").trim());
      const colMap=mapInvCols(headers);   // 기존 인벤토리 업로더의 별칭 매핑 재사용
      if(colMap.product_name===undefined||colMap.supply_price===undefined){
        setOverrideStatus(`'상품명' / '공급가' 컬럼을 찾지 못했습니다 — 헤더: ${headers.filter(Boolean).join(" / ")}`);
        return;
      }
      const rows=[]; const seen=new Set();
      raw.slice(headerRow+1).forEach(r=>{
        const nm=String(r[colMap.product_name]||"").trim();
        const spRaw=String(r[colMap.supply_price]||"").replace(/[^\d.-]/g,"");
        const sp=Math.max(0,Math.round(Number(spRaw)||0));
        if(!nm||sp<=0) return;
        const nz=normProdName(nm);
        if(!nz||seen.has(nz)) return; seen.add(nz);
        rows.push({product_name:nm,norm_name:nz,supply_price:sp,updated_at:new Date().toISOString()});
      });
      if(rows.length===0){ setOverrideStatus("매칭 가능한 행 없음 — 공급가가 0보다 큰 행이 있는지 확인"); return; }
      const db=await getSupabase();
      const BATCH=500;
      for(let i=0;i<rows.length;i+=BATCH){
        const slice=rows.slice(i,i+BATCH);
        const {error}=await db.from("calc_supply_override").upsert(slice,{onConflict:"norm_name"});
        if(error) throw error;
      }
      const nm={...overrideMap};
      rows.forEach(r=>{nm[r.norm_name]=r.supply_price;});
      setOverrideMap(nm);
      setOverrideStatus(`${rows.length}건 저장 완료 · 누적 ${Object.keys(nm).length}건 (헤더 ${headerRow+1}행)`);
    }catch(e){ setOverrideStatus("업로드 실패: "+(e?.message||e)); }
  };
  const clearOverride=async()=>{
    if(!window.confirm(`저장된 업로드 공급가 ${Object.keys(overrideMap).length}건을 모두 삭제하시겠습니까?`)) return;
    try{
      const db=await getSupabase();
      const {error}=await db.from("calc_supply_override").delete().gte("id",0);
      if(error) throw error;
      setOverrideMap({});
      setOverrideStatus("업로드 공급가 전체 삭제됨");
    }catch(e){ setOverrideStatus("삭제 실패: "+(e?.message||e)); }
  };
  useEffect(()=>{
    let alive=true;
    (async()=>{
      try{
        const db=await getSupabase();
        const {data:latest}=await db.from("inventory_snapshot").select("snapshot_date").order("snapshot_date",{ascending:false}).limit(1);
        const d=latest?.[0]?.snapshot_date; if(!d||!alive) return;
        let all=[],from=0;const PAGE=1000;
        while(true){
          const {data,error}=await db.from("inventory_snapshot").select("product_name,product_code,supply_price,selling_price")
            .eq("snapshot_date",d).range(from,from+PAGE-1);
          if(error||!data||data.length===0) break;
          all=all.concat(data);
          if(data.length<PAGE) break;
          from+=PAGE;
        }
        if(!alive) return;
        const m={};
        const productMap={};
        all.forEach(r=>{
          const sp=r.supply_price||0;
          const n=(r.product_name||"").trim();
          const nz=normProdName(n);
          const c=(r.product_code||"").trim();
          if(n&&!m["n:"+n]) m["n:"+n]=sp;
          if(nz&&!m["z:"+nz]) m["z:"+nz]=sp;
          if(c&&!m["c:"+c]) m["c:"+c]=sp;
          if(n&&!productMap[n]){
            productMap[n]={name:n,code:c,selling:r.selling_price||0,supply:sp};
          }
        });
        setInvMap(m);
        setInvProducts(Object.values(productMap).sort((a,b)=>a.name.localeCompare(b.name,"ko")));
      }catch{}
    })();
    return()=>{alive=false;};
  },[normProdName]);
  // 상품명/코드 → 공급가 매칭
  //  0) 사용자 업로드 오버라이드 정규화명 일치 (최우선)
  //  1) 정확 코드 일치
  //  2) 원본명 정확 일치
  //  3) 정규화된 이름 정확 일치 (옵션 접미사·괄호·구분자 제거)
  //  4) 정규화된 이름끼리 부분 포함 (≥ 4글자)
  const supplyOf=useCallback((name,code)=>{
    const raw=(name||"").trim();
    const nz=normProdName(raw);
    if(nz&&overrideMap[nz]) return overrideMap[nz];
    if(code){const v=invMap["c:"+String(code).trim()]; if(v) return v;}
    if(!raw) return 0;
    if(invMap["n:"+raw]) return invMap["n:"+raw];
    if(!nz) return 0;
    if(invMap["z:"+nz]) return invMap["z:"+nz];
    // fallback — 정규화 키끼리 부분 포함 (양방향, 4자 이상)
    const zkeys=Object.keys(invMap).filter(k=>k.startsWith("z:"));
    for(const k of zkeys){
      const kn=k.slice(2);
      if(kn.length>=4&&(kn.includes(nz)||nz.includes(kn))) return invMap[k];
    }
    return 0;
  },[invMap,overrideMap,normProdName]);
  // 행별 baseDisc 수동 오버라이드 (사용자가 직접 수정한 값) — cpn 변경에도 유지
  useEffect(()=>{
    if(!rawRef.current) return;
    setProcessed(prev=>{
      const factorCoupon=1-cpn/100;
      return rawRef.current.map((r,i)=>{
        const manualBase=prev?.[i]?.manualBase;
        // 기본 할인율: 수동값 우선, 없으면 0% (구간/P75 제거)
        const bd=manualBase!=null?manualBase:0;
        const baseFactor=1-bd/100;
        const basePrice=Math.round(r.list*baseFactor/10)*10;
        const finalPrice=Math.round(basePrice*factorCoupon/10)*10;
        const finalDisc=r.list>0?Math.round((1-finalPrice/r.list)*1000)/10:0;
        const supply=supplyOf(r.name,r.code);
        const supplyIncVat=Math.round(supply*1.1);
        const costRatio=finalPrice>0&&supply>0?Math.round(supplyIncVat/finalPrice*1000)/10:0;
        const m=computeMargin(r.list,bd,selectedScenario.items||[],supply);
        return {...r,baseDisc:bd,basePrice,finalPrice,finalDisc,manualBase,supply,supplyIncVat,costRatio,...m};
      });
    });
  },[cpn,supplyOf,computeMargin,selectedScenario]);

  // 사용자가 기본 할인율을 직접 수정하면 프런트 판매가·최종가·원가율 즉시 재계산
  // rowId(원본 엑셀 row 번호) 기반 — 임시 제거된 행이 있어 shown 의 인덱스가
  // processed 인덱스와 어긋나도 안전하게 정확한 행을 찾는다
  const updateBaseDisc=(rowId,newBase)=>{
    const v=Math.max(0,Math.min(100,Number(newBase)||0));
    setProcessed(prev=>prev.map(r=>{
      if(r.row!==rowId) return r;
      const basePrice=Math.round(r.list*(1-v/100)/10)*10;
      const supply=r.supply||0;
      const supplyIncVat=Math.round(supply*1.1);
      const m=computeMargin(r.list,v,selectedScenario.items||[],supply);
      const finalDisc=r.list>0?Math.round((1-m.finalPrice/r.list)*1000)/10:0;
      return {...r,baseDisc:v,basePrice,manualBase:v,supplyIncVat,finalDisc,...m};
    }));
  };
  const resetBaseDisc=(rowId)=>{
    setProcessed(prev=>prev.map(r=>{
      if(r.row!==rowId) return r;
      const bd=0; // 디폴트 0%
      const basePrice=Math.round(r.list*(1-bd/100)/10)*10;
      const finalPrice=Math.round(basePrice*(1-cpn/100)/10)*10;
      const finalDisc=r.list>0?Math.round((1-finalPrice/r.list)*1000)/10:0;
      const supply=r.supply||0;
      const supplyIncVat=Math.round(supply*1.1);
      const m=computeMargin(r.list,bd,selectedScenario.items||[],supply);
      return {...r,baseDisc:bd,basePrice,finalPrice,finalDisc,manualBase:undefined,supplyIncVat,...m};
    }));
  };
  // 전체 기본 할인율 일괄 적용/복귀
  const applyAllBaseDisc=(newBase)=>{
    const v=Math.max(0,Math.min(100,Number(newBase)||0));
    setProcessed(prev=>prev.map(r=>{
      const basePrice=Math.round(r.list*(1-v/100)/10)*10;
      const supply=r.supply||0;
      const supplyIncVat=Math.round(supply*1.1);
      const m=computeMargin(r.list,v,selectedScenario.items||[],supply);
      const finalDisc=r.list>0?Math.round((1-m.finalPrice/r.list)*1000)/10:0;
      return {...r,baseDisc:v,basePrice,manualBase:v,supplyIncVat,finalDisc,...m};
    }));
  };
  const resetAllBaseDisc=()=>{
    setProcessed(prev=>prev.map(r=>{
      const bd=0; // 디폴트 0%
      const basePrice=Math.round(r.list*(1-bd/100)/10)*10;
      const finalPrice=Math.round(basePrice*(1-cpn/100)/10)*10;
      const finalDisc=r.list>0?Math.round((1-finalPrice/r.list)*1000)/10:0;
      const supply=r.supply||0;
      const supplyIncVat=Math.round(supply*1.1);
      const m=computeMargin(r.list,bd,selectedScenario.items||[],supply);
      return {...r,baseDisc:bd,basePrice,finalPrice,finalDisc,manualBase:undefined,supplyIncVat,...m};
    }));
  };
  // 일괄 적용 입력값
  const [bulkBaseDisc,setBulkBaseDisc]=useState("");
  // 결론 도출 — 펼친 묶음 상품 그룹 (기본 세일율 %)
  const [expandedGroup,setExpandedGroup]=useState(null);
  // 예시 파일 저장/로드 — 최근 업로드 1개를 Supabase 에 보관해 다른 세션에서도 즉시 미리보기
  const saveExampleFile=async file=>{
    try{
      const buf=await file.arrayBuffer();
      const bytes=new Uint8Array(buf);
      let bin=""; for(let i=0;i<bytes.length;i++) bin+=String.fromCharCode(bytes[i]);
      const b64=btoa(bin);
      const db=await getSupabase();
      await db.from("calc_last_file").upsert({id:1,filename:file.name,content_b64:b64,uploaded_at:new Date().toISOString()});
    }catch(err){ console.warn("예시 파일 저장 실패",err); }
  };
  const loadExampleFile=async(opts={})=>{
    try{
      const db=await getSupabase();
      const {data,error}=await db.from("calc_last_file").select("filename,content_b64").eq("id",1).maybeSingle();
      if(error) throw error;
      if(!data){
        if(!opts.silent){ setSummary("저장된 예시 파일이 없습니다. 한 번 업로드해 주세요."); setShowResults(true); }
        return;
      }
      const bin=atob(data.content_b64);
      const bytes=new Uint8Array(bin.length);
      for(let i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i);
      const blob=new Blob([bytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const f=new File([blob],data.filename||"example.xlsx");
      await handleFile(f,{skipSave:true});
    }catch(err){
      console.error("예시 파일 로드 실패",err);
      if(!opts.silent){ setSummary("예시 파일 로드 실패: "+(err?.message||err)); setShowResults(true); }
    }
  };
  // 모달 진입 시 마지막 업로드 예시 파일을 디폴트로 자동 로드
  const exampleAutoLoaded=useRef(false);
  useEffect(()=>{
    if(exampleAutoLoaded.current) return;
    exampleAutoLoaded.current=true;
    loadExampleFile({silent:true});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);
  const handleFile=async(file,opts={})=>{
    fnameRef.current=file.name.replace(/\.xlsx?$/i,"");
    try{
      const XLSX=await getXLSX();
      const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellStyles:true});
      wbRef.current=wb;
      if(!opts.skipSave) saveExampleFile(file);
      const sheet=wb.SheetNames.find(n=>/상품가격|할인가|할인/.test(n))||wb.SheetNames[0];
      sheetRef.current=sheet;
      const ws=wb.Sheets[sheet];
      const range=XLSX.utils.decode_range(ws["!ref"]);
      let headerRow=-1;
      for(let r=range.s.r;r<=Math.min(range.s.r+10,range.e.r);r++){
        const cellE=ws[XLSX.utils.encode_cell({r,c:4})];
        if(/정상가/.test(cellE?String(cellE.v||""):"")){headerRow=r;break;}
      }
      if(headerRow===-1) headerRow=3;
      const rows=[];
      for(let r=headerRow+1;r<=range.e.r;r++){
        const code=ws[XLSX.utils.encode_cell({r,c:0})];
        const name=ws[XLSX.utils.encode_cell({r,c:1})];
        const list=ws[XLSX.utils.encode_cell({r,c:4})];
        const listVal=list&&typeof list.v==="number"?list.v:parseFloat(String(list?.v||"").replace(/[^\d.]/g,""));
        if(!listVal||listVal<=0) continue;
        rows.push({row:r,code:code?.v||"",name:name?.v||"",list:Math.round(listVal)});
      }
      rawRef.current=rows; setShowResults(true);
      // 새 업로드 시 제거/체크 상태 초기화
      setRemovedRows(new Set()); setCheckedRows(new Set());
      if(!rows.length){ setSummary("데이터 행을 찾지 못했습니다. E열에 정상가가 있는지 확인하세요."); setProcessed([]); return; }
      setProcessed(rows.map(r=>{
        // 디폴트 기본 할인율 0% (구간/P75 제거)
        const bd=0;
        const basePrice=Math.round(r.list*(1-bd/100)/10)*10;
        const finalPrice=Math.round(basePrice*(1-cpn/100)/10)*10;
        const finalDisc=r.list>0?Math.round((1-finalPrice/r.list)*1000)/10:0;
        const supply=supplyOf(r.name,r.code);
        const supplyIncVat=Math.round(supply*1.1);
        const costRatio=finalPrice>0&&supply>0?Math.round(supplyIncVat/finalPrice*1000)/10:0;
        const m=computeMargin(r.list,bd,selectedScenario.items||[],supply);
        return {...r,baseDisc:bd,basePrice,finalPrice,finalDisc,supply,supplyIncVat,costRatio,...m};
      }));
      setSummary(`${rows.length}개 처리 완료 (시트: ${sheet}, 헤더 ${headerRow+1}행) · 인벤토리 매칭 ${rows.filter(r=>supplyOf(r.name,r.code)>0).length}건`);
    }catch(err){ setShowResults(true); setProcessed([]); setSummary("파일 읽기 실패: "+(err?.message||err)); }
  };
  const download=async()=>{
    if(!wbRef.current||!processed||!processed.length) return;
    const XLSX=await getXLSX();
    const ws=wbRef.current.Sheets[sheetRef.current];
    // 원본 셀 보존 — 기존 서식(z)·스타일(s)은 유지하고 I열 값만 교체. 캐시 표시값(w)·수식(f)만 제거
    // 임시 제거된 행은 I열 갱신 미적용 (원본 그대로 둠)
    processed.forEach(r=>{
      if(removedRows.has(r.row)) return;
      const addr=XLSX.utils.encode_cell({r:r.row,c:8});
      const prev=ws[addr]||{};
      const next={...prev,t:"n",v:r.basePrice};
      delete next.f; delete next.w;
      ws[addr]=next;
    });
    const range=XLSX.utils.decode_range(ws["!ref"]);
    if(range.e.c<8){ range.e.c=8; ws["!ref"]=XLSX.utils.encode_range(range); }
    // I열 변경값을 참조하는 다른 셀 수식이 엑셀에서 자동 재계산되도록 (수식 자체는 보존)
    const wb=wbRef.current;
    wb.Workbook={...(wb.Workbook||{}),CalcPr:{...(wb.Workbook?.CalcPr||{}),fullCalcOnLoad:true}};
    XLSX.writeFile(wb,`${fnameRef.current}_쿠폰${cpn}%_역산.xlsx`,{cellStyles:true});
    // 다운로드 후에도 임시 제거 / 체크 상태 유지 — 모달 재오픈 / 파일 재업로드 시점까지 보존
  };
  // 행 임시 제거 / 복원 / 체크 토글 / 체크 일괄 삭제
  const removeBulkRow=(rowId)=>setRemovedRows(prev=>{const next=new Set(prev);next.add(rowId);return next;});
  const restoreAllBulk=()=>{setRemovedRows(new Set());setCheckedRows(new Set());};
  const toggleCheckBulk=(rowId)=>setCheckedRows(prev=>{const next=new Set(prev);next.has(rowId)?next.delete(rowId):next.add(rowId);return next;});
  const removeCheckedBulk=()=>{
    if(checkedRows.size===0) return;
    setRemovedRows(prev=>{const next=new Set(prev);checkedRows.forEach(i=>next.add(i));return next;});
    setCheckedRows(new Set());
  };
  // 드래그 다중 선택 — 체크박스에서 mouseDown → mode 결정, 다른 행에 mouseEnter 시 적용
  const startDragSelectBulk=(rowId,isChecked)=>{
    dragSelectModeRef.current=isChecked?"remove":"add";
    setCheckedRows(prev=>{const next=new Set(prev);if(dragSelectModeRef.current==="add")next.add(rowId);else next.delete(rowId);return next;});
  };
  const enterDragSelectBulk=(rowId)=>{
    if(!dragSelectModeRef.current) return;
    setCheckedRows(prev=>{const next=new Set(prev);if(dragSelectModeRef.current==="add")next.add(rowId);else next.delete(rowId);return next;});
  };
  useEffect(()=>{
    const onUp=()=>{dragSelectModeRef.current=null;};
    window.addEventListener("mouseup",onUp);
    return ()=>window.removeEventListener("mouseup",onUp);
  },[]);
  const sec={marginBottom:12,border:`1px solid ${D.black}`,borderRadius:10,background:D.surface};
  const summarySty={display:"flex",alignItems:"center",justifyContent:"space-between",padding:"11px 14px",fontSize:11,fontWeight:700,cursor:"pointer",listStyle:"none",color:D.black};
  const inNum={border:`1px solid ${D.border}`,background:D.surface,color:D.text,borderRadius:6,padding:"6px 10px",fontSize:11,width:120,fontFamily:"inherit"};
  const th={padding:"8px 10px",border:`1px solid ${D.border}`,textAlign:"left",fontSize:11};
  const td={padding:"8px 10px",border:`1px solid ${D.border}`,textAlign:"left"};
  const DISP_CAP=500;
  // 표시 행 — 제거된 행은 숨김
  const visibleProcessed=processed?processed.filter(r=>!removedRows.has(r.row)):[];
  const shown=visibleProcessed.slice(0,DISP_CAP);
  return (
    <div onClick={onClose} className="salecalc-overlay" style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2100,
      display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()} className="salecalc"
        style={{background:D.surface,borderRadius:12,width:"min(1440px,80vw)",maxHeight:"92vh",overflowY:"auto",
          border:`1px solid ${D.black}`,
          boxShadow:"0 8px 40px rgba(0,0,0,0.22)",fontFamily:"'Noto Sans KR','Pretendard',sans-serif",fontSize:11,color:D.text}}>
        <style>{`
          body .salecalc, body .salecalc *,
          body .salecalc input, body .salecalc button, body .salecalc select, body .salecalc textarea, body .salecalc pre,
          body .salecalc th, body .salecalc td, body .salecalc span, body .salecalc div, body .salecalc summary, body .salecalc b {
            font-family: 'Pretendard','Noto Sans KR','-apple-system','BlinkMacSystemFont',sans-serif !important;
            font-size: 11px !important;
            line-height: 1.5 !important;
            letter-spacing: normal !important;
            font-stretch: normal !important;
          }
          .salecalc details[open]>summary .chev{transform:rotate(180deg);}
          .salecalc .chev{transition:transform .2s;display:inline-block;}
          .salecalc input[type="number"]::-webkit-inner-spin-button,
          .salecalc input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
          .salecalc input[type="number"] { -moz-appearance: textfield; appearance: textfield; }
          .salecalc table tbody tr:hover td {
            background-image: linear-gradient(rgba(0,0,0,0.045), rgba(0,0,0,0.045));
          }
          @media (max-width: 768px) {
            .salecalc-overlay { padding: 4px !important; align-items: flex-start !important; }
            .salecalc { width: 90% !important; max-width: 90% !important; max-height: calc(100vh - 8px) !important; border-radius: 8px !important; }
            .salecalc .sc-body { padding: 14px 10px 28px !important; }
            .salecalc table th, .salecalc table td { padding: 5px 6px !important; }
          }
        `}</style>
        <div style={{position:"sticky",top:0,background:D.surface,borderBottom:`1px dashed ${D.border}`,
          padding:"14px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",zIndex:5}}>
          <b style={{fontSize:11,color:D.black,fontWeight:700}}>29CM 할인율 계산기</b>
          <button onClick={onClose} style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
            width:32,height:32,cursor:"pointer",fontSize:11,color:D.textMeta}}>✕</button>
        </div>
        <div className="sc-body" style={{padding:"18px 20px 36px"}}>
          <div style={{display:"flex",gap:14,flexWrap:"wrap",fontSize:11,color:D.textSub,padding:"8px 12px",
            background:D.surfaceAlt,borderRadius:6,marginBottom:12}}>
            <span>표본 <b style={{color:D.black}}>2,184개</b></span>
            <span>P75 할인율 = 최종 목표</span>
            <span>판매가 <b style={{color:D.black}}>10원 단위</b></span>
          </div>
          <div style={{padding:"12px 14px",background:D.surface,border:`1px solid ${D.black}`,
            borderRadius:10,marginBottom:16}}>
            {/* 프리셋 — 자주 쓰는 쿠폰 시나리오 3종 */}
            <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:8,marginBottom:12,
              paddingBottom:12,borderBottom:`1px dashed ${D.border}`}}>
              <span style={{fontSize:11,color:D.textMeta,fontWeight:700,letterSpacing:"0.04em"}}>자주 쓰는 시나리오</span>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:10,alignItems:"stretch",width:"100%"}}>
                {[
                  {
                    title:"29CM 지원 쿠폰 15%",
                    detail:"장바구니 쿠폰 15% 단독 · 채널 전액 부담 (29CM 채널 발행 쿠폰이므로 자사 부담 없음). 프런트 판매가에서 15% 차감 후 결제, 차감액은 정산 시 채널이 보전합니다.",
                    apply:()=>{
                      setCoupon(15);setPrimaryType("cart");setPrimaryBurden("channel");setPrimaryShareRate(50);
                      setStackCoupons([]);setScenarioIdx(0);
                    },
                  },
                  {
                    title:"29CM 지원 15% × 브랜드 쿠폰 10%",
                    detail:"장바구니 쿠폰 15% (채널 부담) + 상품 쿠폰 10% (자사 부담) 누적 적용. 상품 쿠폰을 먼저 10% 차감 후 잔여가에서 장바구니 15% 차감. 자사부담분은 마진에서 차감, 채널부담분은 정산 시 보전.",
                    apply:()=>{
                      setCoupon(15);setPrimaryType("cart");setPrimaryBurden("channel");setPrimaryShareRate(50);
                      setStackCoupons([{rate:"10",type:"product",burden:"self",shareRate:50}]);setScenarioIdx(0);
                    },
                  },
                  {
                    title:"이구쿠폰 29%",
                    detail:"분담 쿠폰 29% 단독 적용 (다른 쿠폰과 중복 불가). 차감액의 60% 는 자사 부담(마진 차감), 40% 는 채널 부담(정산 시 보전). 분담 비율은 입력값에 따라 조정 가능.",
                    apply:()=>{
                      setCoupon(29);setPrimaryType("share");setPrimaryBurden("self");setPrimaryShareRate(40);
                      setStackCoupons([]);setScenarioIdx(0);
                    },
                  },
                ].map((p,i)=>{
                  const active=presetActiveIdx===i;
                  return (
                    <button key={i} type="button" onClick={p.apply} title={p.title}
                      style={{minWidth:0,
                        background:active?D.black:"transparent",
                        border:`1px ${active?"solid":"dashed"} ${D.black}`,
                        borderRadius:6,padding:"12px 14px",fontSize:11,cursor:"pointer",
                        color:active?"#fff":D.text,fontFamily:"inherit",letterSpacing:"-0.01em",lineHeight:1.55,textAlign:"left",
                        display:"flex",flexDirection:"column",gap:6,
                        boxShadow:active?"0 2px 8px rgba(0,0,0,0.18)":"none"}}>
                      <span style={{fontWeight:700,fontSize:11,color:active?"#fff":D.black}}>{p.title}</span>
                      <span style={{fontSize:11,color:active?"rgba(255,255,255,0.85)":D.textSub,fontWeight:400}}>{p.detail}</span>
                    </button>
                  );
                })}
              </div>
            </div>
            {/* 기본 쿠폰 행 */}
            <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
              {/* 타입 세그먼트 */}
              <div style={{display:"flex",border:`1px solid ${D.border}`,borderRadius:4,overflow:"hidden"}}>
                {COUPON_TYPES.map(t=>{
                  const active=primaryType===t.key;
                  return <button key={t.key} type="button" onClick={()=>setPrimaryType(t.key)}
                    style={{background:active?t.color:"transparent",color:active?"#fff":D.textMeta,
                      border:"none",padding:"3px 8px",fontSize:11,fontWeight:700,cursor:"pointer",
                      fontFamily:"inherit",lineHeight:1}}>{t.short}</button>;
                })}
              </div>
              {/* 부담 주체 세그먼트 — share/cart 타입은 자사 부담 옵션 없음 (share=shareRate, cart=29CM 채널 발행) */}
              {primaryType==="product"&&(
                <div style={{display:"flex",border:`1px solid ${D.border}`,borderRadius:4,overflow:"hidden"}}>
                  {[{k:"self",l:"자사"},{k:"channel",l:"채널"}].map(b=>{
                    const active=primaryBurden===b.k;
                    return <button key={b.k} type="button" onClick={()=>setPrimaryBurden(b.k)}
                      title={b.k==="self"?"자사부담 → 마진 감소":"채널부담 → 마진 보전"}
                      style={{background:active?D.black:"transparent",color:active?"#fff":D.textMeta,
                        border:"none",padding:"3px 8px",fontSize:11,fontWeight:700,cursor:"pointer",
                        fontFamily:"inherit",lineHeight:1}}>{b.l}</button>;
                  })}
                </div>
              )}
              <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="60" step="1" value={coupon}
                onChange={e=>setCoupon(e.target.value)} style={{...inNum,width:70}}/>
              <span style={{fontSize:11,color:D.blue}}>%</span>
              {primaryType==="share"&&(
                <>
                  <span style={{fontSize:11,color:D.textMeta}}>채널 분담</span>
                  <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="100" step="1" value={primaryShareRate}
                    onChange={e=>setPrimaryShareRate(Number(e.target.value)||0)} style={{...inNum,width:60}}/>
                  <span style={{fontSize:11,color:D.textMeta}}>%</span>
                </>
              )}
              <button onClick={()=>setStackCoupons([...stackCoupons,{rate:"",type:"product",burden:"self",shareRate:50}])}
                style={{background:D.blue,color:"#fff",border:"none",borderRadius:5,
                  padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:600,marginLeft:"auto"}}>
                + 쿠폰 추가
              </button>
            </div>
            {/* 추가 쿠폰 행들 */}
            {stackCoupons.map((sc,i)=>{
              const t=sc.type||"product";
              const burden=sc.burden||"self";
              return (
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginTop:8,flexWrap:"wrap"}}>
                <div style={{display:"flex",border:`1px solid ${D.border}`,borderRadius:4,overflow:"hidden"}}>
                  {COUPON_TYPES.map(typ=>{
                    const active=t===typ.key;
                    return <button key={typ.key} type="button"
                      onClick={()=>{const n=[...stackCoupons];n[i]={...sc,type:typ.key};setStackCoupons(n);}}
                      style={{background:active?typ.color:"transparent",color:active?"#fff":D.textMeta,
                        border:"none",padding:"3px 8px",fontSize:11,fontWeight:700,cursor:"pointer",
                        fontFamily:"inherit",lineHeight:1}}>{typ.short}</button>;
                  })}
                </div>
                {/* 부담 주체 세그먼트 — share/cart 타입은 자사 부담 옵션 없음 */}
                {t==="product"&&(
                  <div style={{display:"flex",border:`1px solid ${D.border}`,borderRadius:4,overflow:"hidden"}}>
                    {[{k:"self",l:"자사"},{k:"channel",l:"채널"}].map(b=>{
                      const active=burden===b.k;
                      return <button key={b.k} type="button"
                        onClick={()=>{const n=[...stackCoupons];n[i]={...sc,burden:b.k};setStackCoupons(n);}}
                        title={b.k==="self"?"자사부담 → 마진 감소":"채널부담 → 마진 보전"}
                        style={{background:active?D.black:"transparent",color:active?"#fff":D.textMeta,
                          border:"none",padding:"3px 8px",fontSize:11,fontWeight:700,cursor:"pointer",
                          fontFamily:"inherit",lineHeight:1}}>{b.l}</button>;
                    })}
                  </div>
                )}
                <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="60" step="1" value={sc.rate}
                  onChange={e=>{const n=[...stackCoupons];n[i]={...sc,rate:e.target.value};setStackCoupons(n);}}
                  style={{...inNum,width:70}}/>
                <span style={{fontSize:11,color:D.blue}}>%</span>
                {t==="share"&&(
                  <>
                    <span style={{fontSize:11,color:D.textMeta}}>채널 분담</span>
                    <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="100" step="1" value={sc.shareRate==null?50:sc.shareRate}
                      onChange={e=>{const n=[...stackCoupons];n[i]={...sc,shareRate:Number(e.target.value)||0};setStackCoupons(n);}}
                      style={{...inNum,width:60}}/>
                    <span style={{fontSize:11,color:D.textMeta}}>%</span>
                  </>
                )}
                <button onClick={()=>setStackCoupons(stackCoupons.filter((_,j)=>j!==i))}
                  style={{background:"transparent",border:`1px solid ${D.blue}55`,borderRadius:5,
                    padding:"3px 9px",fontSize:11,cursor:"pointer",color:D.blue,marginLeft:"auto"}}>✕</button>
              </div>
              );
            })}
            {/* 시나리오 선택 — 타입 규칙 기반 가능한 조합 (프리셋 적용 시 페이드) */}
            {scenarios.length>0&&(
              <div style={{position:"relative",marginTop:12,paddingTop:10,borderTop:`1px solid ${D.blue}33`}}>
                <div style={{fontSize:11,color:D.blue,fontWeight:700,marginBottom:6}}>
                  유효 시나리오 선택
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                  {scenarios.map((s,i)=>{
                    const active=scenarioIdx===i;
                    return <button key={i} type="button" onClick={()=>setScenarioIdx(i)}
                      style={{background:active?D.blue:"#fff",color:active?"#fff":D.blue,
                        border:`1px solid ${D.blue}`,borderRadius:5,padding:"3px 9px",
                        fontSize:11,cursor:"pointer",fontWeight:active?700:500,whiteSpace:"nowrap"}}>
                      Case {s.caseNum} · {s.label} → {s.eff}%
                    </button>;
                  })}
                </div>
                <div style={{fontSize:11,color:D.blue,marginTop:8,fontWeight:700}}>
                  선택 시나리오의 케이스 총합 할인율: {cpn}% (계산기·1·3·4 에 적용)
                </div>
                <div style={{fontSize:11,color:D.blue,marginTop:8,opacity:.85,lineHeight:1.5}}>
                  최종 할인율 = 1 − ∏(1−rate) of 선택 시나리오의 쿠폰들. 시나리오 미선택 시 기본 쿠폰 단독값 사용.
                </div>
                {presetActiveIdx>=0&&(
                  <>
                    <div aria-hidden="true" style={{position:"absolute",inset:0,
                      background:"rgba(255,255,255,0.78)",pointerEvents:"none",zIndex:1,borderRadius:6}}/>
                    <div style={{position:"absolute",inset:0,zIndex:2,
                      display:"flex",alignItems:"center",justifyContent:"center",gap:10,flexWrap:"wrap"}}>
                      <span style={{fontSize:11,color:D.black,fontWeight:700}}>시나리오 선택이 완료되었습니다</span>
                      <button type="button" onClick={resetPreset}
                        style={{background:"transparent",border:`1px solid ${D.black}`,borderRadius:5,
                          padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:700,color:D.black}}>
                        시나리오 취소
                      </button>
                    </div>
                  </>
                )}
              </div>
            )}
          </div>

          <details className="sec" style={sec} open>
            <summary style={summarySty}>1. 가격대별 분류 정의 <span className="chev" style={{color:D.textMeta}}>▾</span></summary>
            <div style={{padding:"4px 14px 14px"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr>
                  <th style={{...th,textAlign:"left",background:D.black,color:"#fff"}}>분류</th>
                  <th style={{...th,background:D.black,color:"#fff"}}>정가 범위</th>
                  <th style={{...th,background:D.black,color:"#fff"}}>표본</th>
                  <th style={{...th,background:D.black,color:"#fff"}}>제안 할인율</th>
                </tr></thead>
                <tbody>
                  {CALC_SLOTS.map(s=>(
                    <tr key={s.id}>
                      <td style={{...td,textAlign:"left",fontWeight:600,color:D.text}}>{s.name}</td>
                      <td style={{...td,fontSize:11,fontWeight:500}}>{s.range}</td>
                      <td style={{...td,fontSize:11}}>{wonFmt(s.n)}</td>
                      <td style={{...td,fontWeight:700,color:D.text}}>{s.disc}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{marginTop:10,fontSize:11,lineHeight:1.65,color:D.text}}>
                <b>할인율 올림 규칙</b> — 29CM 할인 규정상 10% 단위마다 판매수수료가 1%p 낮아지므로,
                기본 할인율의 <b>일의 자리가 6~9이면 다음 10% 단위로 올림</b>합니다 (예: 7%→10%, 16~19%→20%).
                올림된 할인율로 프런트 판매가·최종 노출가를 재계산합니다.
              </div>
            </div>
          </details>

          <details className="sec" style={sec}>
            <summary style={summarySty}>2. 역산 공식 <span className="chev" style={{color:D.textMeta}}>▾</span></summary>
            <div style={{padding:"4px 14px 14px"}}>
              <pre style={{fontFamily:"'SF Mono',Menlo,Consolas,monospace",fontSize:11,padding:12,
                background:D.surfaceAlt,borderRadius:6,lineHeight:1.7,whiteSpace:"pre",overflowX:"auto",color:D.text,margin:0}}>{`// 최종 = 기본 × 쿠폰 (곱연산)
기본 할인율       = 1 − (1 − P75/100) ÷ (1 − 쿠폰율/100)
  └ 일의 자리 6~9면 다음 10%로 올림 (예: 7→10%, 16~19→20%)
    (29CM 규정: 10% 단위마다 판매수수료 1%p 절감)
프런트 판매가 (I열) = 정가 × (1 − 기본 할인율/100)
최종 노출가       = 프런트 판매가 × (1 − 쿠폰율/100)

※ 판매가는 10원 단위 반올림 (29CM 규정)
※ 올림 규칙 적용 시 올림된 할인율로 가격을 재계산하므로
   최종 할인이 P75 목표보다 커질 수 있음
※ 쿠폰율이 P75보다 크면 기본 할인 0%, 정가 그대로 I열 입력`}</pre>
            </div>
          </details>

          <details className="sec" style={sec} open>
            <summary style={summarySty}>3. 단일 정가 시뮬레이션 <span className="chev" style={{color:D.textMeta}}>▾</span></summary>
            <div style={{padding:"4px 14px 14px"}}>
              <div style={{margin:"0 0 10px",fontSize:11,color:D.textSub,lineHeight:1.6}}>
                <b style={{color:D.black}}>기본 할인율을 직접 조정 가능</b>하며, 구간에 적합한 할인율을 다시 보려면 <b style={{color:D.black}}>재검색</b>하세요.
              </div>
              {/* 인벤토리 상품 검색 — 선택 시 정가/공급가 자동 채움 */}
              <div style={{position:"relative",marginBottom:10}}>
                <label style={{fontSize:11,color:D.textMeta,marginBottom:3,display:"block"}}>인벤토리에서 상품 검색 (선택 시 정가·공급가 자동 입력)</label>
                <input type="search" value={singleQuery}
                  onChange={e=>{setSingleQuery(e.target.value);}}
                  placeholder={invProducts.length===0?"인벤토리 데이터 로딩 중…":`상품명 검색 (${invProducts.length}개 인덱스됨)`}
                  style={{...inNum,width:"min(360px,100%)",textAlign:"left"}}/>
                {singleQuery.trim().length>0&&(()=>{
                  const kw=singleQuery.trim().toLowerCase();
                  const matches=invProducts.filter(p=>p.name.toLowerCase().includes(kw)).slice(0,12);
                  if(matches.length===0) return <div style={{marginTop:4,fontSize:11,color:D.textMeta}}>일치 상품 없음</div>;
                  return (
                    <div style={{position:"absolute",zIndex:10,top:"100%",left:0,right:0,marginTop:2,
                      maxHeight:280,overflowY:"auto",background:D.surface,border:`1px solid ${D.border}`,borderRadius:6,
                      boxShadow:"0 4px 16px rgba(0,0,0,0.12)"}}>
                      {matches.map((p,i)=>(
                        <div key={i} onClick={()=>{
                          setSingleSelected(p);
                          setListPrice(p.selling||0);
                          setSingleQuery("");
                        }}
                          style={{padding:"7px 10px",fontSize:11,cursor:"pointer",
                            borderBottom:i<matches.length-1?`1px solid ${D.border}`:"none",
                            display:"flex",justifyContent:"space-between",gap:8}}>
                          <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</span>
                          <span style={{color:D.textMeta,fontSize:11,whiteSpace:"nowrap"}}>
                            ₩{wonFmt(p.selling||0)} · 공급 ₩{wonFmt(p.supply||0)}
                          </span>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",padding:14,background:D.surfaceAlt,borderRadius:6}}>
                <label style={{fontSize:11,color:D.textSub,minWidth:40}}>정가</label>
                <input type="range" min="50000" max="300000" step="1000" value={Math.min(300000,Math.max(50000,listPrice||50000))}
                  onChange={e=>{setListPrice(parseInt(e.target.value)||0);setSingleSelected(null);}} style={{flex:1,minWidth:120,accentColor:D.black}}/>
                <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" step="1000" value={listPrice}
                  onChange={e=>{setListPrice(Math.max(0,parseInt(e.target.value)||0));setSingleSelected(null);}} style={inNum}/>
                <span style={{fontSize:11,color:D.textSub}}>원</span>
                {singleSelected&&(
                  <span style={{marginLeft:"auto",fontSize:11,color:D.text,fontWeight:600,
                    background:D.surfaceAlt,border:`1px solid ${D.borderMid}`,borderRadius:4,padding:"3px 8px"}}>
                    📌 {singleSelected.name} · 공급가 ₩{wonFmt(singleSelected.supply||0)}
                    <button onClick={()=>setSingleSelected(null)}
                      style={{marginLeft:6,background:"none",border:"none",color:D.textMeta,cursor:"pointer",padding:0,fontSize:11}}>✕</button>
                  </span>
                )}
              </div>
              {single&&(()=>{
                // 공급가 매칭 (선택된 상품 우선, 없으면 invMap 에서 자동 매칭)
                const supply=singleSelected?singleSelected.supply:supplyOf(singleSelected?.name||"",singleSelected?.code||"");
                const supplyIncVat=Math.round(supply*1.1);
                const m=computeMargin(listPrice,single.baseDisc,selectedScenario.items||[],supply);
                const costRatio=single.finalPrice>0&&supply>0?Math.round(supplyIncVat/single.finalPrice*1000)/10:0;
                const frontCut=Math.max(0,listPrice-single.basePrice);
                // 쿠폰 단계별 차감 + 부담 분리 (자사/채널)
                const couponSteps=[];
                let curPrice=single.basePrice;
                (selectedScenario.items||[]).forEach(c=>{
                  const tInfo=COUPON_TYPE_BY_KEY[c.type]||COUPON_TYPE_BY_KEY.product;
                  const before=curPrice;
                  const cut=Math.round(before*((c.rate||0)/100));
                  let chPart=0,slfPart=0,burdenDesc;
                  if(c.type==="share"){
                    chPart=Math.round(cut*((c.shareRate||0)/100));
                    slfPart=cut-chPart;
                    burdenDesc=`분담 자사 ${100-(c.shareRate||0)}% / 채널 ${c.shareRate||0}%`;
                  } else if(c.burden==="channel"){
                    chPart=cut; burdenDesc="채널 전액 부담 (29CM)";
                  } else {
                    slfPart=cut; burdenDesc="자사 전액 부담";
                  }
                  couponSteps.push({c,tInfo,before,cut,chPart,slfPart,burdenDesc});
                  curPrice-=cut;
                });
                const channelDetail=couponSteps.filter(s=>s.chPart>0);
                // 공통 스타일 — 라벨 / 금액(왼쪽 중앙) / 계산식(금액 옆) 3컬럼
                const gridT="minmax(0,170px) minmax(160px,210px) minmax(0,1fr)";
                const rowSty={display:"grid",gridTemplateColumns:gridT,alignItems:"baseline",gap:14,
                  padding:"6px 14px",borderTop:`1px solid ${D.border}`,color:D.text,fontSize:11,lineHeight:1.5};
                const labelCol={display:"flex",alignItems:"baseline",gap:6,flexWrap:"wrap",minWidth:0,fontSize:11};
                const amtCol={textAlign:"right",fontWeight:600,whiteSpace:"nowrap",fontSize:11};
                const totalAmt={...amtCol,fontWeight:800,fontSize:11};
                const calcCol={color:D.textMeta,minWidth:0,fontSize:11};
                const totalSty={...rowSty,background:D.surfaceAlt,
                  borderTop:`1.5px solid ${D.borderMid}`,padding:"8px 14px",fontSize:11};
                const groupGap={borderTop:`6px solid ${D.surfaceAlt}`,padding:0,height:0};
                return (
                <>
                  {singleSelected&&(
                    <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",margin:"4px 0 14px",padding:"14px 18px",
                      background:`${MUTE_BLUE}10`,border:`1px solid ${MUTE_BLUE}33`,borderRadius:9}}>
                      <span style={{fontSize:11,fontWeight:700,color:MUTE_BLUE,letterSpacing:".04em",whiteSpace:"nowrap"}}>📊 분석 상품</span>
                      <span style={{fontSize:19,fontWeight:800,color:D.black,letterSpacing:"-0.3px",lineHeight:1.25,
                        flex:"1 1 auto",minWidth:0,wordBreak:"break-word"}}>{singleSelected.name}</span>
                      <span style={{fontSize:12,color:D.textSub,whiteSpace:"nowrap"}}>정가 ₩{wonFmt(listPrice)} · 공급가 ₩{wonFmt(singleSelected.supply||0)}</span>
                    </div>
                  )}
                  {/* 기본할인율 조정 컨트롤 — 5% 단위 다이얼 */}
                  <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",margin:"12px 0 10px"}}>
                    <span style={{marginLeft:"auto",fontSize:11,color:D.textSub,display:"flex",alignItems:"center",gap:6}}>
                      기본 할인율
                      <select
                        value={Math.round((singleManualBase!=null?parseFloat(singleManualBase):single.baseDisc)/5)*5}
                        onChange={e=>setSingleManualBase(e.target.value)}
                        title="5% 단위 다이얼"
                        style={{padding:"4px 8px",fontSize:11,textAlign:"right",
                          border:`1px solid ${singleManualBase!=null?D.blue:D.border}`,
                          borderRadius:4,background:singleManualBase!=null?"#eef3ff":D.surface,
                          color:singleManualBase!=null?D.blue:D.text,
                          fontWeight:singleManualBase!=null?700:400,fontFamily:"inherit"}}>
                        {Array.from({length:21},(_,j)=>j*5).map(v=>(
                          <option key={v} value={v}>{v}%</option>
                        ))}
                      </select>
                      {singleManualBase!=null&&(
                        <button onClick={()=>setSingleManualBase(null)} title="0%로 복귀"
                          style={{background:"transparent",border:"none",cursor:"pointer",fontSize:11,color:D.textMeta,padding:"0 3px"}}>↻</button>
                      )}
                    </span>
                  </div>
                  {/* 단일 정가 시뮬레이션 */}
                  <div style={{border:`1px solid ${D.borderMid}`,borderRadius:6,background:D.surface,overflow:"hidden"}}>
                    {/* 메타 헤더 */}
                    <div style={{padding:"10px 12px",borderBottom:`1px solid ${D.borderMid}`,
                      display:"flex",flexWrap:"wrap",alignItems:"center",gap:8,fontSize:11,color:D.textSub}}>
                      {selectedScenario.caseNum&&(
                        <span style={{background:D.black,color:"#fff",fontSize:11,padding:"1px 6px",borderRadius:3}}>Case {selectedScenario.caseNum}</span>
                      )}
                      <span style={{color:D.text,fontWeight:600}}>{selectedScenario.label||`기본 쿠폰 ${cpn}%`}</span>
                      <span style={{color:D.textMeta}}>케이스 총합 할인율 {cpn}%</span>
                      <span style={{color:D.blue,fontWeight:700,background:"#eef3ff",border:`1px solid ${D.blue}`,borderRadius:3,padding:"1px 6px"}}>
                        기본 할인율 {single.baseDisc}%
                      </span>
                      {singleSelected&&(
                        <>
                          <span style={{color:D.textMeta}}>·</span>
                          <span style={{color:D.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:240}} title={singleSelected.name}>{singleSelected.name}</span>
                        </>
                      )}
                    </div>

                    {(()=>{
                      // 항목 번호 부여 (둥근 숫자) — 계산식에서 참조용
                      const C=["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮"];
                      let p=0; const next=()=>C[p++];
                      const N={};
                      N.list=next();
                      if(single.baseDisc>0) N.baseDisc=next();
                      N.basePrice=next();
                      N.coupons=couponSteps.map(()=>next());
                      N.finalPrice=next();
                      N.fee=next();
                      N.refund=next();
                      N.netSettle=next();
                      if(supply>0){ N.supply=next(); N.margin=next(); }
                      const couponSumExpr=N.coupons.length>0
                        ? `${N.basePrice} − ${N.coupons.join(" − ")}`
                        : N.basePrice;
                      return (
                      <>
                      {/* 가격 차감 흐름 */}
                      {(()=>{
                        const listMu=supplyIncVat>0?Math.round(listPrice/supplyIncVat*100)/100:null;
                        return (
                          <div style={{...rowSty,borderTop:"none"}}>
                            <span style={labelCol}><span>{N.list} 정상 가격</span></span>
                            <span style={amtCol}>₩{wonFmt(listPrice)}{listMu!=null&&<span style={{color:listMu>3?D.green:D.red,fontWeight:600}}> (×{listMu.toFixed(2)})</span>}</span>
                            <span style={calcCol}>{listMu!=null?`정가 ÷ 원가(₩${wonFmt(supplyIncVat)}) = ×${listMu.toFixed(2)} 마크업`:""}</span>
                          </div>
                        );
                      })()}
                      {single.baseDisc>0&&(
                        <div style={rowSty}>
                          <span style={labelCol}><span>{N.baseDisc}-1 기본 할인율 <span style={{color:D.textMeta,fontWeight:400}}>(프런트 할인율)</span></span></span>
                          <span style={{...amtCol,color:singleManualBase!=null?D.blue:D.text}}>{single.baseDisc}%{singleManualBase!=null&&<span style={{color:D.blue,fontWeight:400}}> · 수동</span>}</span>
                          <span style={calcCol}>정상 가격에 적용되는 프런트(기본) 할인율입니다{singleManualBase!=null?" · 직접 변경한 값":""}.</span>
                        </div>
                      )}
                      {single.baseDisc>0&&(
                        <div style={rowSty}>
                          <span style={labelCol}><span>{N.baseDisc}-2 기본 할인 금액</span></span>
                          <span style={{...amtCol,color:D.red}}>−₩{wonFmt(frontCut)}</span>
                          <span style={calcCol}>정상 가격 × {(single.baseDisc/100).toFixed(single.baseDisc%1===0?2:3)} · 자사가 부담하는 프런트 할인액입니다.</span>
                        </div>
                      )}
                      <div style={totalSty}>
                        <span style={labelCol}><span>{N.basePrice} 프런트 판매가</span></span>
                        <span style={totalAmt}>₩{wonFmt(single.basePrice)}</span>
                        <span style={calcCol}>정상 가격에서 기본 할인을 차감한 노출 판매가입니다.</span>
                      </div>
                      {couponSteps.length>0?couponSteps.map((s,i)=>{
                        const burdenSentence = s.c.type==="share"
                          ? `자사 ${100-(s.c.shareRate||0)}% / 채널 ${s.c.shareRate||0}%로 분담합니다. 채널분 ₩${wonFmt(s.chPart)}은 정산 시 채널 보전금액으로 이어집니다.`
                          : (s.c.burden==="channel"
                              ? `29CM 채널이 전액 부담합니다. 정산 시 채널 보전금액으로 이어집니다.`
                              : `자사가 전액 부담하는 할인액입니다.`);
                        const refLabel=i===0?"프런트 판매가":`직전 단계 판매가`;
                        return (
                          <div key={`cs${i}`} style={rowSty}>
                            <span style={labelCol}>
                              <span>{N.coupons[i]} {s.tInfo.label} 금액</span>
                            </span>
                            <span style={{...amtCol,color:D.red}}>−₩{wonFmt(s.cut)} <span style={{color:D.textMeta,fontWeight:400}}>({s.c.rate}%)</span></span>
                            <span style={calcCol}>{refLabel} × {(s.c.rate/100).toFixed(s.c.rate%1===0?2:3)} · {burdenSentence}</span>
                          </div>
                        );
                      }):(
                        <div style={rowSty}>
                          <span style={labelCol}><span style={{color:D.textMeta}}>쿠폰 미적용</span></span>
                          <span style={{...amtCol,color:D.textMeta}}>—</span>
                          <span style={calcCol}>적용된 쿠폰이 없습니다.</span>
                        </div>
                      )}
                      {(()=>{
                        const finalMu=supplyIncVat>0?Math.round(single.finalPrice/supplyIncVat*100)/100:null;
                        return (
                          <div style={totalSty}>
                            <span style={labelCol}><span>{N.finalPrice} 실제 판매 가격</span></span>
                            <span style={totalAmt}>₩{wonFmt(single.finalPrice)} <span style={{color:D.red,fontWeight:600}}>(정상가 대비 −{single.finalDisc}%)</span>{finalMu!=null&&<span style={{color:finalMu>3?D.green:D.red,fontWeight:600}}> (×{finalMu.toFixed(2)})</span>}</span>
                            <span style={calcCol}>쿠폰까지 적용된 최종 노출가로, 고객이 실제 결제하는 금액입니다 (케이스 총합 할인율 {single.finalDisc}%){finalMu!=null?` · 실판매가 ÷ 원가 = ×${finalMu.toFixed(2)} 마크업`:""}.</span>
                          </div>
                        );
                      })()}

                      {/* 정산 흐름 */}
                      <div style={groupGap}></div>
                      <div style={rowSty}>
                        <span style={labelCol}><span>{N.fee} 채널 수수료</span></span>
                        <span style={{...amtCol,color:D.red}}>−₩{wonFmt(m.fee)} <span style={{color:D.textMeta,fontWeight:400}}>({m.feeRate}%)</span></span>
                        <span style={calcCol}>{N.basePrice} 프런트 판매가 × {(m.feeRate/100).toFixed(2)} · 기본 28%에서 기본 할인율 10%당 1%p씩 차감되어 {m.feeRate}%가 부과됩니다.</span>
                      </div>
                      <div style={rowSty}>
                        <span style={labelCol}><span>{N.refund} 쿠폰 채널 보전 금액</span></span>
                        <span style={{...amtCol,color:m.channelBurden>0?D.blue:D.textMeta}}>+₩{wonFmt(m.channelBurden)}</span>
                        <span style={calcCol}>
                          {channelDetail.length>0
                            ? `${channelDetail.map((s)=>`${s.tInfo.label} ${s.c.type==="share"?`(분담 채널 ${s.c.shareRate}%)`:"(채널 전액 부담)"} ₩${wonFmt(s.chPart)}`).join(" / ")} · 채널이 부담한 쿠폰 금액을 자사에 보전합니다.`
                            : "채널이 부담하는 쿠폰이 없어 보전 금액이 발생하지 않습니다."}
                        </span>
                      </div>
                      <div style={totalSty}>
                        <span style={labelCol}><span>{N.netSettle} 정산 금액</span></span>
                        <span style={{...totalAmt,color:D.black}}>₩{wonFmt(m.net)}{supplyIncVat>0&&<span style={{color:m.markup>3?D.green:D.red,fontWeight:600}}> (×{(m.markup||0).toFixed(2)})</span>}</span>
                        <span style={calcCol}>실 판매액에서 수수료를 빼고 채널 보전을 더한, 자사가 실제로 수령하는 정산액입니다{supplyIncVat>0?` · 정산액 ÷ 원가 = ×${(m.markup||0).toFixed(2)} 마크업`:""}.</span>
                      </div>

                      {/* 마진 흐름 */}
                      {supply>0?(
                        <>
                          <div style={groupGap}></div>
                          <div style={rowSty}>
                            <span style={labelCol}><span>{N.supply} 공급가(부가세 합)</span></span>
                            <span style={{...amtCol,color:D.red}}>−₩{wonFmt(supplyIncVat)}</span>
                            <span style={calcCol}>인벤토리 업로드 파일에서 추출한 공급가액 ₩{wonFmt(supply)} × 1.1 · 부가세 10%를 포함한 실 원가입니다.</span>
                          </div>
                          <div style={totalSty}>
                            <span style={labelCol}><span>{N.margin} 마진 금액</span></span>
                            <span style={{...totalAmt,color:m.margin>=0?D.green:D.red}}>₩{wonFmt(m.margin)}</span>
                            <span style={calcCol}>정산 금액에서 공급가를 차감한, 자사가 남기는 이익입니다.</span>
                          </div>
                          <div style={rowSty}>
                            <span style={labelCol}><span>마크업</span></span>
                            <span style={{...amtCol,color:m.markup>3?D.green:D.red}}>×{(m.markup||0).toFixed(2)}</span>
                            <span style={calcCol}>실수령액(정산액) ÷ 원가 — 자사 실수령이 원가의 몇 배인지 보여주는 마크업입니다 (×3 이하 적색).</span>
                          </div>
                          {(()=>{
                            // 최소 마진 기본 세일율 — 마진이 0이 되는 시점의 기본 할인율(0.1% 정밀도)
                            let lastValid=null;
                            for(let bd=0; bd<=80.05; bd=Math.round((bd+0.1)*10)/10){
                              const mm=computeMargin(listPrice,bd,selectedScenario.items||[],supply);
                              if(mm.margin>=0) lastValid=bd;
                            }
                            return (
                              <div style={rowSty}>
                                <span style={labelCol}><span>최소 마진 기본 세일율 {single.baseDisc>0&&<span style={{color:D.textMeta,fontWeight:400}}>(비교대상은 {N.baseDisc}-1)</span>}</span></span>
                                <span style={{...amtCol,color:lastValid!=null?D.text:D.textMeta}}>
                                  {lastValid!=null?`${lastValid}%`:"적용 불가"}
                                </span>
                                <span style={calcCol}>
                                  {(()=>{
                                    const scenarioTag=selectedScenario.caseNum?`Case ${selectedScenario.caseNum} (${selectedScenario.label})`:`기본 쿠폰 ${cpn}%`;
                                    return lastValid!=null
                                      ? `${scenarioTag} 시나리오 기준 역산 — 마진이 0원이 되는 시점의 기본 할인율로, 이 비율을 초과하면 적자가 발생합니다 (현재 기본 할인율 ${single.baseDisc}%).`
                                      : `${scenarioTag} 시나리오 기준 — 어떤 기본 할인율로도 마진이 0 이상이 되지 않습니다.`;
                                  })()}
                                </span>
                              </div>
                            );
                          })()}
                        </>
                      ):(
                        <div style={{padding:"10px 14px",color:D.textMeta,borderTop:`6px solid ${D.surfaceAlt}`}}>
                          공급가 미연동 — 인벤토리 매칭 시 마진 금액 / 마크업 자동 계산
                        </div>
                      )}
                      </>
                      );
                    })()}
                  </div>
                </>
                );
              })()}
            </div>
          </details>

          <details className="sec" style={sec}>
            <summary style={summarySty}>
              공급가 직접 업로드 (xlsx) <span className="chev" style={{color:D.textMeta}}>▾</span>
            </summary>
            <div style={{padding:"4px 14px 14px"}}>
              <div style={{fontSize:11,color:D.textSub,lineHeight:1.6,marginBottom:8}}>
                <b style={{color:D.black}}>상품명</b>·<b style={{color:D.black}}>공급가</b> 컬럼이 있는 xlsx 를 업로드하면 인벤토리에 없는 상품도 공급가가 매칭됩니다.
                Supabase 에 저장되어 다음 세션·다른 기기에서도 그대로 유지되며, <b>인벤토리 공급가보다 우선 적용</b>됩니다.
              </div>
              <div onClick={()=>overrideFileRef.current?.click()}
                onDragOver={e=>{e.preventDefault();setOverrideDrag(true);}}
                onDragLeave={()=>setOverrideDrag(false)}
                onDrop={e=>{e.preventDefault();setOverrideDrag(false);if(e.dataTransfer.files[0])handleSupplyUpload(e.dataTransfer.files[0]);}}
                style={{border:`1px dashed ${overrideDrag?D.blue:D.borderMid}`,borderRadius:6,padding:18,textAlign:"center",
                  cursor:"pointer",background:overrideDrag?"#eef3ff":D.surface}}>
                <div style={{fontSize:11,color:D.text,marginBottom:3}}>xlsx 파일 끌어다 놓거나 클릭해 업로드</div>
                <div style={{fontSize:11,color:D.textMeta}}>첫 번째 시트의 헤더에서 '상품명' / '공급가' 컬럼 자동 인식</div>
                <input ref={overrideFileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                  onChange={e=>{if(e.target.files[0])handleSupplyUpload(e.target.files[0]); e.target.value="";}}/>
              </div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:8,gap:8,flexWrap:"wrap"}}>
                <span style={{fontSize:11,color:D.textSub}}>{overrideStatus||(Object.keys(overrideMap).length>0?`업로드 공급가 ${Object.keys(overrideMap).length}건 적용 중`:"업로드된 공급가 없음")}</span>
                {Object.keys(overrideMap).length>0&&(
                  <button onClick={clearOverride}
                    style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                      padding:"3px 10px",fontSize:11,cursor:"pointer",color:D.textSub,fontWeight:600}}>
                    전체 초기화
                  </button>
                )}
              </div>
            </div>
          </details>

          <div style={sec}>
            <div style={{padding:"11px 14px",fontSize:11,fontWeight:700,borderBottom:`1px solid ${D.border}`,color:D.black,
              display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
              <span>4. 29CM 일괄할인 양식 업로드</span>
              <button onClick={loadExampleFile} title="마지막으로 업로드한 파일을 예시로 불러옵니다"
                style={{background:"transparent",border:`1px solid ${D.borderMid}`,borderRadius:5,
                  padding:"3px 9px",fontSize:11,cursor:"pointer",color:D.textSub,fontWeight:600}}>
                &lt;예시파일&gt;
              </button>
            </div>
            <div style={{padding:14}}>
              <div onClick={()=>fileRef.current?.click()}
                onDragOver={e=>{e.preventDefault();setDragOver(true);}}
                onDragLeave={()=>setDragOver(false)}
                onDrop={e=>{e.preventDefault();setDragOver(false);if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);}}
                style={{border:`1px dashed ${dragOver?D.blue:D.borderMid}`,borderRadius:6,padding:22,textAlign:"center",
                  cursor:"pointer",background:dragOver?"#eef3ff":D.surface}}>
                <div style={{margin:"0 0 4px",fontSize:11,color:D.text}}>29CM 일괄할인 v2 양식을 끌어다 놓거나 클릭해서 업로드</div>
                <div style={{fontSize:11,color:D.textMeta}}>E열 정상가 기준 분류 → 쿠폰율 역산한 프런트 판매가를 I열에 입력</div>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
                  onChange={e=>{if(e.target.files[0])handleFile(e.target.files[0]);e.target.value="";}}/>
              </div>
              {showResults&&(
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"12px 0 8px",flexWrap:"wrap",gap:8}}>
                    <span style={{fontSize:11,color:D.textMeta}}>{summary}{processed&&processed.length>DISP_CAP?` · 처음 ${DISP_CAP}행 표시(다운로드는 전체)`:""}</span>
                    {processed&&processed.length>0&&(
                      <button onClick={download} style={{background:D.black,color:"#fff",border:"none",padding:"9px 18px",
                        fontSize:11,borderRadius:6,cursor:"pointer",fontWeight:600}}>I열 갱신본 다운로드</button>
                    )}
                  </div>
                  {processed&&processed.length>0&&(
                    <div style={{margin:"0 0 10px",fontSize:11,color:D.black,fontWeight:700,lineHeight:1.55}}>
                      적용 시나리오
                      {selectedScenario.caseNum&&<> · Case {selectedScenario.caseNum}</>}
                      {' · '}{selectedScenario.label||`기본 쿠폰 ${cpn}%`}
                      {(selectedScenario.items||[]).length>0&&(
                        <> ({(selectedScenario.items||[]).map(c=>{
                          const tInfo=COUPON_TYPE_BY_KEY[c.type];
                          const burden=c.type==="share"
                            ?`분담 자사${100-(c.shareRate||0)}:채널${c.shareRate||0}`
                            :(c.burden==="channel"?"채널부담":"자사부담");
                          return `${tInfo.label} ${c.rate}% · ${burden}`;
                        }).join(" × ")})</>
                      )}
                      {' · '}케이스 총합 할인율 {cpn}%
                    </div>
                  )}
                  {visibleProcessed.length>0&&(()=>{
                    const matched=visibleProcessed.filter(r=>(r.supply||0)>0);
                    const avg=matched.length>0?matched.reduce((s,r)=>s+(r.markup||0),0)/matched.length:null;
                    const avgBase=visibleProcessed.length>0?visibleProcessed.reduce((s,r)=>s+(r.baseDisc||0),0)/visibleProcessed.length:0;
                    const avgFinal=visibleProcessed.length>0?visibleProcessed.reduce((s,r)=>s+(r.finalDisc||0),0)/visibleProcessed.length:0;
                    const anyManual=visibleProcessed.some(r=>r.manualBase!=null);
                    return (
                      <div style={{margin:"0",padding:"10px 14px",background:D.surfaceAlt,borderRadius:6,
                        display:"flex",alignItems:"center",gap:14,flexWrap:"wrap",fontSize:11,color:D.text,
                        position:"sticky",top:61,zIndex:4,border:`1px solid ${D.borderMid}`}}>
                        <span style={{display:"inline-flex",alignItems:"baseline",gap:6}}>
                          <span style={{fontWeight:700,color:D.black}}>평균 마크업</span>
                          {avg==null
                            ?<span style={{color:D.textMeta}}>공급가 매칭 행 없음</span>
                            :<span style={{fontSize:15,fontWeight:800,color:avg>3?D.green:D.red}}>×{avg.toFixed(2)}</span>}
                        </span>
                        {avg!=null&&<span style={{color:D.textMeta}}>· 매칭 {matched.length.toLocaleString()}/{visibleProcessed.length.toLocaleString()}건</span>}
                        <span style={{color:D.textMeta}}>· 평균 기본 <b style={{color:D.text}}>{(Math.round(avgBase*10)/10)}%</b> · 평균 최종 <b style={{color:D.text}}>{(Math.round(avgFinal*10)/10)}%</b></span>
                        {checkedRows.size>0&&(
                          <button onClick={removeCheckedBulk}
                            style={{background:D.red,color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                            🗑 체크 {checkedRows.size}개 일괄 삭제
                          </button>
                        )}
                        {removedRows.size>0&&(
                          <span style={{display:"inline-flex",alignItems:"center",gap:6,fontSize:11,color:D.textSub,
                            border:`1px dashed ${D.borderMid}`,borderRadius:5,padding:"2px 8px"}}>
                            🗑 {removedRows.size}개 제거 <span style={{color:D.textMeta}}>(다운로드/재업로드 시 복원)</span>
                            <button onClick={restoreAllBulk}
                              style={{background:"transparent",border:"none",cursor:"pointer",color:D.blue,fontSize:11,fontWeight:600,padding:"0 4px"}}>↻ 복원</button>
                          </span>
                        )}
                        <span style={{marginLeft:"auto",display:"inline-flex",alignItems:"center",gap:5}}>
                          <span style={{fontSize:11,color:D.textMeta,fontWeight:600}}>기본 세일율 일괄:</span>
                          {[10,15,20,25,30,40].map(v=>(
                            <button key={v} onClick={()=>applyAllBaseDisc(v)}
                              style={{background:"transparent",border:`1px solid ${D.borderMid}`,borderRadius:4,
                                padding:"3px 8px",fontSize:11,cursor:"pointer",color:D.textSub,fontWeight:600}}>
                              {v}%
                            </button>
                          ))}
                          <input type="number" min="0" max="100" step="0.1"
                            value={bulkBaseDisc} onChange={e=>setBulkBaseDisc(e.target.value)}
                            onKeyDown={e=>{if(e.key==="Enter"&&bulkBaseDisc!==""){applyAllBaseDisc(bulkBaseDisc);setBulkBaseDisc("");}}}
                            placeholder="직접" onWheel={e=>e.currentTarget.blur()}
                            style={{width:44,padding:"3px 5px",fontSize:11,textAlign:"right",
                              border:`1px solid ${D.borderMid}`,borderRadius:4,background:D.surface,color:D.text,fontFamily:"inherit"}}/>
                          <button onClick={()=>{if(bulkBaseDisc!==""){applyAllBaseDisc(bulkBaseDisc);setBulkBaseDisc("");}}}
                            disabled={bulkBaseDisc===""}
                            style={{background:bulkBaseDisc!==""?D.black:D.surface,color:bulkBaseDisc!==""?"#fff":D.textMeta,
                              border:`1px solid ${bulkBaseDisc!==""?D.black:D.border}`,borderRadius:4,
                              padding:"3px 8px",fontSize:11,cursor:bulkBaseDisc!==""?"pointer":"default",fontWeight:700}}>
                            적용
                          </button>
                          {anyManual&&(
                            <button onClick={resetAllBaseDisc} title="모든 행을 역산값으로 복귀"
                              style={{background:"transparent",border:`1px solid ${D.borderMid}`,borderRadius:4,
                                padding:"3px 8px",fontSize:11,cursor:"pointer",color:D.textMeta,fontWeight:600}}>
                              ↻ 복귀
                            </button>
                          )}
                        </span>
                      </div>
                    );
                  })()}
                  {processed&&processed.length>0&&(
                    <div style={{border:`1px solid ${D.border}`,borderRadius:6,marginTop:8}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                        <thead><tr>
                          {["상품명","정가 (E열)","쿠폰율","기본 할인율","프런트 판매가 (I열)","최종 노출가","최종 할인율(쿠폰 포함)","자사부담","수수료","채널보전","자사 정산","공급가 (세포)","마진","마크업"].map((h,i)=>(
                            <th key={i} style={{padding:"7px 8px",borderBottom:`1px solid ${D.borderMid}`,
                              textAlign:i===0?"left":"right",fontWeight:600,color:D.textSub,background:D.surfaceAlt,whiteSpace:"nowrap",
                              position:"sticky",top:108,zIndex:3,boxShadow:`0 1px 0 ${D.borderMid}`}}>
                              {i===0?(
                                <span style={{display:"inline-flex",alignItems:"center",gap:4}}>
                                  <input type="checkbox" title="현재 표시된 행 전체 선택/해제"
                                    checked={shown.length>0&&shown.every(s=>checkedRows.has(s.row))}
                                    onChange={()=>{
                                      const allChecked=shown.length>0&&shown.every(s=>checkedRows.has(s.row));
                                      setCheckedRows(prev=>{
                                        const next=new Set(prev);
                                        if(allChecked) shown.forEach(s=>next.delete(s.row));
                                        else shown.forEach(s=>next.add(s.row));
                                        return next;
                                      });
                                    }}
                                    style={{cursor:"pointer"}}/>
                                  {h}
                                </span>
                              ):h}
                            </th>
                          ))}
                        </tr></thead>
                        <tbody>
                          {shown.map((r,i)=>(
                            <tr key={i}
                              onMouseEnter={()=>enterDragSelectBulk(r.row)}
                              style={{background:checkedRows.has(r.row)?`${D.red}0a`:"transparent"}}>
                              <td title={r.name}
                                onMouseDown={e=>{
                                  if(e.target.tagName==="INPUT"||e.target.tagName==="BUTTON"||e.target.closest("button")) return;
                                  e.preventDefault();
                                  startDragSelectBulk(r.row,checkedRows.has(r.row));
                                }}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"left",
                                  minWidth:160,maxWidth:280,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",cursor:"pointer",userSelect:"none"}}>
                                <input type="checkbox" checked={checkedRows.has(r.row)} onChange={()=>{}}
                                  onMouseDown={e=>{e.preventDefault();startDragSelectBulk(r.row,checkedRows.has(r.row));}}
                                  title="클릭 또는 드래그로 다중 선택"
                                  style={{marginRight:4,cursor:"pointer",verticalAlign:"middle"}}/>
                                <button onClick={()=>removeBulkRow(r.row)} title="이 행을 표에서 임시로 제거 (다운로드/재업로드 시 복원)"
                                  style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:11,padding:"0 4px",marginRight:2,verticalAlign:"middle"}}>✕</button>
                                {r.name}
                              </td>
                              {(()=>{
                                const sv=r.supplyIncVat||Math.round((r.supply||0)*1.1);
                                const listMu=sv>0?Math.round(r.list/sv*100)/100:null;
                                return (
                                  <td title={`정가 ₩${wonFmt(r.list)}${listMu!=null?` · 마크업 ×${listMu.toFixed(2)} (정가 ÷ 원가)`:""}`}
                                    style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",whiteSpace:"nowrap"}}>
                                    ₩{wonFmt(r.list)}
                                    {listMu!=null&&<span style={{marginLeft:4,fontSize:10,fontWeight:700,color:listMu>3?D.green:D.red}}>×{listMu.toFixed(2)}</span>}
                                  </td>
                                );
                              })()}
                              <td style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right"}}>{cpn}%</td>
                              <td style={{padding:"4px 6px",borderBottom:`1px solid ${D.border}`,textAlign:"right",whiteSpace:"nowrap"}}>
                                <select value={Math.round((r.baseDisc||0)/5)*5}
                                  onChange={e=>updateBaseDisc(r.row,e.target.value)}
                                  title={r.manualBase!=null?"수동 변경됨 — ↻ 로 0%로 복귀":"5% 단위 다이얼"}
                                  style={{padding:"3px 6px",fontSize:11,textAlign:"right",
                                    border:`1px solid ${r.manualBase!=null?D.blue:D.border}`,
                                    borderRadius:4,background:r.manualBase!=null?"#eef3ff":D.surface,
                                    color:r.manualBase!=null?D.blue:D.text,
                                    fontWeight:r.manualBase!=null?700:400,fontFamily:"inherit"}}>
                                  {Array.from({length:21},(_,j)=>j*5).map(v=>(
                                    <option key={v} value={v}>{v}%</option>
                                  ))}
                                </select>
                                {r.manualBase!=null&&(
                                  <button onClick={()=>resetBaseDisc(r.row)} title="0%로 복귀"
                                    style={{background:"transparent",border:"none",cursor:"pointer",fontSize:11,color:D.textMeta,padding:"0 3px"}}>↻</button>
                                )}
                              </td>
                              {(()=>{
                                const sv=r.supplyIncVat||Math.round((r.supply||0)*1.1);
                                const baseMu=sv>0?Math.round(r.basePrice/sv*100)/100:null;
                                const finalMu=sv>0?Math.round(r.finalPrice/sv*100)/100:null;
                                return (<>
                                  <td title={`프런트 판매가 = 정상가 ₩${wonFmt(r.list)} × (1 − ${r.baseDisc}%) = ₩${wonFmt(r.basePrice)} (10원 단위 반올림) | I열에 입력되는 노출 판매가${baseMu!=null?` · 마크업 ×${baseMu.toFixed(2)} (기본가 ÷ 원가)`:""}`}
                                    style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",background:"#eef3ff",color:D.blue,fontWeight:600,whiteSpace:"nowrap"}}>
                                    ₩{wonFmt(r.basePrice)}
                                    {baseMu!=null&&<span style={{marginLeft:4,fontSize:10,fontWeight:700,color:baseMu>3?D.green:D.red}}>×{baseMu.toFixed(2)}</span>}
                                  </td>
                                  <td title={`최종 노출가 = 프런트 판매가 ₩${wonFmt(r.basePrice)} × (1 − ${cpn}%) = ₩${wonFmt(r.finalPrice)} | 쿠폰 적용 후 고객 결제 금액${finalMu!=null?` · 마크업 ×${finalMu.toFixed(2)} (실판매가 ÷ 원가)`:""}`}
                                    style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",color:D.textSub,whiteSpace:"nowrap"}}>
                                    ₩{wonFmt(r.finalPrice)}
                                    {finalMu!=null&&<span style={{marginLeft:4,fontSize:10,fontWeight:700,color:finalMu>3?D.green:D.red}}>×{finalMu.toFixed(2)}</span>}
                                  </td>
                                </>);
                              })()}
                              <td title={`최종 할인율 = 1 − (실 판매액 ₩${wonFmt(r.finalPrice)} ÷ 정상가 ₩${wonFmt(r.list)}) = ${r.finalDisc}% (정상가 대비 총 할인)`}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",
                                  background:"#eef3ff",color:D.blue,fontWeight:700,whiteSpace:"nowrap"}}>{r.finalDisc}%</td>
                              <td title={`정상가 ₩${wonFmt(r.list)} × 기본 할인율 ${r.baseDisc}% = ₩${wonFmt(Math.round((r.list||0)*((r.baseDisc||0)/100)))} 프런트 할인분 + 자사부담 쿠폰 차감액 합산 (자사 부담 → 마진 차감 대상)`}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",color:(r.selfBurden||0)>0?D.red:D.textMeta,whiteSpace:"nowrap"}}>
                                {(r.selfBurden||0)>0?`−₩${wonFmt(r.selfBurden)}`:"—"}
                              </td>
                              <td title={`프런트 판매가 ₩${wonFmt(r.basePrice||0)} × 수수료율 ${r.feeRate}% = ₩${wonFmt(r.fee||0)} | 수수료율 = max(0, 28% − floor(기본 할인율 ${r.baseDisc}% / 10)%) = ${r.feeRate}%`}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",color:D.red,whiteSpace:"nowrap"}}>
                                −₩{wonFmt(r.fee||0)} <span style={{fontSize:11,color:D.textMeta}}>({r.feeRate}%)</span>
                              </td>
                              <td title={(r.channelBurden||0)>0?`채널부담 쿠폰 차감액 + 분담 쿠폰의 채널분 합산 = ₩${wonFmt(r.channelBurden)} (정산 시 자사에 +가산되어 보전)`:"채널이 부담한 쿠폰이 없어 보전 금액 없음"}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",color:(r.channelBurden||0)>0?D.blue:D.textMeta,whiteSpace:"nowrap"}}>
                                {(r.channelBurden||0)>0?`+₩${wonFmt(r.channelBurden)}`:"—"}
                              </td>
                              <td title={`자사 정산 = 실 판매액 ₩${wonFmt(r.finalPrice||0)} − 채널 수수료 ₩${wonFmt(r.fee||0)} + 채널 보전 ₩${wonFmt(r.channelBurden||0)} = ₩${wonFmt(r.net||0)} (자사 수령액)${r.supply>0?` · 마크업 ×${(r.markup||0).toFixed(2)} (정산액 ÷ 원가)`:""}`}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",fontWeight:600,whiteSpace:"nowrap"}}>
                                ₩{wonFmt(r.net||0)}
                                {r.supply>0&&<span style={{marginLeft:4,fontSize:10,fontWeight:700,color:(r.markup||0)>3?D.green:D.red}}>×{(r.markup||0).toFixed(2)}</span>}
                              </td>
                              <td title={r.supply>0?`인벤토리 공급가액 ₩${wonFmt(r.supply)} × 1.1 (부가세 10% 포함) = ₩${wonFmt(r.supplyIncVat||Math.round(r.supply*1.1))} (자사 실 원가)`:"인벤토리 매칭 없음 — 공급가 자동 입력 불가"}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",color:r.supply>0?D.text:D.textMeta,whiteSpace:"nowrap"}}>
                                {r.supply>0?`₩${wonFmt(r.supplyIncVat||Math.round(r.supply*1.1))}`:"—"}
                              </td>
                              <td title={r.supply>0?`마진 = 자사 정산 ₩${wonFmt(r.net||0)} − 공급가(세포) ₩${wonFmt(r.supplyIncVat||Math.round((r.supply||0)*1.1))} = ₩${wonFmt(r.margin||0)} | 마크업 ×${(r.markup||0).toFixed(2)} (실수령 ÷ 원가)`:""}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",
                                  color:r.supply>0?((r.margin||0)>=0?D.text:D.red):D.textMeta,whiteSpace:"nowrap",fontWeight:600}}>
                                {r.supply>0?`₩${wonFmt(r.margin||0)}`:"—"}
                              </td>
                              <td title={r.supply>0?`마크업 = 실수령액 ₩${wonFmt(r.net||0)} ÷ 원가 ₩${wonFmt(r.supplyIncVat||Math.round((r.supply||0)*1.1))} = ×${(r.markup||0).toFixed(2)} (정산액이 원가의 몇 배인지 · ×3 이하 적색)`:""}
                                style={{padding:"7px 8px",borderBottom:`1px solid ${D.border}`,textAlign:"right",
                                color:r.supply>0?((r.markup||0)>3?D.green:D.red):D.textMeta,
                                fontWeight:700,whiteSpace:"nowrap"}}>
                                {r.supply>0?`×${(r.markup||0).toFixed(2)}`:"—"}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                  {processed&&processed.length>0&&(()=>{
                    // 기본 할인율 % (5% 단위 버킷) 별 그룹 — 묶인 상품 리스트 보존
                    const groups={};
                    visibleProcessed.forEach(r=>{
                      const k=Math.round((r.baseDisc||0)/5)*5;
                      if(!groups[k]) groups[k]={baseDisc:k,products:[],count:0,matched:0,mSum:0,frSum:0};
                      groups[k].count++;
                      groups[k].frSum+=(r.finalDisc||0);
                      groups[k].products.push(r);
                      if((r.supply||0)>0){groups[k].matched++;groups[k].mSum+=(r.markup||0);}
                    });
                    const rows=Object.values(groups).sort((a,b)=>a.baseDisc-b.baseDisc);
                    return (
                      <div style={{marginTop:18,border:`1px solid ${D.borderMid}`,borderRadius:6,overflow:"hidden"}}>
                        <div style={{padding:"9px 12px",fontSize:11,fontWeight:700,color:D.black,
                          background:D.surfaceAlt,borderBottom:`1px solid ${D.borderMid}`}}>
                          기본 세일율 %대별 결론 — 상품 수 / 평균 최종 할인율 / 평균 마크업 / 묶음 상품
                        </div>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                          <thead><tr>
                            {["기본 세일율","상품 수","평균 최종 할인율","평균 마크업","묶음 상품"].map((h,i)=>(
                              <th key={i} style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,
                                textAlign:i===0?"left":(i===4?"center":"right"),fontWeight:600,color:D.textSub,background:D.surface,whiteSpace:"nowrap"}}>{h}</th>
                            ))}
                          </tr></thead>
                          <tbody>
                            {rows.map((g,i)=>{
                              const avgFr=g.count>0?Math.round(g.frSum/g.count*10)/10:0;
                              const avgM=g.matched>0?Math.round(g.mSum/g.matched*100)/100:null;
                              return (
                                <React.Fragment key={i}>
                                <tr>
                                  <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,
                                    textAlign:"left",fontWeight:700,color:D.black,whiteSpace:"nowrap"}}>
                                    {g.baseDisc}% <span style={{fontSize:10,color:D.textMeta,fontWeight:400,marginLeft:4}}>기본 세일율</span>
                                  </td>
                                  <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"right",whiteSpace:"nowrap"}}>{g.count}개</td>
                                  <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"right",
                                    color:D.text,fontWeight:600,whiteSpace:"nowrap"}}>{avgFr}%</td>
                                  <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"right",fontWeight:700,whiteSpace:"nowrap",
                                    color:avgM==null?D.textMeta:(avgM>3?D.green:D.red)}}>
                                    {avgM==null?<span style={{fontWeight:400,color:D.textMeta}}>공급가 미매칭</span>:`×${avgM.toFixed(2)}`}
                                    {avgM!=null&&g.matched<g.count&&(
                                      <span style={{fontSize:9,color:D.textMeta,fontWeight:400,marginLeft:4}}>
                                        ({g.matched}/{g.count})
                                      </span>
                                    )}
                                  </td>
                                  <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"center"}}>
                                    <button onClick={()=>setExpandedGroup(expandedGroup===g.baseDisc?null:g.baseDisc)}
                                      style={{background:expandedGroup===g.baseDisc?D.black:"transparent",
                                        color:expandedGroup===g.baseDisc?"#fff":D.text,
                                        border:`1px solid ${D.borderMid}`,borderRadius:5,
                                        padding:"3px 9px",fontSize:11,cursor:"pointer",fontWeight:600}}>
                                      {expandedGroup===g.baseDisc?"▾ 닫기":`▸ 묶음 상품 ${g.count}개`}
                                    </button>
                                  </td>
                                </tr>
                                {expandedGroup===g.baseDisc&&(
                                  <tr><td colSpan={5} style={{padding:"0 10px 10px",borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt}}>
                                    <div style={{maxHeight:280,overflow:"auto",border:`1px solid ${D.border}`,borderRadius:4,background:D.surface}}>
                                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
                                        <thead><tr style={{background:D.surfaceAlt,color:D.textMeta}}>
                                          {["상품명","정가","쿠폰율","기본 할인율","프런트 판매가","최종 노출가","최종 할인율","자사부담","수수료","채널보전","자사 정산","공급가","마진","마크업"].map((h,k)=>(
                                            <th key={k} style={{padding:"5px 8px",textAlign:k===0?"left":"right",fontWeight:600,position:"sticky",top:0,background:D.surfaceAlt,whiteSpace:"nowrap"}}>{h}</th>
                                          ))}
                                        </tr></thead>
                                        <tbody>
                                          {g.products.map((p,j)=>{
                                            const sv=p.supplyIncVat||Math.round((p.supply||0)*1.1);
                                            return (
                                            <tr key={j} style={{borderTop:`1px solid ${D.border}`}}>
                                              <td title={p.name} style={{padding:"4px 8px",maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap"}}>₩{wonFmt(p.list)}</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap"}}>{cpn}%</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{p.baseDisc}%</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.blue,background:"#eef3ff",fontWeight:600}}>₩{wonFmt(p.basePrice)}</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.textSub}}>₩{wonFmt(p.finalPrice)}</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.blue,background:"#eef3ff",fontWeight:700}}>{p.finalDisc}%</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.selfBurden||0)>0?D.red:D.textMeta}}>
                                                {(p.selfBurden||0)>0?`−₩${wonFmt(p.selfBurden)}`:"—"}
                                              </td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.red}}>−₩{wonFmt(p.fee||0)} <span style={{fontSize:9,color:D.textMeta}}>({p.feeRate||0}%)</span></td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:(p.channelBurden||0)>0?D.blue:D.textMeta}}>
                                                {(p.channelBurden||0)>0?`+₩${wonFmt(p.channelBurden)}`:"—"}
                                              </td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>₩{wonFmt(p.net||0)}</td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:p.supply>0?D.text:D.textMeta}}>
                                                {p.supply>0?`₩${wonFmt(sv)}`:"—"}
                                              </td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600,
                                                color:p.supply>0?((p.margin||0)>=0?D.text:D.red):D.textMeta}}>
                                                {p.supply>0?`₩${wonFmt(p.margin||0)}`:"—"}
                                              </td>
                                              <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:700,
                                                color:p.supply>0?((p.markup||0)>3?D.green:D.red):D.textMeta}}>
                                                {p.supply>0?`×${(p.markup||0).toFixed(2)}`:"—"}
                                              </td>
                                            </tr>
                                            );
                                          })}
                                        </tbody>
                                      </table>
                                    </div>
                                  </td></tr>
                                )}
                                </React.Fragment>
                              );
                            })}
                          </tbody>
                        </table>
                        {(onCreatePromo||onAttachInlineCalc)&&(
                          <div style={{padding:"12px",borderTop:`1px solid ${D.borderMid}`,display:"flex",justifyContent:"center",background:D.surface}}>
                            <button onClick={()=>{
                              // 기본 세일율 그룹 → products.rows (묶음 상품 + 계산기 모든 필드 보존)
                              const productRows=rows.map(g=>{
                                const avgM=g.matched>0?Math.round(g.mSum/g.matched*100)/100:null;
                                return {
                                  group:`기본 세일율 ${g.baseDisc}%`,
                                  rate:String(g.baseDisc),
                                  markup:avgM!=null?String(avgM):"",
                                  cpn,
                                  products:g.products.map(p=>({
                                    code:p.code||"",name:p.name||"",
                                    list:p.list||0,baseDisc:p.baseDisc||0,
                                    basePrice:p.basePrice||0,finalPrice:p.finalPrice||0,
                                    finalDisc:p.finalDisc||0,markup:p.markup||0,
                                    supply:p.supply||0,supplyIncVat:p.supplyIncVat||Math.round((p.supply||0)*1.1),
                                    selfBurden:p.selfBurden||0,channelBurden:p.channelBurden||0,
                                    fee:p.fee||0,feeRate:p.feeRate||0,
                                    net:p.net||0,margin:p.margin||0,
                                  })),
                                };
                              });
                              const couponRows=allCoupons.map((c,i)=>{
                                const tInfo=COUPON_TYPE_BY_KEY[c.type];
                                return {
                                  ...emptyCouponRow(c.type),
                                  name:`${tInfo.short} 쿠폰 ${i===0?"(기본)":`#${i}`}`,
                                  rate:String(c.rate),
                                  burden:c.burden,
                                  shareRate:c.shareRate,
                                };
                              });
                              const groupLines=rows.map(g=>{
                                const avgFr=g.count>0?Math.round(g.frSum/g.count*10)/10:0;
                                const avgM=g.matched>0?Math.round(g.mSum/g.matched*100)/100:null;
                                return `• 기본 세일율 ${g.baseDisc}% · ${g.count}개 · 최종할인 ${avgFr}%${avgM!=null?` · 평균 마크업 ×${avgM.toFixed(2)} (${g.matched}/${g.count} 매칭)`:" · 공급가 미매칭"}`;
                              }).join("\n");
                              const couponLines=allCoupons.map(c=>{
                                const tInfo=COUPON_TYPE_BY_KEY[c.type];
                                const burdenLabel=c.type==="share"
                                  ?`분담 자사${100-(c.shareRate||0)}:채널${c.shareRate||0}`
                                  :(c.burden==="channel"?"채널부담":"자사부담");
                                return `• ${tInfo.label} ${c.rate}% · ${burdenLabel}`;
                              }).join("\n");
                              const scenarioLine=selectedScenario.caseNum?`Case ${selectedScenario.caseNum} · ${selectedScenario.label} · 케이스 총합 할인율 ${cpn}%`:`기본 쿠폰 ${cpn}%`;
                              const content=`[적용 시나리오]\n${scenarioLine}\n\n[입력된 쿠폰]\n${couponLines}\n\n[기본 세일율별 결론]\n${groupLines}`;
                              const payload={
                                platform:"29CM",
                                content,
                                discount_plan:{
                                  products:{period:{start:"",end:""},rows:productRows},
                                  coupons:couponRows,
                                },
                              };
                              if(onAttachInlineCalc) onAttachInlineCalc(payload);
                              else onCreatePromo(payload);
                            }}
                              style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
                                padding:"9px 22px",fontSize:11,cursor:"pointer",fontWeight:700}}>
                              {onAttachInlineCalc?(attachMode==="fill"?"+ 이 행 묶음으로 채우기":"+ 매트릭스에 묶음 추가"):"+ 29CM 프로모션 추가하기"}
                            </button>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// 29CM 프로모션 추가/편집 시 사용하는 미니 계산기
// — 쿠폰율 입력 → 5개 가격 구간(P75 + 역산 기본 할인율) 선택 → 기본×쿠폰 매트릭스 셀 클릭 시 상품군·쿠폰 자동 입력
// ─────────────────────────────────────────────
// 자사몰 세일율 계산기 (베타) — 카페24 상품 파일 업로드 → 세일율 입력 → 마진/마크업
//   · 멤버십 쿠폰은 서로 교차 불가(하나만 적용) · 가격대 구간/제안 할인율 없음
//   · 엑셀 추출: 상품코드 · 상품명 · 할인 이후 가격 · 할인율
// ─────────────────────────────────────────────
// 자사몰(카페24) 상품 마스터 파서 — 상품코드/상품명/판매가/공급가만 추출
function parseMallProductFile(file,onResult,onError){
  const pickKey=(row,cands)=>{
    const map={};
    Object.keys(row||{}).forEach(k=>{ map[String(k).replace(/^﻿/,"").trim()]=k; });
    for(const c of cands){ if(map[c]!==undefined) return map[c]; }
    return undefined;
  };
  const finish=rows=>{
    if(!rows||!rows.length){ onError("데이터 행이 없습니다"); return; }
    const kCode=pickKey(rows[0],["상품코드","product_code","상품 코드"]);
    const kName=pickKey(rows[0],["상품명","product_name"]);
    const kSell=pickKey(rows[0],["판매가","selling_price"]);
    const kSup =pickKey(rows[0],["공급가","supply_price","원가"]);
    if(kName===undefined||kSell===undefined){
      onError(`'상품명' / '판매가' 컬럼을 찾지 못했습니다 — 헤더: ${Object.keys(rows[0]).map(k=>String(k).replace(/^﻿/,"").trim()).slice(0,30).join(" / ")}`);
      return;
    }
    const out=[];
    // 원본 파일 행 순서 보존 — _origRow 인덱스로 명시 (이후 정렬/필터 시 기준)
    rows.forEach((r,rowIdx)=>{
      const name=String(r[kName]||"").trim();
      const selling=toNum(r[kSell]);
      if(!name||selling<=0) return;
      out.push({
        _origRow:rowIdx,
        code:kCode!==undefined?String(r[kCode]||"").trim():"",
        name,
        selling,
        supply:kSup!==undefined?toNum(r[kSup]):0,
      });
    });
    if(!out.length){ onError("유효한 상품 행이 없습니다 (상품명·판매가 확인)"); return; }
    onResult(out);
  };
  const ext=(file.name||"").toLowerCase();
  if(ext.endsWith(".csv")){
    (async()=>{
      try{
        const Papa=await getPapa();
        const text=await file.text();
        Papa.parse(text,{header:true,skipEmptyLines:true,transformHeader:h=>String(h).replace(/^﻿/,"").trim(),
          complete:res=>finish(res.data),
          error:err=>onError(String(err?.message||err))});
      }catch(err){ onError(String(err?.message||err)); }
    })();
  }else{
    const reader=new FileReader();
    reader.onload=async e=>{
      try{
        const XLSX=await getXLSX();
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        finish(XLSX.utils.sheet_to_json(ws,{defval:""}));
      }catch(err){ onError(String(err?.message||err)); }
    };
    reader.onerror=()=>onError("파일 읽기 도중 시스템 오류가 발생했습니다");
    reader.readAsArrayBuffer(file);
  }
}

function OwnMallSaleCalcModal({ onClose, onCreatePromo, onAttachInlineCalc, attachMode, initialCoupon, initialCouponName }){
  const [products,setProducts]=useState([]);
  const [fileName,setFileName]=useState("");
  const [status,setStatus]=useState("");
  const [dragOver,setDragOver]=useState(false);
  const [rates,setRates]=useState({});           // 상품별 할인율 % (index→값, 기본 10)
  const [coupons,setCoupons]=useState(
    initialCoupon!=null
      ?[{name:initialCouponName||"기존 쿠폰",rate:initialCoupon}]
      :[{name:"멤버십 10%",rate:10}]
  ); // 기본 쿠폰 10% 자동 추가 (삭제 가능)
  const [selCoupon,setSelCoupon]=useState(0);     // 기본 쿠폰 선택 (-1 = 쿠폰 없음)
  const [search,setSearch]=useState("");   // 표 내 검색 (상품명·할인율)
  const [sample,setSample]=useState(null);        // {filename} — Supabase 보관 메타
  const [sampleMsg,setSampleMsg]=useState("");    // 샘플 저장/로드 상태 메시지
  const [removedIdx,setRemovedIdx]=useState(()=>new Set()); // 수기로 제거한 상품 (임시 — 다운로드/재업로드/리로드 시 복원)
  const [checkedIdx,setCheckedIdx]=useState(()=>new Set()); // 일괄 삭제용 체크박스 선택 인덱스
  const [mallExpandedGroup,setMallExpandedGroup]=useState(null); // 결론 표 — 펼친 묶음 그룹 (할인율 %)
  const modalCardRef=useRef(null);
  const inNum={background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"5px 8px",fontSize:12,color:D.text,fontFamily:"inherit"};
  const loadProducts=(rows,name)=>{ setProducts(rows); setFileName(name); setRates({}); setRemovedIdx(new Set()); setCheckedIdx(new Set()); setStatus(`${rows.length.toLocaleString()}개 상품 로드됨`); };
  // 마지막 업로드의 파싱 결과(상품 목록)를 Supabase(mall_calc_last_file)에 보관 → 원본 파일이 커도(수십 MB) 안전.
  //   content_b64 컬럼에 base64 대신 상품 JSON 을 저장한다(텍스트 컬럼이라 호환). 모달 열 때 자동 로드.
  const sampleToProducts=raw=>{ try{ const p=JSON.parse(raw||""); return Array.isArray(p)?p:null; }catch{ return null; } };
  useEffect(()=>{ let alive=true;
    (async()=>{
      try{
        const db=await getSupabase();
        const {data,error}=await db.from("mall_calc_last_file").select("filename,content_b64").eq("id",1).maybeSingle();
        if(error||!data||!alive) return;
        const prods=sampleToProducts(data.content_b64);
        if(!prods||!prods.length) return;
        setSample({filename:data.filename});
        loadProducts(prods,data.filename||"샘플");
        setSampleMsg(`📎 샘플 자동 로드됨 — ${data.filename||""} (${prods.length.toLocaleString()}개)`);
      }catch{}
    })();
    return()=>{alive=false;};
  },[]);
  const saveSample=async(name,rows)=>{
    try{
      const db=await getSupabase();
      const {error}=await db.from("mall_calc_last_file").upsert({id:1,filename:name,content_b64:JSON.stringify(rows),uploaded_at:new Date().toISOString()});
      if(error) throw error;
      setSample({filename:name}); setSampleMsg(`📎 샘플로 저장됨 (${rows.length.toLocaleString()}개) — 다음에 이 모달을 열면 자동 로드됩니다`);
    }catch(err){ setSampleMsg("⚠ 샘플 저장 실패: "+(err?.message||err)); }
  };
  const loadSample=async()=>{
    try{
      const db=await getSupabase();
      const {data,error}=await db.from("mall_calc_last_file").select("filename,content_b64").eq("id",1).maybeSingle();
      if(error) throw error;
      const prods=sampleToProducts(data?.content_b64);
      if(!prods||!prods.length){ setSampleMsg("저장된 샘플이 없습니다 — 파일을 한 번 업로드해 주세요."); return; }
      loadProducts(prods,data.filename||"샘플");
      setSampleMsg(`📎 샘플 불러옴 — ${data.filename||""} (${prods.length.toLocaleString()}개)`);
    }catch(err){ setSampleMsg("샘플 로드 실패: "+(err?.message||err)); }
  };
  const handleFile=f=>{
    if(!f) return;
    setFileName(f.name);setStatus("파싱 중…");setProducts([]);setRates({});
    parseMallProductFile(f,rows=>{ loadProducts(rows,f.name); saveSample(f.name,rows); }, err=>{ setStatus("오류: "+err); });
  };
  const couponRate=selCoupon>=0?Math.max(0,Math.min(100,Number(coupons[selCoupon]?.rate)||0)):0;
  const couponName=selCoupon>=0?(coupons[selCoupon]?.name||`쿠폰 ${selCoupon+1}`):"";
  // 가격 DB(calc_supply_override + 인벤토리 스냅샷) 동기화 — 공급가/정상가 우선 적용
  const {priceOf,ready:priceReady}=useInventoryPricing();
  // 곱연산 체인: 판매가 > 할인율 > 할인금액 > 할인가 > 쿠폰율 > 쿠폰금액 > 쿠폰적용가 > 원가 > 마진 > 마크업
  //   · 공급가는 가격 DB(데이터 입력 > 인벤토리 > 가격 DB) 값을 우선 사용하고, 없으면 파일 값
  //   · 정상가는 파일 값(카페24 판매가)을 우선 사용하고, 없으면 가격 DB 폴백
  const rows=useMemo(()=>products
    .map((p,i)=>({p,i}))
    .filter(({i})=>!removedIdx.has(i))
    // 원본 파일 행 순서 보존: _origRow 가 있으면 그 순으로, 없으면 products 인덱스 순
    .sort((a,b)=>{
      const ao=(a.p&&a.p._origRow!=null)?a.p._origRow:a.i;
      const bo=(b.p&&b.p._origRow!=null)?b.p._origRow:b.i;
      return ao-bo;
    })
    .map(({p,i})=>{
      const rate=Math.max(0,Math.min(100,parseFloat(rates[i]??10)||0));
      const priced=priceOf?priceOf(p.name,p.code):{selling:0,supply:0};
      const supply=priced.supply||p.supply||0;
      const selling=p.selling||priced.selling||0;
      const discAmt=Math.round(selling*rate/100);
      const discPrice=selling-discAmt;
      const couponAmt=Math.round(discPrice*couponRate/100);
      const couponPrice=discPrice-couponAmt;
      const supplyVat=Math.round(supply*1.1);
      const margin=couponPrice-supplyVat;
      // 실수령 마크업 = 실수령액(쿠폰적용가, 자사몰 수수료 0%) ÷ 원가(부가세 포함)
      const markup=supplyVat>0?couponPrice/supplyVat:0;
      const effDisc=selling>0?(1-couponPrice/selling)*100:0;
      return {...p,idx:i,rate,selling,supply,discAmt,discPrice,couponAmt,couponPrice,supplyVat,margin,markup,effDisc,supplyFromDb:priced.supply>0};
    }),[products,rates,couponRate,priceOf,removedIdx]);
  // 수기 제거 / 복원 / 일괄
  const removeRow=(idx)=>setRemovedIdx(prev=>{const next=new Set(prev);next.add(idx);return next;});
  const restoreAll=()=>{setRemovedIdx(new Set());setCheckedIdx(new Set());};
  const toggleCheck=(idx)=>setCheckedIdx(prev=>{const next=new Set(prev);next.has(idx)?next.delete(idx):next.add(idx);return next;});
  const checkedVisibleAll=rowsRef=>rowsRef.length>0&&rowsRef.every(r=>checkedIdx.has(r.idx));
  const removeChecked=()=>{
    if(checkedIdx.size===0) return;
    setRemovedIdx(prev=>{const next=new Set(prev);checkedIdx.forEach(i=>next.add(i));return next;});
    setCheckedIdx(new Set());
  };
  // 드래그 다중 선택 — 체크박스에서 mouseDown → mode 결정, 다른 행에 mouseEnter 시 적용
  const dragSelectModeRef=useRef(null);
  const startDragSelect=(idx,isChecked)=>{
    dragSelectModeRef.current=isChecked?"remove":"add";
    setCheckedIdx(prev=>{const next=new Set(prev);if(dragSelectModeRef.current==="add")next.add(idx);else next.delete(idx);return next;});
  };
  const enterDragSelect=(idx)=>{
    if(!dragSelectModeRef.current) return;
    setCheckedIdx(prev=>{const next=new Set(prev);if(dragSelectModeRef.current==="add")next.add(idx);else next.delete(idx);return next;});
  };
  useEffect(()=>{
    const onUp=()=>{dragSelectModeRef.current=null;};
    window.addEventListener("mouseup",onUp);
    return ()=>window.removeEventListener("mouseup",onUp);
  },[]);
  const dbMatchedCount=useMemo(()=>rows.filter(r=>r.supplyFromDb).length,[rows]);
  const agg=useMemo(()=>{
    if(!rows.length) return null;
    const n=rows.length;
    const matched=rows.filter(r=>r.supplyVat>0);
    return {n,
      avgDisc:rows.reduce((s,r)=>s+r.effDisc,0)/n,
      avgMarkup:matched.length>0?matched.reduce((s,r)=>s+r.markup,0)/matched.length:0,
      matchedCount:matched.length,
      noCost:rows.filter(r=>!r.supplyVat).length,
      neg:rows.filter(r=>r.margin<0).length};
  },[rows]);
  const won=n=>"₩"+Math.round(n||0).toLocaleString();
  const pct=v=>`${(Math.round(v*10)/10).toFixed(1)}%`;
  const setRate=(i,v)=>setRates(prev=>({...prev,[i]:v}));
  const setAllRates=v=>setRates(()=>{const m={};products.forEach((_,i)=>{m[i]=v;});return m;});
  const exportXlsx=async()=>{
    if(!rows.length) return;
    const XLSX=await getXLSX();
    const aoa=[["상품코드","상품명","기본 할인가","기본 할인율(%)"],
      ...rows.map(r=>[r.code,r.name,r.discPrice,r.rate])];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),"세일율");
    XLSX.writeFile(wb,`자사몰_세일율${couponName?"_"+couponName:""}_${dayjs().format("YYYYMMDD")}.xlsx`);
    // 다운로드 후에도 임시 제거 / 체크 상태 유지 — 모달 재오픈 / 파일 재업로드 시점까지 보존
  };
  // + 자사몰 프로모션 추가 / 인라인 묶음 추가 — 할인율(rate) 5% 버킷으로 묶고 묶음 상품 리스트 전달
  const handleCreatePromo=()=>{
    if(!(onCreatePromo||onAttachInlineCalc)||!rows.length) return;
    const groups={};
    rows.forEach(r=>{
      const k=Math.round((r.rate||0)/5)*5;
      if(!groups[k]) groups[k]={baseDisc:k,products:[],count:0,matched:0,mSum:0,effSum:0};
      groups[k].count++;
      groups[k].effSum+=(r.effDisc||0);
      groups[k].products.push(r);
      if((r.supplyVat||0)>0){groups[k].matched++;groups[k].mSum+=(r.markup||0);}
    });
    const groupArr=Object.values(groups).sort((a,b)=>a.baseDisc-b.baseDisc);
    const productRows=groupArr.map(g=>{
      const avgM=g.matched>0?Math.round(g.mSum/g.matched*100)/100:null;
      return {
        group:`상품 할인 ${g.baseDisc}%`,
        rate:String(g.baseDisc),
        markup:avgM!=null?String(avgM):"",
        cpn:couponRate||0,
        products:g.products.map(p=>({
          code:p.code||"",name:p.name||"",
          list:p.selling||0,baseDisc:p.rate||0,
          basePrice:p.discPrice||0,finalPrice:p.couponPrice||0,
          finalDisc:Math.round((p.effDisc||0)*10)/10,markup:p.markup||0,
          supply:p.supply||0,supplyIncVat:p.supplyVat||0,
          selfBurden:0,channelBurden:0,fee:0,feeRate:0,
          net:p.couponPrice||0,margin:p.margin||0,
        })),
      };
    });
    const couponRows=couponRate>0?[{...emptyCouponRow("product"),name:couponName||"멤버십 쿠폰",rate:String(couponRate)}]:[];
    const groupLines=groupArr.map(g=>{
      const avgFr=g.count>0?Math.round(g.effSum/g.count*10)/10:0;
      const avgM=g.matched>0?Math.round(g.mSum/g.matched*100)/100:null;
      return `• 상품 할인 ${g.baseDisc}% · ${g.count}개 · 최종할인 ${avgFr}%${avgM!=null?` · 평균 마크업 ×${avgM.toFixed(2)} (${g.matched}/${g.count} 매칭)`:" · 공급가 미매칭"}`;
    }).join("\n");
    const content=`[자사몰 세일율 계산기]\n쿠폰: ${couponRate>0?`${couponName||"멤버십"} ${couponRate}%`:"없음"}\n\n[기본 세일율별 결론]\n${groupLines}`;
    const payload={
      platform:"자사몰",
      content,
      discount_plan:{
        products:{period:{start:"",end:""},rows:productRows},
        coupons:couponRows,
      },
    };
    if(onAttachInlineCalc) onAttachInlineCalc(payload);
    else onCreatePromo(payload);
  };
  const numCell={padding:"4px 6px",textAlign:"right"};
  // 표 내 검색 — 상품명 또는 할인율(%) 부분일치. 검색 중에는 전체 매칭 표시(limit 미적용).
  const q=search.trim().toLowerCase();
  const filtered=q?rows.filter(r=>(r.name||"").toLowerCase().includes(q)||String(r.rate).includes(q)):rows;
  const shown=filtered; // 전체 로드 (더보기 없음)
  return (
    <div onClick={onClose}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div ref={modalCardRef} onClick={e=>e.stopPropagation()} className="mallcalc"
        style={{background:D.surface,borderRadius:14,width:"90vw",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 8px 40px rgba(0,0,0,0.22)"}}>
        <style>{`.mallcalc input[type="number"]::-webkit-inner-spin-button,.mallcalc input[type="number"]::-webkit-outer-spin-button{-webkit-appearance:none;margin:0;}.mallcalc input[type="number"]{-moz-appearance:textfield;appearance:textfield;}.mallcalc thead th{position:sticky;top:0;background:${D.surface};z-index:2;box-shadow:inset 0 -1px 0 ${D.border};}.mallcalc tbody tr:hover td{background:${D.surfaceAlt};}`}</style>
        <div style={{position:"sticky",top:0,background:D.surface,borderBottom:`1px dashed ${D.border}`,
          padding:"14px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",zIndex:5}}>
          <b style={{fontSize:13,color:D.black,fontWeight:700}}>자사몰 세일율 계산기
            <span style={{fontSize:10,fontWeight:700,color:"#fff",background:D.amber,borderRadius:10,padding:"2px 7px",marginLeft:6}}>베타</span></b>
          <div style={{display:"flex",gap:6}}>
            <CaptureBtn cardRef={modalCardRef} filename={`자사몰세일율_${dayjs().format("YYYYMMDD")}`} DC={{border:D.border,sub:D.textMeta}}/>
            <button onClick={onClose} style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,width:32,height:32,cursor:"pointer",fontSize:11,color:D.textMeta}}>✕</button>
          </div>
        </div>
        <div style={{padding:"18px 20px 36px"}}>
          <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
            onDrop={e=>{e.preventDefault();setDragOver(false);handleFile(e.dataTransfer.files[0]);}}
            onClick={()=>{const inp=document.createElement("input");inp.type="file";inp.accept=".csv,.xlsx,.xls";inp.onchange=ev=>handleFile(ev.target.files[0]);inp.click();}}
            style={{border:`1.5px dashed ${dragOver?D.blue:D.border}`,borderRadius:10,padding:18,textAlign:"center",cursor:"pointer",
              background:dragOver?`${D.blue}08`:D.surfaceAlt,marginBottom:10}}>
            <div style={{fontSize:13,fontWeight:600,color:D.text,marginBottom:4}}>자사몰 상품 파일 (CSV / Excel) 드래그 &amp; 드롭</div>
            <div style={{fontSize:11,color:D.textMeta}}>상품코드 · 상품명 · 판매가 · 공급가 컬럼 사용 (카페24 상품 엑셀)</div>
            {fileName&&<div style={{marginTop:6,fontSize:12,color:D.blue}}>{fileName}</div>}
            {status&&<div style={{marginTop:4,fontSize:11,color:status.startsWith("오류")?D.red:D.textSub}}>{status}</div>}
          </div>

          {(sample||sampleMsg)&&(
            <div style={{marginBottom:14,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              {sample&&(
                <button onClick={loadSample}
                  style={{background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:6,padding:"6px 12px",fontSize:12,cursor:"pointer",color:D.text}}>
                  📎 샘플 파일 불러오기 — {sample.filename}
                </button>
              )}
              {sampleMsg&&<span style={{fontSize:11,color:sampleMsg.startsWith("⚠")?D.red:D.green}}>{sampleMsg}</span>}
            </div>
          )}

          {products.length>0&&(<>
            <div style={{padding:"12px 14px",background:D.surface,border:`1px solid ${D.black}`,borderRadius:10,marginBottom:14}}>
              <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:10}}>
                <span style={{fontSize:12,fontWeight:700,color:D.black}}>할인율</span>
                <span style={{fontSize:11,color:D.textMeta}}>상품별로 입력(기본 10%) · 변경 시 마진·마크업 실시간 반영</span>
                <span style={{fontSize:11,color:D.textMeta,marginLeft:"auto"}}>전체 일괄:</span>
                {[0,10,15,20,30].map(v=>(
                  <button key={v} onClick={()=>setAllRates(v)}
                    style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 8px",fontSize:11,cursor:"pointer",color:D.textSub}}>{v}%</button>
                ))}
              </div>
              <div style={{borderTop:`1px dashed ${D.border}`,paddingTop:10}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
                  <span style={{fontSize:11,fontWeight:700,color:D.textMeta}}>멤버십 쿠폰 (서로 교차 불가 · 하나만 적용 · 상품 할인 후 곱연산)</span>
                  <button onClick={()=>setCoupons(c=>[...c,{name:"",rate:5}])}
                    style={{background:D.blue,color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer",fontWeight:600,marginLeft:"auto"}}>+ 쿠폰 추가</button>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  <label style={{display:"flex",alignItems:"center",gap:8,fontSize:12,cursor:"pointer"}}>
                    <input type="radio" name="mallcoupon" checked={selCoupon===-1} onChange={()=>setSelCoupon(-1)}/>
                    <span style={{color:D.text}}>쿠폰 없음</span>
                  </label>
                  {coupons.map((c,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                      <input type="radio" name="mallcoupon" checked={selCoupon===i} onChange={()=>setSelCoupon(i)}/>
                      <input placeholder={`쿠폰 ${i+1} 이름`} value={c.name}
                        onChange={e=>setCoupons(arr=>arr.map((x,j)=>j===i?{...x,name:e.target.value}:x))} style={{...inNum,width:140}}/>
                      <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="100" step="1" value={c.rate}
                        onChange={e=>setCoupons(arr=>arr.map((x,j)=>j===i?{...x,rate:e.target.value}:x))} style={{...inNum,width:70}}/>
                      <span style={{fontSize:11,color:D.blue}}>%</span>
                      <button onClick={()=>{setCoupons(arr=>arr.filter((_,j)=>j!==i));setSelCoupon(s=>s===i?-1:s>i?s-1:s);}}
                        style={{background:"none",border:"none",color:D.red,cursor:"pointer",fontSize:13}}>✕</button>
                    </div>
                  ))}
                </div>
                {couponRate>0&&<div style={{fontSize:11,color:D.textSub,marginTop:8}}>적용 쿠폰: {couponName||"멤버십"} {couponRate}% (각 상품 할인가에 곱연산)</div>}
              </div>
            </div>

            {agg&&(
              <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:12,alignItems:"center"}}>
                {[["상품 수",`${agg.n.toLocaleString()}개`],["평균 할인율",`${(Math.round(agg.avgDisc*100)/100).toFixed(2)}%`],["평균 마크업",agg.matchedCount>0?`×${agg.avgMarkup.toFixed(2)}`:"—"]].map(([k,v])=>(
                  <div key={k} style={{padding:"8px 14px",background:D.surfaceAlt,borderRadius:8}}>
                    <div style={{fontSize:10,color:D.textMeta,marginBottom:2}}>{k}</div>
                    <div style={{fontSize:15,fontWeight:700,color:D.text}}>{v}</div>
                  </div>
                ))}
                {priceReady&&<span style={{fontSize:11,color:dbMatchedCount>0?D.green:D.textMeta}} title="데이터 입력 > 인벤토리 > 가격 DB 와 동기화된 공급가 건수">🔗 가격 DB 매칭 {dbMatchedCount.toLocaleString()}/{rows.length.toLocaleString()}건</span>}
                {agg.noCost>0&&<span style={{fontSize:11,color:D.amber}}>공급가 미입력 {agg.noCost}개 (마진 과대평가)</span>}
                {agg.neg>0&&<span style={{fontSize:11,color:D.red}}>역마진 {agg.neg}개</span>}
                {checkedIdx.size>0&&(
                  <button onClick={removeChecked}
                    style={{display:"inline-flex",alignItems:"center",gap:4,background:D.red,color:"#fff",border:"none",borderRadius:5,
                      padding:"5px 12px",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                    🗑 체크 {checkedIdx.size}개 일괄 삭제
                  </button>
                )}
                {removedIdx.size>0&&(
                  <span style={{display:"inline-flex",alignItems:"center",gap:6,fontSize:11,color:D.textSub,
                    border:`1px dashed ${D.borderMid}`,borderRadius:5,padding:"3px 8px"}}>
                    🗑 {removedIdx.size}개 제거 <span style={{color:D.textMeta}}>(임시 · 다운로드/재업로드 시 복원)</span>
                    <button onClick={restoreAll}
                      style={{background:"transparent",border:"none",cursor:"pointer",color:D.blue,fontSize:11,fontWeight:600,padding:"0 4px"}}>↻ 전체 복원</button>
                  </span>
                )}
                {(onCreatePromo||onAttachInlineCalc)&&(
                  <button onClick={handleCreatePromo}
                    style={{marginLeft:"auto",background:D.black,color:"#fff",border:"none",borderRadius:6,padding:"8px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                    {onAttachInlineCalc?(attachMode==="fill"?"+ 이 행 묶음으로 채우기":"+ 매트릭스에 묶음 추가"):"+ 자사몰 프로모션 추가하기"}
                  </button>
                )}
                <button onClick={exportXlsx}
                  style={{marginLeft:onCreatePromo?0:"auto",background:D.black,color:"#fff",border:"none",borderRadius:6,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                  엑셀 추출 (코드·상품명·기본 할인가·기본 할인율)
                </button>
              </div>
            )}

            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="상품명 또는 할인율로 검색"
                style={{...inNum,width:280,maxWidth:"60vw"}}/>
              {q&&<span style={{fontSize:11,color:D.textMeta}}>{filtered.length.toLocaleString()}개 검색됨 · 전체 {rows.length.toLocaleString()}</span>}
              {q&&<button onClick={()=>setSearch("")} style={{background:"none",border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 9px",fontSize:11,cursor:"pointer",color:D.textSub}}>✕ 초기화</button>}
            </div>
            <div style={{overflow:"auto",height:"55vh"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,whiteSpace:"nowrap",tableLayout:"fixed"}}>
                <colgroup>
                  <col style={{width:"7%"}}/>
                  <col style={{width:"17%"}}/>
                  <col style={{width:"9%"}}/>
                  <col style={{width:"6%"}}/>
                  <col style={{width:"7%"}}/>
                  <col style={{width:"9%"}}/>
                  <col style={{width:"5%"}}/>
                  <col style={{width:"7%"}}/>
                  <col style={{width:"11%"}}/>
                  <col style={{width:"7%"}}/>
                  <col style={{width:"8%"}}/>
                  <col style={{width:"7%"}}/>
                </colgroup>
                <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                  <th style={{padding:"4px 6px",textAlign:"left",fontWeight:500}}>
                    <input type="checkbox" title="현재 표시된 행 전체 선택/해제"
                      checked={shown.length>0&&shown.every(r=>checkedIdx.has(r.idx))}
                      onChange={()=>{
                        const allChecked=shown.length>0&&shown.every(r=>checkedIdx.has(r.idx));
                        setCheckedIdx(prev=>{
                          const next=new Set(prev);
                          if(allChecked) shown.forEach(r=>next.delete(r.idx));
                          else shown.forEach(r=>next.add(r.idx));
                          return next;
                        });
                      }}
                      style={{marginRight:4,cursor:"pointer"}}/>
                    상품코드
                  </th>
                  <th style={{padding:"4px 6px",textAlign:"left",fontWeight:500}}>상품명</th>
                  <th style={{...numCell,fontWeight:500}}>판매가</th>
                  <th style={{...numCell,fontWeight:500}}>할인율</th>
                  <th style={{...numCell,fontWeight:500}}>할인금액</th>
                  <th style={{...numCell,fontWeight:500}}>할인가</th>
                  <th style={{...numCell,fontWeight:500}}>쿠폰율</th>
                  <th style={{...numCell,fontWeight:500}}>쿠폰금액</th>
                  <th style={{...numCell,fontWeight:500}}>쿠폰적용가 (최종할인)</th>
                  <th style={{...numCell,fontWeight:500}}>원가(VAT)</th>
                  <th style={{...numCell,fontWeight:500}}>마진</th>
                  <th style={{...numCell,fontWeight:500}}>마크업</th>
                </tr></thead>
                <tbody>
                  {shown.map((r)=>{
                    const sellMu=r.supplyVat>0?Math.round(r.selling/r.supplyVat*100)/100:null;
                    const discMu=r.supplyVat>0?Math.round(r.discPrice/r.supplyVat*100)/100:null;
                    const cpnMu=r.supplyVat>0?Math.round(r.couponPrice/r.supplyVat*100)/100:null;
                    const muBadge=(v)=>v==null?null:(
                      <span style={{marginLeft:4,fontSize:10,fontWeight:700,color:v>3?D.green:D.red}}>×{v.toFixed(2)}</span>
                    );
                    return (
                    <tr key={r.code+r.idx}
                      onMouseEnter={()=>enterDragSelect(r.idx)}
                      style={{borderBottom:`1px solid ${D.border}`,background:checkedIdx.has(r.idx)?`${D.red}0a`:"transparent"}}>
                      <td
                        onMouseDown={e=>{
                          if(e.target.tagName==="INPUT"||e.target.tagName==="BUTTON"||e.target.closest("button")) return;
                          e.preventDefault();
                          startDragSelect(r.idx,checkedIdx.has(r.idx));
                        }}
                        style={{padding:"4px 6px",color:D.textMeta,fontFamily:"monospace",cursor:"pointer",userSelect:"none"}}>
                        <input type="checkbox" checked={checkedIdx.has(r.idx)} onChange={()=>{}}
                          onMouseDown={e=>{e.preventDefault();startDragSelect(r.idx,checkedIdx.has(r.idx));}}
                          title="클릭 또는 드래그로 다중 선택"
                          style={{marginRight:4,cursor:"pointer",verticalAlign:"middle"}}/>
                        <button onClick={()=>removeRow(r.idx)} title="이 상품을 표에서 임시로 제거 (다운로드/재업로드 시 복원)"
                          style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:11,padding:"0 4px",marginRight:2}}>✕</button>
                        {r.code}
                      </td>
                      <td
                        onMouseDown={e=>{
                          if(e.target.tagName==="INPUT"||e.target.tagName==="BUTTON"||e.target.closest("button")) return;
                          e.preventDefault();
                          startDragSelect(r.idx,checkedIdx.has(r.idx));
                        }}
                        style={{padding:"4px 6px",color:D.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",cursor:"pointer",userSelect:"none"}} title={r.name}>{r.name}</td>
                      <td style={{...numCell,color:D.textMeta}}>{won(r.selling)}{muBadge(sellMu)}</td>
                      <td style={numCell}>
                        <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="100" step="1"
                          value={rates[r.idx]??10} onChange={e=>setRate(r.idx,e.target.value)}
                          style={{width:46,textAlign:"right",background:"transparent",border:`1px solid ${D.border}`,borderRadius:4,padding:"2px 4px",fontSize:11,color:D.text,fontFamily:"inherit"}}/>
                        <span style={{color:D.textMeta}}> %</span>
                      </td>
                      <td style={{...numCell,color:D.textMeta}}>{won(r.discAmt)}</td>
                      <td style={{...numCell,color:D.textSub}}>{won(r.discPrice)}{muBadge(discMu)}</td>
                      <td style={{...numCell,color:D.textMeta}}>{couponRate?pct(couponRate):"—"}</td>
                      <td style={{...numCell,color:D.textMeta}}>{couponRate?won(r.couponAmt):"—"}</td>
                      <td style={{...numCell,color:D.text,fontWeight:700}}>
                        {won(r.couponPrice)}
                        {muBadge(cpnMu)}
                        {r.effDisc>0&&<span style={{marginLeft:6,fontSize:10,fontWeight:600,color:D.red}}>−{(Math.round(r.effDisc*10)/10).toFixed(1)}%</span>}
                      </td>
                      <td style={{...numCell,color:D.textMeta}}>{r.supplyVat?won(r.supplyVat):"—"}</td>
                      <td style={{...numCell,fontWeight:600,color:r.margin>=0?D.green:D.red}}>{won(r.margin)}</td>
                      <td style={{...numCell,fontWeight:700,color:r.supplyVat>0?(r.markup>3?D.green:D.red):D.textMeta}}>{r.supplyVat>0?`×${r.markup.toFixed(2)}`:"—"}</td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* 하단 결론 — 할인율 5% 버킷별 그룹 (29CM 계산기와 동일 UI 패턴) */}
            {rows.length>0&&(()=>{
              const groups={};
              rows.forEach(r=>{
                const k=Math.round((r.rate||0)/5)*5;
                if(!groups[k]) groups[k]={baseDisc:k,products:[],count:0,matched:0,mSum:0,frSum:0};
                groups[k].count++;
                groups[k].frSum+=(r.effDisc||0);
                groups[k].products.push(r);
                if((r.supplyVat||0)>0){groups[k].matched++;groups[k].mSum+=(r.markup||0);}
              });
              const groupArr=Object.values(groups).sort((a,b)=>a.baseDisc-b.baseDisc);
              return (
                <div style={{marginTop:18,border:`1px solid ${D.borderMid}`,borderRadius:6,overflow:"hidden"}}>
                  <div style={{padding:"9px 12px",fontSize:11,fontWeight:700,color:D.black,
                    background:D.surfaceAlt,borderBottom:`1px solid ${D.borderMid}`}}>
                    할인율 %대별 결론 — 상품 수 / 평균 최종 할인율 / 평균 마크업 / 묶음 상품 <span style={{color:D.textMeta,fontWeight:400}}>· 할인율 수정 시 실시간 반영</span>
                  </div>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                    <thead><tr>
                      {["할인율","상품 수","평균 최종 할인율","평균 마크업","묶음 상품"].map((h,i)=>(
                        <th key={i} style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,
                          textAlign:i===0?"left":(i===4?"center":"right"),fontWeight:600,color:D.textSub,background:D.surface,whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr></thead>
                    <tbody>
                      {groupArr.map((g,i)=>{
                        const avgFr=g.count>0?Math.round(g.frSum/g.count*10)/10:0;
                        const avgM=g.matched>0?Math.round(g.mSum/g.matched*100)/100:null;
                        const isOpen=mallExpandedGroup===g.baseDisc;
                        return (
                          <React.Fragment key={i}>
                            <tr>
                              <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,
                                textAlign:"left",fontWeight:700,color:D.black,whiteSpace:"nowrap"}}>
                                {g.baseDisc}% <span style={{fontSize:10,color:D.textMeta,fontWeight:400,marginLeft:4}}>할인율</span>
                              </td>
                              <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"right",whiteSpace:"nowrap"}}>{g.count}개</td>
                              <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"right",color:D.text,fontWeight:600,whiteSpace:"nowrap"}}>{avgFr}%</td>
                              <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"right",fontWeight:700,whiteSpace:"nowrap",
                                color:avgM==null?D.textMeta:(avgM>3?D.green:D.red)}}>
                                {avgM==null?<span style={{fontWeight:400,color:D.textMeta}}>공급가 미매칭</span>:`×${avgM.toFixed(2)}`}
                                {avgM!=null&&g.matched<g.count&&(
                                  <span style={{fontSize:9,color:D.textMeta,fontWeight:400,marginLeft:4}}>({g.matched}/{g.count})</span>
                                )}
                              </td>
                              <td style={{padding:"7px 10px",borderBottom:`1px solid ${D.border}`,textAlign:"center"}}>
                                <button onClick={()=>setMallExpandedGroup(isOpen?null:g.baseDisc)}
                                  style={{background:isOpen?D.black:"transparent",
                                    color:isOpen?"#fff":D.text,
                                    border:`1px solid ${D.borderMid}`,borderRadius:5,
                                    padding:"3px 9px",fontSize:11,cursor:"pointer",fontWeight:600}}>
                                  {isOpen?"▾ 닫기":`▸ 묶음 상품 ${g.count}개`}
                                </button>
                              </td>
                            </tr>
                            {isOpen&&(
                              <tr><td colSpan={5} style={{padding:"0 10px 10px",borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt}}>
                                <div style={{maxHeight:280,overflow:"auto",border:`1px solid ${D.border}`,borderRadius:4,background:D.surface}}>
                                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
                                    <thead><tr style={{background:D.surfaceAlt,color:D.textMeta}}>
                                      {["상품명","판매가","할인율","할인가","쿠폰적용가","최종할인","원가(VAT)","마진","마크업"].map((h,k)=>(
                                        <th key={k} style={{padding:"5px 8px",textAlign:k===0?"left":"right",fontWeight:600,position:"sticky",top:0,background:D.surfaceAlt,whiteSpace:"nowrap"}}>{h}</th>
                                      ))}
                                    </tr></thead>
                                    <tbody>
                                      {g.products.map((p,j)=>(
                                        <tr key={j} style={{borderTop:`1px solid ${D.border}`}}>
                                          <td title={p.name} style={{padding:"4px 8px",maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap"}}>{won(p.selling)}</td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{p.rate}%</td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.textSub}}>{won(p.discPrice)}</td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:700}}>{won(p.couponPrice)}</td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:D.red,fontWeight:600}}>{pct(p.effDisc)}</td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",color:p.supplyVat>0?D.text:D.textMeta}}>
                                            {p.supplyVat>0?won(p.supplyVat):"—"}
                                          </td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:600,
                                            color:p.supplyVat>0?((p.margin||0)>=0?D.text:D.red):D.textMeta}}>
                                            {p.supplyVat>0?won(p.margin||0):"—"}
                                          </td>
                                          <td style={{padding:"4px 8px",textAlign:"right",whiteSpace:"nowrap",fontWeight:700,
                                            color:p.supplyVat>0?((p.markup||0)>3?D.green:D.red):D.textMeta}}>
                                            {p.supplyVat>0?`×${(p.markup||0).toFixed(2)}`:"—"}
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </td></tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                  {(onCreatePromo||onAttachInlineCalc)&&(
                    <div style={{padding:"12px",borderTop:`1px solid ${D.borderMid}`,display:"flex",justifyContent:"center",background:D.surface}}>
                      <button onClick={handleCreatePromo}
                        style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
                          padding:"9px 22px",fontSize:11,cursor:"pointer",fontWeight:700}}>
                        {onAttachInlineCalc?(attachMode==="fill"?"+ 이 행 묶음으로 채우기":"+ 매트릭스에 묶음 추가"):"+ 자사몰 프로모션 추가하기"}
                      </button>
                    </div>
                  )}
                </div>
              );
            })()}
          </>)}
        </div>
      </div>
    </div>
  );
}

function Promo29CMCalcModal({ initialCoupon=10, onApply, onClose }){
  const [coupon,setCoupon]=useState(String(initialCoupon||10));
  const [stackCoupons,setStackCoupons]=useState([]); // [{rate}] — 중복 쿠폰
  const [recent,setRecent]=useState(null);
  const cpnPrimary=(()=>{const v=Number(coupon);return isNaN(v)||v<0?0:Math.min(v,60);})();
  const stackRates=stackCoupons.map(sc=>{const v=Number(sc.rate);return isNaN(v)||v<0?0:Math.min(v,60);});
  const stackFactor=stackRates.reduce((a,r)=>a*(1-r/100),1);
  const effCpn=Math.round((1-(1-cpnPrimary/100)*stackFactor)*1000)/10;

  // 쿠폰 열 — 입력 기본 쿠폰 중심 ±10% (5% 단위 5개)
  const couponCols=useMemo(()=>{
    const snap=Math.max(0,Math.min(50,Math.round(cpnPrimary/5)*5));
    let lo=snap-10, hi=snap+10;
    if(lo<0){hi+=-lo;lo=0;}
    if(hi>50){lo=Math.max(0,lo-(hi-50));hi=50;}
    const arr=[]; for(let v=lo;v<=hi&&arr.length<5;v+=5) arr.push(v);
    return arr;
  },[cpnPrimary]);
  const hlCol=couponCols.reduce((best,c)=>Math.abs(c-cpnPrimary)<Math.abs(best-cpnPrimary)?c:best,couponCols[0]);

  // (구간, 기본 쿠폰) → {raw 역산값, front 올림 적용 후, final 검증 최종}
  const compute=(slot,primary)=>{
    const factorFinal=1-slot.disc/100, factorCoupon=(1-primary/100)*stackFactor;
    if(factorCoupon<=0) return {raw:0,front:0,final:0};
    let bf=factorFinal/factorCoupon; if(bf>1) bf=1;
    const raw=Math.round((1-bf)*1000)/10;
    const front=roundUpBaseDisc(raw);
    const final=Math.round((1-(1-front/100)*(1-primary/100)*stackFactor)*1000)/10;
    return {raw,front,final};
  };

  const handleApply=(slot,primary)=>{
    const {raw,front,final}=compute(slot,primary);
    onApply({tier:slot,baseDisc:front,primaryCoupon:primary,stackRates});
    setRecent({tierName:slot.name,range:slot.range,baseDisc:front,raw,primaryCoupon:primary,
      stackRates:[...stackRates],finalDisc:final,p75:slot.disc});
    setTimeout(()=>setRecent(r=>r&&r.tierName===slot.name&&r.primaryCoupon===primary?null:r),2500);
  };

  const inNum={border:`1px solid ${D.border}`,background:D.surface,color:D.text,borderRadius:6,padding:"6px 10px",fontSize:13,width:80,fontFamily:"inherit"};

  return(
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2100,
      display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()}
        style={{background:D.surface,borderRadius:12,width:"min(820px,96vw)",maxHeight:"92vh",overflowY:"auto",
          boxShadow:"0 8px 40px rgba(0,0,0,0.22)",fontSize:12,color:D.text,
          fontFamily:"'Noto Sans KR','Pretendard',sans-serif"}}>
        <div style={{position:"sticky",top:0,background:D.surface,borderBottom:`1px solid ${D.border}`,
          padding:"14px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",zIndex:5}}>
          <b style={{fontSize:15,color:D.black,fontWeight:700}}>29CM 프로모션 계산기</b>
          <button onClick={onClose} style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
            width:30,height:30,cursor:"pointer",fontSize:14,color:D.textMeta}}>✕</button>
        </div>
        <div style={{padding:"16px 20px 24px"}}>

          {/* 1. 쿠폰율 입력 */}
          <div style={{padding:"10px 12px",background:"#eef3ff",border:`1px solid ${D.blue}`,
            borderRadius:6,marginBottom:14}}>
            <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <label style={{fontSize:13,fontWeight:700,color:D.blue}}>1. 쿠폰율 입력</label>
              <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="60" step="1" value={coupon}
                onChange={e=>setCoupon(e.target.value)} style={inNum}/>
              <span style={{fontSize:12,color:D.blue}}>% (기본)</span>
              <button onClick={()=>setStackCoupons([...stackCoupons,{rate:""}])}
                style={{background:D.blue,color:"#fff",border:"none",borderRadius:5,
                  padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:600}}>
                + 쿠폰 추가
              </button>
              {stackCoupons.length>0&&(
                <span style={{fontSize:12,color:D.blue,marginLeft:"auto",fontWeight:700}}>
                  유효 쿠폰율 {effCpn}%
                </span>
              )}
            </div>
            {stackCoupons.map((sc,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:10,marginTop:8,flexWrap:"wrap"}}>
                <span style={{fontSize:11,color:D.blue,fontWeight:600,minWidth:80}}>Case {i+1}</span>
                <input type="number" onWheel={e=>e.currentTarget.blur()} min="0" max="60" step="1" value={sc.rate}
                  onChange={e=>{const n=[...stackCoupons];n[i]={...sc,rate:e.target.value};setStackCoupons(n);}}
                  style={inNum}/>
                <span style={{fontSize:12,color:D.blue}}>%</span>
                <button onClick={()=>setStackCoupons(stackCoupons.filter((_,j)=>j!==i))}
                  style={{background:"transparent",border:`1px solid ${D.blue}55`,borderRadius:5,
                    padding:"3px 9px",fontSize:11,cursor:"pointer",color:D.blue}}>✕</button>
              </div>
            ))}
            <div style={{fontSize:10,color:D.blue,marginTop:8,opacity:.85,lineHeight:1.5}}>
              매트릭스 열은 기본 쿠폰 변동값을 보여주고, 추가 Case 쿠폰은 모든 셀에 동일 누적 적용됩니다.
            </div>
          </div>

          {/* 2. 구간 × 쿠폰율 매트릭스 — 셀=도출 프런트 할인율(P75 목표 기준 역산 + 올림) */}
          <div>
            <div style={{fontSize:12,fontWeight:700,color:D.black,marginBottom:6}}>
              2. 구간 × 쿠폰율 — 도출 프런트 할인율
              <span style={{color:D.textMeta,fontWeight:400}}> · 굵은 값=올림 규칙 적용 후 프런트 · raw=올림 전 역산값(다를 때만 노출) · → =쿠폰 적용 후 실제 최종 할인율 · ★=입력 기본 쿠폰 · 셀 클릭 시 자동 입력</span>
            </div>
            <div style={{overflowX:"auto",border:`1px solid ${D.border}`,borderRadius:6}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr>
                  <th style={{padding:"6px 8px",background:D.surfaceAlt,color:D.textSub,fontWeight:600,
                    borderBottom:`1px solid ${D.border}`,fontSize:10,textAlign:"left",whiteSpace:"nowrap"}}>
                    구간 \ 기본 쿠폰
                  </th>
                  {couponCols.map(c=>(
                    <th key={c} style={{padding:"6px 8px",background:c===hlCol?`${D.blue}14`:D.surfaceAlt,
                      color:c===hlCol?D.blue:D.textSub,fontWeight:c===hlCol?700:600,
                      borderBottom:`1px solid ${D.border}`,fontSize:10,textAlign:"center",whiteSpace:"nowrap"}}>
                      {c}%{c===hlCol?" ★":""}
                    </th>
                  ))}
                </tr></thead>
                <tbody>
                  {CALC_SLOTS.map(s=>(
                    <tr key={s.id}>
                      <td style={{padding:"7px 9px",background:s.bg,
                        color:s.color,borderBottom:`1px solid ${D.border}`,
                        fontSize:10,whiteSpace:"nowrap",lineHeight:1.4}}>
                        <div style={{fontWeight:700}}>{s.name} <span style={{fontSize:9,color:D.textMeta,fontWeight:600}}>P75 {s.disc}%</span></div>
                        <div style={{fontSize:9,color:D.textSub,fontWeight:500}}>{s.range}</div>
                      </td>
                      {couponCols.map(c=>{
                        const {raw,front,final:fin}=compute(s,c);
                        const isHL=c===hlCol;
                        const rounded=raw!==front;
                        return(
                          <td key={c} onClick={()=>handleApply(s,c)}
                            style={{padding:"7px 9px",textAlign:"center",cursor:"pointer",
                              borderBottom:`1px solid ${D.border}`,
                              background:isHL?`${s.color}14`:D.surface,
                              outline:isHL?`2px solid ${s.color}`:"none",outlineOffset:"-2px",
                              transition:"background 0.12s"}}>
                            <div style={{fontSize:13,fontWeight:700,color:s.color,lineHeight:1.2}}>{front}%</div>
                            {rounded&&(
                              <div style={{fontSize:9,color:D.textMeta,lineHeight:1.2,opacity:.8}}>raw {raw}%</div>
                            )}
                            <div style={{fontSize:9,color:D.textMeta,lineHeight:1.3}}>→ {fin}%</div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{fontSize:10,color:D.textMeta,marginTop:6,lineHeight:1.5}}>
              셀 클릭 시 상품군(구간 + 금액 범위) 할인율 + 기본 쿠폰{stackCoupons.length>0?` + 추가 Case ${stackCoupons.length}개`:""}이 자동 입력되며,
              이 창은 계속 열린 채로 다른 구간/쿠폰 조합을 추가로 선택할 수 있습니다.
            </div>
          </div>

          {recent&&(
            <div style={{marginTop:12,padding:"10px 12px",background:`${D.green}10`,border:`1px solid ${D.green}55`,
              borderRadius:6,fontSize:12,color:D.text,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
              <span style={{color:D.green,fontWeight:700}}>✓ 추가됨</span>
              <span>
                {recent.tierName} ({recent.range}) · 프런트 {recent.baseDisc}%
                {recent.raw!==undefined&&recent.raw!==recent.baseDisc&&` (raw ${recent.raw}%)`}
                {" · 쿠폰 "}{recent.primaryCoupon}%
                {recent.stackRates&&recent.stackRates.length>0&&` (+중복 ${recent.stackRates.filter(r=>r>0).join("/")}%)`}
                {" → 최종 "}{recent.finalDisc}%
              </span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// 첨부 파일 미리보기 모달 — 엑셀/CSV는 표, 이미지·PDF는 인라인, 그 외는 다운로드 안내
// 핀셋 상품 전체 보기 모달 — 카드에서 10개 초과 시 "+ N개 더보기" 로 진입
function PinnedListModal({ promo, onToggleHighlight, onClose }){
  const pins=promo?.pinned_products||[];
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2050,
      display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()}
        style={{background:D.surface,borderRadius:12,padding:"18px 20px",
          width:"min(720px,95vw)",maxHeight:"80vh",display:"flex",flexDirection:"column",
          boxShadow:"0 8px 40px rgba(0,0,0,0.22)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,marginBottom:10}}>
          <div>
            <b style={{fontSize:14,color:D.black}}>{promo?.name||"프로모션"} · 핀셋 상품</b>
            <span style={{fontSize:11,color:D.textMeta,marginLeft:8}}>총 {pins.length}개 · 클릭 시 강조 토글</span>
          </div>
          <button onClick={onClose}
            style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
              padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
        </div>
        <div style={{overflow:"auto",flex:1,minHeight:0}}>
          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {pins.map((pp,i)=>(
              <span key={i} onClick={()=>onToggleHighlight(i)}
                title={pp.memo?`${pp.name} · ${pp.memo}`:pp.name}
                style={{cursor:"pointer",borderRadius:8,padding:"3px 10px",fontSize:11,maxWidth:240,
                  overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                  background:pp.highlight?D.black:D.surfaceAlt,
                  color:pp.highlight?"#fff":D.textSub,
                  border:`1px solid ${pp.highlight?D.black:D.border}`}}>{pp.name}</span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function FilePreviewModal({ file, onClose }){
  const [aoaList,setAoaList]=useState(null);
  const [sheets,setSheets]=useState([]);
  const [active,setActive]=useState(0);
  const [err,setErr]=useState("");
  const [query,setQuery]=useState("");
  const name=file?.name||"";
  const ext=(name.split(".").pop()||"").toLowerCase();
  const type=file?.type||"";
  const isImage=type.startsWith("image/")||["png","jpg","jpeg","gif","webp","svg","bmp","avif"].includes(ext);
  const isPdf=type.includes("pdf")||ext==="pdf";
  const isSheet=["xlsx","xls","xlsm","csv"].includes(ext)||type.includes("sheet")||type.includes("excel")||type.includes("csv");
  useEffect(()=>{
    if(!file||!isSheet) return;
    let alive=true;
    (async()=>{
      try{
        const XLSX=await getXLSX();
        const b64=(file.data||"").split(",")[1]||"";
        const wb=XLSX.read(b64,{type:"base64"});
        const all=wb.SheetNames.map(nm=>XLSX.utils.sheet_to_json(wb.Sheets[nm],{header:1,defval:""}));
        if(!alive) return;
        setSheets(wb.SheetNames); setAoaList(all); setActive(0);
      }catch{ if(alive) setErr("미리보기를 불러오지 못했습니다."); }
    })();
    return()=>{alive=false;};
  },[file,isSheet]);
  const loading=isSheet&&aoaList===null&&!err;
  const MAXR=300, MAXC=40;
  const aoa=aoaList?.[active]||[];
  const headerRow=aoa[0]||[];
  const dataRows=aoa.length>0?aoa.slice(1):[];
  const q=query.trim().toLowerCase();
  const filteredData=q
    ? dataRows.filter(r=>r.some(c=>String(c??"").toLowerCase().includes(q)))
    : dataRows;
  const displayData=filteredData.slice(0,MAXR);
  const rows=aoa.length>0?[headerRow,...displayData]:[];
  const truncated=filteredData.length>MAXR;
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2100,
      display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()} style={{background:D.surface,borderRadius:12,padding:"18px 20px",
        width:"min(1000px,95vw)",
        // 엑셀/CSV 시트 미리보기에서 검색 시 행이 줄어도 모달이 흔들리지 않도록 최소 높이 확보
        ...(isSheet?{minHeight:"min(720px,85vh)"}:{}),
        maxHeight:"90vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,0.22)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,marginBottom:12}}>
          <b style={{fontSize:14,color:D.black,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={name}>📎 {name}</b>
          <div style={{display:"flex",gap:6,flexShrink:0}}>
            <a href={file.data} download={name} style={{fontSize:12,color:D.textSub,textDecoration:"none",
              border:`1px solid ${D.border}`,borderRadius:6,padding:"4px 10px"}}>다운로드</a>
            <button onClick={onClose} style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
              padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
          </div>
        </div>
        {isSheet&&sheets.length>1&&(
          <div style={{display:"flex",gap:4,marginBottom:8,flexWrap:"wrap"}}>
            {sheets.map((s,i)=>(
              <button key={i} onClick={()=>{setActive(i);setQuery("");}} style={{fontSize:11,padding:"3px 10px",borderRadius:6,cursor:"pointer",
                border:`1px solid ${i===active?D.black:D.border}`,background:i===active?D.black:D.surface,color:i===active?"#fff":D.textSub}}>{s}</button>
            ))}
          </div>
        )}
        {isSheet&&!loading&&!err&&aoa.length>0&&(
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
            <input type="search" value={query} onChange={e=>setQuery(e.target.value)}
              placeholder="셀 내용 검색 (대소문자 무시 · 헤더 제외 데이터 행)"
              style={{flex:"1 1 200px",minWidth:160,padding:"5px 10px",fontSize:12,
                border:`1px solid ${D.border}`,borderRadius:6,color:D.text,background:D.surface,
                fontFamily:"inherit"}}/>
            <span style={{fontSize:11,color:D.textMeta,whiteSpace:"nowrap"}}>
              {q
                ?`매치 ${filteredData.length.toLocaleString()}행 / 전체 ${dataRows.length.toLocaleString()}행`
                :`전체 ${dataRows.length.toLocaleString()}행`}
            </span>
          </div>
        )}
        <div style={{overflow:"auto",flex:1,minHeight:0}}>
          {isImage&&<img src={file.data} alt={name} style={{maxWidth:"100%",height:"auto",display:"block",margin:"0 auto"}}/>}
          {isPdf&&<iframe title={name} src={file.data} style={{width:"100%",height:"72vh",border:"none"}}/>}
          {isSheet&&(
            loading?<div style={{color:D.textMeta,fontSize:12,padding:30,textAlign:"center"}}>불러오는 중…</div>
            :err?<div style={{color:D.red,fontSize:12,padding:30,textAlign:"center"}}>{err}</div>
            :q&&filteredData.length===0
              ?<div style={{color:D.textMeta,fontSize:12,padding:30,textAlign:"center"}}>검색 결과가 없습니다.</div>
              :<table style={{borderCollapse:"collapse",fontSize:11}}>
              <tbody>
                {rows.map((r,ri)=>(
                  <tr key={ri}>
                    {Array.from({length:Math.min(MAXC,Math.max(1,r.length))}).map((_,ci)=>{
                      const val=r[ci]; const Tag=ri===0?"th":"td";
                      return <Tag key={ci} title={String(val??"")} style={{border:`1px solid ${D.border}`,padding:"3px 8px",
                        whiteSpace:"nowrap",maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",textAlign:"left",
                        background:ri===0?D.surfaceAlt:"transparent",fontWeight:ri===0?600:400,
                        color:ri===0?D.textSub:D.text}}>{val===""||val==null?"":String(val)}</Tag>;
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
          {!isImage&&!isPdf&&!isSheet&&(
            <div style={{color:D.textMeta,fontSize:13,padding:"40px 0",textAlign:"center"}}>
              이 형식은 미리보기를 지원하지 않습니다. 다운로드해서 확인해 주세요.
            </div>
          )}
        </div>
        {isSheet&&!loading&&!err&&truncated&&(
          <div style={{fontSize:10,color:D.textMeta,marginTop:6}}>
            ※ 처음 {MAXR}행만 표시 ({q?`매치 ${filteredData.length.toLocaleString()}행 중`:`전체 ${dataRows.length.toLocaleString()}행 중`})
          </div>
        )}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 프로모션 이익률 분석 (베타)
//   정상가 = 인벤토리 selling_price (또는 가격 DB), 원가 = 공급가 × 1.1 (부가세 포함)
//   매출  = 자사몰 주문의 payment_amount (주문 단위, 중복 제거). 취소/교환/반품 제외.
//   마진 = 결제금액 − 원가합, 마진율 = 마진 ÷ 결제금액 (결제금액 분모), 할인율 = (정상가−결제금액)/정상가
// ─────────────────────────────────────────────
// 가격 소스 훅: calc_supply_override(가격 DB, 최우선) → 최근 inventory_snapshot 폴백
function useInventoryPricing(){
  const [invMap,setInvMap]=useState({}); // n:/z:/c: → {selling,supply}
  const [ovMap,setOvMap]=useState({});   // norm_name → {selling,supply}
  const [ready,setReady]=useState(false);
  useEffect(()=>{
    let alive=true;
    (async()=>{
      try{
        const db=await getSupabase();
        const {data:latest}=await db.from("inventory_snapshot").select("snapshot_date").order("snapshot_date",{ascending:false}).limit(1);
        const d=latest?.[0]?.snapshot_date;
        if(!d) return;
        let all=[],from=0;const PAGE=1000;
        while(true){
          const {data,error}=await db.from("inventory_snapshot").select("product_name,product_code,supply_price,selling_price")
            .eq("snapshot_date",d).range(from,from+PAGE-1);
          if(error||!data||data.length===0) break;
          all=all.concat(data);
          if(data.length<PAGE) break;
          from+=PAGE;
        }
        if(!alive) return;
        const m={};
        all.forEach(r=>{
          const v={selling:r.selling_price||0,supply:r.supply_price||0};
          const n=(r.product_name||"").trim();
          const nz=normProdName(n);
          const c=(r.product_code||"").trim();
          if(n&&!m["n:"+n]) m["n:"+n]=v;
          if(nz&&!m["z:"+nz]) m["z:"+nz]=v;
          if(c&&!m["c:"+c]) m["c:"+c]=v;
        });
        setInvMap(m);
      }catch{}
      finally{ if(alive) setReady(true); }
    })();
    return()=>{alive=false;};
  },[]);
  useEffect(()=>{
    let alive=true;
    (async()=>{
      try{
        const db=await getSupabase();
        let all=[],from=0;const PAGE=1000;
        while(true){
          const {data,error}=await db.from("calc_supply_override").select("norm_name,supply_price,selling_price").range(from,from+PAGE-1);
          if(error||!data||data.length===0) break;
          all=all.concat(data);
          if(data.length<PAGE) break;
          from+=PAGE;
        }
        if(!alive) return;
        const m={};
        all.forEach(r=>{ if(r.norm_name) m[r.norm_name]={selling:r.selling_price||0,supply:r.supply_price||0}; });
        setOvMap(m);
      }catch{}
    })();
    return()=>{alive=false;};
  },[]);
  // 상품 → {selling,supply,matched} : 오버라이드(필드별) 우선, 인벤토리 폴백 + 정규화/부분일치
  const priceOf=useCallback((name,code)=>{
    const raw=(name||"").trim();
    const nz=normProdName(raw);
    let selling=0,supply=0;
    const ov=nz?ovMap[nz]:null;
    if(ov){ selling=ov.selling||0; supply=ov.supply||0; }
    const fill=v=>{ if(v){ if(!selling) selling=v.selling||0; if(!supply) supply=v.supply||0; } };
    if((!selling||!supply)&&code) fill(invMap["c:"+String(code).trim()]);
    if((!selling||!supply)&&raw)  fill(invMap["n:"+raw]);
    if((!selling||!supply)&&nz)   fill(invMap["z:"+nz]);
    if((!selling||!supply)&&nz){
      const zk=Object.keys(invMap).filter(k=>k.startsWith("z:"));
      for(const k of zk){ const kn=k.slice(2); if(kn.length>=4&&(kn.includes(nz)||nz.includes(kn))){ fill(invMap[k]); break; } }
    }
    return {selling,supply,matched:selling>0&&supply>0};
  },[invMap,ovMap]);
  return {priceOf,ready};
}

// 순수 계산 — 프로모션 기간 자사몰 주문의 건별/합계 마진·마진율 (결제금액 분모)
function computePromoProfit(orders, promo, priceOf){
  const dayMs=86400000;
  const todayStr=localDate(0), yesterdayStr=localDate(-1);
  const promoStart=String(promo.start_date||"").slice(0,10);
  const promoEndRaw=String(promo.end_date||"").slice(0,10);
  const isOngoing=promoEndRaw>=todayStr;
  const promoEnd=isOngoing?(yesterdayStr>=promoStart?yesterdayStr:promoStart):promoEndRaw;
  const startsToday=promoStart>=todayStr;
  const lenDays=(promoStart&&promoEnd)?Math.max(0,(new Date(promoEnd)-new Date(promoStart))/dayMs)+1:0;
  const period={promoStart,promoEnd,isOngoing,startsToday,lenDays};
  const emptyTotals={regular:0,actual:0,cost:0,profit:0,discRate:null,profitRate:null,orders:0,units:0};
  if(!promoStart||!promoEnd||startsToday||!orders?.length) return {rows:[],excluded:[],period,totals:emptyTotals,abnormal:0,cancelled:0};
  // 자사몰 · 기간 → 주문번호로 그룹. payment_amount 는 주문 단위.
  //   · 동일 주문번호에 취소/교환/반품 라인이 하나라도 있으면(부분취소 포함) 주문 전체를 제외.
  //     (취소 라인은 기간·번호형식과 무관하게 먼저 수집 → 해당 주문번호를 통째 제외)
  //   · 비정상 주문번호(YYYYMMDD-XXXXXXX 형식 아님)도 제외.
  const cancelledNos=new Set();
  orders.forEach(r=>{
    if((r.channel||"")!=="자사몰") return;
    if(isProfitCountable(r)) return;
    const oid=r.order_no||r.order_id; if(oid) cancelledNos.add(String(oid).trim());
  });
  const byOrder={};
  let abnormal=0; const seenAbnormal=new Set();
  orders.forEach(r=>{
    if((r.channel||"")!=="자사몰") return;
    const d=r.order_date; if(!d||d<promoStart||d>promoEnd) return;
    if(!isProfitCountable(r)) return; // 취소 라인 자체는 합산 제외(주문은 cancelledNos 로 통째 제외)
    const oid=String(r.order_no||r.order_id||"").trim(); if(!oid) return;
    if(!/^\d{8}-\d+$/.test(oid)){ if(!seenAbnormal.has(oid)){seenAbnormal.add(oid);abnormal++;} return; }
    if(!byOrder[oid]) byOrder[oid]={order_no:oid,order_date:d,payment:0,lines:[]};
    const pa=r.payment_amount||0; if(pa>byOrder[oid].payment) byOrder[oid].payment=pa;
    byOrder[oid].lines.push({name:r.product_name||"미분류",option:r.option_name||"",qty:r.qty||1});
  });
  const rows=[],excluded=[]; let cancelled=0;
  let tReg=0,tAct=0,tCost=0,tUnits=0;
  Object.values(byOrder).forEach(o=>{
    if(cancelledNos.has(o.order_no)){ cancelled++; return; } // 동일 주문번호에 취소 있음 → 통째 제외
    if(!o.lines.length) return;
    let reg=0,cost=0,units=0,missing=false;
    const lines=o.lines.map(ln=>{
      const p=priceOf(ln.name);
      const supplyVat=Math.round((p.supply||0)*1.1);
      const lineReg=(p.selling||0)*ln.qty;
      const lineCost=supplyVat*ln.qty;
      if(!p.matched) missing=true;
      reg+=lineReg; cost+=lineCost; units+=ln.qty;
      return {...ln,selling:p.selling||0,supplyVat,lineReg,lineCost,matched:p.matched};
    });
    const actual=o.payment||0;
    const rec={order_no:o.order_no,order_date:o.order_date,lines,units,regular:reg,actual,cost,
      discRate:reg>0?(reg-actual)/reg*100:null,
      profit:actual-cost,
      profitRate:actual>0?(actual-cost)/actual*100:null};
    if(missing||reg<=0){ excluded.push(rec); return; }
    rows.push(rec);
    tReg+=reg; tAct+=actual; tCost+=cost; tUnits+=units;
  });
  rows.sort((a,b)=>b.profit-a.profit);
  const totals={regular:tReg,actual:tAct,cost:tCost,profit:tAct-tCost,
    discRate:tReg>0?(tReg-tAct)/tReg*100:null,
    profitRate:tAct>0?(tAct-tCost)/tAct*100:null,
    orders:rows.length,units:tUnits};
  return {rows,excluded,period,totals,abnormal,cancelled};
}

function ProfitCalcModal({ promo, orders=[], onClose }){
  const {priceOf,ready}=useInventoryPricing();
  const {rows,excluded,period,totals,abnormal,cancelled}=useMemo(()=>computePromoProfit(orders,promo,priceOf),[orders,promo,priceOf]);
  const [expanded,setExpanded]=useState(()=>new Set());
  const [limit,setLimit]=useState(100);
  const modalCardRef=useRef(null);
  const ch=promo.platform;
  const won=n=>"₩"+Math.round(n||0).toLocaleString();
  const pct=v=>v==null?"—":`${v.toFixed(1)}%`;
  const toggle=oid=>setExpanded(s=>{const n=new Set(s);n.has(oid)?n.delete(oid):n.add(oid);return n;});
  return (
    <div onClick={onClose}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:2000,
        display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div ref={modalCardRef} onClick={e=>e.stopPropagation()}
        style={{background:D.surface,borderRadius:14,padding:"24px 28px",
          width:"min(960px,96vw)",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 8px 40px rgba(0,0,0,0.22)",
          WebkitTextSizeAdjust:"100%",textSizeAdjust:"100%"}}>
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:8,gap:10}}>
          <div>
            <div style={{fontWeight:700,fontSize:16,color:D.black}}>
              {promo.name}
              <span style={{fontSize:12,color:D.textMeta,fontWeight:500,marginLeft:6}}>· 마진율 분석</span>
              <span style={{marginLeft:8,fontSize:10,fontWeight:700,color:"#fff",background:MUTE_BLUE,
                padding:"2px 7px",borderRadius:10,verticalAlign:"middle"}}>베타</span>
            </div>
            <div style={{fontSize:11,color:D.textMeta,marginTop:5,lineHeight:1.7}}>
              <div>
                <span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:chColor(ch),verticalAlign:"middle",marginRight:5}}/>
                <b style={{color:D.text,fontWeight:600}}>{ch}</b> · 주문일 기준 · 취소/교환/반품 제외
              </div>
              <div>
                <span style={{display:"inline-block",minWidth:80,color:D.textSub,fontWeight:600}}>분석 기간</span>
                {period.promoStart} ~ {period.promoEnd} <span style={{color:D.textSub}}>({period.lenDays}일{period.isOngoing?", 시작일 ~ 어제":""})</span>
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:6,flexShrink:0}}>
            <CaptureBtn cardRef={modalCardRef} filename={`마진율_${promo.name}_${period.promoStart}_${period.promoEnd}`} DC={{border:D.border,sub:D.textMeta}}/>
            <button onClick={onClose}
              style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
          </div>
        </div>

        <div style={{fontSize:11,color:D.textSub,background:`${MUTE_BLUE}10`,border:`1px solid ${MUTE_BLUE}33`,
          borderRadius:6,padding:"7px 10px",marginBottom:12,lineHeight:1.6}}>
          베타 테스트 중 · 자사몰은 <b>주문(주문서 쿠폰 반영) 단위로 비교</b>합니다 — <b>주문 결제금액 vs 주문 상품 원가합</b>(공급가×1.1).
          마진 = 결제금액 − 원가합 · 마진율은 <b>결제금액 분모</b> 기준(실제 판매가 대비), 할인율은 정상가 대비.
        </div>

        {period.startsToday?(
          <div style={{margin:"18px 0",padding:"40px 24px",background:D.surfaceAlt,border:`1px dashed ${D.border}`,borderRadius:8,textAlign:"center"}}>
            <div style={{fontSize:15,fontWeight:700,color:D.text,marginBottom:6}}>아직 집계 전</div>
            <div style={{fontSize:12,color:D.textSub}}>오늘 시작한 프로모션입니다 · 익일부터 주문이 집계됩니다.</div>
          </div>
        ):!ready?(
          <div style={{padding:40,textAlign:"center",color:D.textMeta,fontSize:13}}>가격 데이터 불러오는 중…</div>
        ):(
          <>
            <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
              {[["결제금액 합",won(totals.actual)],["원가합(VAT 포함)",won(totals.cost)],
                ["할인율",pct(totals.discRate)]].map(([k,v])=>(
                <div key={k} style={{padding:"8px 14px",background:D.surfaceAlt,borderRadius:8}}>
                  <div style={{fontSize:10,color:D.textMeta,marginBottom:2}}>{k}</div>
                  <div style={{fontSize:15,fontWeight:700,color:D.text}}>{v}</div>
                </div>
              ))}
              <div style={{padding:"8px 14px",background:totals.profit>=0?`${D.green}12`:`${D.red}12`,borderRadius:8}}>
                <div style={{fontSize:10,color:D.textMeta,marginBottom:2}}>마진</div>
                <div style={{fontSize:15,fontWeight:800,color:totals.profit>=0?D.green:D.red}}>{won(totals.profit)}</div>
              </div>
              <div style={{padding:"8px 14px",background:(totals.profitRate||0)>=0?`${D.green}12`:`${D.red}12`,borderRadius:8}}>
                <div style={{fontSize:10,color:D.textMeta,marginBottom:2}}>마진율 (결제금액 분모)</div>
                <div style={{fontSize:15,fontWeight:800,color:(totals.profitRate||0)>=0?D.green:D.red}}>{pct(totals.profitRate)}</div>
              </div>
            </div>
            <div style={{fontSize:11,color:D.textMeta,marginBottom:14}}>
              집계 주문 {totals.orders.toLocaleString()}건 · 판매수량 {totals.units.toLocaleString()}장
              {excluded.length>0&&<span style={{color:MUTE_BLUE}}> · 가격 미등록 {excluded.length}건 제외</span>}
              {abnormal>0&&<span style={{color:MUTE_BLUE}}> · 비정상 주문번호 {abnormal}건 제외</span>}
              {cancelled>0&&<span style={{color:MUTE_BLUE}}> · 취소 포함 주문 {cancelled}건 제외</span>}
            </div>

            <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"}}>
              주문 단위 계산 — 행 클릭 시 주문 상품 구성·술식 (결제금액 vs 원가합)
            </div>
            {rows.length===0?(
              <div style={{color:D.textMeta,fontSize:12,padding:"30px 0",textAlign:"center",background:D.surfaceAlt,borderRadius:6}}>
                해당 기간·채널의 집계 가능한 주문이 없습니다.
              </div>
            ):(
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500}}>주문번호</th>
                    <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500}}>주문일</th>
                    <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>결제금액</th>
                    <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>원가합</th>
                    <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>할인율</th>
                    <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>마진</th>
                    <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>마진율</th>
                  </tr></thead>
                  <tbody>
                    {rows.slice(0,limit).map(r=>(
                      <React.Fragment key={r.order_no}>
                        <tr onClick={()=>toggle(r.order_no)} style={{borderBottom:`1px solid ${D.border}`,cursor:"pointer"}}>
                          <td style={{padding:"5px 7px",color:D.text}}>{expanded.has(r.order_no)?"▾":"▸"} {r.order_no}</td>
                          <td style={{padding:"5px 7px",color:D.textSub}}>{r.order_date}</td>
                          <td style={{padding:"5px 7px",textAlign:"right",color:D.text,fontWeight:600}}>{won(r.actual)}</td>
                          <td style={{padding:"5px 7px",textAlign:"right",color:D.textSub}}>{won(r.cost)}</td>
                          <td style={{padding:"5px 7px",textAlign:"right",color:D.textMeta}}>{pct(r.discRate)}</td>
                          <td style={{padding:"5px 7px",textAlign:"right",fontWeight:700,color:r.profit>=0?D.green:D.red}}>{won(r.profit)}</td>
                          <td style={{padding:"5px 7px",textAlign:"right",fontWeight:700,color:(r.profitRate||0)>=0?D.green:D.red}}>{pct(r.profitRate)}</td>
                        </tr>
                        {expanded.has(r.order_no)&&(
                          <tr style={{borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt}}>
                            <td colSpan={7} style={{padding:"8px 14px"}}>
                              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,marginBottom:8}}>
                                <thead><tr style={{color:D.textMeta}}>
                                  <th style={{textAlign:"left",fontWeight:500,padding:"2px 6px"}}>상품</th>
                                  <th style={{textAlign:"right",fontWeight:500,padding:"2px 6px"}}>수량</th>
                                  <th style={{textAlign:"right",fontWeight:500,padding:"2px 6px"}}>정상가</th>
                                  <th style={{textAlign:"center",fontWeight:500,padding:"2px 6px"}}>실판매액(결제)</th>
                                  <th style={{textAlign:"right",fontWeight:500,padding:"2px 6px"}}>공급가×1.1</th>
                                  <th style={{textAlign:"center",fontWeight:500,padding:"2px 6px"}}>원가합계</th>
                                </tr></thead>
                                <tbody>
                                  {r.lines.map((ln,i)=>(
                                    <tr key={i}>
                                      <td style={{padding:"2px 6px",color:D.text}}>{ln.name}{ln.option?` · ${ln.option}`:""}</td>
                                      <td style={{padding:"2px 6px",textAlign:"right",color:D.textSub}}>{ln.qty}</td>
                                      <td style={{padding:"2px 6px",textAlign:"right",color:D.textSub}}>{won(ln.selling)}</td>
                                      {i===0&&<td rowSpan={r.lines.length} style={{padding:"2px 6px",textAlign:"center",verticalAlign:"middle",color:D.text,fontWeight:700,borderLeft:`1px solid ${D.border}`,borderRight:`1px solid ${D.border}`}}>{won(r.actual)}</td>}
                                      <td style={{padding:"2px 6px",textAlign:"right",color:D.textSub}}>{won(ln.supplyVat)}</td>
                                      {i===0&&<td rowSpan={r.lines.length} style={{padding:"2px 6px",textAlign:"center",verticalAlign:"middle",color:D.text,fontWeight:700,borderLeft:`1px solid ${D.border}`}}>{won(r.cost)}</td>}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                              <div style={{fontSize:11,color:D.textSub,lineHeight:1.9}}>
                                <div>실판매액(결제) = {won(r.actual)} · 원가합 = {won(r.cost)}</div>
                                <div>할인율 = (정상가 {won(r.regular)} − 결제 {won(r.actual)}) ÷ {won(r.regular)} = <b>{pct(r.discRate)}</b></div>
                                <div>마진 = 결제금액 {won(r.actual)} − 원가합 {won(r.cost)} = <b style={{color:r.profit>=0?D.green:D.red}}>{won(r.profit)}</b></div>
                                <div>마진율 = 마진 {won(r.profit)} ÷ 결제금액 {won(r.actual)} = <b style={{color:(r.profitRate||0)>=0?D.green:D.red}}>{pct(r.profitRate)}</b></div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
                {rows.length>limit&&(
                  <div style={{textAlign:"center",marginTop:10}}>
                    <button onClick={()=>setLimit(l=>l+100)}
                      style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,padding:"5px 14px",fontSize:12,cursor:"pointer",color:D.textSub}}>
                      더 보기 ({(rows.length-limit).toLocaleString()}건 남음)
                    </button>
                  </div>
                )}
              </div>
            )}

            {excluded.length>0&&(
              <div style={{marginTop:16}}>
                <div style={{fontSize:12,fontWeight:600,color:MUTE_BLUE,marginBottom:6}}>가격 미등록 제외 ({excluded.length}건)</div>
                <div style={{fontSize:11,color:D.textMeta,lineHeight:1.7}}>
                  인벤토리/가격 DB에 정상가·공급가가 없어 제외된 주문입니다. 인벤토리 업로더의 "가격 DB" 모드로 해당 상품 가격을 등록하면 집계됩니다.
                  <div style={{marginTop:4,maxHeight:120,overflowY:"auto"}}>
                    {[...new Set(excluded.flatMap(e=>e.lines.filter(l=>!l.matched).map(l=>l.name)))].slice(0,40).map((nm,i)=>(
                      <span key={i} style={{display:"inline-block",background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:4,padding:"1px 6px",margin:"2px",fontSize:10}}>{nm}</span>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function PromoImpactModal({ promo, onClose, revenues=[], storeSales=[], orders=[] }) {
  const ch=promo.platform;
  const dayMs=86400000;
  const todayStr=localDate(0);
  const yesterdayStr=localDate(-1);
  const promoStart=String(promo.start_date||"").slice(0,10);
  const promoEndRaw=String(promo.end_date||"").slice(0,10);
  // 종료일이 미래거나 당일(오늘) → 오늘 데이터는 미완성이므로 분석 종료일을 전일(어제)로 클램프
  //   (직전 동기간도 lenDays 기준이라 같은 만큼 하루 당겨짐) · 종료 다음날부터는 전체 기간 집계
  //   - 어제가 시작일보다 이르면(시작 당일) 시작일로 클램프
  const isOngoing=promoEndRaw>=todayStr;
  // 오늘 시작(혹은 미래 시작) — 비교/그래프가 의미 없으므로 "아직 집계 전" 안내로 대체
  const startsToday=promoStart>=todayStr;
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
      .sort((a,b)=>b.qty-a.qty||b.orders-a.orders).slice(0,20);
  },[ch,promoStart,promoEnd,orders]);

  // 핀셋 상품 — 프로모션 전/후 주문 수량 비교 (해당 채널 한정, order_date 기준, 모든 상태 = 주문 수량)
  const pinned=promo.pinned_products||[];
  const pinnedNames=useMemo(()=>new Set(pinned.map(p=>p.name)),[pinned]);
  const pinnedComparison=useMemo(()=>{
    if(!pinned.length) return [];
    const OFFLINE=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
    const matchesCh=r=> ch==="오프라인 스토어" ? OFFLINE.has(r.channel||"") : (r.channel||"")===ch;
    const map={};
    pinned.forEach(p=>{map[p.name]={name:p.name,memo:p.memo||"",prev:0,promo:0};});
    orders.forEach(r=>{
      const nm=r.product_name||"";
      if(!map[nm]||!matchesCh(r)) return;
      const d= ch==="오프라인 스토어" ? (r.sale_date||r.order_date) : r.order_date;
      if(!d) return;
      const q=r.qty||1;
      if(d>=promoStart&&d<=promoEnd) map[nm].promo+=q;
      else if(d>=prevStart&&d<=prevEnd) map[nm].prev+=q;
    });
    return pinned.map(p=>{
      const o=map[p.name];
      const chg=o.prev>0?((o.promo-o.prev)/o.prev*100):null;
      return {...o,chg};
    });
  },[pinned,ch,orders,promoStart,promoEnd,prevStart,prevEnd]);

  const {prevTotal,promoTotal,chg}=promoRevenueChg(promo,revenues,storeSales);

  // 주문 장수 — 주문·배송 데이터 소스(orders)에서 채널·기간 일치 + order_id 유니크 집계
  const {prevOrders,promoOrders,orderChg}=useMemo(()=>{
    if(!orders||!orders.length||!promoStart||!promoEnd) return {prevOrders:0,promoOrders:0,orderChg:null};
    const OFFLINE=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
    const matchesCh=r=> ch==="오프라인 스토어" ? OFFLINE.has(r.channel||"") : (r.channel||"")===ch;
    const prevSeen=new Set(); let prevExtra=0;
    const promoSeen=new Set(); let promoExtra=0;
    orders.forEach(r=>{
      if(!matchesCh(r)) return;
      const d= ch==="오프라인 스토어" ? (r.sale_date||r.order_date) : r.order_date;
      if(!d) return;
      const oid=r.order_no||r.order_id;
      if(d>=promoStart&&d<=promoEnd){
        if(oid) promoSeen.add(oid); else promoExtra++;
      } else if(d>=prevStart&&d<=prevEnd){
        if(oid) prevSeen.add(oid); else prevExtra++;
      }
    });
    const prevCount=prevSeen.size+prevExtra;
    const promoCount=promoSeen.size+promoExtra;
    const oc=prevCount>0?((promoCount-prevCount)/prevCount*100):null;
    return {prevOrders:prevCount,promoOrders:promoCount,orderChg:oc};
  },[ch,orders,promoStart,promoEnd,prevStart,prevEnd]);

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
          {!isOngoing&&(
            <span style={{marginLeft:"auto",alignSelf:"center",flexShrink:0,fontSize:12,fontWeight:700,color:D.red,background:`${D.red}14`,border:`1px solid ${D.red}40`,borderRadius:999,padding:"3px 11px",whiteSpace:"nowrap"}}>종료된 프로모션</span>
          )}
          <div style={{display:"flex",gap:6,flexShrink:0}}>
            <CaptureBtn cardRef={modalCardRef} filename={`임팩트분석_${promo.name}_${promoStart}_${promoEnd}`} DC={{border:D.border,sub:D.textMeta}}/>
            <button onClick={onClose}
              style={{background:"none",border:`1px solid ${D.border}`,borderRadius:6,
                padding:"4px 10px",fontSize:12,cursor:"pointer",color:D.textMeta}}>✕ 닫기</button>
          </div>
        </div>

        {startsToday?(
          <div style={{margin:"18px 0 22px",padding:"40px 24px",background:D.surfaceAlt,
            border:`1px dashed ${D.border}`,borderRadius:8,textAlign:"center"}}>
            <div style={{fontSize:15,fontWeight:700,color:D.text,marginBottom:6}}>아직 집계 전</div>
            <div style={{fontSize:12,color:D.textSub,lineHeight:1.7}}>
              오늘 시작한 프로모션입니다 · 익일부터 일자별 매출이 집계되어 비교 그래프가 표시됩니다.
            </div>
          </div>
        ):(
          <>
            {/* 매출 요약 */}
            <div style={{display:"flex",flexDirection:"column",gap:8,marginTop:14,marginBottom:10,fontSize:12}}>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <div style={{padding:"7px 12px",background:D.surfaceAlt,borderRadius:6}}>
                  <span style={{color:D.textMeta}}>직전 매출</span> <b style={{marginLeft:6}}>{fmtWonShort(prevTotal)}</b>
                </div>
                <div style={{padding:"7px 12px",background:D.surfaceAlt,borderRadius:6}}>
                  <span style={{color:D.textMeta}}>프로모션 매출</span> <b style={{marginLeft:6}}>{fmtWonShort(promoTotal)}</b>
                </div>
                {chg!==null&&(
                  <div style={{padding:"7px 12px",background:chg>=0?`${D.green}12`:`${D.red}12`,borderRadius:6,color:chg>=0?D.green:D.red}}>
                    <span>매출 증감</span> <b style={{marginLeft:6}}>{chg>=0?"+":""}{chg.toFixed(1)}%</b>
                  </div>
                )}
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <div style={{padding:"7px 12px",background:D.surfaceAlt,borderRadius:6}} title="주문·배송 데이터 기준 · 동일 채널 · order_id 유니크 집계">
                  <span style={{color:D.textMeta}}>직전 주문 장수</span> <b style={{marginLeft:6}}>{prevOrders.toLocaleString()}건</b>
                </div>
                <div style={{padding:"7px 12px",background:D.surfaceAlt,borderRadius:6}} title="주문·배송 데이터 기준 · 동일 채널 · order_id 유니크 집계">
                  <span style={{color:D.textMeta}}>프로모션 주문 장수</span> <b style={{marginLeft:6}}>{promoOrders.toLocaleString()}건</b>
                </div>
                {orderChg!==null&&(
                  <div style={{padding:"7px 12px",background:orderChg>=0?`${D.green}12`:`${D.red}12`,borderRadius:6,color:orderChg>=0?D.green:D.red}}>
                    <span>주문 증감</span> <b style={{marginLeft:6}}>{orderChg>=0?"+":""}{orderChg.toFixed(1)}%</b>
                  </div>
                )}
              </div>
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
          </>
        )}

        {/* 할인율 매트릭스 — 그래프 아래 (프로모션 카드와 동일 뷰 재사용, 뱃지는 생략) */}
        {computeDiscountMatrix(promo.discount_plan||{}).hasGroup&&(
          <div style={{marginBottom:18}}>
            <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"}}>
              할인율 매트릭스
            </div>
            <DiscountPlanView plan={promo.discount_plan} marks={promo.discount_marks||{}} compact={false} showBadges={false}/>
          </div>
        )}

        {/* 행 호버 강조 — 핀셋 비교 + Top 20 공용 */}
        <style>{`.impact-hoverable tbody tr{transition:background 0.12s;}.impact-hoverable tbody tr:hover{background:${D.surfaceAlt};}`}</style>

        {/* 핀셋 상품 — 전/후 비교 (핀셋 상품 있을 때만 노출) */}
        {pinned.length>0&&(
          <div style={{marginBottom:18}}>
            <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"}}>
              핀셋 상품 — 프로모션 전/후 주문량 비교 ({ch}, 주문 수량 기준)
            </div>
            <div style={{overflowX:"auto"}}>
              <table className="impact-hoverable" style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                  <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500}}>상품명</th>
                  <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>직전 ({lenDays}일)</th>
                  <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>프로모션 ({lenDays}일)</th>
                  <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>증감</th>
                </tr></thead>
                <tbody>
                  {pinnedComparison.map((p,i)=>(
                    <tr key={p.name+i} style={{borderBottom:`1px solid ${D.border}`}}>
                      <td style={{padding:"5px 7px",color:D.text,maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={p.memo?`${p.name} · ${p.memo}`:p.name}>{p.name}</td>
                      <td style={{padding:"5px 7px",textAlign:"right",color:D.textSub}}>{p.prev.toLocaleString()}장</td>
                      <td style={{padding:"5px 7px",textAlign:"right",color:D.text,fontWeight:600}}>{p.promo.toLocaleString()}장</td>
                      <td style={{padding:"5px 7px",textAlign:"right",fontWeight:700,
                        color:p.chg==null?D.textMeta:p.chg>=0?D.green:D.red}}>
                        {p.chg==null?(p.promo>0?"신규":"—"):`${p.chg>=0?"+":""}${p.chg.toFixed(1)}%`}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{fontSize:10,color:D.textMeta,marginTop:5}}>
              주문일 기준 · 프로모션 채널({ch}) 한정 · 주문 수량(장) 합계 · 직전 동일기간 대비
            </div>
          </div>
        )}

        {/* Top 20 */}
        <div>
          <div style={{fontSize:12,fontWeight:600,color:D.textSub,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"}}>
            프로모션 기간 판매 Top 20 ({ch}, 주문일 기준)
          </div>
          {top20.length===0?(
            <div style={{color:D.textMeta,fontSize:12,padding:"30px 0",textAlign:"center",background:D.surfaceAlt,borderRadius:6}}>
              해당 기간·채널의 배송 데이터가 없습니다.
            </div>
          ):(
            <div style={{overflowX:"auto"}}>
              <table className="impact-hoverable" style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                  <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500,width:30}}>#</th>
                  <th style={{padding:"5px 7px",textAlign:"left",fontWeight:500}}>상품명</th>
                  <th style={{padding:"5px 7px",textAlign:"right",fontWeight:500}}>판매 수량(장)</th>
                </tr></thead>
                <tbody>
                  {top20.map((p,i)=>(
                    <tr key={p.name+i} style={{borderBottom:`1px solid ${D.border}`}}>
                      <td style={{padding:"5px 7px",color:D.textMeta}}>{i+1}</td>
                      <td style={{padding:"5px 7px",color:D.text,maxWidth:380}} title={p.name}>
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",minWidth:0}}>{p.name}</span>
                          {pinnedNames.has(p.name)&&<span style={{flexShrink:0,fontSize:9,fontWeight:700,color:D.blue,
                            background:`${D.blue}14`,border:`1px solid ${D.blue}55`,borderRadius:4,padding:"1px 5px",whiteSpace:"nowrap"}}>핀셋 상품</span>}
                        </div>
                      </td>
                      <td style={{padding:"5px 7px",textAlign:"right",color:D.text,fontWeight:600}}>{p.qty.toLocaleString()}</td>
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
  const today=localDate(0);
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
  const today=localDate(0);
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
  const today=localDate(0);
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
            // ── 상태 분류 (파싱 집계 규칙) ──────────────────────────
            //   주문 KPI    = 모든 행 (상태 무관)
            //   배송(총 출고) = 상태 컬럼 = "배송"  (CS 무관)
            //   상태=배송 + CS 키워드: 교환→"교환", 취소→"반품" (둘 다 '배송'에 포함되는 부분집합)
            //   상태≠배송(접수/송장 등): CS 취소→"취소"(판매 Top 제외), 그 외→"주문"  (미출고)
            const csN=csRaw.toLowerCase().replace(/\s/g,"");
            let status;
            if(statusCol&&statusRaw){
              if(statusRaw==="배송"){
                status = csN.includes("교환") ? "교환"
                       : csN.includes("취소") ? "반품"
                       : "배송";
              } else {
                status = csN.includes("취소") ? "취소" : "주문";
              }
            } else {
              // 상태 컬럼 없는 파일(구 형식) — CS 기반 호환 분류
              status = csRaw ? normCS(csRaw) : "배송";
            }
            // CORD prefix = 29CM 취소 주문 → 반품으로 강제
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
  const [uploadDates,setUploadDates]=useState([]);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [conflictCount,setConflictCount]=useState(0);

  // 파일 파싱 완료 후 기존 데이터 수 조회 (업로드 파일에 포함된 날짜만 대상)
  useEffect(()=>{
    if(!preview||!uploadDates.length){setConflictCount(0);return;}
    (async()=>{
      const db=await getSupabase();
      let total=0;
      for(let i=0;i<uploadDates.length;i+=100){
        const batch=uploadDates.slice(i,i+100);
        const{count}=await db.from("store_sales").select("*",{count:"exact",head:true})
          .in("sale_date",batch);
        total+=count||0;
      }
      setConflictCount(total);
    })();
  },[preview,uploadDates]);

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
          setUploadDates(dates);
          setPreview(rows); setStep(1);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e?.message||String(e)}));
  },[]);

  const handleUpload=async()=>{
    if(!preview?.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    // 업로드 파일에 포함된 날짜만 삭제 (날짜 범위 전체가 아님 — 사이의 미포함 날짜 데이터 보존)
    for(let i=0;i<uploadDates.length;i+=100){
      const batch=uploadDates.slice(i,i+100);
      const{error:de}=await db.from("store_sales").delete().in("sale_date",batch);
      if(de){setResult({type:"error",msg:"기존 데이터 삭제 실패: "+de.message});setLoading(false);return;}
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

  const reset=()=>{setStep(0);setPreview(null);setFileName("");setResult(null);setConflictCount(0);setUploadDates([]);};

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
                  업로드 파일에 포함된 {uploadDates.length}개 날짜({uploadDates.length<=3?uploadDates.join(", "):`${uploadDates[0]} 외 ${uploadDates.length-1}개`})의 기존 데이터만 삭제되고 새 데이터로 교체됩니다. 그 외 날짜의 데이터는 보존됩니다.
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

// ─────────────────────────────────────────────
// DATA INPUT — Claude 자동화 스크립트 파일함 (업로드 / 수정 / 다운로드)
// ─────────────────────────────────────────────
const _scriptId=()=>(globalThis.crypto?.randomUUID?.()||(Date.now().toString(36)+Math.random().toString(36).slice(2)));
const _fmtBytes=n=>{const b=Number(n)||0;return b<1024?`${b} B`:b<1024*1024?`${(b/1024).toFixed(1)} KB`:`${(b/1048576).toFixed(1)} MB`;};

function ClaudeScriptPanel(){
  const KEY="claude_scripts";
  const [files,setFiles]=useState([]);
  const [msg,setMsg]=useState(null);
  const [busy,setBusy]=useState(false);
  const [editing,setEditing]=useState(null); // {id,name,content,isNew} | null
  const uploadRef=useRef(null);

  const cacheLocal=arr=>localStorage.setItem(KEY,JSON.stringify(arr));
  const sortFiles=arr=>[...arr].sort((a,b)=>String(b.updated_at||"").localeCompare(String(a.updated_at||"")));

  useEffect(()=>{
    let alive=true;
    (async()=>{
      let local=[];
      try{local=JSON.parse(localStorage.getItem(KEY)||"[]");}catch{/* ignore */}
      if(alive&&local.length) setFiles(sortFiles(local));
      try{
        const db=await getSupabase();
        const{data,error}=await db.from("claude_scripts").select("*");
        if(error||!data) return; // 오프라인/미연결 → 로컬 유지
        if(!alive) return;
        if(data.length>0){
          const rows=sortFiles(data);
          setFiles(rows);cacheLocal(rows);
        } else if(local.length>0){
          await db.from("claude_scripts").upsert(local,{onConflict:"id"}); // 1회 마이그레이션
        }
      }catch{/* 로컬 유지 */}
    })();
    return()=>{alive=false;};
  },[]);

  const persist=async row=>{
    let saved={...row,updated_at:new Date().toISOString()};
    try{
      const db=await getSupabase();
      const{data,error}=await db.from("claude_scripts").upsert({id:row.id,name:row.name,content:row.content},{onConflict:"id"}).select().maybeSingle();
      if(error) throw error;
      if(data) saved={...saved,...data};
      setMsg({type:"success",msg:`"${saved.name}" 저장 완료 — Supabase 동기화.`});
    }catch{
      setMsg({type:"warn",msg:`"${saved.name}" 로컬에만 저장됨 (Supabase 연결 실패).`});
    }
    setFiles(prev=>{const next=sortFiles([...prev.filter(f=>f.id!==saved.id),saved]);cacheLocal(next);return next;});
    return saved;
  };

  const handleUpload=async e=>{
    const list=[...(e.target.files||[])];
    e.target.value="";
    if(!list.length) return;
    setBusy(true);
    for(const f of list){
      try{
        const text=await f.text();
        await persist({id:_scriptId(),name:f.name,content:text});
      }catch{setMsg({type:"error",msg:`"${f.name}" 읽기 실패.`});}
    }
    setBusy(false);
  };

  const download=f=>{
    const blob=new Blob([f.content||""],{type:"text/plain;charset=utf-8"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;a.download=f.name||"claude-script.txt";
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const del=async f=>{
    if(!window.confirm(`"${f.name}" 파일을 삭제할까요?`)) return;
    setFiles(prev=>{const next=prev.filter(x=>x.id!==f.id);cacheLocal(next);return next;});
    if(editing?.id===f.id) setEditing(null);
    try{const db=await getSupabase();await db.from("claude_scripts").delete().eq("id",f.id);}catch{/* 로컬 반영됨 */}
    setMsg({type:"info",msg:`"${f.name}" 삭제됨.`});
  };

  const saveEdit=async()=>{
    if(!editing) return;
    if(!editing.name.trim()){setMsg({type:"error",msg:"파일명을 입력하세요."});return;}
    setBusy(true);
    await persist({id:editing.id,name:editing.name.trim(),content:editing.content});
    setBusy(false);
    setEditing(null);
  };

  const linkBtn={background:"transparent",border:"none",cursor:"pointer",fontSize:12,padding:"4px 6px",borderRadius:5};

  return (
    <Card>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <span style={{fontWeight:600,fontSize:13,color:D.black}}>Claude 자동화 스크립트 파일함</span>
        <span style={{fontSize:11,color:D.textMeta,flex:1}}>업로드한 파일은 Supabase에 저장되어 어디서든 다운로드할 수 있습니다.</span>
        <input ref={uploadRef} type="file" multiple onChange={handleUpload} style={{display:"none"}}
          accept=".txt,.md,.py,.js,.ts,.sh,.json,.yaml,.yml,.toml,.csv,.html,.css,text/*"/>
        <Btn variant="ghost" onClick={()=>uploadRef.current?.click()} disabled={busy}>업로드</Btn>
        <Btn variant="primary" onClick={()=>setEditing({id:_scriptId(),name:"",content:"",isNew:true})}>+ 새 파일</Btn>
      </div>

      {files.length===0?(
        <div style={{padding:"40px 0",textAlign:"center",color:D.textMeta,fontSize:13,
          border:`1px dashed ${D.border}`,borderRadius:8}}>
          저장된 파일이 없습니다. 파일을 업로드하거나 새로 만드세요.
        </div>
      ):(
        <div style={{border:`1px solid ${D.border}`,borderRadius:8,overflow:"hidden"}}>
          {files.map((f,i)=>(
            <div key={f.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",
              borderTop:i?`1px solid ${D.border}`:"none",background:D.surface}}>
              <span style={{fontSize:14}}>📄</span>
              <button onClick={()=>setEditing({...f,isNew:false})} title="수정"
                style={{flex:1,minWidth:0,textAlign:"left",background:"transparent",border:"none",cursor:"pointer",
                  fontSize:13,fontWeight:500,color:D.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                {f.name}
              </button>
              <span style={{fontSize:11,color:D.textMeta,whiteSpace:"nowrap"}}>
                {_fmtBytes(new Blob([f.content||""]).size)}
                {f.updated_at?` · ${dayjs(f.updated_at).format("YY.MM.DD HH:mm")}`:""}
              </span>
              <button onClick={()=>download(f)} style={{...linkBtn,color:D.blue,fontWeight:600}}>다운로드</button>
              <button onClick={()=>del(f)} style={{...linkBtn,color:D.red}}>삭제</button>
            </div>
          ))}
        </div>
      )}

      {editing&&(
        <div style={{marginTop:14,border:`1px solid ${D.borderMid}`,borderRadius:8,padding:14,background:D.surfaceAlt}}>
          <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:10,flexWrap:"wrap"}}>
            <input value={editing.name} onChange={e=>setEditing(s=>({...s,name:e.target.value}))}
              placeholder="파일명 (예: automation.py)" autoFocus
              style={{flex:"1 1 240px",minWidth:200,border:`1px solid ${D.border}`,borderRadius:7,
                padding:"8px 12px",fontSize:13,color:D.text,background:D.surface,outline:"none"}}/>
            <Btn variant="primary" onClick={saveEdit} disabled={busy}>{busy?"저장 중…":"저장"}</Btn>
            <Btn variant="ghost" onClick={()=>setEditing(null)}>취소</Btn>
          </div>
          <textarea value={editing.content} onChange={e=>setEditing(s=>({...s,content:e.target.value}))} spellCheck={false}
            placeholder="스크립트 내용을 입력하거나 파일을 업로드하세요."
            style={{width:"100%",minHeight:300,boxSizing:"border-box",resize:"vertical",
              border:`1px solid ${D.border}`,borderRadius:8,padding:"12px 14px",
              fontSize:12.5,lineHeight:1.6,color:D.text,background:D.surface,outline:"none",
              fontFamily:"'SFMono-Regular',Consolas,'Liberation Mono',Menlo,monospace",
              tabSize:2,whiteSpace:"pre"}}/>
        </div>
      )}

      <Alert type={msg?.type} msg={msg?.msg}/>
    </Card>
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
    {key:"inventory",name:"인벤토리"},
    {key:"cs",name:"CS"},
    {key:"script",name:"자동화 스크립트"},
    {key:"delete",name:"데이터 삭제"},
  ];

  const GUIDES={
    revenue:"KPI 카드의 매출, 매출 점유율, 판매처별 매출의 소스입니다.\n매출 금액은 취소/환불이 포함된 금액이며, 엑셀 다운로드 시 각 채널 어드민의 통계에서 확인하세요.\n*매일 전날의 데이터를 업로드하세요.",
    stock:"KPI 카드의 입고 수량, 물류 플로우 섹션 전체의 데이터 소스입니다.\n*매일 전날의 데이터를 업로드하세요.",
    orders:"KPI 카드의 배송·반품 수, 판매처 상세의 배송·반품 수, 판매·반품 TOP, 플랫폼 별 선호·반품 옵션 랭킹, 객단가 계산의 데이터 소스입니다.\n필요 컬럼: 주문번호 · 주문일 · 배송일 · 판매처 · 상품명 · 옵션 · 수량 · 판매가(29CM·무신사 AOV) · 결제금액(자사몰 AOV) · CS처리\n*매일 최근 한달 데이터(주문건 반품 정보 업데이트)를 업로드하세요.",
    store:"KPI 카드의 매출(오프라인 스토어) 합산, 랭크 지표 내 오프라인 스토어 항목의 데이터 소스입니다.\n*매일 최근 한달의 데이터를 업로드하세요.",
    inventory:"데이터 컴페어(SKU Risk · Aging Trend)와 리오더 계산기의 공통 데이터 소스입니다.\n엑셀에 가용재고 · 입고대기 · 1주발주합계 · 4주발주합계 컬럼이 있으면 리오더 데이터가 자동 계산됩니다.",
    cs:"반품 랭크 상품의 주요 반품 사유 데이터 소스로 매칭됩니다.\n*매일 전날 데이터를 업로드하세요.",
  };
  const DC_LIGHT={bg:"#f8f8f6",card:"#ffffff",border:"#e0e0da",text:"#111111",sub:"#444444",dim:"#888888"};

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
      {tab==="inventory"&&(
        <div style={{background:DC_LIGHT.card,border:`1px solid ${DC_LIGHT.border}`,borderRadius:12,padding:"20px 20px 24px"}}>
          <div style={{fontWeight:600,fontSize:16,color:DC_LIGHT.text,letterSpacing:"-0.2px",marginBottom:16}}>Inventory 업로더</div>
          <InventoryUploader DC={DC_LIGHT} onUploaded={()=>onDataChange?.()} onReorderDone={()=>onDataChange?.()}/>
        </div>
      )}
      {tab==="cs"&&<CSDataInput/>}
      {tab==="script"&&<ClaudeScriptPanel/>}
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
      // 데이터날짜 컬럼은 더 이상 필수 아님 — 업로드 시점에 사용자가 직접 선택. 파일에 있으면 기본값으로 활용.
      const FIELD_LABELS={product_name:"상품명",current_stock_qty:"현재고",first_inbound_date:"처음입고일"};
      const required=["product_name","current_stock_qty","first_inbound_date"];
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
      }).filter(r=>r.product_name&&r.first_inbound_date);
      onResult(rows);
    }catch(err){onError(uploadErrParse(String(err?.message||err)));}
  };
  reader.onerror=()=>onError(uploadErrParse("파일 읽기 도중 시스템 오류가 발생했습니다"));
  reader.readAsArrayBuffer(file);
}

// 가격 데이터베이스용 파서 — 날짜/재고 없이 상품명·판매가(정상가)·공급가만 파싱.
// calc_supply_override(norm_name 유니크)에 업서트되어 이익률 계산·할인율 계산기의 가격 소스로 쓰인다.
function parsePriceDbFile(file,onResult,onError){
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
      if(colMap.product_name===undefined||(colMap.selling_price===undefined&&colMap.supply_price===undefined)){
        onError(uploadErrColumns({
          missing:[colMap.product_name===undefined?"상품명":null,(colMap.selling_price===undefined&&colMap.supply_price===undefined)?"판매가 또는 공급가":null].filter(Boolean),
          required:["상품명","판매가","공급가"],
          headers,
        }));
        return;
      }
      const get=(r,f)=>colMap[f]!==undefined?String(r[colMap[f]]||"").trim():"";
      const getWon=(r,f)=>colMap[f]!==undefined?Math.max(0,Math.round(parseFloat(String(r[colMap[f]]||"0").replace(/,/g,""))||0)):0;
      const seen=new Set();const rows=[];
      raw.slice(1).forEach(r=>{
        const name=get(r,"product_name");if(!name) return;
        const selling=getWon(r,"selling_price");
        const supply=getWon(r,"supply_price");
        if(selling<=0&&supply<=0) return;
        const nz=normProdName(name);if(!nz||seen.has(nz)) return;seen.add(nz);
        rows.push({product_name:name,norm_name:nz,selling_price:selling,supply_price:supply,updated_at:new Date().toISOString()});
      });
      if(rows.length===0){onError("매칭 가능한 행이 없습니다 — 판매가 또는 공급가가 0보다 큰 행이 있는지 확인하세요");return;}
      onResult(rows);
    }catch(err){onError(uploadErrParse(String(err?.message||err)));}
  };
  reader.onerror=()=>onError(uploadErrParse("파일 읽기 도중 시스템 오류가 발생했습니다"));
  reader.readAsArrayBuffer(file);
}

// 한글 CSV 디코딩 — UTF-8(BOM) 우선, 깨지면 EUC-KR(CP949)로 폴백. 카페24 내보내기는 보통 EUC-KR.
function decodeKoreanBytes(u8){
  if(u8.length>=3&&u8[0]===0xEF&&u8[1]===0xBB&&u8[2]===0xBF){
    try{ return new TextDecoder("utf-8").decode(u8.subarray(3)); }catch{/* fall through */}
  }
  let utf8="";
  try{ utf8=new TextDecoder("utf-8",{fatal:false}).decode(u8); }catch{ utf8=""; }
  if(!utf8||utf8.includes("�")){
    for(const enc of ["euc-kr","cp949","windows-949"]){
      try{ const t=new TextDecoder(enc).decode(u8); if(t&&!t.includes("�")) return t; }catch{/* try next */}
    }
  }
  return utf8;
}

// 카페24 상품코드 파일 파서 — 상품명 + 상품코드만 추출 (날짜·가격 불필요).
// cafe24_product_codes(norm_name 유니크)에 업서트되어 SKU Risk 다운로드의 카페24 코드 매칭에 쓰인다.
// 카페24 CSV는 EUC-KR·탭 구분이 많아, CSV는 인코딩 폴백+구분자 자동감지(PapaParse)로 처리한다.
function parseCafe24CodeFile(file,onResult,onError){
  const buildRows=raw=>{
    if(!raw||raw.length<2) return {err:uploadErrParse("파일에 데이터 행이 없습니다 (헤더 행만 있거나 비어있음)")};
    const headers=raw[0].map(h=>String(h||"").replace(/^﻿/,"").trim());
    const colMap=mapInvCols(headers);
    if(colMap.product_name===undefined||colMap.product_code===undefined){
      return {err:uploadErrColumns({
        missing:[colMap.product_name===undefined?"상품명":null,colMap.product_code===undefined?"상품코드":null].filter(Boolean),
        required:["상품명","상품코드"],
        headers,
      })};
    }
    const get=(r,f)=>colMap[f]!==undefined?String(r[colMap[f]]??"").trim():"";
    const seen=new Set();const rows=[];
    raw.slice(1).forEach(r=>{
      const name=get(r,"product_name");if(!name) return;
      const code=get(r,"product_code");if(!code) return;
      const nz=normCafe24Name(name);if(!nz||seen.has(nz)) return;seen.add(nz);
      rows.push({product_name:name,norm_name:nz,product_code:code,updated_at:new Date().toISOString()});
    });
    if(rows.length===0) return {err:"매칭 가능한 행이 없습니다 — 상품명·상품코드가 모두 있는 행이 있는지 확인하세요"};
    return {rows};
  };
  const finish=raw=>{ const {rows,err}=buildRows(raw); if(err) onError(err); else onResult(rows); };
  const ext=(file.name||"").toLowerCase();
  const reader=new FileReader();
  if(ext.endsWith(".csv")||ext.endsWith(".tsv")||ext.endsWith(".txt")){
    // CSV/TSV: 인코딩 폴백 후 PapaParse 구분자 자동감지(쉼표/탭 등)
    reader.onload=async e=>{
      try{
        const Papa=await getPapa();
        const text=decodeKoreanBytes(new Uint8Array(e.target.result));
        Papa.parse(text,{header:false,skipEmptyLines:true,
          complete:res=>finish(res.data),
          error:err=>onError(uploadErrParse(String(err?.message||err)))});
      }catch(err){onError(uploadErrParse(String(err?.message||err)));}
    };
  }else{
    reader.onload=async e=>{
      try{
        const XLSX=await getXLSX();
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        finish(XLSX.utils.sheet_to_json(ws,{header:1,raw:false}));
      }catch(err){onError(uploadErrParse(String(err?.message||err)));}
    };
  }
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
  const [dateModalOpen,setDateModalOpen]=useState(false);
  const [histFilter,setHistFilter]=useState("");
  const [selDates,setSelDates]=useState(new Set());
  const [delConfirm,setDelConfirm]=useState(false);
  const [dragOver,setDragOver]=useState(false);
  const [uploadMode,setUploadMode]=useState("snapshot"); // "snapshot"(날짜별 스냅샷) | "priceDb"(가격 데이터베이스용) | "cafe24"(카페24 상품코드)
  const [priceRows,setPriceRows]=useState([]); // 가격 DB 모드 파싱 결과
  const [cafe24Rows,setCafe24Rows]=useState([]); // 카페24 상품코드 모드 파싱 결과

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
    setFile(f);setUploadStatus("parsing");setStatusMsg("파일 파싱 중...");setParsedRows([]);setPriceRows([]);setCafe24Rows([]);setSnapDate(null);
    if(uploadMode==="priceDb"){
      parsePriceDbFile(f,rows=>{
        setUploadStatus(null);setStatusMsg("");
        setPriceRows(rows);
      },err=>{setUploadStatus("error");setStatusMsg(err);});
      return;
    }
    if(uploadMode==="cafe24"){
      parseCafe24CodeFile(f,rows=>{
        setUploadStatus(null);setStatusMsg("");
        setCafe24Rows(rows);
      },err=>{setUploadStatus("error");setStatusMsg(err);});
      return;
    }
    parseInvFile(f,parsed=>{
      setUploadStatus(null);setStatusMsg("");
      if(!parsed.length){setUploadStatus("error");setStatusMsg("유효한 데이터 행이 없습니다");return;}
      // 파일에 데이터날짜 컬럼이 있으면 기본값으로, 없으면 오늘 날짜를 기본값으로
      const dates=[...new Set(parsed.map(r=>r.snapshot_date).filter(Boolean))].sort();
      const today=localDate(0);
      setSnapDate(dates[dates.length-1]||today);
      setParsedRows(parsed);
    },err=>{setUploadStatus("error");setStatusMsg(err);});
  },[uploadMode]);

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
      // 모든 행의 snapshot_date 를 사용자가 선택한 snapDate 로 통일 (파일 컬럼은 더 이상 신뢰하지 않음)
      const invRows=parsedRows.map(({_r_avail,_r_incoming,_r_weekly,_r_monthly,...rest})=>({...rest,snapshot_date:snapDate}));
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

  // 가격 DB 업서트 — calc_supply_override (norm_name 유니크) 에 정상가/공급가 갱신
  const doPriceDbUpload=useCallback(async()=>{
    if(!priceRows.length) return;
    setUploadStatus("uploading");setStatusMsg("가격 DB 저장 중...");
    try{
      const db=await getSupabase();
      const CHUNK=500;
      for(let i=0;i<priceRows.length;i+=CHUNK){
        const{error}=await db.from("calc_supply_override").upsert(priceRows.slice(i,i+CHUNK),{onConflict:"norm_name"});
        if(error) throw new Error(error.message);
      }
      setUploadStatus("done");setStatusMsg(`가격 DB ${priceRows.length.toLocaleString()}건 갱신 완료 (마진율 계산에 사용)`);
      setFile(null);setPriceRows([]);
    }catch(err){setUploadStatus("error");setStatusMsg(String(err));}
  },[priceRows]);

  // 카페24 상품코드 업서트 — cafe24_product_codes (norm_name 유니크) 에 갱신 (last-write-wins)
  const doCafe24Upload=useCallback(async()=>{
    if(!cafe24Rows.length) return;
    setUploadStatus("uploading");setStatusMsg("카페24 상품코드 저장 중...");
    try{
      const db=await getSupabase();
      const CHUNK=500;
      for(let i=0;i<cafe24Rows.length;i+=CHUNK){
        const{error}=await db.from("cafe24_product_codes").upsert(cafe24Rows.slice(i,i+CHUNK),{onConflict:"norm_name"});
        if(error) throw new Error(error.message);
      }
      setUploadStatus("done");setStatusMsg(`카페24 상품코드 ${cafe24Rows.length.toLocaleString()}건 갱신 완료 (SKU Risk 다운로드에 사용)`);
      setFile(null);setCafe24Rows([]);
    }catch(err){setUploadStatus("error");setStatusMsg(String(err));}
  },[cafe24Rows]);

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
      {/* 업로드 모드 토글 — 스냅샷(날짜별) / 가격 DB(판매가·공급가) / 카페24 상품코드 */}
      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {[["snapshot","스냅샷 (날짜별)"],["priceDb","가격 DB (판매가·공급가)"],["cafe24","카페24 상품코드"]].map(([m,label])=>{
          const accent=m==="cafe24"?"#9E92C8":"#7EC8A4";
          return (
          <button key={m} onClick={()=>{setUploadMode(m);setFile(null);setParsedRows([]);setPriceRows([]);setCafe24Rows([]);setUploadStatus(null);setStatusMsg("");setSnapDate(null);}}
            style={{flex:1,background:uploadMode===m?accent:"transparent",color:uploadMode===m?(m==="cafe24"?"#fff":"#0a1a12"):DC.sub,
              border:`1px solid ${uploadMode===m?accent:DC.border}`,borderRadius:6,padding:"6px 10px",
              fontSize:12,fontWeight:700,cursor:"pointer"}}>{label}</button>
          );
        })}
      </div>
      {uploadMode==="priceDb"&&(
        <div style={{fontSize:11,color:DC.dim,lineHeight:1.6,marginBottom:8}}>
          날짜 입력 없이 상품의 <b style={{color:DC.sub}}>판매가(정상가)·공급가</b>만 저장합니다. 마진율 계산(베타)의 가격 소스로 사용되며, 같은 상품명은 최신 값으로 갱신됩니다.
        </div>
      )}
      {uploadMode==="cafe24"&&(
        <div style={{fontSize:11,color:DC.dim,lineHeight:1.6,marginBottom:8}}>
          날짜 입력 없이 <b style={{color:DC.sub}}>상품명·카페24 상품코드</b>만 영구 저장합니다. SKU Risk 다운로드의 카페24 상품코드 매칭에 사용되며, 같은 상품명은 최신 값으로 갱신됩니다.
        </div>
      )}
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
        {uploadMode==="snapshot"?(
        <div style={{fontSize:12,lineHeight:1.9,textAlign:"left",display:"inline-block",width:"100%"}}>
          <div style={{marginBottom:6}}>
            <span style={{color:"#7EC8A4",fontWeight:700,fontSize:13}}>인벤토리 트렌드</span>
            <div style={{display:"flex",flexWrap:"wrap",gap:"2px 8px",marginTop:3}}>
              {["상품명","상품코드","옵션","판매가","공급가","현재고","처음입고일","처음입고수량","누적입고","마지막입고일","마지막입고수량","마지막배송일","누적배송수량"].map(c=>(
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
        ):uploadMode==="priceDb"?(
        <div style={{fontSize:12,lineHeight:1.9,textAlign:"left",display:"inline-block",width:"100%"}}>
          <div style={{marginBottom:6}}>
            <span style={{color:"#C8A87B",fontWeight:700,fontSize:13}}>가격 데이터베이스</span>
            <div style={{display:"flex",flexWrap:"wrap",gap:"2px 8px",marginTop:3}}>
              {["상품명","판매가","공급가"].map(c=>(
                <span key={c} style={{background:"rgba(200,168,123,0.1)",border:"1px solid rgba(200,168,123,0.25)",
                  borderRadius:4,padding:"1px 6px",fontSize:12,color:"#C8A87B",fontFamily:"monospace"}}>{c}</span>
              ))}
            </div>
          </div>
          <div style={{color:DC.dim,fontSize:12}}>상품명 + 판매가/공급가만 있으면 됩니다 (날짜·재고 불필요).</div>
        </div>
        ):(
        <div style={{fontSize:12,lineHeight:1.9,textAlign:"left",display:"inline-block",width:"100%"}}>
          <div style={{marginBottom:6}}>
            <span style={{color:"#9E92C8",fontWeight:700,fontSize:13}}>카페24 상품코드</span>
            <div style={{display:"flex",flexWrap:"wrap",gap:"2px 8px",marginTop:3}}>
              {["상품명","상품코드"].map(c=>(
                <span key={c} style={{background:"rgba(158,146,200,0.1)",border:"1px solid rgba(158,146,200,0.25)",
                  borderRadius:4,padding:"1px 6px",fontSize:12,color:"#9E92C8",fontFamily:"monospace"}}>{c}</span>
              ))}
            </div>
          </div>
          <div style={{color:DC.dim,fontSize:12}}>상품명 + 카페24 상품코드만 있으면 됩니다 (카페24 상품 엑셀 그대로 사용 가능).</div>
        </div>
        )}
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

      {/* Upload action — 스냅샷 모드 */}
      {uploadMode==="snapshot"&&parsedRows.length>0&&uploadStatus!=="uploading"&&(
        <div style={{marginTop:10,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:DC.sub}}>
            {`${parsedRows.length.toLocaleString()}개 SKU 준비됨`}
          </span>
          <button onClick={()=>setDateModalOpen(true)}
            style={{background:"#7EC8A4",color:"#0a1a12",border:"none",borderRadius:6,
              padding:"6px 18px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
            업로드
          </button>
        </div>
      )}

      {/* Upload action — 가격 DB 모드 */}
      {uploadMode==="priceDb"&&priceRows.length>0&&uploadStatus!=="uploading"&&(
        <div style={{marginTop:10,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:DC.sub}}>
            {`가격 ${priceRows.length.toLocaleString()}건 준비됨 (판매가 ${priceRows.filter(r=>r.selling_price>0).length.toLocaleString()} · 공급가 ${priceRows.filter(r=>r.supply_price>0).length.toLocaleString()})`}
          </span>
          <button onClick={doPriceDbUpload}
            style={{background:"#C8A87B",color:"#1a1206",border:"none",borderRadius:6,
              padding:"6px 18px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
            가격 DB 저장
          </button>
        </div>
      )}

      {/* Upload action — 카페24 상품코드 모드 */}
      {uploadMode==="cafe24"&&cafe24Rows.length>0&&uploadStatus!=="uploading"&&(
        <div style={{marginTop:10,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:DC.sub}}>
            {`카페24 상품코드 ${cafe24Rows.length.toLocaleString()}건 준비됨`}
          </span>
          <button onClick={doCafe24Upload}
            style={{background:"#9E92C8",color:"#fff",border:"none",borderRadius:6,
              padding:"6px 18px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
            카페24 상품코드 저장
          </button>
        </div>
      )}

      {/* 데이터 날짜 선택 모달 — 업로드 버튼 클릭 시 노출 */}
      {dateModalOpen&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#1a1a1a",border:"1px solid #333",borderRadius:12,padding:28,maxWidth:380,width:"90%"}}>
            <div style={{fontWeight:700,fontSize:15,color:"#F0F0F0",marginBottom:8}}>데이터 날짜 선택</div>
            <div style={{fontSize:12,color:"#888",lineHeight:1.7,marginBottom:16}}>
              업로드할 인벤토리 스냅샷의 기준 날짜를 선택해 주세요. 모든 행에 동일한 날짜가 적용됩니다.
              <br/>{parsedRows.length.toLocaleString()}개 SKU
            </div>
            <input type="date" value={snapDate||""} onChange={e=>setSnapDate(e.target.value)}
              style={{width:"100%",boxSizing:"border-box",background:"transparent",border:"1px solid #333",
                borderRadius:6,padding:"8px 10px",fontSize:13,color:"#F0F0F0",colorScheme:"dark",
                fontFamily:"inherit",marginBottom:18}}/>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button onClick={()=>setDateModalOpen(false)}
                style={{background:"transparent",color:"#888",border:"1px solid #333",borderRadius:6,padding:"6px 16px",fontSize:12,cursor:"pointer"}}>취소</button>
              <button onClick={()=>{
                if(!snapDate) return;
                setDateModalOpen(false);
                handleUploadClick();
              }} disabled={!snapDate}
                style={{background:snapDate?"#7EC8A4":"#3a4a40",color:"#0a1a12",border:"none",borderRadius:6,
                  padding:"6px 16px",fontSize:12,fontWeight:700,cursor:snapDate?"pointer":"default"}}>
                확인
              </button>
            </div>
          </div>
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

// 카페24 상품코드 DB 전체 로드 (페이지네이션). 없거나 실패하면 빈 배열.
async function loadCafe24CodeRows(){
  const out=[];
  try{
    const db=await getSupabase();
    let from=0;const PAGE=1000;
    while(true){
      const{data,error}=await db.from("cafe24_product_codes").select("norm_name,product_name,product_code").range(from,from+PAGE-1);
      if(error||!data||data.length===0) break;
      out.push(...data);
      if(data.length<PAGE) break;
      from+=PAGE;
    }
  }catch{/* 테이블 없거나 로드 실패 시 빈 배열 */}
  return out;
}

// 카페24 코드 매처 생성 — 상품명 매칭 함수 cafe24Of(invRow) 반환.
// ① 색상 포함 키 정확 매칭
// ② 색상-hex 매칭: 카페24 [WHITE] ↔ 어드민 옵션 '화이트' (KR↔EN 동의어). base+색상hex 로 양방향 매칭.
// ③ 색상 제거 base 키 폴백(양쪽 다 색상 없을 때). base가 여러 색상으로 갈리면 ②·③은 오매칭 방지로 비활성.
function makeCafe24Matcher(cafeRows){
  const byColorKey={};            // norm_name(색상 포함) → code
  const baseToCodes={};           // base(색상 제거) → Set(code)
  const baseFirstCode={};         // base → 처음 본 code
  const byBaseColor={};           // base + "|" + colorHex → code (카페24 [색상]에서 추출)
  (cafeRows||[]).forEach(r=>{
    const code=String(r.product_code||"").trim();
    if(!r.norm_name||!code) return;
    byColorKey[r.norm_name]=code;
    const b=String(r.norm_name).replace(/\[[^\]]*\]/g,"");
    if(!baseToCodes[b]) baseToCodes[b]=new Set();
    baseToCodes[b].add(code);
    if(baseFirstCode[b]===undefined) baseFirstCode[b]=code;
    const hex=extractColorHex(r.norm_name,r.product_name);
    if(hex){ const bk=b+"|"+hex; if(byBaseColor[bk]===undefined) byBaseColor[bk]=code; }
  });
  return r=>{
    const ck=normCafe24Name(r.product_name);
    if(byColorKey[ck]) return byColorKey[ck];                 // ① 색상 포함 정확 매칭
    const b=cafe24BaseKey(r.product_name);
    const hex=extractColorHex(r.product_name,r.option_name);  // ② 어드민 색상(상품명[..] 또는 옵션)을 hex로
    if(hex&&byBaseColor[b+"|"+hex]) return byBaseColor[b+"|"+hex];
    if(baseToCodes[b]&&baseToCodes[b].size===1) return baseFirstCode[b]; // ③ 색상 단일일 때만 base 폴백
    return "";
  };
}

// 카페24 코드 미매칭 상품 확인·연결 모달 — SKU Risk 의 인벤토리 상품 중
// 카페24 상품코드가 매칭되지 않은 것을 모아 보여주고, 코드를 직접 입력해 연결(영구 저장)한다.
function Cafe24UnmatchedModal({ rows, onClose }){
  const DC={bg:"#f8f8f6",card:"#ffffff",border:"#e0e0da",text:"#111111",sub:"#444444",dim:"#888888"};
  const [loading,setLoading]=useState(true);
  const [unmatched,setUnmatched]=useState([]); // {key,product_name,option_name,qty,codeInput,saved}
  const [search,setSearch]=useState("");
  const [savingAll,setSavingAll]=useState(false);
  const [msg,setMsg]=useState("");

  const compute=useCallback(async()=>{
    setLoading(true);
    const cafeRows=await loadCafe24CodeRows();
    const cafe24Of=makeCafe24Matcher(cafeRows);
    const seen=new Map();
    (rows||[]).forEach(r=>{
      if(cafe24Of(r)) return; // 이미 매칭됨
      const key=(r.product_name||"")+"__"+(r.option_name||"");
      if(seen.has(key)){ seen.get(key).qty+=(r.current_stock_qty||0); return; }
      seen.set(key,{key,product_name:r.product_name||"",option_name:r.option_name||"",
        qty:(r.current_stock_qty||0),codeInput:"",saved:false});
    });
    const list=[...seen.values()].sort((a,b)=>b.qty-a.qty);
    setUnmatched(list);setLoading(false);
  },[rows]);
  useEffect(()=>{compute();},[compute]);

  const setCode=(key,v)=>setUnmatched(prev=>prev.map(u=>u.key===key?{...u,codeInput:v}:u));
  // 색상-aware norm_name: 상품명에 [색상]이 없고 옵션이 색상이면 [색상] 합성 후 정규화 → 다음 다운로드에서 매칭
  const normForSave=(name,opt)=>{
    const hasBracket=/\[[^\]]+\]/.test(name||"");
    const optTok=String(opt||"").trim();
    const composed=(!hasBracket&&optTok&&colorToHex(optTok))?`${name} [${optTok}]`:name;
    return normCafe24Name(composed);
  };
  const saveOne=async(u)=>{
    const code=String(u.codeInput||"").trim();
    if(!code) return false;
    const db=await getSupabase();
    const row={product_name:u.product_name,norm_name:normForSave(u.product_name,u.option_name),
      product_code:code,updated_at:new Date().toISOString()};
    const{error}=await db.from("cafe24_product_codes").upsert(row,{onConflict:"norm_name"});
    return !error;
  };
  const handleSaveAll=async()=>{
    const targets=unmatched.filter(u=>!u.saved&&String(u.codeInput||"").trim());
    if(!targets.length){setMsg("입력된 코드가 없습니다.");return;}
    setSavingAll(true);setMsg("저장 중...");
    let ok=0;
    for(const u of targets){ if(await saveOne(u)){ok++; setUnmatched(prev=>prev.map(x=>x.key===u.key?{...x,saved:true}:x));} }
    setSavingAll(false);setMsg(`${ok}건 저장 완료 — 다음 다운로드부터 매칭됩니다.`);
  };

  const q=search.trim().toLowerCase();
  const shown=q?unmatched.filter(u=>u.product_name.toLowerCase().includes(q)||(u.option_name||"").toLowerCase().includes(q)):unmatched;
  const pending=unmatched.filter(u=>!u.saved).length;

  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:2000,
      display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()} style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,
        width:"min(720px,96vw)",maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,0.22)"}}>
        {/* 헤더 */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,padding:"16px 18px",borderBottom:`1px solid ${DC.border}`}}>
          <div style={{display:"flex",alignItems:"baseline",gap:10,flexWrap:"wrap"}}>
            <span style={{fontSize:16,fontWeight:700,color:DC.text}}>카페24 미매칭 상품</span>
            <span style={{fontSize:12,color:DC.sub}}>{loading?"확인 중…":`${unmatched.length.toLocaleString()}건 · 코드 입력 후 저장하면 다음 다운로드부터 자동 매칭`}</span>
          </div>
          <button onClick={onClose} style={{background:"none",border:`1px solid ${DC.border}`,borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",color:DC.sub}}>✕ 닫기</button>
        </div>
        {/* 툴바 */}
        <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 18px",borderBottom:`1px solid ${DC.border}`,flexWrap:"wrap"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="상품명·옵션 검색"
            style={{flex:"1 1 180px",background:"transparent",border:`1px solid ${DC.border}`,borderRadius:6,padding:"6px 10px",fontSize:13,color:DC.text,outline:"none",fontFamily:"inherit"}}/>
          <span style={{fontSize:12,color:DC.dim}}>미저장 {pending.toLocaleString()}</span>
          <button onClick={handleSaveAll} disabled={savingAll}
            style={{background:"#9E92C8",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:savingAll?"default":"pointer",opacity:savingAll?0.6:1}}>
            입력한 코드 일괄 저장
          </button>
        </div>
        {msg&&<div style={{padding:"8px 18px",fontSize:12,color:"#5E81AC",borderBottom:`1px solid ${DC.border}`}}>{msg}</div>}
        {/* 목록 */}
        <div style={{overflowY:"auto",padding:"4px 0"}}>
          {loading?(
            <div style={{textAlign:"center",padding:"48px 0",color:DC.dim,fontSize:13}}>인벤토리·카페24 코드 대조 중…</div>
          ):unmatched.length===0?(
            <div style={{textAlign:"center",padding:"48px 0",color:DC.sub,fontSize:14}}>🎉 모든 상품이 카페24 코드와 매칭되었습니다.</div>
          ):(
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{position:"sticky",top:0,background:DC.bg}}>
                {["상품명","옵션","현재고","카페24 상품코드 입력"].map((h,i)=>(
                  <th key={h} style={{textAlign:i>=2?"center":"left",padding:"7px 12px",fontSize:11,fontWeight:600,color:DC.sub,borderBottom:`1px solid ${DC.border}`,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {shown.map(u=>(
                  <tr key={u.key} style={{borderBottom:`1px solid ${DC.border}`,background:u.saved?"rgba(126,200,164,0.12)":"transparent"}}>
                    <td style={{padding:"6px 12px",color:DC.text,maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={u.product_name}>{u.product_name}</td>
                    <td style={{padding:"6px 12px",color:DC.sub}}>{u.option_name||"—"}</td>
                    <td style={{padding:"6px 12px",textAlign:"center",color:DC.sub}}>{u.qty.toLocaleString()}</td>
                    <td style={{padding:"6px 12px",textAlign:"center"}}>
                      {u.saved
                        ?<span style={{color:"#5B9A7B",fontWeight:700,fontSize:12}}>✓ {u.codeInput}</span>
                        :<input value={u.codeInput} onChange={e=>setCode(u.key,e.target.value)} placeholder="예: P0000ABC"
                           style={{width:140,background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,padding:"4px 8px",fontSize:12,color:DC.text,outline:"none",fontFamily:"monospace"}}/>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

// SKU Risk Bubble — 상태별 상품 엑셀 다운로드 (모달 표시 항목 포함)
async function exportSkuRiskXlsx(rows){
  if(!rows||!rows.length){alert("다운로드할 데이터가 없습니다. 날짜를 선택해 인벤토리를 불러오세요.");return;}
  const XLSX=await getXLSX();
  const cafe24Of=makeCafe24Matcher(await loadCafe24CodeRows());
  // 컬럼: …상품코드 · 카페24 상품코드 · …판매가(F) · 공급가(G) · 세일율(H, 입력칸) · 세일가(I, 수식) …
  const colDefs=[
    ["상품코드",      r=>r.product_code||""],
    ["카페24 상품코드", cafe24Of],
    ["상품명",        r=>r.product_name||""],
    ["옵션",          r=>r.option_name||""],
    ["상태",          r=>INV_AGING_DEFS[r.agingKey]?.label||r.agingKey||""],
    ["판매가",        r=>r.selling_price||0],
    ["공급가",        r=>r.supply_price||0],
    ["세일율(%)",     ()=>""],   // 사용자 입력 칸 (비워둠)
    ["세일가",        ()=>""],   // 수식 주입 (아래) — 판매가×(1−세일율/100)
    ["수량(현재고)",   r=>r.current_stock_qty||0],
    ["현재 재고 금액", r=>r.currentInventoryValue||0],
    ["미판매 일수",    r=>r.noSalesDays||0],
    ["SKU 운영기간",  r=>r.skuAge||0],
    ["최근입고 후",    r=>r.postRestockDays||0],
    ["누적배송수량",   r=>r.cumulative_delivery_qty||0],
    ["판매효율(STP)", r=>r.sellThroughProxy||0],
  ];
  const SELL_COL="F",RATE_COL="H",SALE_COL_IDX=8; // 판매가=F, 세일율=H, 세일가=I(0-based 8)
  const toAOA=list=>[colDefs.map(c=>c[0]),...list.map(r=>colDefs.map(c=>c[1](r)))];
  const makeSheet=list=>{
    const ws=XLSX.utils.aoa_to_sheet(toAOA(list));
    // 세일가 = 세일율 비었으면 판매가, 입력 시 판매가×(1−세일율/100) — 엑셀에서 즉시 재계산
    list.forEach((r,i)=>{
      const xlRow=i+2; // 1행=헤더
      const addr=XLSX.utils.encode_cell({c:SALE_COL_IDX,r:i+1});
      ws[addr]={t:"n",v:r.selling_price||0,
        f:`IF(${RATE_COL}${xlRow}="",${SELL_COL}${xlRow},ROUND(${SELL_COL}${xlRow}*(1-${RATE_COL}${xlRow}/100),0))`};
    });
    return ws;
  };
  const wb=XLSX.utils.book_new();
  const sorted=[...rows].sort((a,b)=>
    INV_AGING_KEYS.indexOf(a.agingKey)-INV_AGING_KEYS.indexOf(b.agingKey)
    ||(b.current_stock_qty||0)-(a.current_stock_qty||0));
  XLSX.utils.book_append_sheet(wb,makeSheet(sorted),"전체");
  INV_AGING_KEYS.forEach(k=>{
    const list=rows.filter(r=>r.agingKey===k);
    if(list.length) XLSX.utils.book_append_sheet(wb,makeSheet(list),INV_AGING_DEFS[k].label.slice(0,31));
  });
  const date=rows[0]?.snapshot_date||dayjs().format("YYYY-MM-DD");
  XLSX.writeFile(wb,`SKU_Risk_${date}.xlsx`);
}

// ─────────────────────────────────────────────
// INV BUBBLE SCATTER PLOT
// ─────────────────────────────────────────────
function InvBubblePlot({DC,snapshotDates,stopRef,onExportData}){
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
      const computed=all.map(calcInvRow);
      setData(computed);
      onExportData?.(computed);
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
            <input type="number" onWheel={e=>e.currentTarget.blur()} min={1} value={minStock} onChange={e=>setMinStock(Math.max(1,parseInt(e.target.value)||1))}
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
          .select("snapshot_date,current_stock_qty,selling_price,supply_price,last_delivery_date,first_inbound_date")
          .gte("snapshot_date",rangeStart)
          .order("snapshot_date",{ascending:true})
          .range(from,from+PAGE-1);
        if(!rows||rows.length===0) break;
        rows.forEach(r=>{
          if(rangeEnd&&r.snapshot_date>rangeEnd) return;
          const c=calcInvRow(r);
          if(!map[r.snapshot_date]) map[r.snapshot_date]={};
          if(!map[r.snapshot_date][c.agingKey]) map[r.snapshot_date][c.agingKey]={count:0,qty:0,value:0,cost:0};
          map[r.snapshot_date][c.agingKey].count++;
          map[r.snapshot_date][c.agingKey].qty+=r.current_stock_qty||0;
          map[r.snapshot_date][c.agingKey].value+=c.currentInventoryValue||0;
          map[r.snapshot_date][c.agingKey].cost+=(r.current_stock_qty||0)*(r.supply_price||0);
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
    let total=0,totalQty=0,totalVal=0,totalCost=0;
    INV_AGING_KEYS.forEach(k=>{total+=(d[k]?.count||0);totalQty+=(d[k]?.qty||0);totalVal+=(d[k]?.value||0);totalCost+=(d[k]?.cost||0);});
    const deadQty=d["DEAD"]?.qty||0;const healthyQty=d["HEALTHY"]?.qty||0;
    const qtyByKey={};
    INV_AGING_KEYS.forEach(k=>{
      const qty=d[k]?.qty||0;const val=d[k]?.value||0;const count=d[k]?.count||0;
      qtyByKey[k]={qty,pct:totalQty?(qty/totalQty*100).toFixed(1):"0.0",val,valPct:totalVal?(val/totalVal*100).toFixed(1):"0.0",count,skuPct:total?(count/total*100).toFixed(1):"0.0"};
    });
    const deadCount=d["DEAD"]?.count||0;const healthyCount=d["HEALTHY"]?.count||0;
    return{total,totalQty,totalVal,totalCost,
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
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:18}}>
          {[
            {label:"Dead Stock 비율",value:`${kpi.deadPct}%`,color:"#C87B7B"},
            {label:"Healthy 비율",value:`${kpi.healthyPct}%`,color:"#7EC8A4"},
            {label:"총 현재고",value:`${kpi.totalQty.toLocaleString()}개`,color:DC.text},
            {label:"총 재고 금액",value:fmtVal(kpi.totalVal)+"원",color:"#C8A87B",sub:"판매가 기준"},
            {label:"총 재고 원가",value:fmtVal(kpi.totalCost)+"원",color:MUTE_BLUE,sub:kpi.totalVal>0?`공급가 기준 · 재고금액의 ${Math.round(kpi.totalCost/kpi.totalVal*100)}%`:"공급가 기준"},
          ].map(c=>(
            <div key={c.label} style={{background:DC.bg,border:`1px solid ${DC.border}`,borderRadius:8,padding:"13px 15px"}}>
              <div style={{fontSize:12,color:DC.sub,marginBottom:5}}>{c.label}</div>
              <div style={{fontSize:18,fontWeight:700,color:c.color,letterSpacing:"-0.3px"}}>{c.value}</div>
              {c.sub&&<div style={{fontSize:10,color:DC.dim,marginTop:4}}>{c.sub}</div>}
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
  const [copied,setCopied]=useState(false);
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

  const exportCols=[
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
  const buildExportHtml=(target,cols=exportCols)=>{
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
    return `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"/></head><body><table style="border-collapse:collapse;"><thead><tr>${headerHtml}</tr></thead><tbody>${bodyHtml}</tbody></table></body></html>`;
  };
  const buildExportTsv=(target,cols=exportCols)=>{
    const cell=v=>String(v??"").replace(/[\t\r\n]+/g," ").trim();
    return [cols.map(c=>c.label).join("\t"),...target.map(r=>cols.map(c=>cell(c.get(r))).join("\t"))].join("\n");
  };
  const downloadCSV=(source)=>{
    const html=buildExportHtml(source||filtered);
    const blob=new Blob(["﻿"+html],{type:"application/vnd.ms-excel;charset=utf-8;"});
    const a=Object.assign(document.createElement("a"),{href:URL.createObjectURL(blob),download:`reorder_${localDate(0)}.xls`});
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
  // 선택 복사용 컬럼 — 상품명 · 옵션 · 추천 리오더 수만 (상품코드 제외)
  const copyCols=exportCols.filter(c=>["상품명","옵션","추천 리오더"].includes(c.label));
  const copySelected=async()=>{
    const target=filtered.filter(r=>selected.has(rowKey(r)));
    if(!target.length) return;
    const tsv=buildExportTsv(target,copyCols);
    const html=buildExportHtml(target,copyCols);
    const fallback=()=>{
      const ta=document.createElement("textarea");
      ta.value=tsv;ta.style.position="fixed";ta.style.opacity="0";
      document.body.appendChild(ta);ta.focus();ta.select();
      try{document.execCommand("copy");}catch{/* noop */}
      document.body.removeChild(ta);
    };
    try{
      if(navigator.clipboard&&typeof window.ClipboardItem==="function"){
        await navigator.clipboard.write([new window.ClipboardItem({
          "text/plain":new Blob([tsv],{type:"text/plain"}),
          "text/html":new Blob([html],{type:"text/html"}),
        })]);
      }else if(navigator.clipboard&&navigator.clipboard.writeText){
        await navigator.clipboard.writeText(tsv);
      }else{
        fallback();
      }
    }catch{
      fallback();
    }
    setCopied(true);
    setTimeout(()=>setCopied(false),1500);
  };

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
                  <button onClick={copySelected} title="상품명 · 옵션 · 추천 리오더 수를 표 형식으로 복사 — 엑셀/구글시트에 바로 붙여넣기"
                    style={{background:copied?"#7B9EC8":"transparent",color:copied?"#fff":"#7B9EC8",border:"1px solid #7B9EC8",borderRadius:5,
                      padding:"4px 12px",fontSize:13,cursor:"pointer",fontWeight:600}}>{copied?"복사됨 ✓":"⧉ 선택 복사"}</button>
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
            <style>{`.reordercalc tbody tr{transition:background 0.12s;}.reordercalc tbody tr:hover td{background:#f4f4f2;}`}</style>
            <table className="reordercalc" style={{width:"100%",borderCollapse:"collapse",fontSize:13,tableLayout:"auto"}}>
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
                    <tr key={r.reorder_id||i} onClick={()=>toggleRow(k)}
                      style={{borderBottom:`1px solid ${DC.border}`,background:isSel?"rgba(126,200,164,0.08)":"transparent",cursor:"pointer",userSelect:"none"}}>
                      <td style={{padding:"4px 6px",textAlign:"center"}} onClick={e=>e.stopPropagation()}>
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
      filterEnd=localDate(0);
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
    const el=cardRef.current;
    // Hide all capture buttons inside the card before snapshot
    const btns=el.querySelectorAll("[data-capture-hide]");
    btns.forEach(b=>{b._prevVis=b.style.visibility;b.style.visibility="hidden";});
    // 스크롤 컨테이너(모달 등)는 보이는 영역만 잡히므로, 높이 제약을 잠시 풀어
    // 폭(현재 뷰 비율)은 그대로 두고 전체 스크롤 높이로 펼쳐 캡처 후 복원
    const prevStyle={maxHeight:el.style.maxHeight,height:el.style.height,overflow:el.style.overflow};
    const prevScroll=el.scrollTop;
    el.style.maxHeight="none";el.style.height="auto";el.style.overflow="visible";
    const fullH=el.scrollHeight;
    const restore=()=>{
      el.style.maxHeight=prevStyle.maxHeight;el.style.height=prevStyle.height;el.style.overflow=prevStyle.overflow;
      el.scrollTop=prevScroll;
      btns.forEach(b=>{b.style.visibility=b._prevVis||"";});
    };
    try{
      const {default:html2canvas}=await import("html2canvas");
      const canvas=await html2canvas(el,{scale:2,useCORS:true,backgroundColor:null,logging:false,height:fullH,windowHeight:fullH});
      restore();
      const fname=`${filename}_${localDate(0)}.png`;
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
    }catch(e){restore();console.error(e);}
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

function DataCompare({revenues,storeSales=[],orders=[],stocks=[],ts={}}){
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
  const [svgW,setSvgW]=useState(760);
  const [snapshotDates,setSnapshotDates]=useState([]);
  const [invRefreshKey]=useState(0);
  const [agingDate,setAgingDate]=useState(null);
  const [bubbleRows,setBubbleRows]=useState([]); // SKU Risk Bubble 현재 로드된 SKU (엑셀 다운로드용)
  const [unmatchedOpen,setUnmatchedOpen]=useState(false); // 카페24 미매칭 확인 모달

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
        res.push({label:`${cur.getFullYear()}.${cur.getMonth()+1}`,start:ymd(cur),end:ymd(e)});
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
          res.push({label:`${cur.getFullYear()}.${cur.getMonth()+1}`,start:ymd(cur),end:ymd(eC)});
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

      {/* Inventory 업로더는 '데이터 입력 > 인벤토리' 탭으로 이동 (snapshotDates는 마운트 시 로드) */}

      {/* ② SKU Risk Bubble — 다크 카드 */}
      <div ref={bubbleCardRef} style={sectionCard}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,gap:8}}>
          <div style={{fontWeight:600,fontSize:16,color:DC.text,letterSpacing:"-0.2px"}}>SKU Risk Bubble</div>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <button onClick={()=>setUnmatchedOpen(true)} title="카페24 상품코드가 매칭되지 않은 상품 확인·연결"
              style={{background:"transparent",border:`1px solid #9E92C8`,borderRadius:6,
                padding:"5px 12px",fontSize:12,fontWeight:600,color:"#9E92C8",cursor:"pointer",
                display:"flex",alignItems:"center",gap:5}}>
              ⚠ 카페24 미매칭 확인
            </button>
            <button onClick={()=>exportSkuRiskXlsx(bubbleRows)} title="상태별 상품 엑셀 다운로드"
              style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:6,
                padding:"5px 12px",fontSize:12,fontWeight:600,color:DC.text,cursor:"pointer",
                display:"flex",alignItems:"center",gap:5}}>
              ⬇ 상품 다운로드
            </button>
            <CaptureBtn cardRef={bubbleCardRef} filename="SKU_Risk_Bubble" DC={DC}/>
          </div>
        </div>
        <InvBubblePlot DC={DC} snapshotDates={snapshotDates} stopRef={agingTrendSecRef} onExportData={setBubbleRows}/>
      </div>
      {unmatchedOpen&&<Cafe24UnmatchedModal rows={bubbleRows} onClose={()=>setUnmatchedOpen(false)}/>}

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

      {/* 리오더 계산기는 '리오더 계산기' 탭으로 분리됨 — 업로더(위)는 그대로 유지되어 계산 소스 로직 보존 */}
      <div ref={reorderSecRef} style={{height:1}}/>

      {/* ⑤ 물류 플로우 — 기존 '물류 플로우' 탭에서 이동 */}
      <LogisticsFlow orders={orders} stocks={stocks} ts={ts}/>
    </div>
  );
}

// ─────────────────────────────────────────────
// GMV 계산기 — 채널별 목표 매출 → 최종 세일율 역산
// ─────────────────────────────────────────────
const GMV_CHANNELS=["자사몰","29CM","오프라인 스토어"];
const GMV_FEE_DEFAULT={"자사몰":3,"29CM":28,"오프라인 스토어":28};
// 채널 수수료율 반영 세일율→마진 (실수령 net 기준, 자사몰 실수령=쿠폰적용가 / 29CM·오프라인=정산액)
function gmvCompute(list,saleRate,supply,feeRate,couponRate=0){
  const r=Math.max(0,Math.min(100,saleRate||0));
  const saleAmt=Math.round((list||0)*r/100);
  const salePrice=(list||0)-saleAmt;                 // 세일 후 노출가
  const cpn=Math.max(0,Math.min(100,couponRate||0));
  const couponAmt=Math.round(salePrice*cpn/100);
  const finalPrice=salePrice-couponAmt;              // 쿠폰까지 적용된 실판매가
  const fee=Math.round(finalPrice*(feeRate||0)/100); // 채널 수수료
  const net=finalPrice-fee;                          // 자사 실수령(정산액)
  const supplyVat=Math.round((supply||0)*1.1);       // 원가(부가세 10% 포함)
  const margin=net-supplyVat;
  const marginRate=net>0?Math.round(margin/net*1000)/10:0;
  const effDisc=(list>0)?Math.round((1-finalPrice/list)*1000)/10:0;
  return {saleAmt,salePrice,couponAmt,finalPrice,fee,net,supplyVat,margin,marginRate,effDisc};
}
// 채널 목표 매출 달성 세일율 = (1 − 목표매출 ÷ 정가GMV) × 100  (0~90 클램프)
function gmvTargetRate(targetRev,regularGmv){
  if(!regularGmv||regularGmv<=0) return null;
  if(!targetRev||targetRev<=0) return 0;
  const r=(1-targetRev/regularGmv)*100;
  return Math.max(0,Math.min(90,Math.round(r*10)/10));
}

function GmvCalculator({orders=[],revenues=[],storeSales=[],stocks=[]}){
  const DC={bg:"#f8f8f6",card:"#ffffff",border:"#e0e0da",text:"#111111",sub:"#444444",dim:"#888888"};
  const [invRows,setInvRows]=useState([]);       // 최신 inventory_snapshot SKU
  const [reorderRows,setReorderRows]=useState([]); // reorder_recommendations
  const [loading,setLoading]=useState(true);
  // 채널별 목표 매출은 입력하지 않음 — 월 고정금액 + 목표 이익금으로 자동 산출(읽기 전용).
  // 월 고정금액 / 목표 이익금 / 채널 수수료율 — Supabase(gmv_settings) 단일 행에 저장, localStorage 폴백
  const [fixedCost,setFixedCost]=useState(()=>{try{return localStorage.getItem("gmv_fixed_cost")||"";}catch{return "";}});
  const [targetProfit,setTargetProfit]=useState(()=>{try{return localStorage.getItem("gmv_target_profit")||"";}catch{return "";}});
  const [feeRates,setFeeRates]=useState({...GMV_FEE_DEFAULT});
  const settingsLoaded=useRef(false);
  // 마운트 시 gmv_settings 로드(있으면 입력값 복원)
  useEffect(()=>{(async()=>{
    try{
      const db=await getSupabase();
      const{data}=await db.from("gmv_settings").select("fixed_cost,target_profit,fee_rates").eq("id",1).maybeSingle();
      if(data){
        if(data.fixed_cost!=null) setFixedCost(String(data.fixed_cost||""));
        if(data.target_profit!=null) setTargetProfit(String(data.target_profit||""));
        if(data.fee_rates&&typeof data.fee_rates==="object") setFeeRates(p=>({...p,...data.fee_rates}));
      }
    }catch{/* 테이블 없거나 로드 실패 시 localStorage 값 유지 */}
    settingsLoaded.current=true;
  })();},[]);
  // 변경 시 저장(디바운스) — localStorage 즉시 + Supabase 업서트
  useEffect(()=>{
    try{localStorage.setItem("gmv_fixed_cost",fixedCost);localStorage.setItem("gmv_target_profit",targetProfit);}catch{/* noop */}
    if(!settingsLoaded.current) return; // 초기 로드 전에는 저장 안 함
    const t=setTimeout(async()=>{
      try{
        const db=await getSupabase();
        await db.from("gmv_settings").upsert({id:1,
          fixed_cost:parseInt(String(fixedCost).replace(/[^0-9]/g,""),10)||0,
          target_profit:parseInt(String(targetProfit).replace(/[^0-9]/g,""),10)||0,
          fee_rates:feeRates,updated_at:new Date().toISOString()},{onConflict:"id"});
      }catch{/* noop */}
    },600);
    return()=>clearTimeout(t);
  },[fixedCost,targetProfit,feeRates]);
  const [search,setSearch]=useState("");
  const [expanded,setExpanded]=useState(null);    // 펼친 상품 key
  const [cycleCh,setCycleCh]=useState("자사몰");   // 사이클 다이어그램 채널
  const [srcOpen,setSrcOpen]=useState(null);      // 채널 카드 소스 상세 펼침
  const [showCount,setShowCount]=useState(40);
  const cardRef=useRef(null);

  // ── 데이터 로드: 최신 inventory_snapshot + reorder_recommendations
  const load=useCallback(async()=>{
    setLoading(true);
    const db=await getSupabase();
    // 최신 스냅샷 날짜
    const{data:latest}=await db.from("inventory_snapshot").select("snapshot_date").order("snapshot_date",{ascending:false}).limit(1);
    const snap=latest?.[0]?.snapshot_date;
    let inv=[];
    if(snap){
      let from=0;const PAGE=1000;
      while(true){
        const{data,error}=await db.from("inventory_snapshot")
          .select("product_name,option_name,product_code,selling_price,supply_price,current_stock_qty,cumulative_delivery_qty,last_delivery_date,first_inbound_date,snapshot_date")
          .eq("snapshot_date",snap).range(from,from+PAGE-1);
        if(error||!data||data.length===0) break;
        inv=inv.concat(data);
        if(data.length<PAGE) break;
        from+=PAGE;
      }
    }
    let reo=[];
    {let from=0;const PAGE=1000;
      while(true){
        const{data,error}=await db.from("reorder_recommendations").select("*").range(from,from+PAGE-1);
        if(error||!data||data.length===0) break;
        reo=reo.concat(data);
        if(data.length<PAGE) break;
        from+=PAGE;
      }}
    setInvRows(inv);setReorderRows(reo);setLoading(false);
  },[]);
  useEffect(()=>{load();},[load]);

  // ── 채널 매출·점유율 — 데이터 컴페어/대시보드와 동일 방식으로 직접 집계(채널명 정확 매칭)
  //    온라인: revenues(채널별 매출 − 환불), 오프라인: store_sales(배송 − 반품). 점유율 분모=GMV 채널 합.
  const revByCh=useMemo(()=>{
    const onlineByCh={};
    (revenues||[]).forEach(r=>{
      const ch=normChannel(r.channel);
      onlineByCh[ch]=(onlineByCh[ch]||0)+((r.amount||0)-(r.refund_amount||0));
    });
    let offline=0;
    (storeSales||[]).forEach(r=>{
      if(r.status==="배송") offline+=(r.amount||0);
      else if(r.status==="반품") offline-=(r.amount||0);
    });
    const rev={
      "자사몰":Math.max(0,onlineByCh["자사몰"]||0),
      "29CM":Math.max(0,onlineByCh["29CM"]||0),
      "오프라인 스토어":Math.max(0,offline),
    };
    const total=GMV_CHANNELS.reduce((s,ch)=>s+rev[ch],0)||1;
    const m={};
    GMV_CHANNELS.forEach(ch=>{ m[ch]={revenue:rev[ch],share:Math.round(rev[ch]/total*1000)/10}; });
    return m;
  },[revenues,storeSales]);

  // ── 입고 흐름: stock_uploads(입고 CSV)에서 SKU별 최신 입고 + 최근 입고일 집계
  const inboundBySku=useMemo(()=>{
    const latest={};
    (stocks||[]).forEach(r=>{
      const k=normProdName(r.product_name)+"|"+String(r.option_name||"").trim().toLowerCase();
      if(!latest[k]||(r.upload_date||"")>(latest[k].upload_date||"")) latest[k]={qty:r.qty||0,upload_date:r.upload_date||""};
    });
    return latest;
  },[stocks]);

  // ── 상품 모집단: 인벤토리 + 4주 판매수량(reorder_monthly_sales) + 입고 흐름 → 예상수량·에이징
  //    ※ '4주 판매수량'은 실제 리오더 수량이 아니라 최근 4주간 판매(발주합계) 수량 → 향후 판매 추정치로 사용
  const products=useMemo(()=>{
    const rkey=(n,o)=>normProdName(n)+"|"+String(o||"").trim().toLowerCase();
    const reoMap={};
    reorderRows.forEach(r=>{reoMap[rkey(r.reorder_product_name,r.reorder_option_name)]=r;});
    const out=invRows.map(r=>{
      const k=rkey(r.product_name,r.option_name);
      // 정가 = 인벤토리 업로더 데이터 파일의 판매가(selling_price) — 전 채널 공통 정가 기준
      const list=r.selling_price||0, supply=r.supply_price||0;
      const reo=reoMap[k];
      const qty4w=reo?reo.reorder_monthly_sales||0:0;   // 최근 4주 판매수량
      const qtyDeliv=r.cumulative_delivery_qty||0;
      const inb=inboundBySku[k]||null;
      const expectedQty=qty4w>0?qty4w:0;                // 4주 판매수량 우선(향후 판매 추정)
      const aging=getAgingKey(calcInvRow(r).noSalesDays);
      return {key:k,name:r.product_name,option:r.option_name||"",code:r.product_code||"",
        list,supply,qty4w,qtyDeliv,inboundQty:inb?inb.qty:0,inboundDate:inb?inb.upload_date:"",
        expectedQty,aging,matched:list>0&&supply>0};
    });
    // 스테디셀러: 4주 판매수량 상위 + Healthy
    const sortedQty=[...out].filter(p=>p.qty4w>0).sort((a,b)=>b.qty4w-a.qty4w);
    const topCut=sortedQty[Math.floor(sortedQty.length*0.2)]?.qty4w||0; // 상위 20% 컷
    out.forEach(p=>{p.steady=(p.aging==="HEALTHY"&&p.qty4w>0&&p.qty4w>=topCut);});
    return out;
  },[invRows,reorderRows,inboundBySku]);

  // ── 입고 흐름 요약 (최근 입고일별 입고 수량 추이)
  const inboundFlow=useMemo(()=>{
    const byDate={};
    (stocks||[]).forEach(r=>{const d=r.upload_date||"";if(!d)return;byDate[d]=(byDate[d]||0)+(r.qty||0);});
    const dates=Object.keys(byDate).sort();
    const recent=dates.slice(-6).map(d=>({date:d,qty:byDate[d]}));
    const totalInbound=Object.values(inboundBySku).reduce((s,v)=>s+(v.qty||0),0);
    return {recent,totalInbound,latestDate:dates[dates.length-1]||"—"};
  },[stocks,inboundBySku]);

  // ── 최근 한 달 실적: 가장 최근 주문일 기준 30일간 배송 주문(채널별) → 실제 판매수량·실판매가·실효 세일율·마진
  //   채널별 실판매 매출 산정(주문배송 데이터 기준):
  //   · 자사몰(MERRYON): 동일 order_no의 [결제금액] 1개가 실제 결제금액(주문 단위, 1회만).
  //     상품별 세일율은 결제금액을 주문 내 정가비중으로 배분해 추정.
  //   · 29CM: 동일 order_no 내 [판매가] 합이 실제 결제금액 → 상품(라인)별 판매가 그대로 사용.
  //   · 오프라인: store_sales 실판매금액(amount, 라인) 그대로.
  const recentActuals=useMemo(()=>{
    const dated=(orders||[]).filter(o=>o.order_date);
    if(!dated.length) return {byKeyCh:{},start:"",end:"",chTotals:{}};
    const end=dated.reduce((mx,o)=>o.order_date>mx?o.order_date:mx,"");
    const start=new Date(new Date(end+"T00:00:00").getTime()-29*86400000).toISOString().slice(0,10);
    const supplyByKey={};products.forEach(p=>{supplyByKey[p.key]={supply:p.supply,list:p.list};});
    const inWin=o=>o.order_date&&o.order_date>=start&&o.order_date<=end&&o.status==="배송";
    const byKeyCh={}; const chTotals={};
    GMV_CHANNELS.forEach(c=>{chTotals[c]={qty:0,revenue:0,margin:0,listGmv:0};});
    // 자사몰: order_no별 결제금액(1회) + 라인들의 정가비중으로 실판매 배분
    const ownOrders={}; // order_no → {payment, lines:[{k,meta,qty}]}
    const addRow=(ch,k,meta,qty,lineRev)=>{
      const fee=feeRates[ch]||0;
      const lineList=meta.list*qty;
      const lineNet=Math.round(lineRev*(1-fee/100));
      const lineMargin=lineNet-Math.round((meta.supply||0)*1.1)*qty;
      const id=k+"@@"+ch;
      if(!byKeyCh[id]) byKeyCh[id]={qty:0,revenue:0,margin:0,listGmv:0,list:meta.list,supply:meta.supply};
      byKeyCh[id].qty+=qty; byKeyCh[id].revenue+=lineRev; byKeyCh[id].margin+=lineMargin; byKeyCh[id].listGmv+=lineList;
      chTotals[ch].qty+=qty; chTotals[ch].revenue+=lineRev; chTotals[ch].margin+=lineMargin; chTotals[ch].listGmv+=lineList;
    };
    (orders||[]).forEach(o=>{
      if(!inWin(o)) return;
      const ch=normChannel(o.channel);
      if(!GMV_CHANNELS.includes(ch)) return;
      const k=normProdName(o.product_name)+"|"+String(o.option_name||"").trim().toLowerCase();
      const meta=supplyByKey[k]; if(!meta) return; // 인벤토리 매칭분만
      const qty=o.qty||0;
      if(ch==="자사몰"){
        const ono=o.order_no||o.order_id||""; if(!ono) return;
        if(!ownOrders[ono]) ownOrders[ono]={payment:0,lines:[]};
        if((o.payment_amount||0)>ownOrders[ono].payment) ownOrders[ono].payment=o.payment_amount||0; // 주문당 1회(MAX)
        ownOrders[ono].lines.push({k,meta,qty});
      }else{
        // 29CM·오프라인: 라인 판매가(sale_price=상품 판매가/매장 실판매금액) 그대로
        addRow(ch,k,meta,qty,(o.sale_price||0)>0?o.sale_price:meta.list*qty);
      }
    });
    // 자사몰 배분: 결제금액을 주문 내 정가매출(정가×수량) 비중으로 라인에 배분
    Object.values(ownOrders).forEach(ord=>{
      const totalList=ord.lines.reduce((s,l)=>s+l.meta.list*l.qty,0);
      const pay=ord.payment>0?ord.payment:totalList; // 결제금액 없으면 정가(세일 0)
      ord.lines.forEach(l=>{
        const w=totalList>0?(l.meta.list*l.qty)/totalList:0;
        addRow("자사몰",l.k,l.meta,l.qty,Math.round(pay*w));
      });
    });
    return {byKeyCh,start,end,chTotals};
  },[orders,products,feeRates]);

  // ── 현재 이익금(최근 30일 실판매 마진) + 목표 배수 r = (목표이익금 + 월고정금액) ÷ 현재이익금
  const targetProfitN=parseInt(String(targetProfit).replace(/[^0-9]/g,""),10)||0;
  const fixedCostN=parseInt(String(fixedCost).replace(/[^0-9]/g,""),10)||0;
  const reqMargin=targetProfitN+fixedCostN; // 필요 총마진(목표 이익금 + 월 고정금액)
  const currentTotalMargin=GMV_CHANNELS.reduce((s,ch)=>s+(recentActuals.chTotals[ch]?.margin||0),0);
  // 목표 배수 = 목표 이익금 ÷ 현재 이익금 (월 고정금액은 필요 총마진·최종 이익금에만 반영, 배수엔 미포함)
  const targetMultiplier=(currentTotalMargin>0&&targetProfitN>0)?targetProfitN/currentTotalMargin:1;
  // 최근 한달 재입고비(현재) = Σ(최근 30일 입고 품목 수량 × 공급가) — 입고(stock_uploads) 기준
  const currentTotalRestock=useMemo(()=>{
    if(!(stocks||[]).length) return 0;
    const latest=stocks.reduce((mx,r)=>(r.upload_date||"")>mx?r.upload_date:mx,"");
    if(!latest) return 0;
    const from=new Date(new Date(latest+"T00:00:00").getTime()-29*86400000).toISOString().slice(0,10);
    const supplyByKey={};products.forEach(p=>{supplyByKey[p.key]=p.supply;});
    let r=0;
    stocks.forEach(s=>{
      const d=s.upload_date||""; if(d<from||d>latest) return;
      const k=normProdName(s.product_name)+"|"+String(s.option_name||"").trim().toLowerCase();
      const supply=supplyByKey[k]||0;
      r+=supply*(s.qty||0); // 입고 수량 × 공급가 (재입고 매입원가)
    });
    return r;
  },[stocks,products]);

  // ── 채널별: 현재 세일율 → 목표 매출(자동) → 권장 세일율 → Δ
  const channelCalc=useMemo(()=>{
    return GMV_CHANNELS.map(ch=>{
      const t=recentActuals.chTotals[ch]||{qty:0,revenue:0,margin:0,listGmv:0};
      const fee=feeRates[ch]||0;
      const curRev=t.revenue||0;                 // 최근 30일 실판매 매출(실효가 기준)
      const listGmv=t.listGmv||0;                // 정가 GMV (분모 고정)
      const curRate=listGmv>0?Math.max(0,Math.min(95,Math.round((1-curRev/listGmv)*1000)/10)):null; // 현재 실효 세일율
      const curMargin=t.margin||0;
      const curMarginRate=curRev>0?Math.round(curMargin/(curRev*(1-fee/100))*1000)/10:0;
      // 목표 매출 = 현재 매출 × 목표 배수 (점유율 자동 보존)
      const targetRev=Math.round(curRev*targetMultiplier);
      // 권장 세일율 = 1 − (목표매출 ÷ 정가GMV), 현재와 동일 분모 → Δ 직관적
      const recRate=listGmv>0?Math.max(0,Math.min(90,Math.round((1-targetRev/listGmv)*1000)/10)):null;
      const delta=(recRate!=null&&curRate!=null)?Math.round((recRate-curRate)*10)/10:null;
      // 권장 세일율 적용 시 예상 마진 (현재 판매수량 고정 가정 → 정가GMV·수량 그대로)
      // 예상 재입고비 = 예측 판매량(최근 30일 실판매 수량) × 공급가×1.1 — 판 만큼 다시 매입
      let expMargin=0,restock=0;
      products.forEach(p=>{
        const id=p.key+"@@"+ch; const a=recentActuals.byKeyCh[id];
        if(!a||a.qty<=0) return;
        const m=gmvCompute(p.list,recRate||0,p.supply,fee);
        expMargin+=m.margin*a.qty;
        restock+=Math.round(p.supply*1.1)*a.qty; // 예측 판매량 기준 예상 재입고비
      });
      const expMarginRate=targetRev>0?Math.round(expMargin/(targetRev*(1-fee/100))*1000)/10:0;
      const share=revByCh[ch]?.share||0;
      return {ch,fee,curRev,listGmv,curRate,curMargin,curMarginRate,targetRev,recRate,delta,expMargin,expMarginRate,restock,share};
    });
  },[products,recentActuals,feeRates,targetMultiplier,revByCh]);

  const totalTargetRev=channelCalc.reduce((s,c)=>s+c.targetRev,0);
  const totalExpMargin=channelCalc.reduce((s,c)=>s+c.expMargin,0);
  const totalRestock=channelCalc.reduce((s,c)=>s+c.restock,0);
  const netContribution=totalExpMargin-reqMargin; // 예상 마진 − 필요 총마진

  // ── 표시 상품 (스테디셀러·매칭 우선 + 검색)
  const shownProducts=useMemo(()=>{
    const q=search.trim().toLowerCase();
    let list=products.filter(p=>p.matched);
    if(q) list=list.filter(p=>p.name.toLowerCase().includes(q)||(p.option||"").toLowerCase().includes(q));
    list=[...list].sort((a,b)=>(b.steady?1:0)-(a.steady?1:0)||b.qty4w-a.qty4w);
    return list;
  },[products,search]);
  const unmatchedCount=products.filter(p=>!p.matched).length;

  // ── 데이터 출처 각주 (최신 날짜)
  const maxOf=(arr,f)=>{let mx="";arr.forEach(x=>{const v=f(x);if(v&&v>mx)mx=v;});return mx||"—";};
  const footnotes=[
    {label:"인벤토리 스냅샷",date:maxOf(invRows,r=>r.snapshot_date),note:"공급가·판매가·재고"},
    {label:"리오더(4주 판매수량)",date:maxOf(reorderRows,r=>r.reorder_data_date),note:"최근 4주 판매=향후 추정"},
    {label:"입고(stock_uploads)",date:inboundFlow.latestDate,note:"현재 입고 흐름"},
    {label:"주문·배송",date:maxOf(orders,r=>r.order_date),note:"채널 매출(이지어드민)"},
    {label:"매장 판매",date:maxOf(storeSales,r=>r.sale_date),note:"오프라인 매출"},
    {label:"채널 매출",date:maxOf(revenues,r=>r.date),note:"온라인 채널 일자 매출"},
  ];

  const won=n=>"₩"+Math.round(n||0).toLocaleString();
  // 금액 입력 보조: 숫자만 추출 → 천단위 콤마 표시 / 억천만 읽기
  const digitsOf=v=>String(v||"").replace(/[^0-9]/g,"");
  const commaOf=v=>{const d=digitsOf(v);return d?Number(d).toLocaleString():"";};
  const eokManOf=v=>{const n=parseInt(digitsOf(v),10)||0;if(n<10000)return n>0?n.toLocaleString()+"원":"";const s=fmtEokMan(n);return s.startsWith("0억")?s.slice(2):s;};
  const lbl={fontSize:11,color:DC.sub,marginBottom:4,display:"block"};
  const inBox={background:"transparent",border:`1px solid ${DC.border}`,borderRadius:6,padding:"7px 10px",fontSize:13,color:DC.text,fontFamily:"inherit",width:"100%",outline:"none"};

  return(
    <div ref={cardRef} style={{background:DC.bg,minHeight:"100%",padding:"28px 28px 48px"}}>
      <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:4,flexWrap:"wrap"}}>
        <span style={{fontWeight:700,fontSize:20,color:DC.text,letterSpacing:"-0.3px"}}>GMV 계산기</span>
        <span style={{fontSize:12,color:DC.sub}}>채널별 목표 매출 → 적정 최종 세일율·마진율 역산 (베타)</span>
        {loading&&<span style={{fontSize:11,color:MUTE_BLUE}}>데이터 로딩 중…</span>}
      </div>
      <div style={{fontSize:12,color:DC.dim,marginBottom:18}}>데이터 입력 탭의 업로더로 새 데이터가 올라오면 자동으로 다시 계산됩니다.</div>

      {/* ① 목표 입력 — 월 고정금액 + 목표 이익금만(채널별 목표 매출은 자동 산출) */}
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"18px 20px",marginBottom:16}}>
        <div style={{fontSize:13,fontWeight:700,color:DC.text,marginBottom:4}}>① 목표 설정 <span style={{fontSize:11,color:DC.dim,fontWeight:400}}>· 월 고정금액·목표 이익금만 입력하면 채널별 목표 매출은 현재 점유율대로 자동 산출</span></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginTop:10}}>
          <div><label style={lbl}>월 고정금액 <span style={{color:DC.dim,fontWeight:400}}>· 최근 입력값 저장</span></label>
            <input value={commaOf(fixedCost)} onChange={e=>setFixedCost(digitsOf(e.target.value))} inputMode="numeric" placeholder="예: 10,000,000" style={inBox}/>
            {digitsOf(fixedCost)&&<div style={{fontSize:11,color:MUTE_BLUE,marginTop:3,fontWeight:600}}>{eokManOf(fixedCost)}</div>}</div>
          <div><label style={lbl}>목표 이익금 <span style={{color:DC.dim,fontWeight:400}}>· 최근 입력값 저장</span></label>
            <input value={commaOf(targetProfit)} onChange={e=>setTargetProfit(digitsOf(e.target.value))} inputMode="numeric" placeholder="예: 20,000,000" style={inBox}/>
            {digitsOf(targetProfit)&&<div style={{fontSize:11,color:MUTE_BLUE,marginTop:3,fontWeight:600}}>{eokManOf(targetProfit)}</div>}</div>
        </div>
        {/* 채널 수수료율 (편집) */}
        <div style={{display:"flex",gap:16,flexWrap:"wrap",marginTop:12}}>
          {GMV_CHANNELS.map(ch=>(
            <span key={ch} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:DC.sub}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:chColor(ch),display:"inline-block"}}/>{ch} 수수료
              <input value={feeRates[ch]} onChange={e=>setFeeRates(p=>({...p,[ch]:Math.max(0,Math.min(100,parseFloat(e.target.value)||0))}))}
                type="number" style={{width:46,textAlign:"right",background:"transparent",border:`1px solid ${DC.border}`,borderRadius:4,padding:"2px 5px",fontSize:12,color:DC.text}}/>%
            </span>
          ))}
        </div>
        <div style={{fontSize:11,color:DC.dim,marginTop:10,lineHeight:1.7}}>
          필요 총마진 = 목표 이익금 + 월 고정금액 = <b style={{color:DC.sub}}>{won(reqMargin)}</b><br/>
          현재 이익금(최근 30일 실판매 마진) <b style={{color:DC.sub}}>{won(currentTotalMargin)}</b> · 목표 배수 <b style={{color:MUTE_BLUE}}>×{targetMultiplier.toFixed(2)}</b> <span style={{color:DC.dim}}>(목표 이익금 ÷ 현재 이익금)</span><br/>
          재입고비(최근 30일 입고 수량 × 공급가) <b style={{color:"#C8A87B"}}>−{won(currentTotalRestock)}</b><br/>
          최종 이익금(현재 이익금 − 재입고비) <b style={{color:(currentTotalMargin-currentTotalRestock)>=0?"#1a7a4f":"#c0392b"}}>{won(currentTotalMargin-currentTotalRestock)}</b>
          {" · 채널별 목표 매출 = 현재 매출 × 배수 (점유율 유지)"}
          {currentTotalMargin<=0&&<span style={{color:"#C8A87B"}}> · ⚠ 최근 실적 마진이 0이라 배수 산출 불가(데이터 보강 필요)</span>}
        </div>
      </div>

      {/* 채널별 조정 카드 — 현재 운용 → 목표 도달, 세일율 조정 방향(Δ) + 현재/예상 동시 표기 */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:16}}>
        {channelCalc.map(c=>{
          const cur30={revenue:recentActuals.chTotals[c.ch]?.revenue||0,margin:recentActuals.chTotals[c.ch]?.margin||0};
          const curRestock=(()=>{let r=0;products.forEach(p=>{const a=recentActuals.byKeyCh[p.key+"@@"+c.ch];if(a&&a.qty>0)r+=Math.round(p.supply*1.1)*a.qty;});return r;})();
          const up=c.delta!=null&&c.delta>0;   // 세일율 ↑ = 더 할인(빨강)
          const dn=c.delta!=null&&c.delta<0;   // 세일율 ↓ = 덜 할인(초록)
          const dCol=up?"#c0392b":dn?"#1a7a4f":DC.dim;
          const cell=(k,curV,expV,col)=>(
            <div style={{display:"grid",gridTemplateColumns:"1fr auto auto auto",gap:6,alignItems:"baseline",fontSize:11,padding:"2px 0"}}>
              <span style={{color:DC.sub}}>{k}</span>
              <span style={{color:DC.dim}}>{curV}</span>
              <span style={{color:DC.dim}}>→</span>
              <span style={{fontWeight:700,color:col||DC.text}}>{expV}</span>
            </div>
          );
          return(
          <div key={c.ch} style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:10,padding:"14px 16px"}}>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10}}>
              <span style={{width:9,height:9,borderRadius:"50%",background:chColor(c.ch),display:"inline-block"}}/>
              <span style={{fontSize:13,fontWeight:700,color:DC.text}}>{c.ch}</span>
              <span style={{fontSize:10,color:DC.dim,marginLeft:"auto"}}>수수료 {c.fee}%</span>
            </div>
            {/* 세일율: 현재 → 권장 (Δ) */}
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
              <div style={{textAlign:"center"}}>
                <div style={{fontSize:9,color:DC.dim}}>현재 세일율</div>
                <div style={{fontSize:19,fontWeight:800,color:DC.sub}}>{c.curRate==null?"—":`${c.curRate}%`}</div>
              </div>
              <div style={{fontSize:18,color:dCol,fontWeight:800}}>▶</div>
              <div style={{textAlign:"center"}}>
                <div style={{fontSize:9,color:DC.dim}}>권장 세일율</div>
                <div style={{fontSize:24,fontWeight:800,color:MUTE_BLUE,letterSpacing:"-0.5px"}}>{c.recRate==null?"—":`${c.recRate}%`}</div>
              </div>
              {c.delta!=null&&Math.abs(c.delta)>=0.05&&(
                <div style={{marginLeft:"auto",textAlign:"right"}}>
                  <div style={{fontSize:11,fontWeight:800,color:dCol}}>{up?"▲":"▼"} {Math.abs(c.delta)}%p</div>
                  <div style={{fontSize:9,color:DC.dim}}>{up?"더 할인":"덜 할인"}</div>
                </div>
              )}
            </div>
            <div style={{fontSize:10,color:dCol,marginBottom:8,minHeight:13}}>
              {c.delta==null?"":Math.abs(c.delta)<0.05?"현재 수준 유지 시 목표 도달":`세일율을 ${Math.abs(c.delta)}%p ${up?"높여야":"낮춰야"} 목표 달성`}
            </div>
            {/* 현재 → 예상 동시 표기 */}
            <div style={{borderTop:`1px solid ${DC.border}`,paddingTop:6}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr auto auto auto",gap:6,fontSize:9,color:DC.dim,marginBottom:2}}>
                <span/><span>최근 한달</span><span/><span>예상</span>
              </div>
              {cell("매출",won(cur30.revenue),won(c.targetRev),DC.text)}
              {cell("마진",won(cur30.margin),won(c.expMargin),c.expMargin>=0?"#1a7a4f":"#c0392b")}
              {cell("마진율",`${c.curMarginRate}%`,`${c.expMarginRate}%`,c.expMargin>=0?"#1a7a4f":"#c0392b")}
              {cell("재입고비",won(curRestock),won(c.restock),"#C8A87B")}
            </div>
            {/* 데이터 소스 (계산 근거 원본) */}
            <button onClick={()=>setSrcOpen(srcOpen===c.ch?null:c.ch)}
              style={{marginTop:8,background:"transparent",border:"none",color:DC.dim,fontSize:10,cursor:"pointer",padding:0,textDecoration:"underline"}}>
              {srcOpen===c.ch?"▾ 데이터 소스 닫기":"▸ 데이터 소스·계산 근거"}
            </button>
            {srcOpen===c.ch&&(
              <div style={{marginTop:6,padding:"8px 10px",background:DC.bg,borderRadius:6,fontSize:10,color:DC.sub,lineHeight:1.7}}>
                <div style={{color:DC.dim,marginBottom:3}}>출처: 주문·배송(이지어드민) status=배송, 최근 30일 · 정가·공급가=인벤토리 업로더 데이터(판매가/공급가)</div>
                <div>실판매 매출 = Σ(판매가×수량) = <b>{won(c.curRev)}</b> · 판매수량 = <b>{(recentActuals.chTotals[c.ch]?.qty||0).toLocaleString()}개</b></div>
                <div>정가 GMV = Σ(정가×수량) = <b>{won(c.listGmv)}</b></div>
                <div>현재 세일율 = 1 − (실판매 ÷ 정가GMV) = 1 − ({won(c.curRev)} ÷ {won(c.listGmv)}) = <b>{c.curRate==null?"—":`${c.curRate}%`}</b></div>
                <div>현재 마진 = Σ((정산액−공급가×1.1)×수량) = <b>{won(c.curMargin)}</b> · 수수료 {c.fee}%</div>
                <div>목표 매출 = 현재 매출 × 배수({targetMultiplier.toFixed(2)}) = <b>{won(c.targetRev)}</b></div>
                <div>권장 세일율 = 1 − (목표매출 ÷ 정가GMV) = <b>{c.recRate==null?"—":`${c.recRate}%`}</b></div>
              </div>
            )}
          </div>
          );
        })}
      </div>

      {/* 필요 마진 가드 */}
      <div style={{background:netContribution>=0?"rgba(26,122,79,0.07)":"rgba(192,57,43,0.07)",
        border:`1px solid ${netContribution>=0?"#1a7a4f":"#c0392b"}40`,borderRadius:10,padding:"12px 16px",marginBottom:20,fontSize:12,color:DC.text}}>
        예상 총마진 <b>{won(totalExpMargin)}</b> {netContribution>=0?"≥":"<"} 필요 총마진 <b>{won(reqMargin)}</b> →{" "}
        <b style={{color:netContribution>=0?"#1a7a4f":"#c0392b"}}>{netContribution>=0?`여유 ${won(netContribution)}`:`부족 ${won(-netContribution)} — 목표 이익금/고정비를 조정하세요`}</b>
        <span style={{color:DC.dim}}> · 총 목표 매출 {won(totalTargetRev)} · 예상 재입고비 {won(totalRestock)}</span>
      </div>

      {/* 상품별×채널별 매트릭스 — 2단계: 최근 한 달 실적 / 목표 도달 설정 */}
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"16px 18px",marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:13,fontWeight:700,color:DC.text}}>② 상품별 · 채널별 세일율 / 마진율 <span style={{fontSize:11,color:DC.dim,fontWeight:400}}>· ★=스테디셀러 · 행 클릭 시 계산식</span></span>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="상품명·옵션 검색" style={{...inBox,width:200}}/>
        </div>
        <div style={{fontSize:11,color:DC.dim,marginBottom:8}}>
          왼쪽 <b style={{color:DC.sub}}>현재</b>(최근 30일 실판매: 수량·세일율·마진율){recentActuals.start?` (${recentActuals.start}~${recentActuals.end})`:""} ▶ 오른쪽 <b style={{color:MUTE_BLUE}}>권장</b>(목표 도달 세일율·마진율, Δ)을 한 셀에 함께 표시
        </div>
        {unmatchedCount>0&&<div style={{fontSize:11,color:"#C8A87B",marginBottom:8}}>⚠ 공급가/판매가 미매칭 {unmatchedCount.toLocaleString()}건은 합계·표에서 제외됨 (인벤토리 가격 DB 보강 시 반영)</div>}
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{borderBottom:`2px solid ${DC.border}`}}>
              <th style={{textAlign:"left",padding:"6px 8px",color:DC.sub,fontWeight:600}}>상품</th>
              <th style={{textAlign:"center",padding:"6px 8px",color:DC.sub,fontWeight:600}}>에이징</th>
              <th style={{textAlign:"right",padding:"6px 8px",color:DC.sub,fontWeight:600}}>정가</th>
              <th style={{textAlign:"right",padding:"6px 8px",color:DC.sub,fontWeight:600}} title="최근 4주 판매수량(향후 판매 추정)">4주 판매</th>
              <th style={{textAlign:"right",padding:"6px 8px",color:DC.sub,fontWeight:600}} title="최신 입고 수량(stock_uploads)">입고</th>
              {GMV_CHANNELS.map(ch=>(
                <th key={ch} style={{textAlign:"center",padding:"6px 8px",color:chColor(ch),fontWeight:700,minWidth:150}}>{ch}<div style={{fontSize:9,color:DC.dim,fontWeight:400}}>현재 세일율·마진율 ▶ 권장 (Δ)</div></th>
              ))}
            </tr></thead>
            <tbody>
              {shownProducts.slice(0,showCount).map(p=>{
                const isExp=expanded===p.key;
                return(
                  <React.Fragment key={p.key}>
                    <tr onClick={()=>setExpanded(isExp?null:p.key)} style={{borderBottom:`1px solid ${DC.border}`,cursor:"pointer",background:isExp?"rgba(94,129,172,0.06)":"transparent"}}>
                      <td style={{padding:"6px 8px",color:DC.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {p.steady&&<span style={{color:"#C8A020",marginRight:3}}>★</span>}{p.name}{p.option?<span style={{color:DC.dim}}> / {p.option}</span>:null}
                      </td>
                      <td style={{padding:"6px 8px",textAlign:"center"}}><span style={{fontSize:10,fontWeight:600,color:INV_AGING_DEFS[p.aging]?.color}}>{INV_AGING_DEFS[p.aging]?.label||"—"}</span></td>
                      <td style={{padding:"6px 8px",textAlign:"right",color:DC.sub}}>{won(p.list)}</td>
                      <td style={{padding:"6px 8px",textAlign:"right",color:DC.sub}}>{p.qty4w.toLocaleString()}</td>
                      <td style={{padding:"6px 8px",textAlign:"right",color:DC.dim}} title={p.inboundDate?`입고일 ${p.inboundDate}`:""}>{p.inboundQty?p.inboundQty.toLocaleString():"—"}</td>
                      {channelCalc.map(c=>{
                        const a=recentActuals.byKeyCh[p.key+"@@"+c.ch];
                        const hasCur=a&&a.qty>0;
                        const avg=hasCur?a.revenue/a.qty:0;
                        const curDisc=hasCur&&p.list>0?Math.round((1-avg/p.list)*1000)/10:null;
                        const curMR=hasCur&&a.revenue>0?Math.round(a.margin/(a.revenue*(1-(c.fee||0)/100))*1000)/10:null;
                        const rec=gmvCompute(p.list,c.recRate||0,p.supply,c.fee);
                        const pDelta=(c.recRate!=null&&curDisc!=null)?Math.round((c.recRate-curDisc)*10)/10:null;
                        const pUp=pDelta!=null&&pDelta>0;
                        return(
                          <td key={c.ch} style={{padding:"6px 8px",textAlign:"center"}}>
                            <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
                              <div style={{minWidth:42}}>
                                <div style={{fontSize:9,color:DC.dim}}>{hasCur?`${a.qty}개`:"—"}</div>
                                <div style={{fontWeight:600,color:DC.sub}}>{curDisc==null?"—":`${curDisc}%`}</div>
                                <div style={{fontSize:10,color:curMR==null?DC.dim:(a.margin>=0?"#1a7a4f":"#c0392b")}}>{curMR==null?"":`${curMR}%`}</div>
                              </div>
                              <span style={{color:pDelta==null?DC.dim:(pUp?"#c0392b":"#1a7a4f"),fontWeight:800}}>▶</span>
                              <div style={{minWidth:42}}>
                                <div style={{fontSize:9,color:pDelta==null?DC.dim:(pUp?"#c0392b":"#1a7a4f")}}>{pDelta==null?"":`${pUp?"▲":"▼"}${Math.abs(pDelta)}%p`}</div>
                                <div style={{fontWeight:700,color:MUTE_BLUE}}>{c.recRate==null?"—":`${c.recRate}%`}</div>
                                <div style={{fontSize:10,fontWeight:600,color:rec.margin>=0?"#1a7a4f":"#c0392b"}}>{rec.marginRate}%</div>
                              </div>
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                    {isExp&&(
                      <tr style={{background:"rgba(94,129,172,0.04)"}}>
                        <td colSpan={5+GMV_CHANNELS.length} style={{padding:"4px 8px 12px"}}>
                          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:10}}>
                            {channelCalc.map(c=>{
                              const m=gmvCompute(p.list,c.recRate||0,p.supply,c.fee);
                              const a=recentActuals.byKeyCh[p.key+"@@"+c.ch];
                              const soldQty=a?a.qty:0;
                              const restock=Math.round(p.supply*1.1)*soldQty; // 예측 판매량(최근 30일) 기준
                              const calcRow=(k,v,expr)=>(
                                <div style={{display:"grid",gridTemplateColumns:"86px 92px 1fr",gap:8,fontSize:11,padding:"2px 0",alignItems:"baseline"}}>
                                  <span style={{color:DC.sub}}>{k}</span><span style={{textAlign:"right",fontWeight:600,color:DC.text}}>{v}</span><span style={{color:DC.dim}}>{expr}</span>
                                </div>
                              );
                              return(
                                <div key={c.ch} style={{border:`1px solid ${DC.border}`,borderRadius:8,padding:"10px 12px",background:DC.card}}>
                                  <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:6}}>
                                    <span style={{width:8,height:8,borderRadius:"50%",background:chColor(c.ch),display:"inline-block"}}/>
                                    <span style={{fontSize:12,fontWeight:700,color:DC.text}}>{c.ch}</span>
                                    <span style={{fontSize:10,color:DC.dim,marginLeft:"auto"}}>권장 세일율 {c.recRate==null?"—":`${c.recRate}%`}</span>
                                  </div>
                                  {calcRow("① 정가",won(p.list),"")}
                                  {calcRow("② 세일",`−${won(m.saleAmt)}`,`정가 × ${c.recRate||0}%`)}
                                  {calcRow("③ 실판매가",won(m.finalPrice),`정가 × (1−${c.recRate||0}%)`)}
                                  {calcRow("④ 수수료",`−${won(m.fee)}`,`실판매가 × ${c.fee}%`)}
                                  {calcRow("⑤ 정산액",won(m.net),"실판매가 − 수수료")}
                                  {calcRow("⑥ 공급가",`−${won(m.supplyVat)}`,`${won(p.supply)} × 1.1`)}
                                  {calcRow("⑦ 마진",won(m.margin),"정산액 − 공급가")}
                                  {calcRow("⑦ 마진율",`${m.marginRate}%`,"마진 ÷ 정산액")}
                                  {calcRow("⑧ 예상 재입고비",won(restock),`예측 판매 ${soldQty}개 × 공급가×1.1`)}
                                  {p.inboundQty>0&&calcRow("입고 흐름",`${p.inboundQty.toLocaleString()}개`,`최신 입고 ${p.inboundDate||""}`)}
                                </div>
                              );
                            })}
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
        {shownProducts.length>showCount&&(
          <button onClick={()=>setShowCount(c=>c+40)} style={{marginTop:10,background:"transparent",border:`1px solid ${DC.border}`,borderRadius:6,padding:"6px 14px",fontSize:12,color:DC.sub,cursor:"pointer"}}>
            더 보기 ({(shownProducts.length-showCount).toLocaleString()})
          </button>
        )}
      </div>

      {/* 입고 흐름 (stock_uploads) */}
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"16px 18px",marginBottom:16}}>
        <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:10,flexWrap:"wrap"}}>
          <span style={{fontSize:13,fontWeight:700,color:DC.text}}>③ 입고 흐름</span>
          <span style={{fontSize:11,color:DC.dim}}>입고 데이터(stock_uploads)로 본 현재 입고 흐름 · 최신 {inboundFlow.latestDate} · 누적 입고 {inboundFlow.totalInbound.toLocaleString()}개</span>
        </div>
        {inboundFlow.recent.length>0?(
          <div style={{display:"flex",alignItems:"flex-end",gap:10,flexWrap:"wrap"}}>
            {inboundFlow.recent.map((d,i)=>{
              const mx=Math.max(...inboundFlow.recent.map(x=>x.qty))||1;
              return(
                <div key={d.date} style={{flex:"1 1 80px",minWidth:70,textAlign:"center"}}>
                  <div style={{fontSize:11,fontWeight:700,color:DC.text,marginBottom:4}}>{d.qty.toLocaleString()}</div>
                  <div style={{height:Math.max(6,Math.round(d.qty/mx*70)),background:i===inboundFlow.recent.length-1?MUTE_BLUE:"#7B9EC8",borderRadius:4}}/>
                  <div style={{fontSize:10,color:DC.dim,marginTop:4}}>{d.date.slice(5)}</div>
                </div>
              );
            })}
          </div>
        ):(
          <div style={{fontSize:12,color:DC.dim,padding:"16px 0",textAlign:"center"}}>입고 데이터가 없습니다 — 데이터 입력 &gt; 입고 업로더로 등록하면 표시됩니다.</div>
        )}
        <div style={{fontSize:10,color:DC.dim,marginTop:8}}>입고일별 입고 수량 합계(최근 6개 업로드일). 상품별 최신 입고 수량은 ② 표의 ‘입고’ 컬럼에 표시됩니다.</div>
      </div>

      {/* 사이클 다이어그램: 원가 → 실판매가 → 마진 → 재입고 */}
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"16px 18px",marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:13,fontWeight:700,color:DC.text}}>③ 판매 → 마진 → 재입고 사이클</span>
          <div style={{display:"flex",gap:4}}>
            {GMV_CHANNELS.map(ch=>(
              <button key={ch} onClick={()=>setCycleCh(ch)} style={{background:cycleCh===ch?chColor(ch):"transparent",color:cycleCh===ch?"#fff":DC.sub,
                border:`1px solid ${cycleCh===ch?chColor(ch):DC.border}`,borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>{ch}</button>
            ))}
          </div>
        </div>
        {(()=>{
          const c=channelCalc.find(x=>x.ch===cycleCh)||channelCalc[0];
          if(!c) return null;
          // 원가 합계 = 예측 판매량 × 공급가×1.1 (= 예상 재입고비와 동일 모집단)
          const costTotal=c.restock;
          const nodes=[
            {k:"원가 합계",v:costTotal,color:"#c0392b",desc:"예측 판매량 × 공급가×1.1"},
            {k:"실제 판매가(목표)",v:c.targetRev,color:chColor(cycleCh),desc:`정가 GMV × (1−권장 ${c.recRate||0}%)`},
            {k:"마진",v:c.expMargin,color:c.expMargin>=0?"#1a7a4f":"#c0392b",desc:"정산액 − 원가"},
            {k:"재입고",v:c.restock,color:"#C8A87B",desc:"예측 판매량 × 공급가×1.1"},
          ];
          return(
            <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",justifyContent:"center"}}>
              {nodes.map((n,i)=>(
                <React.Fragment key={n.k}>
                  <div style={{flex:"1 1 130px",minWidth:120,maxWidth:180,textAlign:"center",border:`2px solid ${n.color}`,borderRadius:12,padding:"14px 10px",background:`${n.color}0d`}}>
                    <div style={{fontSize:11,color:DC.sub,marginBottom:3}}>{n.k}</div>
                    <div style={{fontSize:17,fontWeight:800,color:n.color,letterSpacing:"-0.3px"}}>{won(n.v)}</div>
                    <div style={{fontSize:9,color:DC.dim,marginTop:3}}>{n.desc}</div>
                  </div>
                  <div style={{color:DC.dim,fontSize:18,flexShrink:0}}>→</div>
                </React.Fragment>
              ))}
              <div style={{flex:"1 1 100%",textAlign:"center",fontSize:11,color:DC.dim,marginTop:4}}>↻ 재입고가 다시 원가로 순환 — 판매·마진으로 다음 재고를 매입하는 구조</div>
            </div>
          );
        })()}
      </div>

      {/* 채널 워크벤치 (업로드→자동산출→조정→재다운로드) */}
      <GmvWorkbench DC={DC} channelCalc={channelCalc} feeRates={feeRates}/>

      {/* 각주 */}
      <div style={{marginTop:18,padding:"12px 16px",background:DC.card,border:`1px solid ${DC.border}`,borderRadius:10}}>
        <div style={{fontSize:11,fontWeight:700,color:DC.sub,marginBottom:8,letterSpacing:".04em"}}>데이터 출처 · 최신 기준일</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:"6px 18px"}}>
          {footnotes.map(f=>(
            <div key={f.label} style={{fontSize:11,color:DC.sub}}>
              <b style={{color:DC.text}}>{f.label}</b> <span style={{color:MUTE_BLUE}}>{f.date}</span> <span style={{color:DC.dim}}>· {f.note}</span>
            </div>
          ))}
          <div style={{fontSize:11,color:DC.dim}}>카페24 매칭은 자사몰 코드 연결에 사용</div>
        </div>
        <div style={{fontSize:10,color:DC.dim,marginTop:8,lineHeight:1.6}}>
          · 채널 매출·점유율은 대시보드와 동일 소스(analyze)입니다. · 오프라인 POS 금액은 정산후(net)지만, 요청에 따라 29CM과 동일한 28% 수수료 모델로 비교합니다.
          · 마진율 = (정산액 − 공급가×1.1) ÷ 정산액. 자사몰은 수수료 0%(실수령=쿠폰적용가).
        </div>
      </div>
    </div>
  );
}

// GMV 채널 워크벤치 — 양식 업로드 → 세일율 자동/직접 조정 → 재다운로드 (29CM/자사몰/오프라인 토글)
function GmvWorkbench({DC,channelCalc,feeRates}){
  const [ch,setCh]=useState("29CM");
  const [rows,setRows]=useState([]);     // {code,name,list(selling),supply, rate, coupon}
  const [fileName,setFileName]=useState("");
  const [status,setStatus]=useState("");
  const fee=feeRates[ch]||0;
  const chRate=channelCalc.find(c=>c.ch===ch)?.recRate||0; // GMV 권장 세일율

  const loadRows=(prods)=>{
    setRows(prods.map(p=>({code:p.code||"",name:p.name,list:p.selling||0,supply:p.supply||0,rate:chRate,coupon:0})));
    setStatus(`${prods.length.toLocaleString()}개 상품 로드 · 세일율 GMV 추천값(${chRate}%)으로 자동 채움`);
  };
  const handleFile=f=>{
    if(!f) return;
    setFileName(f.name);setStatus("파싱 중…");
    parseMallProductFile(f,prods=>loadRows(prods),err=>setStatus("오류: "+err));
  };
  // 채널 변경/추천세일율 변경 시 아직 손대지 않은 행은 추천값 추적
  const setRate=(i,v)=>setRows(prev=>prev.map((r,idx)=>idx===i?{...r,rate:v,_touched:true}:r));
  const setCoupon=(i,v)=>setRows(prev=>prev.map((r,idx)=>idx===i?{...r,coupon:v}:r));
  const applyAll=v=>setRows(prev=>prev.map(r=>({...r,rate:v,_touched:true})));

  // 계산
  const calc=rows.map(r=>{
    const m=gmvCompute(r.list,parseFloat(r.rate)||0,r.supply,fee,parseFloat(r.coupon)||0);
    const recM=gmvCompute(r.list,chRate,r.supply,fee);
    return {...r,...m,recDisc:recM.effDisc,deviation:Math.round((m.effDisc-chRate)*10)/10};
  });
  const sumRev=calc.reduce((s,r)=>s+r.finalPrice,0); // 단가 합(수량 미반영, 상대 비교용)
  const sumMargin=calc.reduce((s,r)=>s+r.margin,0);

  const exportXlsx=async()=>{
    if(!calc.length) return;
    const XLSX=await getXLSX();
    let aoa;
    if(ch==="자사몰"){
      aoa=[["상품코드","상품명","할인 이후 가격","할인율(%)"],...calc.map(r=>[r.code,r.name,r.finalPrice,r.effDisc])];
    }else if(ch==="29CM"){
      aoa=[["상품코드","상품명","정가","세일율(%)","세일가","마진율(%)"],...calc.map(r=>[r.code,r.name,r.list,r.rate,r.finalPrice,r.marginRate])];
    }else{ // 오프라인 — 인벤토리 트렌드 형식 + 세일율·세일가·마진율
      aoa=[["상품코드","상품명","판매가","공급가","세일율(%)","세일가","마진율(%)"],...calc.map(r=>[r.code,r.name,r.list,r.supply,r.rate,r.finalPrice,r.marginRate])];
    }
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),ch.slice(0,20));
    XLSX.writeFile(wb,`GMV_${ch}_세일율_${dayjs().format("YYYYMMDD")}.xlsx`);
  };

  const inSm={width:56,textAlign:"right",background:"transparent",border:`1px solid ${DC.border}`,borderRadius:4,padding:"3px 6px",fontSize:12,color:DC.text,fontFamily:"inherit"};
  const won=n=>"₩"+Math.round(n||0).toLocaleString();

  return(
    <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"16px 18px",marginBottom:16}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
        <span style={{fontSize:13,fontWeight:700,color:DC.text}}>④ 채널 세일율 워크벤치</span>
        <span style={{fontSize:11,color:DC.dim}}>양식 업로드 → GMV 추천 세일율 자동 → 직접 조정 → 재다운로드</span>
        <div style={{display:"flex",gap:4,marginLeft:"auto"}}>
          {GMV_CHANNELS.map(c=>(
            <button key={c} onClick={()=>{setCh(c);setRows([]);setFileName("");setStatus("");}}
              style={{background:ch===c?chColor(c):"transparent",color:ch===c?"#fff":DC.sub,border:`1px solid ${ch===c?chColor(c):DC.border}`,
                borderRadius:6,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>{c}</button>
          ))}
        </div>
      </div>
      <div style={{fontSize:11,color:DC.sub,marginBottom:10}}>
        {ch==="29CM"&&"29CM 일괄할인/상품 양식(상품코드·상품명·판매가·공급가) 업로드 → 다운로드"}
        {ch==="자사몰"&&"자사몰 상품 파일(CSV/Excel: 상품코드·상품명·판매가·공급가) 업로드 → 다운로드"}
        {ch==="오프라인 스토어"&&"이지어드민 인벤토리 트렌드 양식 업로드 → 세일율·세일가 명시 파일 다운로드"}
        {" · 수수료 "}{fee}% · GMV 추천 세일율 <b style={{color:MUTE_BLUE}}>{chRate}%</b>
      </div>
      {/* 업로드 */}
      <div onClick={()=>{const inp=document.createElement("input");inp.type="file";inp.accept=".xlsx,.xls,.csv";inp.onchange=e=>handleFile(e.target.files[0]);inp.click();}}
        style={{border:`1.5px dashed ${DC.border}`,borderRadius:8,padding:"14px",textAlign:"center",cursor:"pointer",marginBottom:10,fontSize:12,color:DC.sub}}>
        ⬆ 파일 드래그&드롭 또는 클릭 — {fileName||"상품 파일 업로드"}
      </div>
      {status&&<div style={{fontSize:11,color:MUTE_BLUE,marginBottom:10}}>{status}</div>}
      {calc.length>0&&(
        <>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexWrap:"wrap"}}>
            <span style={{fontSize:11,color:DC.sub}}>전체 세일율 일괄:</span>
            {[0,5,10,15,20,chRate].filter((v,i,a)=>a.indexOf(v)===i).map(v=>(
              <button key={v} onClick={()=>applyAll(v)} style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:4,padding:"3px 9px",fontSize:11,color:DC.sub,cursor:"pointer"}}>{v}%</button>
            ))}
            <button onClick={exportXlsx} style={{marginLeft:"auto",background:chColor(ch),color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>↓ {ch} 형식 다운로드</button>
          </div>
          <div style={{display:"flex",gap:14,marginBottom:8,fontSize:11,color:DC.sub,flexWrap:"wrap"}}>
            <span>단가합 {won(sumRev)}</span><span>마진합 {won(sumMargin)}</span>
            <span style={{color:DC.dim}}>GMV 추천 세일율 {chRate}% 대비 이탈을 행별로 표시</span>
          </div>
          <div style={{overflowX:"auto",maxHeight:420,overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{position:"sticky",top:0,background:DC.bg}}>
                {["상품명","정가","세일율(%)","쿠폰(%)","세일가","목표 대비","마진","마진율"].map((h,i)=>(
                  <th key={h} style={{textAlign:i===0?"left":i<=1?"right":"center",padding:"6px 8px",fontSize:10,color:DC.sub,fontWeight:600,borderBottom:`1px solid ${DC.border}`,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {calc.map((r,i)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${DC.border}`}}>
                    <td style={{padding:"5px 8px",color:DC.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:DC.sub}}>{won(r.list)}</td>
                    <td style={{padding:"5px 8px",textAlign:"center"}}><input value={r.rate} onChange={e=>setRate(i,e.target.value)} type="number" style={inSm}/></td>
                    <td style={{padding:"5px 8px",textAlign:"center"}}><input value={r.coupon} onChange={e=>setCoupon(i,e.target.value)} type="number" style={inSm}/></td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:DC.text,fontWeight:600}}>{won(r.finalPrice)}</td>
                    <td style={{padding:"5px 8px",textAlign:"center",color:Math.abs(r.deviation)<0.05?DC.dim:(r.deviation>0?"#c0392b":"#1a7a4f")}}>{r.deviation>0?"+":""}{r.deviation}%p</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:r.margin>=0?"#1a7a4f":"#c0392b"}}>{won(r.margin)}</td>
                    <td style={{padding:"5px 8px",textAlign:"center",fontWeight:600,color:r.margin>=0?"#1a7a4f":"#c0392b"}}>{r.marginRate}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{fontSize:10,color:DC.dim,marginTop:8}}>
            쿠폰은 세일율 다음에 곱연산(자사몰 멤버십 쿠폰 방식). 29CM 채널부담 시나리오·자사몰 멤버십 쿠폰 모델을 동일 적용. 목표 대비 = 행 실효 세일율 − GMV 추천 세일율.
          </div>
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// 리오더 계산기 — 독립 페이지 (데이터 컴페어에서 분리)
// 계산 소스(computeAndSaveReorder)는 '데이터 입력 > 인벤토리' 업로더가 그대로 트리거
// ─────────────────────────────────────────────
function ReorderPage(){
  const reorderCardRef=useRef(null);
  const [snapshotDates,setSnapshotDates]=useState([]);
  const DC={bg:"#f8f8f6",card:"#ffffff",border:"#e0e0da",text:"#111111",sub:"#444444",dim:"#888888"};
  useEffect(()=>{(async()=>{
    const db=await getSupabase();
    const{data}=await db.from("inventory_snapshot").select("snapshot_date").order("snapshot_date",{ascending:false});
    if(data) setSnapshotDates([...new Set(data.map(r=>r.snapshot_date))]);
  })();},[]);
  const latestSnap=snapshotDates.length?[...snapshotDates].sort()[snapshotDates.length-1]:null;
  return(
    <div style={{background:"#f8f8f6",minHeight:"100%",padding:"28px 28px 40px"}}>
      <div style={{fontSize:13,color:"#888",marginBottom:2}}>
        데이터 소스: <b style={{color:"#444"}}>데이터 입력 &gt; 인벤토리</b> — 엑셀 업로드 시 리오더 데이터가 자동 계산됩니다.
      </div>
      <div ref={reorderCardRef} style={{position:"relative"}}>
        <div style={{position:"absolute",top:20,right:20,zIndex:10}}>
          <CaptureBtn cardRef={reorderCardRef} filename="리오더_계산기" DC={DC}/>
        </div>
        <ReorderCalculator DC={DC} refreshKey={0} latestSnapDate={latestSnap}/>
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
  const validPages=["dashboard","promo","input","compare","impact","reorder","gmv"];
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
      sale_price:r.amount||0,   // 매장 실판매금액(라인) — GMV 현재 세일율(정가 대비) 산출용
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
    {key:"impact",label:"콘텐츠 임팩트"},
    {key:"input",label:"데이터 입력"},
    {key:"reorder",label:"리오더 계산기"},
    {key:"gmv",label:"GMV 계산기"},
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
        {page==="promo"&&<PromoFlow revenues={revenues} storeSales={storeSales} orders={orders}/>}
        {page==="compare"&&<DataCompare revenues={revenues} storeSales={storeSales} orders={orders} stocks={stocks} ts={ts}/>}
        {page==="impact"&&<ContentImpact orders={orders} revenues={revenues} storeSales={storeSales}/>}
        {page==="reorder"&&<ReorderPage/>}
        {page==="gmv"&&<GmvCalculator orders={orders} revenues={revenues} storeSales={storeSales} stocks={stocks}/>}
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
