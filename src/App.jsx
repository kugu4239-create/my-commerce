import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend,
} from "recharts";

// ─────────────────────────────────────────────
// NOTE: Supabase 테이블 컬럼 추가 필요:
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS order_count integer DEFAULT 0;
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS refund_amount integer DEFAULT 0;
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS refund_count integer DEFAULT 0;
//   ALTER TABLE orders ADD COLUMN IF NOT EXISTS amount integer DEFAULT 0;
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

const toDate = raw => {
  if (!raw) return null;
  const s = String(raw).trim();
  // YYYY. M. D (점+공백 포함, 예: "2025. 10. 16")
  const m0 = s.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (m0) return `${m0[1]}-${m0[2].padStart(2,"0")}-${m0[3].padStart(2,"0")}`;
  const m1 = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m1) return `${m1[1]}-${m1[2].padStart(2,"0")}-${m1[3].padStart(2,"0")}`;
  const m2 = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m2) return `${m2[3]}-${m2[2].padStart(2,"0")}-${m2[1].padStart(2,"0")}`;
  const m3 = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m3) return `${m3[1]}-${m3[2]}-${m3[3]}`;
  return null;
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
  if (v.includes("취소")) return "반품";
  if (v.includes("교환")) return "교환";
  return "배송";
};

const fmtWon = n => {
  if (!n) return "—";
  if (n>=1e8) return "₩"+(n/1e8).toFixed(1)+"억";
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

function filterByDate(rows, dateField, period, customStart, customEnd) {
  if (period === "all") return rows;
  const today = localDate(0);
  if (period === "week") {
    const now = new Date();
    const dow = now.getDay() || 7;
    const monday = new Date(now);
    monday.setDate(now.getDate() - dow + 1);
    const cutStr = [monday.getFullYear(),String(monday.getMonth()+1).padStart(2,'0'),String(monday.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cutStr && r[dateField] <= today);
  }
  if (period === "yd") {
    const yStr = localDate(-1);
    return rows.filter(r => r[dateField] === yStr);
  }
  if (period === "7d") {
    const c = new Date(); c.setDate(c.getDate()-7);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "14d") {
    const c = new Date(); c.setDate(c.getDate()-14);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "1m") {
    const d=new Date(); d.setMonth(d.getMonth()-1);
    const cut=[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cut);
  }
  if (period === "3m") {
    const d=new Date(); d.setMonth(d.getMonth()-3);
    const cut=[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cut);
  }
  if (period === "6m") {
    const d=new Date(); d.setMonth(d.getMonth()-6);
    const cut=[d.getFullYear(),String(d.getMonth()+1).padStart(2,'0'),String(d.getDate()).padStart(2,'0')].join('-');
    return rows.filter(r => r[dateField] >= cut);
  }
  if (period === "custom" && customStart && customEnd) {
    return rows.filter(r => r[dateField] >= customStart && r[dateField] <= customEnd);
  }
  return rows;
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
function Card({ children, style={} }) {
  return (
    <div style={{ background:D.surface, border:`1px solid ${D.border}`,
      borderRadius:10, padding:"16px 18px", ...style }}>
      {children}
    </div>
  );
}
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
      padding:"9px 13px", color:c, fontSize:12, marginTop:9, lineHeight:1.5 }}>
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
const parseAnyFile=(file,opts,completeCb,errorCb)=>{
  const ext=file.name.split(".").pop().toLowerCase();
  if(ext==="xlsx"||ext==="xls"){
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        let data=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});
        if(opts.transformHeader) data=data.map(row=>{const nr={};Object.keys(row).forEach(k=>{nr[opts.transformHeader(k)]=row[k];});return nr;});
        completeCb({data});
      }catch(err){if(errorCb)errorCb(err);}
    };
    reader.readAsArrayBuffer(file);
  }else{
    Papa.parse(file,{...opts,complete:completeCb,error:errorCb});
  }
};

function DropZone({ onFile, label="파일을 드래그 앤 드롭 또는 클릭하여 선택", fileName="", required="", optional="" }) {
  const [hover,setHover]=useState(false);
  const handle=useCallback(e=>{
    e.preventDefault();
    const file=e.dataTransfer?.files?.[0]||e.target.files?.[0];
    if(file) onFile(file);
  },[onFile]);
  const cols=[...required.split("·").map(s=>s.trim()).filter(Boolean),...optional.split("·").map(s=>s.trim()).filter(Boolean)].filter(Boolean);
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
    const { error } = await db.from(table).delete().gte(dateField,start).lte(dateField,end);
    setLoading(false);
    if (error) { setResult({type:"error",msg:error.message}); setStep(0); }
    else {
      setResult({type:"success",msg:`${start} ~ ${end} 데이터 삭제 완료`});
      setStep(0); onDone?.();
    }
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
      입고 CSV 또는 이지어드민 CSV를 업로드하면<br/>상품별 물류 흐름이 표시됩니다
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
  // 매출 입력 데이터 기반 KPI
  const onlineRevenue   = revenueRows.reduce((s,r)=>s+(r.amount||0)-(r.refund_amount||0),0);
  const offlineRevenue  = storeRows.reduce((s,r)=>r.status==="배송"?s+(r.amount||0):r.status==="반품"?s-(r.amount||0):s,0);
  const totalRevenue    = onlineRevenue + offlineRevenue;
  const totalOrderCount = revenueRows.reduce((s,r)=>s+(r.order_count||0),0);
  const totalRefundAmt  = revenueRows.reduce((s,r)=>s+(r.refund_amount||0),0);
  const totalRefundCount= revenueRows.reduce((s,r)=>s+(r.refund_count||0),0);
  // 이지어드민 CSV 기반 KPI
  const totalShipped  = orderRows.filter(r=>r.status==="배송").length;
  const totalReturned = orderRows.filter(r=>r.status==="반품").length;
  const returnRate    = totalShipped>0?(totalReturned/totalShipped*100).toFixed(1):"0.0";
  const totalStock    = stockRows.reduce((s,r)=>s+(r.qty||0),0);

  // 판매처별 집계
  const byChannel={};
  revenueRows.forEach(r=>{
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    byChannel[ch].revenue+=(r.amount||0)-(r.refund_amount||0);
    byChannel[ch].orderCount+=(r.order_count||0);
    byChannel[ch].refundCount+=(r.refund_count||0);
  });
  // 채널별 배송 주문 집계: shipped/returned 카운트 + 객단가용 주문금액 맵
  // chOrderAmt[ch][oid] = 주문금액 (자사몰/무신사: 상품별 합산, 29CM: 최초값만)
  const chOrderAmt={};  // 객단가 계산용
  const chOrderIds={};  // uniqueOrders 카운트용 (offline 합산에도 사용)
  // 29CM은 주문번호당 총금액이 각 행에 중복 표시 → 첫 번째 값만
  const CH_SUM_ALL=new Set(["자사몰","무신사"]);
  orderRows.forEach(r=>{
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    if(r.status==="배송") byChannel[ch].shipped++;
    if(r.status==="반품") byChannel[ch].returned++;
    if(r.status!=="배송") return;
    const raw=(r.order_id||"");
    const oid=raw.split("||")[0]||raw;
    if(!chOrderIds[ch]) chOrderIds[ch]=new Set();
    chOrderIds[ch].add(oid);
    if(!chOrderAmt[ch]) chOrderAmt[ch]={};
    const amt=r.amount||0;
    if(CH_SUM_ALL.has(ch)){
      chOrderAmt[ch][oid]=(chOrderAmt[ch][oid]||0)+amt; // 개별 금액 합산
    } else {
      if(chOrderAmt[ch][oid]===undefined) chOrderAmt[ch][oid]=amt; // 29CM: 최초값만
    }
  });
  // 판교점+일산점 → 오프라인 스토어 합산
  const OFFLINE_CHS=new Set(["판교점","일산점","오프라인스토어","오프라인","오프라인 스토어"]);
  const offlineAgg={name:"오프라인 스토어",revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
  const offlineOrderIds=new Set();
  const offlineBreakdown={};  // 판교점/일산점 개별 보존
  let hasOffline=false;
  Object.keys(byChannel).forEach(ch=>{
    if(OFFLINE_CHS.has(ch)){
      hasOffline=true;
      offlineAgg.revenue+=byChannel[ch].revenue;
      offlineAgg.orderCount+=byChannel[ch].orderCount;
      offlineAgg.refundCount+=byChannel[ch].refundCount;
      offlineAgg.shipped+=byChannel[ch].shipped;
      offlineAgg.returned+=byChannel[ch].returned;
      (chOrderIds[ch]||new Set()).forEach(id=>offlineOrderIds.add(id));
      if(ch!=="오프라인스토어"&&ch!=="오프라인"&&ch!=="오프라인 스토어")
        offlineBreakdown[ch]={...byChannel[ch]};
      delete byChannel[ch];
      delete chOrderIds[ch];
    }
  });
  if(hasOffline){
    byChannel["오프라인 스토어"]=offlineAgg;
    chOrderIds["오프라인 스토어"]=offlineOrderIds;
  }
  // 매장 판매 CSV 기반 오프라인 매출 및 객단가 주문 ID
  if(storeRows.length){
    const storeByStore={};
    storeRows.forEach(r=>{
      const st=r.store_name||"오프라인 스토어";
      if(!storeByStore[st]) storeByStore[st]={revenue:0,orderIds:new Set()};
      if(r.status==="배송"){storeByStore[st].revenue+=(r.amount||0);if(r.order_id)storeByStore[st].orderIds.add(r.order_id);}
      else if(r.status==="반품") storeByStore[st].revenue-=(r.amount||0);
    });
    let storeTotal=0; const storeAllIds=new Set();
    Object.values(storeByStore).forEach(d=>{storeTotal+=d.revenue;d.orderIds.forEach(id=>storeAllIds.add(id));});
    if(!byChannel["오프라인 스토어"]) byChannel["오프라인 스토어"]={name:"오프라인 스토어",revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    byChannel["오프라인 스토어"].revenue=storeTotal;
    const existingIds=chOrderIds["오프라인 스토어"]||new Set();
    storeAllIds.forEach(id=>existingIds.add(id));
    chOrderIds["오프라인 스토어"]=existingIds;
    // store_sales 기반 판교점/일산점 breakdown 덮어쓰기
    Object.entries(storeByStore).forEach(([st,d])=>{
      if(st!=="오프라인 스토어") offlineBreakdown[st]={name:st,revenue:d.revenue};
    });
  }
  const channelList=Object.values(byChannel).sort((a,b)=>b.revenue-a.revenue||b.shipped-a.shipped);
  const totalRev=channelList.reduce((s,c)=>s+c.revenue,0)||1;
  channelList.forEach(c=>{
    c.share=((c.revenue||0)/totalRev*100).toFixed(1);
    c.returnRate=c.shipped>0?(c.returned/c.shipped*100).toFixed(1):"0.0";
    const uq=(chOrderIds[c.name]||new Set()).size||c.shipped;
    c.uniqueOrders=uq;
    if(c.name==="오프라인 스토어"){
      // 오프라인: 기존 방식 (store_sales CSV 매출 / unique 주문ID)
      c.avgOrderValue=(uq>0&&c.revenue>0)?Math.round(c.revenue/uq):0;
    } else {
      // 온라인: 이지어드민 CSV 주문금액 합 / unique 주문번호 수
      const orderMap=chOrderAmt[c.name]||{};
      const totalAmt=Object.values(orderMap).reduce((s,a)=>s+a,0);
      const orderCount=Object.keys(orderMap).length||uq;
      if(orderCount>0&&totalAmt>0){
        c.avgOrderValue=Math.round(totalAmt/orderCount);
      } else {
        // 주문 amount 미입력 시 revenues CSV 매출 기반 폴백
        c.avgOrderValue=(uq>0&&c.revenue>0)?Math.round(c.revenue/uq):0;
      }
    }
  });

  // 월별 배송/반품
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

  // 주간 상품 랭킹 (상품명 기준, 옵션 합산)
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

  // 상품명 기준 합산 (옵션 미구분)
  const byProd={};
  weekRows.forEach(r=>{
    const key=r.product_name||"미분류";
    if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
    byProd[key].qty+=(r.qty||0);
    byProd[key].orders++;
    if(r.status==="반품") byProd[key].returned++;
  });
  const prodList=Object.values(byProd);
  const weekBest=[...prodList].sort((a,b)=>b.qty-a.qty).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));
  const weekWorst=[...prodList].filter(p=>p.returned>0).sort((a,b)=>b.returned-a.returned).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));

  return {
    totalRevenue,totalOrderCount,totalRefundAmt,totalRefundCount,returnRate,
    totalShipped,totalReturned,totalStock,
    channelList,offlineBreakdown,monthlyData,weekBest,weekWorst,latestWeek,weekRows,
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

function Dashboard({ orders, stocks, revenues, storeSales=[], ts, onRefresh }) {
  const isMobile=useWindowWidth()<=1080;
  const [period,setPeriod]=useState("1m");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [deleteAll,setDeleteAll]=useState(false);
  const [shippingPeriod,setShippingPeriod]=useState("1m");
  const [returnPeriod,setReturnPeriod]=useState("1m");
  const [rankBestPeriod,setRankBestPeriod]=useState("yd");
  const [rankBestChannel,setRankBestChannel]=useState("전체");
  const [rankInfoOpen,setRankInfoOpen]=useState(false);
  const [rankBestCustomStart,setRankBestCustomStart]=useState("");
  const [rankBestCustomEnd,setRankBestCustomEnd]=useState("");
  const [rankWorstPeriod,setRankWorstPeriod]=useState("1m");
  const [rankWorstChannel,setRankWorstChannel]=useState("전체");
  const [rankWorstCustomStart,setRankWorstCustomStart]=useState("");
  const [rankWorstCustomEnd,setRankWorstCustomEnd]=useState("");
  const [chSort,setChSort]=useState({key:"revenue",dir:"desc"});
  const [optionPeriod,setOptionPeriod]=useState("1m");
  const [returnOptionPeriod,setReturnOptionPeriod]=useState("1m");
  const [offlineExpanded,setOfflineExpanded]=useState(false);
  const [kpiModal,setKpiModal]=useState(null); // "revenue"|"shipped"|"returnRate"|"stock"

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
  const optionFilteredOrders=useMemo(()=>filterByDate(orders,"order_date",optionPeriod,"",""),[orders,optionPeriod]);
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
  const returnOptionFilteredOrders=useMemo(()=>filterByDate(orders,"order_date",returnOptionPeriod,"",""),[orders,returnOptionPeriod]);
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
    filterByDate(orders,"order_date",rankBestPeriod,rankBestCustomStart,rankBestCustomEnd),
    [rankBestPeriod,orders,rankBestCustomStart,rankBestCustomEnd]);

  const bestRows=useMemo(()=>{
    const base=bestFilteredOrders;
    const rows=rankBestChannel==="전체"?base:base.filter(r=>r.channel===rankBestChannel);
    const byProd={};
    rows.forEach(r=>{
      const key=r.product_name||"미분류";
      if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
      byProd[key].qty+=(r.qty||0); byProd[key].orders++;
      if(r.status==="반품") byProd[key].returned++;
    });
    const totalQty=Object.values(byProd).reduce((s,p)=>s+p.qty,0)||1;
    return Object.values(byProd).sort((a,b)=>b.qty-a.qty).slice(0,20)
      .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0",
        share:(p.qty/totalQty*100).toFixed(1)}));
  },[bestFilteredOrders,rankBestChannel]);

  // 반품 Top 랭킹
  const worstFilteredOrders=useMemo(()=>
    filterByDate(orders,"order_date",rankWorstPeriod,rankWorstCustomStart,rankWorstCustomEnd),
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
    const today=localDate(0);
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
    const c=new Date(); c.setMonth(c.getMonth()-3);
    const cut=c.toISOString().slice(0,10);
    const byMonth={};
    orders.filter(r=>r.order_date>=cut).forEach(r=>{
      const ym=r.order_date?.slice(0,7); if(!ym) return;
      if(!byMonth[ym]) byMonth[ym]={date:ym,shipped:0};
      if(r.status==="배송") byMonth[ym].shipped++;
    });
    return Object.values(byMonth).sort((a,b)=>a.date>b.date?1:-1);
  },[orders,shippingPeriod]);

  // 일별 반품 by 채널 차트
  const returnChartData=useMemo(()=>{
    let start;
    if(returnPeriod==="yd"){
      start=localDate(-1);
    } else if(returnPeriod==="7d"){
      const d=new Date(); d.setDate(d.getDate()-7);
      start=d.toISOString().slice(0,10);
    } else {
      const d=new Date(); d.setMonth(d.getMonth()-(returnPeriod==="1m"?1:3));
      start=d.toISOString().slice(0,10);
    }
    const filteredRet=orders.filter(r=>r.order_date>=start&&r.channel!=="오프라인 스토어");
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
  },[orders,returnPeriod]);

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
          <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
            {PERIOD_TABS.map(({key,label})=><PeriodBtn key={key} k={key} l={label}/>)}
            {period==="custom"&&(
              <div style={{display:"flex",gap:4,alignItems:"center"}}>
                <input type="date" value={customStart} onChange={e=>setCustomStart(e.target.value)}
                  style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 7px",fontSize:11,color:D.text}}/>
                <span style={{color:D.textMeta,fontSize:11}}>—</span>
                <input type="date" value={customEnd} onChange={e=>setCustomEnd(e.target.value)}
                  style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"3px 7px",fontSize:11,color:D.text}}/>
              </div>
            )}
          </div>
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

      {/* KPI 카드 - 총 매출/주문/반품은 매출입력, 배송은 이지어드민 */}
      <div style={{display:"flex",gap:9,marginBottom:20,flexWrap:"wrap",minHeight:82}}>
        <KPI label="총 매출" value={fmtWon(stats.totalRevenue)} accent={D.black} onClick={()=>setKpiModal("revenue")}/>
        <KPI label="배송" value={stats.totalShipped.toLocaleString()+"건"} accent={D.green} onClick={()=>setKpiModal("shipped")}/>
        {!["yd","7d"].includes(period)&&<KPI label="반품률" value={stats.totalShipped>0?(stats.totalReturned/stats.totalShipped*100).toFixed(1)+"%":"0.0%"}
          sub={stats.totalReturned.toLocaleString()+"건"}
          accent={stats.totalShipped>0&&(stats.totalReturned/stats.totalShipped)>0.1?D.red:D.textSub}
          onClick={()=>setKpiModal("returnRate")}/>}
        <KPI label="입고 수량" value={stats.totalStock.toLocaleString()+"개"} accent={D.blue} onClick={()=>setKpiModal("stock")}/>
      </div>

      {/* 판매처 점유율 + 판매처별 매출 */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"280px 1fr",gap:10,marginBottom:20,minHeight:220}}>
        <Card>
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
        </Card>
        <Card>
          <SecTitle ts={ts.orders}>판매처별 매출</SecTitle>
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
                      <div style={{color:D.text,marginBottom:2}}>합계: <strong>{total.toLocaleString()}원</strong></div>
                    )}
                    {entries.map((p,i)=>(
                      <div key={i} style={{color:p.color||D.text}}>
                        {p.name}: <strong>{(p.value||0).toLocaleString()}원</strong>
                      </div>
                    ))}
                  </div>
                );
              };
              return(
                <BarChart data={chartData} layout="vertical" barCategoryGap="28%">
                  <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
                  <XAxis type="number" tick={axTick} tickFormatter={v=>v>=1e4?(v/1e4).toFixed(0)+"만":v}/>
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
        </Card>
      </div>

      {/* 판매처 상세 */}
      <Card style={{marginBottom:20,minHeight:380}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <SecTitle ts={ts.orders}>판매처 상세</SecTitle>
          <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
            {[["revenue","매출"],["share","점유율"],["shipped","배송"],...(!["yd","7d"].includes(period)?[["returned","반품"],["rate","반품률"]]:[]),["aov","객단가"]].map(([k,l])=>(
              <button key={k} onClick={()=>setChSort({key:k,dir:"desc"})}
                style={{background:chSort.key===k?D.black:"transparent",
                  color:chSort.key===k?"#fff":D.textSub,
                  border:`1px solid ${chSort.key===k?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer",
                  fontWeight:600,minWidth:36,boxSizing:"border-box"}}>
                {l}
              </button>
            ))}
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
          const cols=[
            {key:"name",   label:"판매처",     left:true, val:c=>c.name,               w:hasRet?"22%":"26%"},
            {key:"share",  label:"점유율",                val:c=>parseFloat(c.share),  w:hasRet?"8%":"9%"},
            {key:"revenue",label:"매출",                  val:c=>c.revenue,            w:hasRet?"15%":"18%"},
            {key:"cmp",    label:"동기간 비교",            val:c=>0,                    w:hasRet?"14%":"17%"},
            {key:"shipped",label:"배송",                  val:c=>c.shipped,            w:hasRet?"8%":"10%"},
            ...(hasRet?[
              {key:"returned",label:"반품",               val:c=>c.returned,           w:"8%"},
              {key:"rate",   label:"반품률",              val:c=>c.shipped>0?c.returned/c.shipped:0, w:"9%"},
            ]:[]),
            {key:"aov",    label:"객단가",                val:c=>c.avgOrderValue||0,   w:hasRet?"16%":"20%", tooltip:"개발 중인 지표입니다.\n온라인 채널의 경우 '배송 단위'의 객단가이므로 오늘의 매출과 연관이 없습니다.", wip:true},
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
                      <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.revenue>0?fmtWon(c.revenue):"—"}</td>
                      <td style={{textAlign:"right",padding:"7px 9px"}}>
                        {!c.isSubRow&&prevPeriod?fmtChg(c.revenue,prev)||<span style={{color:D.textMeta,fontSize:10}}>—</span>:<span style={{color:D.textMeta,fontSize:10}}>—</span>}
                        {!c.isSubRow&&prevPeriod&&<div style={{fontSize:9,color:"#bbb",marginTop:1}}>{prevPeriod.start}~{prevPeriod.end}</div>}
                      </td>
                      <td style={{textAlign:"right",padding:"7px 9px",color:D.green}}>{(c.shipped||0).toLocaleString()}</td>
                      {hasRet&&<td style={{textAlign:"right",padding:"7px 9px"}}>{(c.returned||0).toLocaleString()}</td>}
                      {hasRet&&<td style={{textAlign:"right",padding:"7px 9px",fontWeight:600}}>
                        {c.shipped>0?(c.returned/c.shipped*100).toFixed(1):"0.0"}%</td>}
                      <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.avgOrderValue>0?fmtWon(c.avgOrderValue):"—"}</td>
                    </tr>
                  );
                })}
                <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                  <td style={{padding:"7px 9px"}}>합계</td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>100%</td>
                  <td style={{textAlign:"right",padding:"7px 9px"}}>{fmtWon(stats.totalRevenue)}</td>
                  <td/>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.green}}>{stats.totalShipped.toLocaleString()}</td>
                  {hasRet&&<td style={{textAlign:"right",padding:"7px 9px"}}>{stats.totalReturned.toLocaleString()}</td>}
                  {hasRet&&<td style={{textAlign:"right",padding:"7px 9px"}}>{stats.totalShipped>0?(stats.totalReturned/stats.totalShipped*100).toFixed(1):"0.0"}%</td>}
                  <td/>
                </tr>
              </tbody>
            </table>
            </div>
          );
        })()}
      </Card>

      {/* 월별 배송량 (독립 기간) */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:10,marginBottom:20}}>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <SecTitle ts={ts.orders}>배송량</SecTitle>
            <div style={{display:"flex",gap:4}}>
              {[["yd","어제"],["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([v,l])=>(
                <SmPeriodBtn key={v} val={v} cur={shippingPeriod} onChange={setShippingPeriod} label={l}/>
              ))}
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
        </Card>

        {/* 판매처별 일자 반품 */}
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <SecTitle ts={ts.orders}>판매처별 반품 추이</SecTitle>
            <div style={{display:"flex",gap:4}}>
              {[["yd","어제"],["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([v,l])=>(
                <SmPeriodBtn key={v} val={v} cur={returnPeriod} onChange={setReturnPeriod} label={l}/>
              ))}
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
        </Card>
      </div>

      {/* 판매 Top */}
      <Card style={{marginBottom:20,minHeight:660}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <SecTitle ts={ts.orders}>판매 Top</SecTitle>
            <InfoBtn onClick={()=>setRankInfoOpen(true)}/>
            {["전체",...activeChannels].map(ch=>(
              <button key={ch} onClick={()=>setRankBestChannel(ch)}
                style={{background:rankBestChannel===ch?D.black:"transparent",
                  color:rankBestChannel===ch?"#fff":D.textSub,
                  border:`1px solid ${rankBestChannel===ch?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>{ch}</button>
            ))}
          </div>
          <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
            {[["yd","어제"],["7d","최근 7일"],["14d","최근 14일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([k,l])=>(
              <button key={k} onClick={()=>setRankBestPeriod(k)}
                style={{background:rankBestPeriod===k?D.black:"transparent",
                  color:rankBestPeriod===k?"#fff":D.textSub,
                  border:`1px solid ${rankBestPeriod===k?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>{l}</button>
            ))}
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:14,alignItems:"start"}}>
          <div style={{minHeight:546,overflowY:"auto"}}>
            <RankTable data={bestRows} cols={[
              {key:"name",label:"상품명",maxW:190,bold:true,color:"#2d2d2d"},
              {key:"qty",label:"배송량",right:true,bold:true,fmt:v=>v.toLocaleString()},
              {key:"share",label:"배송 점유율",right:true,color:D.textMeta,fmt:v=>v+"%"},
            ]}/>
          </div>
          <ResponsiveContainer width="100%" height={546}>
            <BarChart data={bestRows.slice(0,12)} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
              <XAxis type="number" tick={axTick}/>
              <YAxis type="category" dataKey="name" width={180} tick={<NoWrapTick/>}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey="qty" name="배송량" radius={[0,3,3,0]}>
                {bestRows.slice(0,12).map((_,i)=>(
                  <Cell key={i} fill={D.SANKEY[i%D.SANKEY.length]}/>
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>

      {/* 플랫폼별 선호 옵션 */}
      {optionStats.length>0&&(
        <Card style={{marginBottom:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <SecTitle ts={ts.orders}>플랫폼별 선호 옵션</SecTitle>
            <div style={{display:"flex",gap:4}}>
              {[["1m","1달"],["3m","3달"],["6m","6달"]].map(([k,l])=>(
                <button key={k} onClick={()=>setOptionPeriod(k)}
                  style={{background:optionPeriod===k?D.black:"transparent",color:optionPeriod===k?"#fff":D.textSub,
                    border:`1px solid ${optionPeriod===k?D.black:D.border}`,borderRadius:5,
                    padding:"3px 9px",fontSize:10,cursor:"pointer",fontWeight:optionPeriod===k?600:400}}>{l}</button>
              ))}
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
        </Card>
      )}

      {rankInfoOpen&&(
        <div onClick={()=>setRankInfoOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.35)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",borderRadius:12,padding:28,maxWidth:340,width:"90%"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>판매 Top 안내</div>
            <p style={{fontSize:12,color:"#555",lineHeight:1.7,margin:"0 0 16px"}}>정확한 데이터 수집을 위해 판매 수치는 배송까지 완료된 건으로 카운트 됩니다.</p>
            <button onClick={()=>setRankInfoOpen(false)} style={{width:"100%",padding:"8px",borderRadius:7,border:"1px solid #e0e0e0",cursor:"pointer",fontSize:12}}>닫기</button>
          </div>
        </div>
      )}
      {/* 반품 탑 */}
      <Card style={{marginBottom:20,minHeight:660}}>
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
          <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
            {[["1m","최근 한달"],["3m","최근 3개월"]].map(([k,l])=>(
              <button key={k} onClick={()=>setRankWorstPeriod(k)}
                style={{background:rankWorstPeriod===k?D.black:"transparent",
                  color:rankWorstPeriod===k?"#fff":D.textSub,
                  border:`1px solid ${rankWorstPeriod===k?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>{l}</button>
            ))}
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
      </Card>

      {/* 플랫폼별 반품률 높은 옵션 */}
      {returnOptionStats.length>0&&(
        <Card style={{marginBottom:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <SecTitle ts={ts.orders}>플랫폼별 반품률 높은 옵션</SecTitle>
            <div style={{display:"flex",gap:4}}>
              {[["1m","1달"],["3m","3달"],["6m","6달"]].map(([k,l])=>(
                <button key={k} onClick={()=>setReturnOptionPeriod(k)}
                  style={{background:returnOptionPeriod===k?D.black:"transparent",color:returnOptionPeriod===k?"#fff":D.textSub,
                    border:`1px solid ${returnOptionPeriod===k?D.black:D.border}`,borderRadius:5,
                    padding:"3px 9px",fontSize:10,cursor:"pointer",fontWeight:returnOptionPeriod===k?600:400}}>{l}</button>
              ))}
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
                        <div style={{fontSize:9,color:D.textMeta,marginTop:1}}>{returned}건 반품 / {total}건</div>
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
                        <div style={{fontSize:9,color:D.textMeta,marginTop:1}}>{returned}건 반품 / {total}건</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}

      {/* 전체 데이터 삭제 */}
      <div style={{marginTop:24,paddingTop:16,borderTop:`1px solid ${D.border}`,display:"flex",justifyContent:"flex-end"}}>
        {!deleteAll?(
          <button onClick={()=>setDeleteAll(true)}
            style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
              padding:"6px 14px",fontSize:11,cursor:"pointer",color:D.textMeta}}>
            ⚠ 전체 데이터 초기화
          </button>
        ):(
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontSize:11,color:D.red}}>모든 주문·입고·매출 데이터가 삭제됩니다. 확인하시겠습니까?</span>
            <button onClick={async()=>{
              const db=await getSupabase();
              await Promise.all([
                db.from("orders").delete().gte("order_date","2000-01-01"),
                db.from("stock_uploads").delete().gte("upload_date","2000-01-01"),
                db.from("revenues").delete().gte("date","2000-01-01"),
                db.from("store_sales").delete().gte("sale_date","2000-01-01"),
                db.from("cs_data").delete().gte("id",1),
              ]);
              try{localStorage.removeItem("cs_data");}catch{}
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
                        <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWon(d.revenue)}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{d.refund>0?fmtWon(d.refund):"—"}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",fontWeight:700}}>{fmtWon(d.revenue-d.refund)}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWon(chRows.reduce((s,[,d])=>s+d.revenue,0))}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{fmtWon(chRows.reduce((s,[,d])=>s+d.refund,0))}</td>
                      <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWon(stats.totalRevenue)}</td>
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
                        <td style={{textAlign:"right",padding:"5px 7px"}}>{fmtWon(v.revenue)}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{v.refund>0?fmtWon(v.refund):"—"}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",fontWeight:600}}>{fmtWon(v.revenue-v.refund)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            </div>
          );
        }

        /* ── 배송 ── */
        else if(kpiModal==="shipped"){
          modalTitle="배송 소스";
          const shipped=filteredOrders.filter(r=>r.status==="배송");
          const byCh={};
          shipped.forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={cnt:0,oids:new Set()};
            byCh[ch].cnt++;
            const oid=(r.order_id||"").split("||")[0]||r.order_id||"";
            if(oid) byCh[ch].oids.add(oid);
          });
          const chRows=Object.entries(byCh).sort((a,b)=>b[1].cnt-a[1].cnt);
          const byDate={};
          shipped.forEach(r=>{
            const d=r.order_date||"—";
            if(!byDate[d]) byDate[d]=0;
            byDate[d]++;
          });
          const dateRows=Object.entries(byDate).sort((a,b)=>b[0]>a[0]?1:-1).slice(0,30);
          modalContent=(
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>채널별</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>채널</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송건수</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>주문수(고유)</th>
                  </tr></thead>
                  <tbody>
                    {chRows.map(([ch,d])=>(
                      <tr key={ch} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",fontWeight:600}}>{ch}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{d.cnt.toLocaleString()}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.textSub}}>{d.oids.size.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{stats.totalShipped.toLocaleString()}</td>
                      <td/>
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
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송건수</th>
                  </tr></thead>
                  <tbody>
                    {dateRows.map(([d,cnt])=>(
                      <tr key={d} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.textMeta}}>{d}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.green,fontWeight:600}}>{cnt.toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            </div>
          );
        }

        /* ── 반품률 ── */
        else if(kpiModal==="returnRate"){
          modalTitle="반품 소스";
          const byCh={};
          filteredOrders.filter(r=>r.status==="배송"||r.status==="반품").forEach(r=>{
            const ch=normCh(r.channel);
            if(!byCh[ch]) byCh[ch]={shipped:0,returned:0};
            if(r.status==="배송") byCh[ch].shipped++;
            if(r.status==="반품") byCh[ch].returned++;
          });
          const chRows=Object.entries(byCh).sort((a,b)=>b[1].returned-a[1].returned);
          // top return products
          const byProd={};
          filteredOrders.filter(r=>r.status==="반품").forEach(r=>{
            const k=(r.product_name||"미분류")+(r.option_name?" / "+r.option_name:"");
            if(!byProd[k]) byProd[k]=0;
            byProd[k]++;
          });
          const prodRows=Object.entries(byProd).sort((a,b)=>b[1]-a[1]).slice(0,20);
          modalContent=(
            <div>
            <div style={{color:D.blue,fontSize:11,marginBottom:14,lineHeight:1.6}}>
              반품률의 경우 발송일 기준으로는 0%, 배송 완료 이후 시점부터 반품 접수가 시작되므로 최근 한달 또는 최근 3개월 데이터를 보는 것이 가장 정확합니다.
            </div>
            <div style={{display:"grid",gridTemplateColumns:["yd","7d"].includes(period)?"1fr":"1fr 1fr",gap:20}}>
              <div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>채널별 반품률</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>채널</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>배송</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품률</th>
                  </tr></thead>
                  <tbody>
                    {chRows.map(([ch,d])=>{
                      const rate=d.shipped>0?(d.returned/d.shipped*100):0;
                      return(
                        <tr key={ch} style={{borderBottom:`1px solid ${D.border}`}}>
                          <td style={{padding:"5px 7px",fontWeight:600}}>{ch}</td>
                          <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{d.shipped.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{d.returned.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"5px 7px",fontWeight:700,color:rate>10?D.red:D.textSub}}>{rate.toFixed(1)}%</td>
                        </tr>
                      );
                    })}
                    <tr style={{borderTop:`2px solid ${D.border}`,fontWeight:700}}>
                      <td style={{padding:"5px 7px"}}>합계</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.green}}>{stats.totalShipped.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px",color:D.red}}>{stats.totalReturned.toLocaleString()}</td>
                      <td style={{textAlign:"right",padding:"5px 7px"}}>{stats.totalShipped>0?(stats.totalReturned/stats.totalShipped*100).toFixed(1):"0.0"}%</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              {!["yd","7d"].includes(period)&&<div>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:D.textSub,letterSpacing:"0.08em",textTransform:"uppercase"}}>반품 Top 상품</div>
                <div style={{maxHeight:360,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{borderBottom:`1px solid ${D.border}`,color:D.textMeta}}>
                    <th style={{textAlign:"left",padding:"5px 7px",fontWeight:500}}>상품/옵션</th>
                    <th style={{textAlign:"right",padding:"5px 7px",fontWeight:500}}>반품수</th>
                  </tr></thead>
                  <tbody>
                    {prodRows.map(([k,cnt])=>(
                      <tr key={k} style={{borderBottom:`1px solid ${D.border}`}}>
                        <td style={{padding:"5px 7px",color:D.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k}</td>
                        <td style={{textAlign:"right",padding:"5px 7px",color:D.red,fontWeight:600}}>{cnt.toLocaleString()}</td>
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
          if(!data.length) throw new Error("데이터가 없습니다");
          const cols=Object.keys(data[0]);
          const lc=cols.map(c=>c.toLowerCase().replace(/[\s_]/g,""));
          const find=(...kws)=>{const i=lc.findIndex(c=>kws.some(k=>c.includes(k)));return i>=0?cols[i]:null;};
          const prodCol=find("상품명","상품","product","품명");
          const optCol=find("옵션","option","사이즈","색상");
          const qtyCol=find("수량","qty","quantity","개수","재고");
          const recCol=find("처음입고일","입고일","최초입고","first","입고");
          const shipCol=find("마지막배송일","최종배송일","마지막출고","last","배송일","출고일");
          if(!prodCol) throw new Error("상품명 컬럼을 찾을 수 없습니다");
          if(!recCol)  throw new Error("처음입고일 컬럼을 찾을 수 없습니다");
          const rows=data.map(r=>({
            product_name:String(r[prodCol]||"").trim(),
            option_name:optCol?String(r[optCol]||"").trim():"",
            qty:toNum(r[qtyCol]||"0"),
            first_received_date:toDate(r[recCol]),
            last_shipped_date:shipCol&&r[shipCol]?toDate(r[shipCol]):null,
            diagnosis_date:diagDate,
          })).filter(r=>r.product_name&&r.first_received_date);
          if(!rows.length) throw new Error("유효한 데이터가 없습니다");
          setPreview(rows);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e.message}));
  },[diagDate]);

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
    setResult({type:"success",msg:`${preview.length}건 저장 완료`});
    setPreview(null); setFileName("");
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
        required="상품명 · 처음입고일"
        optional="옵션 · 수량 · 마지막배송일"/>
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
      {result&&<div style={{marginTop:8,fontSize:11,color:result.type==="error"?D.red:D.green}}>{result.msg}</div>}
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


  const PBtn=({k,l})=>(
    <button onClick={()=>setPeriod(k)}
      style={{background:period===k?D.black:"transparent",
        color:period===k?"#fff":D.textSub,
        border:`1px solid ${period===k?D.black:D.border}`,
        borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:period===k?600:400}}>
      {l}
    </button>
  );

  return (
    <div style={{padding:"20px 24px",maxWidth:1600,margin:"0 auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
        <div>
          <div style={{color:D.black,fontWeight:600,fontSize:15,display:"flex",alignItems:"center",gap:8}}>
            물류 플로우 <UpdatedAt ts={ts.orders||ts.stock}/>
          </div>
        </div>
        <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
          {[["1m","1개월"],["3m","3개월"],["6m","6개월"],["all","전체"]].map(([k,l])=><PBtn key={k} k={k} l={l}/>)}
          {[["custom","기간 선택"]].map(([k,l])=><PBtn key={k} k={k} l={l}/>)}
          {period==="custom"&&(
            <div style={{display:"flex",gap:4,alignItems:"center"}}>
              <input type="date" value={customStart} onChange={e=>setCustomStart(e.target.value)}
                style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"4px 8px",fontSize:11,color:D.text}}/>
              <span style={{color:D.textMeta}}>—</span>
              <input type="date" value={customEnd} onChange={e=>setCustomEnd(e.target.value)}
                style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"4px 8px",fontSize:11,color:D.text}}/>
            </div>
          )}
        </div>
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

function PromoFlow({ revenues }) {
  const [promos,setPromos]=useState(getPromosCache);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({name:"",platform:"자사몰",start_date:"",end_date:"",memo:"",content:"",files:[]});
  const today=new Date().toISOString().slice(0,10);
  const [viewStart,setViewStart]=useState(()=>{const d=new Date();d.setDate(d.getDate()-30);return d.toISOString().slice(0,10);});
  const [viewEnd,setViewEnd]=useState(()=>{const d=new Date();d.setDate(d.getDate()+30);return d.toISOString().slice(0,10);});

  const [hoveredPromo,setHoveredPromo]=useState(null);
  const [fileAddTarget,setFileAddTarget]=useState(null);
  const fileInputRef=useRef(null);
  const [isDragging,setIsDragging]=useState(false);
  const dragRef=useRef(null);
  const formFileRef=useRef(null);
  const [formFileDragOver,setFormFileDragOver]=useState(false);
  const [tableFileDragOver,setTableFileDragOver]=useState(null);
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
  const [editingPromoId,setEditingPromoId]=useState(null);
  const [editPromoForm,setEditPromoForm]=useState({});
  const startEditPromo=p=>{setEditingPromoId(p.id);setEditPromoForm({name:p.name,platform:p.platform,start_date:p.start_date,end_date:p.end_date,content:p.content||p.memo||""});};
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
        const rows=data.map(p=>({...p,files:p.files||(p.file?[p.file]:[])}));
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
    setForm({name:"",platform:"자사몰",start_date:"",end_date:"",memo:"",content:"",files:[]});
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
    return Object.values(byDate).sort((a,b)=>a.date>b.date?1:-1);
  },[revenues,viewStart,viewEnd]);

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
          <input type="date" value={viewStart} onChange={e=>setViewStart(e.target.value)}
            style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"5px 8px",fontSize:13,color:D.text}}/>
          <span style={{color:D.textMeta}}>—</span>
          <input type="date" value={viewEnd} onChange={e=>setViewEnd(e.target.value)}
            style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"5px 8px",fontSize:13,color:D.text}}/>
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
                <DateButtonPicker value={form[field]} onChange={v=>setForm(f=>({...f,[field]:v}))}/>
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
                const activePromos=promos.filter(p=>
                  p.start_date.slice(0,10)<=fullDate&&p.end_date.slice(0,10)>=fullDate
                );
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
                    {activePromos.length>0&&(
                      <div style={{marginTop:8,paddingTop:6,borderTop:`1px solid ${D.border}`}}>
                        <div style={{fontSize:11,color:D.textMeta,marginBottom:4,letterSpacing:"0.05em"}}>진행 중인 프로모션</div>
                        {activePromos.map(p=>(
                          <div key={p.id} style={{display:"flex",alignItems:"center",gap:5,marginBottom:2}}>
                            <div style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),flexShrink:0}}/>
                            <span style={{color:D.textSub,fontSize:12}}>{p.platform}</span>
                            <span style={{fontWeight:600,fontSize:12,marginLeft:2}}>{p.name}</span>
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
      {promos.length>0&&(
        <Card>
          <div style={{fontWeight:600,fontSize:14,marginBottom:12,color:D.black}}>등록된 프로모션</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:D.surfaceAlt}}>
                {["채널","프로모션명","기간","상세 내용","첨부 파일","",""].map(h=>(
                  <th key={h} style={{padding:"5px 8px",textAlign:"left",fontWeight:600,
                    color:D.textSub,borderBottom:`1px solid ${D.border}`,fontSize:12,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...promos].sort((a,b)=>a.start_date>b.start_date?1:-1).map(p=>{
                const ended=isEnded(p);
                const isEditing=editingPromoId===p.id;
                const td={style:{padding:"6px 8px",borderBottom:`1px solid ${D.border}`,
                  color:ended?"#ccc":D.text,textDecoration:"none"}};
                const inp3={background:"transparent",border:`1px solid ${D.border}`,borderRadius:5,
                  padding:"7px 10px",fontSize:15,color:D.text,width:"100%",boxSizing:"border-box",
                  fontFamily:"'Pretendard','Noto Sans KR',sans-serif"};
                if(isEditing) return (
                  <tr key={p.id}>
                    <td colSpan={7} style={{padding:"10px 8px",borderBottom:`1px solid ${D.border}`,background:D.surfaceAlt}}>
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
  const [product,setProduct]=useState("");
  const [reason,setReason]=useState("");
  const [channel,setChannel]=useState("자사몰");
  const [filterProd,setFilterProd]=useState("");
  const [csvResult,setCsvResult]=useState(null);
  const [editCell,setEditCell]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [selected,setSelected]=useState(new Set());
  const [delConfirm,setDelConfirm]=useState(false);

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

  const inp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
    padding:"7px 10px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box"};

  const save=async()=>{
    if(!product.trim()||!reason.trim())return;
    const newR={id:Date.now(),date,product_name:product.trim(),return_reason:reason.trim(),channel};
    const next=[newR,...csData];
    saveCSData(next);setCSData(next);
    setProduct("");setReason("");
    const db=await getSupabase();
    await db.from("cs_data").insert(newR);
  };

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
          if(!prodCol)throw new Error("[상품] 컬럼을 찾을 수 없습니다. 헤더 확인: "+cols.join(", "));
          if(!reasonCol)throw new Error("[반품 사유] 컬럼을 찾을 수 없습니다. 헤더 확인: "+cols.join(", "));

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

  const filtered=csData.filter(r=>!filterProd||(r.product_name||"").includes(filterProd)||(r.date||"").includes(filterProd));

  return (
    <div style={{display:"grid",gridTemplateColumns:"320px 1fr",gap:14}}>
      <Card>
        <div style={{fontWeight:600,marginBottom:14,fontSize:13}}>CS 반품 사유 입력</div>
        <div style={{marginBottom:10}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:4}}>날짜</div>
          <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={inp}/>
        </div>
        <div style={{marginBottom:10}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:4}}>판매처</div>
          <div style={{display:"flex",gap:4}}>
            {["자사몰","29CM","무신사"].map(c=>(
              <button key={c} onClick={()=>setChannel(c)}
                style={{flex:1,background:channel===c?D.black:"transparent",
                  color:channel===c?"#fff":D.textSub,
                  border:`1px solid ${channel===c?D.black:D.border}`,
                  borderRadius:5,padding:"6px 4px",fontSize:11,cursor:"pointer"}}>
                {c}
              </button>
            ))}
          </div>
        </div>
        <div style={{marginBottom:10}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:4}}>상품명</div>
          <input value={product} onChange={e=>setProduct(e.target.value)} style={inp} placeholder="상품명 입력"/>
        </div>
        <div style={{marginBottom:14}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:4}}>반품 사유</div>
          <input value={reason} onChange={e=>setReason(e.target.value)} style={inp} placeholder="예: 사이즈 불일치, 불량, 단순 변심"/>
        </div>
        <button onClick={save}
          style={{width:"100%",background:D.black,color:"#fff",border:"none",borderRadius:6,
            padding:"10px",fontSize:12,cursor:"pointer",fontWeight:600}}>
          저장
        </button>
        <div style={{marginTop:16,paddingTop:14,borderTop:`1px solid ${D.border}`}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6,letterSpacing:"0.06em",textTransform:"uppercase"}}>CSV 일괄 업로드</div>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:8,lineHeight:1.6}}>
            필수 컬럼: <strong>[상품]</strong> · <strong>[반품 사유]</strong><br/>
            선택: [날짜] [판매처]
          </div>
          <DropZone onFile={handleCSVFile} label="반품 CS 파일 업로드"
            required="상품명 · 반품사유"
            optional="날짜 · 판매처"/>
          {csvResult&&<Alert type={csvResult.type} msg={csvResult.msg}/>}
        </div>
      </Card>
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <div style={{fontWeight:600,fontSize:13}}>반품 사유 내역</div>
          <input value={filterProd} onChange={e=>{setFilterProd(e.target.value);setDelConfirm(false);}}
            style={{...inp,width:200,fontSize:11,padding:"5px 8px"}} placeholder="날짜·상품명·판매처 검색"/>
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

function RevenueForm({ onUpdate }) {
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
    const {data}=await db.from("revenues").select("*").order("date",{ascending:false});
    // 대시보드와 동일하게 dedup (date+channel 기준 최신 id만)
    const revMap={};
    (data||[]).forEach(r=>{const k=`${r.date}__${r.channel}`;if(!revMap[k]||r.id>revMap[k].id)revMap[k]=r;});
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
    setResult({type:"success",msg:`${dates.length}일 저장 완료`,ts:ts2});
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
        setCsvPreview({error:`필수 컬럼 없음. 헤더: ${cols.join(", ")}`});return;
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
    setResult({type:"success",msg:`${toUpload.length}건 저장 완료`,ts:ts2});
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
          <div style={{display:"flex",gap:4,marginBottom:6}}>
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
            ? <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={inp}/>
            : <div style={{display:"flex",gap:4,alignItems:"center"}}>
                <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{...inp,width:"50%"}}/>
                <span style={{color:D.textMeta,flexShrink:0}}>~</span>
                <input type="date" value={dateEnd} onChange={e=>setDateEnd(e.target.value)} style={{...inp,width:"50%"}}/>
              </div>
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

        {amt&&<div style={{color:D.textSub,fontSize:11,marginBottom:10}}>
          ₩{Number(amt.replace(/,/g,"")||0).toLocaleString()}
        </div>}
        <Btn onClick={handleSave} disabled={loading} style={{width:"100%"}}>
          {loading?"저장 중...":"저장"}
        </Btn>
        {result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}

        <div style={{marginTop:14,borderTop:`1px solid ${D.border}`,paddingTop:14}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6}}>CSV 일괄 업로드</div>
          <DropZone onFile={handleCsvFile} label="매출 파일 업로드"
            required="날짜 · 매출금액"
            optional="판매처 · 주문수 · 환불금 · 환불수"/>
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
function StockUploader({ onUpdate }) {
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
    const {data}=await db.from("stock_uploads").select("*").order("upload_date",{ascending:false}).order("product_name");
    setHistory(data||[]); setHistTs(nowStr()); setSelected(new Set()); setDeleteConfirm(false);
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
          const f=detectFields(Object.keys(data[0]||{}));
          const rows=data.filter(r=>f.product&&r[f.product]).map(r=>({
            product_name:String(r[f.product]||"").trim(),
            option_name:String(r[f.option]||"").trim(),
            qty:toNum(r[f.qty]),
            memo:String(r[f.memo]||"").trim(),
          }));
          setPreview(rows); setStep(2);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e.message}));
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
            <DateRange start={startDate} end={endDate} onStart={setStartDate} onEnd={setEndDate}/>
            <div style={{color:D.red,fontSize:10,marginBottom:20}}>⚠ 확정 시 해당 기간 DB 데이터 전체 교체</div>
            <Btn onClick={confirmDate} disabled={!dateValid||loading} style={{width:"100%"}}>
              {loading?"조회 중...":"기간 확정"}
            </Btn>
          </>}
          {step===1&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>파일 업로드</div>
            <StatRow items={[{label:"삭제 예정",value:`${existing?.length||0}건`,color:D.red}]}/>
            <DropZone onFile={handleFile} fileName={fileName} label="입고 파일 업로드"
              required="상품명"
              optional="옵션 · 수량 · 메모"/>
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
                              fontSize:11,background:"transparent",color:D.text}}/>
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
function DataHistoryPanel({ table, dateField, searchFields, cols, editableCols=[], onChanged, placeholder="날짜·품목 검색" }) {
  const [rows,setRows]=useState([]);
  const [loading,setLoading]=useState(true);
  const [filter,setFilter]=useState("");
  const [selected,setSelected]=useState(new Set());
  const [editCell,setEditCell]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [deleteConfirm,setDeleteConfirm]=useState(false);
  const [result,setResult]=useState(null);

  useEffect(()=>{
    (async()=>{
      setLoading(true);
      const db=await getSupabase();
      let all=[];let from=0;const PAGE=1000;
      while(true){
        const {data,error}=await db.from(table).select("*")
          .order(dateField,{ascending:false}).range(from,from+PAGE-1);
        if(error||!data||data.length===0) break;
        all=all.concat(data);
        if(data.length<PAGE) break;
        from+=PAGE;
      }
      setRows(all);
      setLoading(false);
    })();
  },[table,dateField]);

  const filtered=filter
    ?rows.filter(r=>[...searchFields,dateField].some(f=>String(r[f]||"").includes(filter)))
    :rows;

  const toggleSelect=id=>setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>{
    const ids=filtered.map(r=>r.id);
    const allSel=ids.every(id=>selected.has(id));
    setSelected(s=>{const n=new Set(s);ids.forEach(id=>allSel?n.delete(id):n.add(id));return n;});
  };

  const handleDelete=async()=>{
    const db=await getSupabase();
    const cnt=selected.size;
    const {error}=await db.from(table).delete().in("id",[...selected]);
    if(error){setResult({type:"error",msg:error.message});return;}
    setRows(r=>r.filter(row=>!selected.has(row.id)));
    setSelected(new Set()); setDeleteConfirm(false);
    setResult({type:"success",msg:`${cnt}건 삭제 완료`});
    onChanged?.();
  };

  const startEdit=(id,field,val)=>{setEditCell({id,field});setEditVal(String(val??""));};
  const saveEdit=async()=>{
    if(!editCell) return;
    const db=await getSupabase();
    const {error}=await db.from(table).update({[editCell.field]:editVal}).eq("id",editCell.id);
    if(!error) setRows(rows=>rows.map(r=>r.id===editCell.id?{...r,[editCell.field]:editVal}:r));
    setEditCell(null);
  };

  const inp2={border:`1px solid ${D.border}`,borderRadius:6,padding:"5px 10px",fontSize:11,background:"transparent",color:D.text};

  return (
    <Card style={{marginTop:14}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:8,flexWrap:"wrap"}}>
        <span style={{fontWeight:600,fontSize:13}}>업로드 내역</span>
        <div style={{display:"flex",gap:8,alignItems:"center",flex:1,justifyContent:"flex-end"}}>
          <input placeholder={placeholder} value={filter} onChange={e=>{setFilter(e.target.value);setDeleteConfirm(false);}}
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
                  <input type="checkbox" checked={filtered.length>0&&filtered.every(r=>selected.has(r.id))} onChange={toggleAll}/>
                </th>
                {cols.map(c=>(
                  <th key={c.key} style={{padding:"5px 7px",textAlign:"left",color:D.textMeta,fontWeight:400,whiteSpace:"nowrap"}}>{c.label}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0,500).map(r=>(
                <tr key={r.id} style={{borderBottom:`1px solid ${D.border}`,background:selected.has(r.id)?"#f5f5f5":"transparent"}}>
                  <td style={{padding:"4px 7px"}}><input type="checkbox" checked={selected.has(r.id)} onChange={()=>toggleSelect(r.id)}/></td>
                  {cols.map(c=>{
                    const isEditing=editCell?.id===r.id&&editCell?.field===c.key;
                    const editable=editableCols.includes(c.key);
                    return(
                      <td key={c.key} style={{padding:"4px 7px",color:c.color||D.black,fontWeight:c.bold?600:400,
                        maxWidth:c.maxW||200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                        cursor:editable?"pointer":"default",userSelect:editable?"none":"auto"}}
                        title={editable?"더블클릭하여 수정":undefined}
                        onDoubleClick={editable?()=>startEdit(r.id,c.key,r[c.key]):undefined}>
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
      {result&&<Alert type={result.type} msg={result.msg}/>}
    </Card>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT — 이지어드민 CSV (배송일 기준)
// ─────────────────────────────────────────────
function EasyAdminUploader({ onUpdate }) {
  const today=new Date().toISOString().slice(0,10);
  const [startDate,setStartDate]=useState(today);
  const [endDate,setEndDate]=useState(today);
  const [step,setStep]=useState(0);
  const [fileName,setFileName]=useState("");
  const [parsedFile,setParsedFile]=useState(null); // 파일 파싱 결과 (업로드 전)
  const [allRows,setAllRows]=useState([]);
  const [inRange,setInRange]=useState([]);
  const [outRows,setOutRows]=useState([]);
  const [dupInfo,setDupInfo]=useState(null);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const dateValid=startDate&&endDate&&startDate<=endDate;

  // Step 0→1: 기간 확정
  const confirmDate=()=>{ if(dateValid) setStep(1); };

  // Step 1: 파일 선택 (파싱만, 업로드 X)
  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    parseAnyFile(file,{header:true,skipEmptyLines:true,transformHeader:h=>h.trim()},({data})=>{
        try{
          if(!data.length) throw new Error("데이터가 없습니다");
          const f=detectFields(Object.keys(data[0]));

          // 배송일 컬럼 명시적 우선 탐색 (Unicode 정규화 + exact match 우선)
          const allCols=Object.keys(data[0]);
          const nrm=s=>String(s).trim().normalize("NFC");
          const findCol=(...names)=>{
            // 1순위: exact match
            for(const n of names){ const c=allCols.find(h=>nrm(h)===nrm(n)); if(c) return c; }
            // 2순위: includes match
            for(const n of names){ const c=allCols.find(h=>nrm(h).includes(nrm(n))); if(c) return c; }
            return null;
          };
          const dateCol = findCol("배송일","배송일시","배송날짜","delivery_date")
                       || findCol("주문일","주문일시","주문날짜","order_date","날짜","date")
                       || f.date;
          const orderIdCol = findCol("주문번호","orderid") || findCol("관리번호","order_id") || f.orderId;
          const channelCol = findCol("판매처","channel","플랫폼","채널") || f.channel;
          const productCol = findCol("상품명","product","품명") || f.product;
          const optionCol  = findCol("옵션명","옵션","option") || f.option;
          const csCol      = findCol("CS","cs처리","cs상태","cs") || f.cs;
          const statusCol  = findCol("상태","status") || f.status;
          const qtyCol     = findCol("주문수량","수량","qty","quantity") || f.qty;
          const amtCol     = findCol("결제금액","판매금액","주문금액","실판매가","판매가","금액","amount","price") || f.revenue;

          if(!orderIdCol) throw new Error("관리번호 컬럼을 찾을 수 없습니다");
          if(!dateCol)    throw new Error(`배송일 컬럼을 찾을 수 없습니다 (컬럼: ${allCols.join(", ")})`);

          // 관리번호+상품명+옵션 기준 중복 합산
          const grouped={};
          data.filter(r=>r[orderIdCol]&&String(r[channelCol]||"").trim()!=="MERRYONOVERSEA").forEach(r=>{
            const oid=String(r[orderIdCol]).trim();
            const prod=String(r[productCol]||"").trim();
            const opt=String(r[optionCol]||"").trim();
            const ch=normChannel(r[channelCol]);
            const rawDate=r[dateCol];
            const dateVal=toDate(rawDate);
            const csRaw=csCol?String(r[csCol]||"").trim():"";
            const statusRaw=statusCol?String(r[statusCol]||"").trim():"";
            const status=csRaw?normCS(csRaw):(statusRaw?normCS(statusRaw):"배송");
            const qty=toNum(r[qtyCol])||1;
            const amt=amtCol?toNum(r[amtCol]):0;
            // 관리번호+상품명+옵션 조합을 DB key로 사용 → 같은 관리번호 내 여러 상품 허용
            const dbKey=`${oid}||${prod}||${opt}`;
            if(!grouped[dbKey]){
              grouped[dbKey]={order_id:dbKey,order_date:dateVal,channel:ch,
                product_name:prod,option_name:opt,qty:0,amount:0,status,raw_status:csRaw||statusRaw};
            }
            grouped[dbKey].qty+=qty;
            grouped[dbKey].amount+=amt;
            grouped[dbKey].status=status;
          });
          const parsed=Object.values(grouped);
          setParsedFile(parsed);
          setResult({type:"info",msg:`날짜 컬럼: "${dateCol}" | 관리번호: "${orderIdCol}" | ${parsed.length}행 파싱 완료`});
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e.message}));
  },[]);

  // Step 1→2: 미리보기 (파싱 결과만 확인, DB 조회 없음)
  const handlePreview=async()=>{
    if(!parsedFile?.length) {setResult({type:"error",msg:"파일을 먼저 선택해주세요"});return;}
    setLoading(true);
    const inR=parsedFile.filter(r=>r.order_date&&r.order_date>=startDate&&r.order_date<=endDate);
    const outR=parsedFile.filter(r=>!r.order_date||r.order_date<startDate||r.order_date>endDate);
    setInRange(inR); setOutRows(outR);
    setDupInfo({total:inR.length,newCount:inR.length,updateCount:0,sameCount:0});
    setLoading(false);
    setStep(2);
  };

  // Step 2→3: 확정 업로드 (기간 내 전체 삭제 후 재삽입 → 재업로드 시 완전 교체)
  const handleUpload=async()=>{
    if(!inRange.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    const {error:delErr}=await db.from("orders").delete().gte("order_date",startDate).lte("order_date",endDate);
    if(delErr){setResult({type:"error",msg:"삭제 실패: "+delErr.message});setLoading(false);return;}
    for(let i=0;i<inRange.length;i+=500){
      const {error}=await db.from("orders").upsert(inRange.slice(i,i+500),{onConflict:"order_id"});
      if(error){setResult({type:"error",msg:"삽입 실패: "+error.message});setLoading(false);return;}
    }
    const ts2=nowStr();
    await db.from("upload_logs").insert({
      upload_type:"orders",file_name:fileName,
      row_count:parsedFile?.length||0,
      inserted:inRange.length,updated:0,
      skipped:outRows.length,date_start:startDate,date_end:endDate,
    });
    setStep(3);
    setResult({type:"success",msg:`기간 내 ${inRange.length}건 등록 / 기간 외 ${outRows.length}건 제외`,ts:ts2});
    onUpdate(ts2); setLoading(false);
  };

  const reset=()=>{setStep(0);setAllRows([]);setInRange([]);setOutRows([]);setDupInfo(null);setFileName("");setParsedFile(null);setResult(null);};

  return (
    <div>
      <Steps current={step} steps={["배송일 기간","파일 선택","미리보기 확인","완료"]}/>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>배송일 기간 선택</div>
            <DateRange start={startDate} end={endDate} onStart={setStartDate} onEnd={setEndDate}/>
            <div style={{color:D.blue,fontSize:10,marginBottom:12,lineHeight:1.7}}>
              배송일 기준 · 관리번호 기준 upsert<br/>신규→추가 / 기존→상태업데이트
            </div>
            <Btn onClick={confirmDate} disabled={!dateValid} style={{width:"100%"}}>다음</Btn>
          </>}
          {step===1&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>파일 선택</div>
            <div style={{color:D.textMeta,fontSize:11,marginBottom:10,lineHeight:1.6}}>
              배송일 {startDate} ~ {endDate}
            </div>
            <DropZone onFile={handleFile} fileName={fileName} label="이지어드민 파일 선택"
              required="관리번호 · 배송일(또는 주문일)"
              optional="판매처 · 상품명 · 옵션 · 수량 · 결제금액 · CS처리"/>
            {result&&<Alert type={result.type} msg={result.msg}/>}
            <div style={{display:"flex",flexDirection:"column",gap:7,marginTop:12}}>
              <Btn onClick={handlePreview} disabled={!parsedFile||loading} style={{width:"100%"}}>
                {loading?"분석 중...":"미리보기"}
              </Btn>
              <button onClick={()=>setStep(0)}
                style={{background:"transparent",border:"none",color:D.textMeta,
                  fontSize:11,cursor:"pointer",padding:"5px"}}>← 기간 다시 선택</button>
            </div>
          </>}
          {step===2&&<>
            <div style={{fontWeight:600,marginBottom:12,fontSize:13}}>미리보기 확인</div>
            {dupInfo&&<StatRow items={[
              {label:"기간 내 등록",value:`${dupInfo.total}건`,color:D.green},
              {label:"기간 외 제외",value:`${outRows.length}건`,color:D.textMeta},
            ]}/>}
            {outRows.length>0&&<Alert type="warn" msg={`기간 밖 ${outRows.length}건 제외`}/>}
            <div style={{display:"flex",flexDirection:"column",gap:7,marginTop:12}}>
              <Btn onClick={handleUpload} disabled={loading||!inRange.length} style={{width:"100%"}}>
                {loading?"처리 중...":`확정 업로드 (${dupInfo?.total||0}건)`}
              </Btn>
              <button onClick={()=>setStep(1)}
                style={{background:"transparent",border:"none",color:D.textMeta,
                  fontSize:11,cursor:"pointer",padding:"5px"}}>← 파일 다시 선택</button>
            </div>
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
          </>}
          {step===3&&<div style={{textAlign:"center"}}>
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
                {key:"order_id",label:"관리번호",color:D.textMeta,maxW:90},
                {key:"order_date",label:"배송일",color:D.textMeta},
                {key:"channel",label:"판매처",bold:true},
                {key:"product_name",label:"상품명",maxW:140},
                {key:"option_name",label:"옵션",color:D.textMeta},
                {key:"qty",label:"수량",bold:true},
                {key:"status",label:"상태",fmt:v=>(
                  <span style={{color:v==="반품"?D.red:v==="교환"?D.amber:D.green,fontWeight:500}}>{v}</span>
                )},
              ]}
            />
          ):<div style={{color:D.textMeta,textAlign:"center",padding:80,fontSize:12}}>
            기간 선택 후 CSV 파일을 선택하고 미리보기를 누르세요
          </div>}
        </Card>
      </div>
      <DataHistoryPanel
        table="orders" dateField="order_date"
        searchFields={["product_name","channel","order_id","option_name"]}
        placeholder="날짜·상품명·판매처 검색"
        editableCols={["channel","status","product_name","option_name"]}
        cols={[
          {key:"order_date",label:"배송일",color:D.textMeta},
          {key:"channel",label:"판매처",bold:true},
          {key:"product_name",label:"상품명",maxW:180},
          {key:"option_name",label:"옵션",color:D.textMeta},
          {key:"qty",label:"수량"},
          {key:"status",label:"상태",fmt:v=><span style={{color:v==="반품"?D.red:v==="교환"?D.amber:D.green,fontWeight:500}}>{v}</span>},
          {key:"order_id",label:"관리번호",color:D.textMeta,maxW:120},
        ]}
        onChanged={()=>onUpdate(nowStr())}
      />
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT — 매장 판매 CSV
// ─────────────────────────────────────────────
function StoreUploader({ onUpdate }) {
  const [step,setStep]=useState(0);
  const [fileName,setFileName]=useState("");
  const [preview,setPreview]=useState(null);
  const [dateRange,setDateRange]=useState({start:"",end:""});
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);

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
          const rows=data.filter(r=>{
            const refPrice=parseKRW(r["기준판매가"]);
            return refPrice!==0; // 사은품·0원 상품 제외
          }).map(r=>{
            const qty=parseKRW(r["수량"]);
            const amount=parseKRW(r["실판매금액"]);
            return{
              sale_date:(r["구매일자"]||"").trim().slice(0,10),
              store_name:(r["매장"]||"").trim(),
              product_name:(r["상품명"]||"").trim(),
              option_name:(r["옵션"]||"").trim(),
              qty:Math.abs(qty),
              amount:Math.abs(amount),
              order_id:(r["ID"]||"").trim(),
              status:qty<0?"반품":"배송",
            };
          }).filter(r=>r.sale_date&&r.product_name&&r.qty>0);
          const dates=[...new Set(rows.map(r=>r.sale_date))].sort();
          setDateRange({start:dates[0]||"",end:dates[dates.length-1]||""});
          setPreview(rows); setStep(1);
        }catch(e){setResult({type:"error",msg:e.message});}
      },e=>setResult({type:"error",msg:e.message}));
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
    setStep(2);setResult({type:"success",msg:`${preview.length}건 등록 완료`,ts:ts2});
    onUpdate(ts2);setLoading(false);
  };

  const reset=()=>{setStep(0);setPreview(null);setFileName("");setResult(null);};

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<>
            <div style={{fontWeight:600,marginBottom:10,fontSize:13}}>매장 판매 CSV 업로드</div>
            <div style={{fontSize:11,color:D.textMeta,marginBottom:16,lineHeight:1.7}}>
              POS 시스템 판매 데이터를 업로드합니다.<br/>
              인식 컬럼: <b>구매일자</b>, <b>매장</b>, <b>상품명</b>, <b>옵션</b>,<br/>
              <b>수량</b> (괄호=반품), <b>실판매금액</b>, <b>ID</b> (객단가 분모)<br/>
              기준판매가=0인 사은품 행 자동 제외
            </div>
            <DropZone onFile={handleFile} fileName={fileName} label="매장 판매 파일 업로드"
              required="기준판매가 · 구매일자 · 상품명 · 수량"
              optional="매장 · 옵션 · 실판매금액 · ID"/>
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
          </>}
          {step===1&&<>
            <div style={{fontWeight:600,marginBottom:10,fontSize:13}}>업로드 확인</div>
            <StatRow items={[
              {label:"파일",value:fileName.slice(0,20)},
              {label:"기간",value:`${dateRange.start}~${dateRange.end}`},
              {label:"행수",value:`${preview?.length||0}건`,color:D.green},
            ]}/>
            <Btn onClick={handleUpload} disabled={loading} style={{width:"100%",marginBottom:7}}>
              {loading?"처리 중...":"업로드"}
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
                {key:"status",label:"상태",color:D.textMeta},
                {key:"order_id",label:"주문ID",color:D.textMeta},
              ]}/>
            </>
          )}
        </Card>
      </div>
      <DataHistoryPanel
        table="store_sales" dateField="sale_date"
        searchFields={["product_name","store_name","option_name"]}
        placeholder="날짜·상품명·매장 검색"
        editableCols={["store_name","product_name","option_name","amount","status"]}
        cols={[
          {key:"sale_date",label:"날짜",color:D.textMeta},
          {key:"store_name",label:"매장",bold:true},
          {key:"product_name",label:"상품명",maxW:180},
          {key:"option_name",label:"옵션",color:D.textMeta},
          {key:"qty",label:"수량"},
          {key:"amount",label:"금액"},
          {key:"status",label:"상태",fmt:v=><span style={{color:v==="반품"?D.red:D.green,fontWeight:500}}>{v}</span>},
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

  const lastDate=(arr,field)=>{
    const d=arr.map(r=>r[field]).filter(Boolean).sort().at(-1);
    return d?<span style={{fontSize:10,color:D.textMeta,fontWeight:400,marginLeft:4}}>({d})</span>:null;
  };

  const tabs=[
    {key:"revenue",name:"매출 입력",extra:lastDate(revenues,"date")},
    {key:"stock",name:"입고",extra:lastDate(stocks,"upload_date")},
    {key:"orders",name:"배송",extra:lastDate(orders,"order_date")},
    {key:"store",name:"매장 판매",extra:lastDate(storeSales,"sale_date")},
    {key:"cs",name:"CS"},
    {key:"delete",name:"데이터 삭제"},
  ];

  const GUIDES={
    revenue:"KPI 카드의 매출, 매출 점유율, 판매처별 매출의 소스입니다.\n매출 금액은 취소/환불이 포함된 금액이며, 엑셀 다운로드 시 각 채널 어드민의 통계에서 확인하세요.\n*매일 전날의 데이터를 업로드하세요.",
    stock:"KPI 카드의 입고 수량, 물류 플로우 섹션 전체의 데이터 소스입니다.\n*매일 전날의 데이터를 업로드하세요.",
    orders:"KPI 카드의 배송·반품 수, 판매처 상세의 배송·반품 수, 판매·반품 TOP, 플랫폼 별 선호·반품 옵션 랭킹, 객단가 계산의 데이터 소스입니다.\n*매일 최근 한달 데이터(주문건 반품 정보 업데이트)를 업로드하세요.",
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

      {tab==="revenue"&&<RevenueForm onUpdate={ts=>onUpdate("revenue",ts)}/>}
      {tab==="stock"&&<StockUploader onUpdate={ts=>onUpdate("stock",ts)}/>}
      {tab==="orders"&&<EasyAdminUploader onUpdate={ts=>{onUpdate("orders",ts);onDataChange?.();}}/>}
      {tab==="store"&&<StoreUploader onUpdate={ts=>{onUpdate("store",ts);onDataChange?.();}}/>}
      {tab==="cs"&&<CSDataInput/>}
      {tab==="delete"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:14}}>
          <DataDeleteSection table="revenues" dateField="date" label="매출 입력" onDone={()=>onDataChange?.()}/>
          <DataDeleteSection table="stock_uploads" dateField="upload_date" label="입고" onDone={()=>onDataChange?.()}/>
          <DataDeleteSection table="orders" dateField="order_date" label="배송" onDone={()=>onDataChange?.()}/>
          <DataDeleteSection table="store_sales" dateField="sale_date" label="매장 판매" onDone={()=>onDataChange?.()}/>
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
// DATA COMPARE
// ─────────────────────────────────────────────
const COMPARE_CH_COLOR={
  "자사몰":"#1D4ED8",
  "29CM":"#15803D",
  "무신사":"#6D28D9",
  "오프라인 스토어":"#B45309",
};
const COMPARE_CHANNELS=["자사몰","29CM","오프라인 스토어","무신사"];

function RevenueSankeyChart({periods,svgW}){
  const wrapRef=useRef(null);
  const [hoveredCh,setHoveredCh]=useState(null);
  const [selNodes,setSelNodes]=useState([]);   // max 2 [{key,pi,ch,amt,label}]
  const [modal,setModal]=useState(null);       // {x,y,a,b}

  const SVG_H=480,PAD_T=42,PAD_B=52,PAD_H=28,NODE_W=40,GAP=3,AVAIL_H=SVG_H-PAD_T-PAD_B;
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

  // 베지어 곡선 리본
  const links=useMemo(()=>{
    const res=[];
    for(let pi=0;pi<cols.length-1;pi++){
      COMPARE_CHANNELS.forEach(ch=>{
        const ln=cols[pi].nodes.find(n=>n.ch===ch);
        const rn=cols[pi+1].nodes.find(n=>n.ch===ch);
        if(!ln||!rn||ln.h<1||rn.h<1) return;
        const x1=ln.x+NODE_W,x2=rn.x,mx=(x1+x2)/2;
        const path=[
          `M${x1} ${ln.y}C${mx} ${ln.y},${mx} ${rn.y},${x2} ${rn.y}`,
          `L${x2} ${rn.y+rn.h}C${mx} ${rn.y+rn.h},${mx} ${ln.y+ln.h},${x1} ${ln.y+ln.h}Z`,
        ].join(" ");
        res.push({ch,path,color:COMPARE_CH_COLOR[ch]});
      });
    }
    return res;
  },[cols]);

  const fmtAmt=a=>{
    if(a>=1e8) return `${(a/1e8).toFixed(1)}억`;
    if(a>=1e4) return `${Math.round(a/1e4)}만`;
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
        // clamp modal so it stays in view
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
      onClick={()=>{setSelNodes([]);setModal(null);}}>
      <svg width={svgW} height={SVG_H} style={{overflow:"visible",display:"block"}}>
        <defs>
          {COMPARE_CHANNELS.map(ch=>{
            const id=`cg2_${ch.replace(/[^a-z0-9]/gi,"_")}`;
            const c=COMPARE_CH_COLOR[ch];
            return(
              <linearGradient key={ch} id={id} x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stopColor={c} stopOpacity={0.55}/>
                <stop offset="50%" stopColor={c} stopOpacity={0.2}/>
                <stop offset="100%" stopColor={c} stopOpacity={0.55}/>
              </linearGradient>
            );
          })}
        </defs>

        {/* 직각 리본 */}
        {links.map((l,i)=>(
          <path key={i} d={l.path}
            fill={`url(#cg2_${l.ch.replace(/[^a-z0-9]/gi,"_")})`}
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
                {/* 직각 노드 rx=0 */}
                <rect x={n.x} y={n.y} width={NODE_W} height={n.h}
                  fill={n.color} rx={0}
                  opacity={hoveredCh===null?1:hoveredCh===n.ch?1:0.18}
                  stroke={isSel(pi,n.ch)?"#fff":"none"} strokeWidth={2}
                  style={{transition:"opacity .15s"}}/>
                {/* 선택 번호 뱃지 */}
                {isSel(pi,n.ch)&&(
                  <text x={n.x+NODE_W/2} y={n.y+n.h/2} textAnchor="middle"
                    dominantBaseline="middle" fontSize={11} fontWeight={800} fill="#fff"
                    style={{pointerEvents:"none",userSelect:"none"}}>
                    {selIdx(pi,n.ch)+1}
                  </text>
                )}
                {/* 첫 번째 컬럼에만 채널명 (노드 왼쪽) */}
                {pi===0&&n.h>=10&&(
                  <text x={n.x-6} y={n.y+n.h/2}
                    textAnchor="end" dominantBaseline="middle"
                    fontSize={10} fontWeight={600} fill="#fff"
                    style={{pointerEvents:"none",userSelect:"none"}}>
                    {n.ch==="오프라인 스토어"?"오프라인":n.ch}
                  </text>
                )}
              </g>
            ))}
            {/* 일정 라벨 — 흰색 */}
            <text x={col.colX+NODE_W/2} y={SVG_H-PAD_B+18}
              textAnchor="middle" fontSize={10} fill="#fff" style={{pointerEvents:"none"}}>
              {col.label}
            </text>
            {/* 매출 합계 라벨 — 오른쪽 세로 방향 (노드 두께 = 금액) */}
            {col.total>0&&(()=>{
              const vis=col.nodes.filter(n=>n.h>0);
              if(!vis.length) return null;
              const top=Math.min(...vis.map(n=>n.y));
              const bot=Math.max(...vis.map(n=>n.y+n.h));
              const midY=(top+bot)/2;
              const rx=col.colX+NODE_W+11;
              return(
                <text x={rx} y={midY}
                  transform={`rotate(90 ${rx} ${midY})`}
                  textAnchor="middle" dominantBaseline="middle"
                  fontSize={9} fill="#fff" style={{pointerEvents:"none",userSelect:"none"}}>
                  {fmtAmt(col.total)}
                </text>
              );
            })()}
          </g>
        ))}
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
              background:"#0e0e0e",border:"1px solid #2a2a2a",borderRadius:10,
              padding:"14px 16px 12px",minWidth:192,
              boxShadow:"0 8px 32px #000c",zIndex:20,pointerEvents:"auto"}}>
            <button onClick={()=>{setSelNodes([]);setModal(null);}}
              style={{position:"absolute",top:7,right:9,background:"none",border:"none",
                color:"#555",cursor:"pointer",fontSize:15,lineHeight:1}}>✕</button>
            <div style={{fontSize:10,color:"#555",marginBottom:10,letterSpacing:"0.08em",textTransform:"uppercase"}}>매출 비교</div>
            {[a,b].map((nd,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <span style={{width:18,height:18,borderRadius:3,background:COMPARE_CH_COLOR[nd.ch]||"#444",
                  display:"inline-flex",alignItems:"center",justifyContent:"center",
                  fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>{i+1}</span>
                <div>
                  <div style={{fontSize:10,color:"#777"}}>{nd.label} · {nd.ch}</div>
                  <div style={{fontSize:13,color:"#fff",fontWeight:700}}>₩{nd.amt.toLocaleString()}</div>
                </div>
              </div>
            ))}
            <div style={{borderTop:"1px solid #222",marginTop:8,paddingTop:8,display:"flex",alignItems:"baseline",gap:8}}>
              <span style={{fontSize:15,fontWeight:800,color:up?"#4ade80":"#f87171"}}>
                {up?"▲":"▼"} {pct!==null?`${Math.abs(pct).toFixed(1)}%`:"—"}
              </span>
              <span style={{fontSize:11,color:"#666"}}>
                ({up?"+":""}{diff.toLocaleString()}원)
              </span>
            </div>
          </div>
        );
      })()}

      {/* 사용 안내 */}
      <div style={{marginTop:14,textAlign:"center",fontSize:15,color:"#fff",letterSpacing:"0.02em",userSelect:"none"}}>
        노드를 순서대로 두 번 탭 · 클릭하면 기간/채널 간 매출 증감률을 비교할 수 있습니다
      </div>
    </div>
  );
}

function DataCompare({revenues,storeSales=[]}){
  // 전체 매출 볼륨 전용 필터 (초기값: 월단위)
  const [volUnit,setVolUnit]=useState("month");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const containerRef=useRef(null);
  const [svgW,setSvgW]=useState(760);

  useEffect(()=>{
    const obs=new ResizeObserver(es=>setSvgW(Math.max(380,es[0].contentRect.width-48)));
    if(containerRef.current) obs.observe(containerRef.current);
    return()=>obs.disconnect();
  },[]);

  const volPeriods=useMemo(()=>{
    const today=new Date();
    const res=[];
    // 직접 날짜 선택 모드: 선택 범위를 현재 단위로 분할
    const rangeStart=customStart?new Date(customStart):null;
    const rangeEnd=customEnd?new Date(customEnd):null;
    if(rangeStart&&rangeEnd&&rangeStart<=rangeEnd){
      if(volUnit==="week"){
        let cur=new Date(rangeStart);
        while(cur<=rangeEnd){
          const s=cur.toISOString().slice(0,10);
          const e=new Date(cur);e.setDate(e.getDate()+6);
          const eClamp=e>rangeEnd?rangeEnd:e;
          res.push({label:`${cur.getMonth()+1}/${cur.getDate()}`,start:s,end:eClamp.toISOString().slice(0,10)});
          cur.setDate(cur.getDate()+7);
        }
      } else if(volUnit==="quarter"){
        let y=rangeStart.getFullYear(),q=Math.floor(rangeStart.getMonth()/3);
        while(true){
          const sm=q*3; const s=new Date(y,sm,1); const e=new Date(y,sm+3,0);
          if(s>rangeEnd) break;
          const eClamp=e>rangeEnd?rangeEnd:e;
          res.push({label:`${y}Q${q+1}`,start:s.toISOString().slice(0,10),end:eClamp.toISOString().slice(0,10)});
          q++; if(q>3){q=0;y++;}
          if(res.length>20) break;
        }
      } else {
        let cur=new Date(rangeStart.getFullYear(),rangeStart.getMonth(),1);
        while(cur<=rangeEnd){
          const endM=new Date(cur.getFullYear(),cur.getMonth()+1,0);
          const eClamp=endM>rangeEnd?rangeEnd:endM;
          res.push({label:`${cur.getFullYear()}.${cur.getMonth()+1}`,start:cur.toISOString().slice(0,10),end:eClamp.toISOString().slice(0,10)});
          cur=new Date(cur.getFullYear(),cur.getMonth()+1,1);
          if(res.length>24) break;
        }
      }
      return res;
    }
    // 기본 모드
    if(volUnit==="week"){
      for(let i=12;i>=0;i--){
        const end=new Date(today);end.setDate(end.getDate()-i*7);
        const start=new Date(end);start.setDate(start.getDate()-6);
        res.push({label:`${start.getMonth()+1}/${start.getDate()}`,start:start.toISOString().slice(0,10),end:end.toISOString().slice(0,10)});
      }
    } else if(volUnit==="quarter"){
      const curQ=Math.floor(today.getMonth()/3);
      const curY=today.getFullYear();
      for(let i=3;i>=0;i--){
        let q=curQ-i,y=curY;
        while(q<0){q+=4;y--;}
        const sm=q*3;
        const s=new Date(y,sm,1);const e=new Date(y,sm+3,0);
        res.push({label:`${y}Q${q+1}`,start:s.toISOString().slice(0,10),end:e.toISOString().slice(0,10)});
      }
    } else {
      for(let i=3;i>=0;i--){
        const d=new Date(today.getFullYear(),today.getMonth()-i,1);
        const endD=new Date(d.getFullYear(),d.getMonth()+1,0);
        res.push({label:`${d.getFullYear()}.${d.getMonth()+1}`,start:d.toISOString().slice(0,10),end:endD.toISOString().slice(0,10)});
      }
    }
    return res;
  },[volUnit,customStart,customEnd]);

  const revenueData=useMemo(()=>volPeriods.map(p=>{
    const byChannel={};
    COMPARE_CHANNELS.forEach(ch=>{byChannel[ch]=0;});
    revenues.filter(r=>r.date>=p.start&&r.date<=p.end).forEach(r=>{
      if(COMPARE_CHANNELS.includes(r.channel)) byChannel[r.channel]+=(r.amount||0);
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
  const minSvgW=volUnit==="week"?Math.max(svgW,820):svgW;

  const DC={bg:"#0A0A0A",card:"#141414",border:"#242424",text:"#F0F0F0",sub:"#888",dim:"#444"};

  return(
    <div style={{background:DC.bg,minHeight:"100%",padding:"28px 28px 40px"}}>
      <div style={{fontWeight:700,fontSize:20,color:DC.text,letterSpacing:"-0.3px",marginBottom:24}}>데이터 컴페어</div>

      {/* 전체 매출 볼륨 카드 — 개별 필터 */}
      <div style={{background:DC.card,border:`1px solid ${DC.border}`,borderRadius:12,padding:"20px 20px 16px"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
          <div style={{fontWeight:600,fontSize:14,color:DC.text}}>전체 매출 볼륨</div>
          <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
            {[["week","주"],["month","월"],["quarter","분기"]].map(([u,lbl])=>(
              <button key={u} onClick={()=>setVolUnit(u)}
                style={{background:volUnit===u?"#fff":"transparent",
                  color:volUnit===u?"#000":DC.sub,
                  border:`1px solid ${volUnit===u?"#fff":DC.border}`,
                  borderRadius:6,padding:"4px 10px",fontSize:11,
                  cursor:"pointer",fontWeight:600,transition:"all .12s"}}>
                {lbl}
              </button>
            ))}
            <span style={{color:DC.border,fontSize:14,margin:"0 4px"}}>|</span>
            <input type="date" value={customStart} onChange={e=>setCustomStart(e.target.value)}
              style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,
                padding:"3px 7px",fontSize:11,color:DC.text,colorScheme:"dark"}}/>
            <span style={{color:DC.sub,fontSize:11}}>~</span>
            <input type="date" value={customEnd} onChange={e=>setCustomEnd(e.target.value)}
              style={{background:"transparent",border:`1px solid ${DC.border}`,borderRadius:5,
                padding:"3px 7px",fontSize:11,color:DC.text,colorScheme:"dark"}}/>
            {(customStart||customEnd)&&(
              <button onClick={()=>{setCustomStart("");setCustomEnd("");}}
                style={{background:"none",border:"none",color:DC.sub,cursor:"pointer",fontSize:14,padding:"0 2px"}}>✕</button>
            )}
          </div>
        </div>
        {/* 채널 범례 */}
        <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:16}}>
          {COMPARE_CHANNELS.map(ch=>(
            <span key={ch} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:DC.sub}}>
              <span style={{width:8,height:8,background:COMPARE_CH_COLOR[ch],display:"inline-block",flexShrink:0}}/>
              {ch}
            </span>
          ))}
        </div>
        <div ref={containerRef} style={{width:"100%",overflowX:"auto"}}>
          {hasData
            ?<RevenueSankeyChart periods={revenueData} svgW={minSvgW}/>
            :<div style={{textAlign:"center",padding:"80px 0",color:DC.dim,fontSize:13}}>
              매출 데이터를 업로드하면 그래프가 표시됩니다
            </div>
          }
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// APP ROOT
// ─────────────────────────────────────────────
export default function App() {
  const validPages=["dashboard","flow","promo","input","compare"];
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
    const t0=Date.now();
    const db=await getSupabase();
    // 전체 orders 페이지네이션 (Supabase 기본 1000행 제한 우회)
    let allOrders=[];
    let from=0;
    const PAGE=1000;
    while(true){
      const {data,error}=await db.from("orders").select("*").order("order_date",{ascending:true}).range(from,from+PAGE-1);
      if(error||!data||data.length===0) break;
      allOrders=allOrders.concat(data);
      if(data.length<PAGE) break;
      from+=PAGE;
    }
    let allStocks=[];
    let sf=0;
    while(true){
      const {data:sd,error:se}=await db.from("stock_uploads").select("*").order("upload_date",{ascending:false}).range(sf,sf+PAGE-1);
      if(se||!sd||sd.length===0) break;
      allStocks=allStocks.concat(sd);
      if(sd.length<PAGE) break;
      sf+=PAGE;
    }
    let allRevenues=[]; let rf=0;
    while(true){
      const {data:rd}=await db.from("revenues").select("*").order("date",{ascending:false}).range(rf,rf+PAGE-1);
      if(!rd||rd.length===0) break;
      allRevenues=allRevenues.concat(rd);
      if(rd.length<PAGE) break;
      rf+=PAGE;
    }
    // 중복 제거: 같은 date+channel은 id 가장 큰 것(최신)만 유지
    {const revMap={};
    allRevenues.forEach(r=>{const k=`${r.date}__${r.channel}`;if(!revMap[k]||r.id>revMap[k].id)revMap[k]=r;});
    allRevenues=Object.values(revMap);}
    let allStoreSales=[];
    let ssf=0;
    while(true){
      const {data:ssd}=await db.from("store_sales").select("*").order("sale_date",{ascending:true}).range(ssf,ssf+PAGE-1);
      if(!ssd||ssd.length===0) break;
      allStoreSales=allStoreSales.concat(ssd);
      if(ssd.length<PAGE) break;
      ssf+=PAGE;
    }
    // store_sales → 주문 호환 rows (채널은 "오프라인 스토어"로 정규화)
    const storeOrderRows=allStoreSales.map(r=>({
      order_date:r.sale_date,
      channel:"오프라인 스토어",
      product_name:r.product_name,
      option_name:r.option_name,
      qty:r.qty,
      status:r.status,
      order_id:r.order_id,
    }));
    setOrders([...allOrders.map(o=>({...o,channel:normChannel(o.channel)})),...storeOrderRows]);
    setStocks(allStocks);
    setRevenues(allRevenues);
    setStoreSales(allStoreSales);
    const{data:tsData}=await db.from("upload_ts").select("*").order("id",{ascending:true}).limit(1);
    if(tsData&&tsData.length>0){
      const t=tsData[0];
      const next={orders:t.orders||null,stock:t.stock||null,revenue:t.revenue||null,store:t.store||null};
      setTs(next);try{localStorage.setItem("merryon_ts",JSON.stringify(next));}catch{}
    }
    if(firstLoad.current){
      const elapsed=Date.now()-t0;
      if(elapsed<3000) await new Promise(res=>setTimeout(res,3000-elapsed));
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
    {key:"flow",label:"물류 플로우"},
    {key:"promo",label:"프로모션 플로우"},
    {key:"compare",label:"데이터 컴페어"},
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

      {/* top bar */}
      <div style={{ background:isDark?DK.surface:D.surface,
        borderBottom:`1px solid ${isDark?DK.border:D.border}`,
        padding:"0 24px", display:"flex", alignItems:"center", gap:24, height:48, flexShrink:0,
        transition:"background 0.2s ease" }}>
        <div style={{ display:"flex", flexDirection:"column", lineHeight:1.1, marginRight:8 }}>
          <span style={{ fontWeight:800, fontSize:13, letterSpacing:"0.08em", color:isDark?"#fff":D.black }}>MERRYON</span>
          <span style={{ fontSize:10, color:isDark?DK.sub:D.textMeta, letterSpacing:"0.06em" }}>COMMERCE · Made by Jihoon</span>
        </div>
        <nav style={{ display:"flex", gap:2, flex:1 }}>
          {nav.map(n=>(
            <button key={n.key} onClick={()=>setPage(n.key)}
              style={{ background:page===n.key?(isDark?DK.active:D.surfaceAlt):"transparent",
                color:page===n.key?(isDark?"#fff":D.black):(isDark?DK.sub:D.textSub),
                border:"none", borderRadius:6, padding:"6px 14px",
                cursor:"pointer", fontSize:12,
                fontWeight:page===n.key?600:400 }}>
              {n.label}
            </button>
          ))}
        </nav>
        <div style={{ color:isDark?DK.sub:D.textMeta, fontSize:11, flexShrink:0 }}>
          {new Date().toLocaleDateString("ko-KR",{year:"numeric",month:"long",day:"numeric"})}
        </div>
      </div>

      {/* main content */}
      <div style={{ flex:1, overflowY:"auto" }}>
        {page==="dashboard"&&(
          <Dashboard orders={orders} stocks={stocks} revenues={revenues} storeSales={storeSales} ts={ts}
            onRefresh={loadData}/>
        )}
        {page==="flow"&&<LogisticsFlow orders={orders} stocks={stocks} ts={ts}/>}
        {page==="promo"&&<PromoFlow revenues={revenues}/>}
        {page==="compare"&&<DataCompare revenues={revenues} storeSales={storeSales}/>}
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