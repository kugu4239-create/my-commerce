import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import Papa from "papaparse";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend,
} from "recharts";

// ─────────────────────────────────────────────
// NOTE: Supabase revenues 테이블에 아래 컬럼 추가 필요:
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS order_count integer DEFAULT 0;
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS refund_amount integer DEFAULT 0;
//   ALTER TABLE revenues ADD COLUMN IF NOT EXISTS refund_count integer DEFAULT 0;
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
  const m1 = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m1) return `${m1[1]}-${m1[2].padStart(2,"0")}-${m1[3].padStart(2,"0")}`;
  const m2 = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m2) return `${m2[3]}-${m2[2].padStart(2,"0")}-${m2[1].padStart(2,"0")}`;
  return null;
};

// 판매처 이름 정규화
const normChannel = raw => {
  if (!raw) return "미분류";
  const v = String(raw).trim();
  if (v === "MERRYON") return "자사몰";
  if (v === "예약거래") return "오프라인스토어";
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
    date:         f("배송일","주문일","날짜","date","order_date","주문날짜","reg_date","delivery_date"),
    orderId:      f("관리번호","order_id","주문번호","orderid"),
    memo:         f("메모","memo","비고","note"),
    revenue:      f("금액","revenue","sales","매출","price","가격","결제금액","주문금액"),
  };
}

// 기간 필터 유틸
function filterByDate(rows, dateField, period, customStart, customEnd) {
  if (period === "all") return rows;
  const today = new Date().toISOString().slice(0,10);
  if (period === "week") {
    const now = new Date();
    const dow = now.getDay() || 7;
    const monday = new Date(now);
    monday.setDate(now.getDate() - dow + 1);
    const cutStr = monday.toISOString().slice(0,10);
    return rows.filter(r => r[dateField] >= cutStr && r[dateField] <= today);
  }
  if (period === "7d") {
    const c = new Date(); c.setDate(c.getDate()-7);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "1m") {
    const c = new Date(); c.setMonth(c.getMonth()-1);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "3m") {
    const c = new Date(); c.setMonth(c.getMonth()-3);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "6m") {
    const c = new Date(); c.setMonth(c.getMonth()-6);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "custom" && customStart && customEnd) {
    return rows.filter(r => r[dateField] >= customStart && r[dateField] <= customEnd);
  }
  return rows;
}

// ─────────────────────────────────────────────
// SHARED UI
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
function KPI({ label, value, sub, accent="#111" }) {
  return (
    <div style={{ background:D.surface, border:`1px solid ${D.border}`,
      borderRadius:9, padding:"14px 16px", flex:1, minWidth:110 }}>
      <div style={{ color:D.textMeta, fontSize:10, letterSpacing:"0.09em", textTransform:"uppercase", marginBottom:5 }}>{label}</div>
      <div style={{ color:accent, fontSize:20, fontWeight:600 }}>{value}</div>
      {sub&&<div style={{ color:D.textMeta, fontSize:10, marginTop:3 }}>{sub}</div>}
    </div>
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
function DropZone({ onFile, label="CSV 드래그 또는 클릭", fileName="" }) {
  const [hover,setHover]=useState(false);
  const handle=useCallback(e=>{
    e.preventDefault();
    const file=e.dataTransfer?.files?.[0]||e.target.files?.[0];
    if(file) onFile(file);
  },[onFile]);
  return (
    <label onDragOver={e=>{e.preventDefault();setHover(true);}}
      onDragLeave={()=>setHover(false)} onDrop={e=>{setHover(false);handle(e);}}
      style={{ display:"flex", flexDirection:"column", alignItems:"center",
        justifyContent:"center", height:100,
        border:`1.5px dashed ${hover?D.black:D.border}`, borderRadius:9,
        cursor:"pointer", background:hover?D.surfaceAlt:D.surface, transition:"all 0.13s" }}>
      <input type="file" accept=".csv" style={{display:"none"}} onChange={handle}/>
      <div style={{ color:D.textSub, fontSize:13 }}>{label}</div>
      {fileName
        ?<div style={{color:D.textMeta,fontSize:11,marginTop:3}}>{fileName}</div>
        :<div style={{color:D.textMeta,fontSize:11,marginTop:3}}>.csv 형식</div>}
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
  const vw = useWindowWidth();
  // SVG 좌표계 기준 폰트 크기 → 화면 너비에 맞게 보정해 실제 렌더 크기를 유지
  const _fs = (px) => Math.round(px * SVG_W / Math.max(vw, 800));
  const hdrFs  = _fs(15);  // 컬럼 헤더
  const lblFs  = _fs(13);  // 블록 메인 레이블
  const subFs  = _fs(11);  // 블록 서브 텍스트

  const filteredOrders = useMemo(() => {
    return filterByDate(orderRows, "order_date", period, customStart, customEnd);
  }, [orderRows, period, customStart, customEnd]);

  const data = useMemo(() => {
    const prodMap = {};
    stockRows.forEach(r => {
      const key = r.product_name || "미분류";
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, shipped:0, returned:0, exchanged:0, byChannel:{} };
      prodMap[key].stock += (r.qty||0);
    });
    filteredOrders.forEach(r => {
      const key = r.product_name || "미분류";
      const ch = r.channel || "미분류";
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, shipped:0, returned:0, exchanged:0, byChannel:{} };
      if (!prodMap[key].byChannel[ch]) prodMap[key].byChannel[ch] = { shipped:0, returned:0, exchanged:0 };
      if (r.status==="배송")  { prodMap[key].shipped++;   prodMap[key].byChannel[ch].shipped++; }
      if (r.status==="반품")  { prodMap[key].returned++;  prodMap[key].byChannel[ch].returned++; }
      if (r.status==="교환")  { prodMap[key].exchanged++; prodMap[key].byChannel[ch].exchanged++; }
    });
    const prods = Object.values(prodMap)
      .filter(p => p.shipped>0||p.stock>0)
      .sort((a,b)=>(b.stock||0)-(a.stock||0)||(b.shipped||0)-(a.shipped||0))
      .slice(0, limit);
    const chanMap = {};
    filteredOrders.forEach(r => {
      const ch = r.channel||"미분류";
      if (!chanMap[ch]) chanMap[ch] = { name:ch, shipped:0, returned:0, exchanged:0 };
      if (r.status==="배송") chanMap[ch].shipped++;
      if (r.status==="반품") chanMap[ch].returned++;
      if (r.status==="교환") chanMap[ch].exchanged++;
    });
    const channels = Object.values(chanMap).sort((a,b)=>b.shipped-a.shipped);
    const totalReturned  = filteredOrders.filter(r=>r.status==="반품").length;
    const totalExchanged = filteredOrders.filter(r=>r.status==="교환").length;
    return { prods, channels, totalReturned, totalExchanged };
  }, [stockRows, filteredOrders]);

  if (!data.prods.length) return (
    <div style={{ textAlign:"center", padding:80, color:D.textMeta, fontSize:13 }}>
      입고 CSV 또는 이지어드민 CSV를 업로드하면<br/>상품별 물류 흐름이 표시됩니다
    </div>
  );

  const { prods, channels, totalReturned, totalExchanged } = data;
  const n = prods.length;

  // ── 레이아웃 상수 ──
  const PAD_T=36, PAD_H=8, ROW_GAP=6, MIN_H=24;
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

  // viewBox 높이: SVG_W(1400)보다 낮게 유지 → 130vh 컨테이너에서 width 기준으로 꽉 참
  const TARGET_H = Math.max(n*MIN_H, Math.min(900, n*42));
  const rawH = prods.map(p => p.stock>0 ? Math.max(MIN_H, (p.stock/totalStock)*TARGET_H) : MIN_H);
  const rawSum = rawH.reduce((s,h)=>s+h,0);
  const scale = rawSum > TARGET_H ? TARGET_H/rawSum : 1;
  const prodH = rawH.map(h => Math.max(MIN_H, Math.round(h*scale)));

  const yPos = [];
  let cumY = PAD_T+16;
  prodH.forEach(h => { yPos.push(cumY); cumY += h+ROW_GAP; });
  const blockTotalH = cumY - ROW_GAP - (PAD_T+16);
  const totalSvgH   = cumY + 30;

  const chanYOf = {};
  let cy = PAD_T+16;
  channels.forEach(ch=>{
    const h = Math.max(MIN_H, (ch.shipped/chanTotal)*blockTotalH - ROW_GAP);
    chanYOf[ch.name] = cy + h/2;
    cy += h + ROW_GAP;
  });

  // 컬럼2: 반품/교환 블록 높이 분할
  const totalRE = (totalReturned + totalExchanged) || 1;
  const retBlockH  = totalReturned  > 0 ? Math.max(MIN_H, Math.round((totalReturned /totalRE)*blockTotalH) - ROW_GAP) : 0;
  const exchBlockH = totalExchanged > 0 ? Math.max(MIN_H, Math.round((totalExchanged/totalRE)*blockTotalH) - ROW_GAP) : 0;
  const retBlockY  = PAD_T+16;
  const exchBlockY = retBlockY + retBlockH + ROW_GAP;
  const retCenterY  = retBlockY  + retBlockH/2;
  const exchCenterY = exchBlockY + exchBlockH/2;

  const maxStroke    = Math.min(20, Math.max(4, Math.round(400/n)));
  const maxRetStroke = Math.min(18, Math.max(3, Math.round(360/n)));

  const headers = ["입고","판매처별 배송","반품/교환"];

  return (
    <div style={{ width:"100%", height:"100vh" }}>
      <svg width="100%" height="100%" viewBox={`0 0 ${SVG_W} ${totalSvgH}`}
        preserveAspectRatio="xMidYMid meet" style={{ display:"block" }}>

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
          return (
            <g key={p.name}>
              <rect x={COLS_X[0]} y={y} width={NODE_W} height={h} rx={3} fill={col} opacity={0.09}/>
              <rect x={COLS_X[0]} y={y} width={3} height={h} rx={1} fill={col}/>
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
            const h=Math.max(MIN_H,(ch.shipped/chanTotal)*blockTotalH-ROW_GAP);
            const y=ry; ry+=h+ROW_GAP;
            const col=D.SANKEY[(ci+5)%D.SANKEY.length];
            chanYOf[ch.name]=y+h/2;
            return (
              <g key={ch.name}>
                <rect x={COLS_X[1]} y={y} width={NODE_W} height={h} rx={4} fill={col} opacity={0.12}/>
                <rect x={COLS_X[1]} y={y} width={4} height={h} rx={2} fill={col}/>
                <text x={COLS_X[1]+12} y={y+h/2-(h>40?10:0)} dominantBaseline="middle"
                  fill={col} fontSize={lblFs}>{ch.name}</text>
                {h>=40&&<text x={COLS_X[1]+12} y={y+h/2+lblFs+2} dominantBaseline="middle"
                  fill={D.textMeta} fontSize={subFs}>{ch.shipped.toLocaleString()}건</text>}
              </g>
            );
          });
        })()}

        {/* 컬럼2: 반품 블록 */}
        {totalReturned>0&&(
          <g>
            <rect x={COLS_X[2]} y={retBlockY} width={NODE_W} height={retBlockH} rx={4} fill={D.red} opacity={0.1}/>
            <rect x={COLS_X[2]} y={retBlockY} width={4} height={retBlockH} rx={2} fill={D.red}/>
            <text x={COLS_X[2]+12} y={retCenterY} dominantBaseline="middle"
              fill={D.red} fontSize={lblFs}>반품 {totalReturned}건</text>
          </g>
        )}

        {/* 컬럼2: 교환 블록 */}
        {totalExchanged>0&&(
          <g>
            <rect x={COLS_X[2]} y={exchBlockY} width={NODE_W} height={exchBlockH} rx={4} fill={D.amber} opacity={0.12}/>
            <rect x={COLS_X[2]} y={exchBlockY} width={4} height={exchBlockH} rx={2} fill={D.amber}/>
            <text x={COLS_X[2]+12} y={exchCenterY} dominantBaseline="middle"
              fill={D.amber} fontSize={lblFs}>교환 {totalExchanged}건</text>
          </g>
        )}
      </svg>
    </div>
  );
}

const getCSData=()=>{try{return JSON.parse(localStorage.getItem("cs_data")||"[]");}catch{return[];}};
const saveCSData=d=>localStorage.setItem("cs_data",JSON.stringify(d));
const getPromos=()=>{try{return JSON.parse(localStorage.getItem("promotions")||"[]");}catch{return[];}};
const savePromos=d=>localStorage.setItem("promotions",JSON.stringify(d));

// ─────────────────────────────────────────────
// ANALYTICS ENGINE
// ─────────────────────────────────────────────
function analyze(orderRows, stockRows, revenueRows) {
  // 매출 입력 데이터 기반 KPI
  const totalRevenue    = revenueRows.reduce((s,r)=>s+(r.amount||0)-(r.refund_amount||0),0);
  const totalOrderCount = revenueRows.reduce((s,r)=>s+(r.order_count||0),0);
  const totalRefundAmt  = revenueRows.reduce((s,r)=>s+(r.refund_amount||0),0);
  const totalRefundCount= revenueRows.reduce((s,r)=>s+(r.refund_count||0),0);
  const returnRate      = totalOrderCount>0?(totalRefundCount/totalOrderCount*100).toFixed(1):"0.0";

  // 이지어드민 CSV 기반 KPI
  const totalShipped  = orderRows.filter(r=>r.status==="배송").length;
  const totalReturned = orderRows.filter(r=>["반품","교환"].includes(r.status)).length;
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
  orderRows.forEach(r=>{
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    if(r.status==="배송") byChannel[ch].shipped++;
    if(["반품","교환"].includes(r.status)) byChannel[ch].returned++;
  });
  const channelList=Object.values(byChannel).filter(c=>c.name!=="오프라인스토어").sort((a,b)=>b.revenue-a.revenue||b.shipped-a.shipped);
  const totalRev=channelList.reduce((s,c)=>s+c.revenue,0)||1;
  channelList.forEach(c=>{
    c.share=((c.revenue||0)/totalRev*100).toFixed(1);
    c.returnRate=c.orderCount>0?(c.refundCount/c.orderCount*100).toFixed(1):"0.0";
  });

  // 월별 배송/반품
  const byMonth={};
  orderRows.forEach(r=>{
    const ym=r.order_date?r.order_date.slice(0,7):null;
    if(!ym) return;
    if(!byMonth[ym]) byMonth[ym]={month:ym,shipped:0,returned:0};
    if(r.status==="배송") byMonth[ym].shipped++;
    if(["반품","교환"].includes(r.status)) byMonth[ym].returned++;
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
    if(["반품","교환"].includes(r.status)) byProd[key].returned++;
  });
  const prodList=Object.values(byProd);
  const weekBest=[...prodList].sort((a,b)=>b.qty-a.qty).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));
  const weekWorst=[...prodList].filter(p=>p.returned>0).sort((a,b)=>b.returned-a.returned).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));

  return {
    totalRevenue,totalOrderCount,totalRefundAmt,totalRefundCount,returnRate,
    totalShipped,totalReturned,totalStock,
    channelList,monthlyData,weekBest,weekWorst,latestWeek,weekRows,
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
  {key:"7d",label:"최근 7일"},
  {key:"1m",label:"최근 한달"},
  {key:"3m",label:"최근 3개월"},
  {key:"custom",label:"기간 선택"},
];

function Dashboard({ orders, stocks, revenues, ts, onRefresh }) {
  const isMobile=useWindowWidth()<=1080;
  const [period,setPeriod]=useState("7d");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [deleteAll,setDeleteAll]=useState(false);
  const [shippingPeriod,setShippingPeriod]=useState("7d");
  const [returnPeriod,setReturnPeriod]=useState("7d");
  const [rankBestPeriod,setRankBestPeriod]=useState("7d");
  const [rankBestChannel,setRankBestChannel]=useState("전체");
  const [rankBestCustomStart,setRankBestCustomStart]=useState("");
  const [rankBestCustomEnd,setRankBestCustomEnd]=useState("");
  const [rankWorstPeriod,setRankWorstPeriod]=useState("7d");
  const [rankWorstChannel,setRankWorstChannel]=useState("전체");
  const [rankWorstCustomStart,setRankWorstCustomStart]=useState("");
  const [rankWorstCustomEnd,setRankWorstCustomEnd]=useState("");
  const [chSort,setChSort]=useState({key:"revenue",dir:"desc"});

  const axTick={fill:D.textMeta,fontSize:10};
  const NoWrapTick=({x,y,payload})=>(
    <text x={x} y={y} dy={4} textAnchor="end" fill={D.textMeta} fontSize={10} style={{whiteSpace:"nowrap"}}>
      {payload.value?.length>22?payload.value.slice(0,22)+"…":payload.value}
    </text>
  );

  const filteredOrders=useMemo(()=>filterByDate(orders,"order_date",period,customStart,customEnd),[orders,period,customStart,customEnd]);
  const filteredRevenues=useMemo(()=>filterByDate(revenues,"date",period,customStart,customEnd),[revenues,period,customStart,customEnd]);
  const stats=useMemo(()=>analyze(filteredOrders,stocks,filteredRevenues),[filteredOrders,stocks,filteredRevenues]);

  // 직전 동일 기간 채널별 순매출

  // 플랫폼별 선호 옵션 (컬러/사이즈)
  const optionStats=useMemo(()=>{
    const map={};
    filteredOrders.filter(r=>r.status==="배송"&&r.channel!=="오프라인스토어").forEach(r=>{
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
    }));
  },[filteredOrders]);

  // 플랫폼별 반품률 높은 옵션
  const returnOptionStats=useMemo(()=>{
    const map={};
    filteredOrders.filter(r=>r.channel!=="오프라인스토어").forEach(r=>{
      const ch=r.channel||"미분류";
      if(!map[ch]) map[ch]={};
      const {color,size}=parseOption(r.product_name,r.option_name);
      const isShipped=r.status==="배송";
      const isReturned=["반품","교환"].includes(r.status);
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
    }).filter(d=>d.colors.length>0||d.sizes.length>0);
  },[filteredOrders]);

  // 판매처 채널 목록 (전체 orders 기준, 오프라인스토어 제외)
  const activeChannels=useMemo(()=>{
    const fixed=["자사몰","29CM","무신사"];
    const inData=new Set(orders.map(r=>r.channel||"미분류").filter(Boolean));
    return fixed.filter(c=>inData.has(c));
  },[orders]);

  // 판매 Top 랭킹
  const bestFilteredOrders=useMemo(()=>
    filterByDate(orders,"order_date",rankBestPeriod,rankBestCustomStart,rankBestCustomEnd),
    [rankBestPeriod,orders,rankBestCustomStart,rankBestCustomEnd]);

  const bestRows=useMemo(()=>{
    const base=bestFilteredOrders.filter(r=>r.channel!=="오프라인스토어");
    const rows=rankBestChannel==="전체"?base:base.filter(r=>r.channel===rankBestChannel);
    const byProd={};
    rows.forEach(r=>{
      const key=r.product_name||"미분류";
      if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
      byProd[key].qty+=(r.qty||0); byProd[key].orders++;
      if(["반품","교환"].includes(r.status)) byProd[key].returned++;
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
    const base=worstFilteredOrders.filter(r=>r.channel!=="오프라인스토어");
    const rows=rankWorstChannel==="전체"?base:base.filter(r=>r.channel===rankWorstChannel);
    const byProd={};
    rows.forEach(r=>{
      const key=r.product_name||"미분류";
      if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
      byProd[key].qty+=(r.qty||0); byProd[key].orders++;
      if(["반품","교환"].includes(r.status)) byProd[key].returned++;
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
    return Object.values(byProd).filter(p=>p.returned>0)
      .sort((a,b)=>(b.returned/b.orders)-(a.returned/a.orders)).slice(0,20)
      .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0",
        topReason:topReason(p.name)}));
  },[worstFilteredOrders,rankWorstChannel]);

  // 월별 배송량 차트 데이터
  const shippingChartData=useMemo(()=>{
    const today=new Date().toISOString().slice(0,10);
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
    if(returnPeriod==="7d"){
      const d=new Date(); d.setDate(d.getDate()-7);
      start=d.toISOString().slice(0,10);
    } else {
      const d=new Date(); d.setMonth(d.getMonth()-(returnPeriod==="1m"?1:3));
      start=d.toISOString().slice(0,10);
    }
    const filteredRet=orders.filter(r=>r.order_date>=start&&r.channel!=="오프라인스토어");
    const retByCh={};
    filteredRet.forEach(r=>{
      if(["반품","교환"].includes(r.status)){
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
      if(["반품","교환"].includes(r.status)) byDate[d][ch]=(byDate[d][ch]||0)+1;
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
        <button onClick={onRefresh}
          style={{background:D.surfaceAlt,border:`1px solid ${D.border}`,borderRadius:7,
            padding:"5px 13px",fontSize:12,cursor:"pointer",color:D.textSub,
            display:"flex",alignItems:"center",gap:5}}>
          ↺ 새로고침
        </button>
      </div>

      {/* KPI 카드 - 총 매출/주문/반품은 매출입력, 배송은 이지어드민 */}
      <div style={{display:"flex",gap:9,marginBottom:20,flexWrap:"wrap"}}>
        <KPI label="총 매출" value={fmtWon(stats.totalRevenue)} accent={D.black}/>
        <KPI label="배송" value={stats.totalShipped.toLocaleString()+"건"} accent={D.green}/>
        <KPI label="반품률" value={stats.totalShipped>0?(stats.totalReturned/stats.totalShipped*100).toFixed(1)+"%":"0.0%"}
          sub={stats.totalReturned.toLocaleString()+"건"}
          accent={stats.totalShipped>0&&(stats.totalReturned/stats.totalShipped)>0.1?D.red:D.textSub}/>
        <KPI label="입고 수량" value={stats.totalStock.toLocaleString()+"개"} accent={D.blue}/>
      </div>

      {/* 판매처 점유율 + 판매처별 매출 */}
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"280px 1fr",gap:10,marginBottom:20}}>
        <Card>
          <SecTitle ts={ts.orders}>매출 점유율</SecTitle>
          <ResponsiveContainer width="100%" height={160}>
            {(()=>{
              const sorted=[...stats.channelList.slice(0,6)].sort((a,b)=>b.revenue-a.revenue);
              return (
                <PieChart>
                  <Pie data={sorted.map(c=>({name:c.name,value:c.revenue}))}
                    dataKey="value" nameKey="name" cx="50%" cy="50%"
                    innerRadius={38} outerRadius={60} paddingAngle={2}>
                    {sorted.map((c,i)=>(<Cell key={i} fill={chColor(c.name)}/>))}
                  </Pie>
                  <Tooltip formatter={(v,n,p)=>{
                    const total=sorted.reduce((s,c)=>s+c.revenue,0)||1;
                    const pct=(v/total*100).toFixed(1);
                    return [`₩${v.toLocaleString()} (${pct}%)`,n];
                  }} contentStyle={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:7,fontSize:11}}/>
                  <Legend iconSize={8} iconType="circle" wrapperStyle={{fontSize:10,paddingTop:6}}/>
                </PieChart>
              );
            })()}
          </ResponsiveContainer>
        </Card>
        <Card>
          <SecTitle ts={ts.orders}>판매처별 매출</SecTitle>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={stats.channelList.slice(0,7)} layout="vertical" barCategoryGap="28%">
              <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
              <XAxis type="number" tick={axTick} tickFormatter={v=>v>=1e4?(v/1e4).toFixed(0)+"만":v}/>
              <YAxis type="category" dataKey="name" width={76} tick={axTick}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey="revenue" name="매출" radius={[0,3,3,0]}>
                {stats.channelList.slice(0,7).map((c,i)=>(
                  <Cell key={i} fill={chColor(c.name)}/>
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* 판매처 상세 */}
      <Card style={{marginBottom:20}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <SecTitle ts={ts.orders}>판매처 상세</SecTitle>
          <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
            {[["revenue","매출"],["share","점유율"],["shipped","배송"],["returned","반품"],["rate","반품률"],["aov","객단가"]].map(([k,l])=>(
              <button key={k} onClick={()=>setChSort({key:k,dir:"desc"})}
                style={{background:chSort.key===k?D.black:"transparent",
                  color:chSort.key===k?"#fff":D.textSub,
                  border:`1px solid ${chSort.key===k?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer",
                  fontWeight:chSort.key===k?600:400}}>
                {l}{chSort.key===k?" ↓":""}
              </button>
            ))}
          </div>
        </div>
        {(()=>{
          const cols=[
            {key:"name",   label:"판매처", left:true,  val:c=>c.name},
            {key:"share",  label:"점유율", val:c=>parseFloat(c.share)},
            {key:"revenue",label:"매출",   val:c=>c.revenue},
            {key:"shipped",label:"배송",   val:c=>c.shipped},
            {key:"returned",label:"반품",  val:c=>c.returned},
            {key:"rate",   label:"반품률", val:c=>c.shipped>0?c.returned/c.shipped:0},
            {key:"aov",    label:"객단가", val:c=>c.orderCount>0?c.revenue/c.orderCount:0},
          ];
          const sorted=[...stats.channelList].sort((a,b)=>{
            const col=cols.find(c=>c.key===chSort.key);
            if(!col) return 0;
            const va=col.val(a), vb=col.val(b);
            return chSort.dir==="desc"?(vb>va?1:vb<va?-1:0):(va>vb?1:va<vb?-1:0);
          });
          return (
            <div style={{minHeight:260}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
                {cols.map(({key,label,left})=>(
                  <th key={key} style={{padding:"7px 9px",textAlign:left?"left":"right",
                    color:chSort.key===key?D.black:D.textMeta,
                    fontWeight:chSort.key===key?600:400,whiteSpace:"nowrap"}}>
                    {label}
                  </th>
                ))}
              </tr></thead>
              <tbody>
                {sorted.map(c=>(
                  <tr key={c.name} style={{borderBottom:`1px solid ${D.border}`}}>
                    <td style={{padding:"7px 9px",fontWeight:600}}>{c.name}</td>
                    <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub,fontWeight:chSort.key==="share"?700:400}}>{c.share}%</td>
                    <td style={{textAlign:"right",padding:"7px 9px",fontWeight:chSort.key==="revenue"?700:600}}>{c.revenue>0?fmtWon(c.revenue):"—"}</td>
                    <td style={{textAlign:"right",padding:"7px 9px",color:D.green,fontWeight:chSort.key==="shipped"?700:400}}>{c.shipped.toLocaleString()}</td>
                    <td style={{textAlign:"right",padding:"7px 9px",color:D.red,fontWeight:chSort.key==="returned"?700:400}}>{c.returned.toLocaleString()}</td>
                    <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600,
                      color:c.shipped>0&&(c.returned/c.shipped)>0.1?D.red:D.textSub}}>
                      {c.shipped>0?(c.returned/c.shipped*100).toFixed(1):0}%</td>
                    <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub,fontWeight:chSort.key==="aov"?700:400}}>{c.orderCount>0?fmtWon(Math.round(c.revenue/c.orderCount)):"—"}</td>
                  </tr>
                ))}
                <tr style={{borderTop:`1px solid ${D.borderMid}`}}>
                  <td style={{padding:"7px 9px",fontWeight:700}}>합계</td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>100%</td>
                  <td style={{textAlign:"right",padding:"7px 9px",fontWeight:700}}>{fmtWon(stats.totalRevenue)}</td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.green,fontWeight:600}}>{stats.totalShipped.toLocaleString()}</td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.red,fontWeight:600}}>{stats.totalReturned.toLocaleString()}</td>
                  <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600,
                    color:stats.totalShipped>0&&(stats.totalReturned/stats.totalShipped)>0.1?D.red:D.textSub}}>
                    {stats.totalShipped>0?(stats.totalReturned/stats.totalShipped*100).toFixed(1):"0.0"}%</td>
                  <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub,fontWeight:600}}>{stats.totalOrderCount>0?fmtWon(Math.round(stats.totalRevenue/stats.totalOrderCount)):"—"}</td>
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
              {[["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([v,l])=>(
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
              {[["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([v,l])=>(
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
      <Card style={{marginBottom:20}}>
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
          <div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center"}}>
            {[["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([k,l])=>(
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
          <SecTitle ts={ts.orders}>플랫폼별 선호 옵션</SecTitle>
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

      {/* 반품 탑 */}
      <Card style={{marginBottom:20}}>
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
            {[["7d","최근 7일"],["1m","최근 한달"],["3m","최근 3개월"]].map(([k,l])=>(
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
              <Bar dataKey="returnRate" name="반품률" radius={[0,3,3,0]} fill={D.red}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>

      {/* 플랫폼별 반품률 높은 옵션 */}
      {returnOptionStats.length>0&&(
        <Card style={{marginBottom:20}}>
          <SecTitle ts={ts.orders}>플랫폼별 반품률 높은 옵션</SecTitle>
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
              ]);
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
    </div>
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

  const filteredOrders=useMemo(()=>filterByDate(orders,"order_date",period,customStart,customEnd),[orders,period,customStart,customEnd]);

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
      {filteredOrders.length>0&&(
        <Card>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,flexWrap:"wrap",gap:8}}>
            <SecTitle ts={ts.orders}>상품별 흐름 요약</SecTitle>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
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
                  stocks.forEach(r=>{
                    const k=r.product_name||"미분류";
                    if(!prodMap[k]) prodMap[k]={name:k,stock:0,shipped:0,returned:0};
                    prodMap[k].stock+=(r.qty||0);
                  });
                  filteredOrders.forEach(r=>{
                    const k=r.product_name||"미분류";
                    if(!prodMap[k]) prodMap[k]={name:k,stock:0,shipped:0,returned:0};
                    if(r.status==="배송") prodMap[k].shipped++;
                    if(["반품","교환"].includes(r.status)) prodMap[k].returned++;
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
const PROMO_PLATFORMS=["자사몰","29CM","무신사"];

function PromoFlow({ revenues }) {
  const [promos,setPromos]=useState(getPromos);
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState({name:"",platform:"자사몰",start_date:"",end_date:"",memo:""});
  const today=new Date().toISOString().slice(0,10);
  const twoMonthsAgo=(()=>{const d=new Date();d.setMonth(d.getMonth()-2);return d.toISOString().slice(0,10);})();
  const [viewStart,setViewStart]=useState(twoMonthsAgo);
  const [viewEnd,setViewEnd]=useState(today);

  const updatePromos=data=>{savePromos(data);setPromos(data);};
  const addPromo=()=>{
    if(!form.name||!form.start_date||!form.end_date)return;
    updatePromos([...promos,{...form,id:Date.now()}]);
    setForm({name:"",platform:"자사몰",start_date:"",end_date:"",memo:""});
    setShowForm(false);
  };
  const delPromo=id=>updatePromos(promos.filter(p=>p.id!==id));

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
      byDate[key]={date:key.slice(5),...Object.fromEntries(PROMO_PLATFORMS.map(p=>[p,null]))};
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
    padding:"7px 10px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box"};

  return (
    <div style={{padding:"20px 24px",maxWidth:1600,margin:"0 auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <div style={{fontWeight:600,fontSize:15,color:D.black}}>프로모션 플로우</div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <input type="date" value={viewStart} onChange={e=>setViewStart(e.target.value)}
            style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"5px 8px",fontSize:11,color:D.text}}/>
          <span style={{color:D.textMeta}}>—</span>
          <input type="date" value={viewEnd} onChange={e=>setViewEnd(e.target.value)}
            style={{border:`1px solid ${D.border}`,borderRadius:5,padding:"5px 8px",fontSize:11,color:D.text}}/>
          <button onClick={()=>setShowForm(v=>!v)}
            style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
              padding:"6px 14px",fontSize:11,cursor:"pointer",fontWeight:600}}>
            {showForm?"취소":"+ 프로모션 추가"}
          </button>
        </div>
      </div>

      {showForm&&(
        <Card style={{marginBottom:20}}>
          <div style={{fontWeight:600,fontSize:12,marginBottom:20}}>프로모션 추가</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr auto",gap:8,alignItems:"end"}}>
            <div>
              <div style={{fontSize:10,color:D.textMeta,marginBottom:4}}>프로모션명</div>
              <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={inp} placeholder="예: 오픈 기념 할인"/>
            </div>
            <div>
              <div style={{fontSize:10,color:D.textMeta,marginBottom:4}}>플랫폼</div>
              <div style={{display:"flex",gap:4}}>
                {PROMO_PLATFORMS.map(p=>(
                  <button key={p} onClick={()=>setForm(f=>({...f,platform:p}))}
                    style={{flex:1,background:form.platform===p?chColor(p):"transparent",
                      color:form.platform===p?"#fff":D.textSub,
                      border:`1px solid ${form.platform===p?chColor(p):D.border}`,
                      borderRadius:5,padding:"6px 4px",fontSize:11,cursor:"pointer"}}>
                    {p}
                  </button>
                ))}
              </div>
            </div>
            <div>
              <div style={{fontSize:10,color:D.textMeta,marginBottom:4}}>시작일</div>
              <input type="date" value={form.start_date} onChange={e=>setForm(f=>({...f,start_date:e.target.value}))} style={inp}/>
            </div>
            <div>
              <div style={{fontSize:10,color:D.textMeta,marginBottom:4}}>종료일</div>
              <input type="date" value={form.end_date} onChange={e=>setForm(f=>({...f,end_date:e.target.value}))} style={inp}/>
            </div>
            <button onClick={addPromo}
              style={{background:D.black,color:"#fff",border:"none",borderRadius:6,
                padding:"9px 16px",fontSize:12,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap"}}>
              저장
            </button>
          </div>
          <div style={{marginTop:8}}>
            <div style={{fontSize:10,color:D.textMeta,marginBottom:4}}>메모</div>
            <input value={form.memo} onChange={e=>setForm(f=>({...f,memo:e.target.value}))} style={inp} placeholder="설명 (선택)"/>
          </div>
        </Card>
      )}

      {/* 플랫폼별 가로 캘린더 바 */}
      <Card style={{marginBottom:20}}>
        <div style={{fontWeight:600,fontSize:12,marginBottom:12,color:D.black}}>플랫폼별 프로모션 일정</div>
        {/* 날짜 눈금 */}
        <div style={{position:"relative",height:16,marginBottom:4,paddingLeft:70}}>
          {[0,25,50,75,100].map(pct=>{
            const ms=startMs+(endMs-startMs)*pct/100;
            const d=new Date(ms);
            const label=`${d.getMonth()+1}/${d.getDate()}`;
            return <span key={pct} style={{position:"absolute",left:`${pct}%`,transform:"translateX(-50%)",
              fontSize:9,color:D.textMeta}}>{label}</span>;
          })}
        </div>
        {PROMO_PLATFORMS.map(plat=>{
          const bars=promos.filter(p=>p.platform===plat&&p.end_date>=viewStart&&p.start_date<=viewEnd);
          return (
            <div key={plat} style={{display:"flex",alignItems:"center",marginBottom:8,gap:8}}>
              <div style={{width:62,fontSize:11,color:D.textSub,flexShrink:0,textAlign:"right"}}>{plat}</div>
              <div style={{flex:1,position:"relative",height:28,background:D.surfaceAlt,borderRadius:4}}>
                {bars.map(promo=>{
                  const l=datePct(promo.start_date);
                  const r=datePct(promo.end_date);
                  const w=Math.max(0.5,r-l);
                  return (
                    <div key={promo.id}
                      title={`${promo.name}\n${promo.start_date} ~ ${promo.end_date}${promo.memo?"\n"+promo.memo:""}`}
                      style={{position:"absolute",left:`${l}%`,width:`${w}%`,height:"100%",
                        background:chColor(plat),borderRadius:4,display:"flex",alignItems:"center",
                        padding:"0 6px",fontSize:10,color:"#fff",overflow:"hidden",
                        boxSizing:"border-box",cursor:"pointer",minWidth:4}}>
                      <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{promo.name}</span>
                    </div>
                  );
                })}
              </div>
              <button onClick={()=>{
                const target=bars[bars.length-1];
                if(target)delPromo(target.id);
              }} style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",
                fontSize:10,padding:"2px 6px",flexShrink:0,visibility:bars.length?"visible":"hidden"}}>
                ✕
              </button>
            </div>
          );
        })}
        {/* 프로모션 목록 */}
        {promos.length>0&&(
          <div style={{marginTop:12,borderTop:`1px solid ${D.border}`,paddingTop:10}}>
            <div style={{fontSize:10,color:D.textMeta,marginBottom:6}}>등록된 프로모션</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {promos.map(p=>(
                <div key={p.id} style={{display:"flex",alignItems:"center",gap:4,
                  background:D.surfaceAlt,borderRadius:4,padding:"3px 8px",fontSize:10}}>
                  <div style={{width:6,height:6,borderRadius:"50%",background:chColor(p.platform),flexShrink:0}}/>
                  <span style={{color:D.textSub}}>{p.platform}</span>
                  <span style={{fontWeight:600}}>{p.name}</span>
                  <span style={{color:D.textMeta}}>{p.start_date}~{p.end_date}</span>
                  <button onClick={()=>delPromo(p.id)}
                    style={{background:"transparent",border:"none",color:D.textMeta,
                      cursor:"pointer",padding:0,fontSize:11}}>✕</button>
                </div>
              ))}
            </div>
          </div>
        )}
      </Card>

      {/* 기간별 플랫폼 매출 그래프 */}
      <Card>
        <div style={{fontWeight:600,fontSize:12,marginBottom:12,color:D.black}}>기간별 플랫폼 매출</div>
        {revenueData.length>0&&revenues.some(r=>r.date>=viewStart&&r.date<=viewEnd)?(
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={revenueData}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="date" tick={{fill:D.textMeta,fontSize:9}}/>
              <YAxis tick={{fill:D.textMeta,fontSize:9}} tickFormatter={v=>v>=10000?(v/10000).toFixed(0)+"만":v}/>
              <Tooltip formatter={(v,n)=>[`₩${v.toLocaleString()}`,n]}/>
              <Legend iconSize={8} wrapperStyle={{fontSize:10}}/>
              {PROMO_PLATFORMS.map(p=>(
                <Line key={p} type="monotone" dataKey={p} name={p}
                  stroke={chColor(p)} strokeWidth={2} dot={false} connectNulls={false}/>
              ))}
            </LineChart>
          </ResponsiveContainer>
        ):(
          <div style={{textAlign:"center",padding:40,color:D.textMeta,fontSize:12}}>
            해당 기간에 매출 데이터가 없습니다
          </div>
        )}
      </Card>
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

  const inp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
    padding:"7px 10px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box"};

  const save=()=>{
    if(!product.trim()||!reason.trim())return;
    const next=[{id:Date.now(),date,product_name:product.trim(),return_reason:reason.trim(),channel},...csData];
    saveCSData(next);setCSData(next);
    setProduct("");setReason("");
  };

  const handleCSVFile=useCallback(file=>{
    if(!file)return;
    setCsvResult(null);
    Papa.parse(file,{header:true,skipEmptyLines:true,
      complete:({data})=>{
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
          const newEntries=data.filter(r=>r[prodCol]&&r[reasonCol]).map(r=>({
            id:Date.now()+Math.random(),
            date:dateCol&&toDate(r[dateCol])?toDate(r[dateCol]):today,
            product_name:String(r[prodCol]||"").trim(),
            return_reason:String(r[reasonCol]||"").trim(),
            channel:chCol?normChannel(r[chCol]):"미분류",
          }));
          if(!newEntries.length)throw new Error("유효한 데이터 행이 없습니다");
          const next=[...newEntries,...csData];
          saveCSData(next);setCSData(next);
          setCsvResult({type:"success",msg:`${newEntries.length}건 추가 완료`});
        }catch(e){setCsvResult({type:"error",msg:e.message});}
      },
      error:e=>setCsvResult({type:"error",msg:e.message}),
    });
  },[csData,today]);

  const del=id=>{
    const next=csData.filter(r=>r.id!==id);
    saveCSData(next);setCSData(next);
  };

  const filtered=csData.filter(r=>!filterProd||r.product_name.includes(filterProd));

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
          <DropZone onFile={handleCSVFile} label="반품 CS CSV 업로드"/>
          {csvResult&&<Alert type={csvResult.type} msg={csvResult.msg}/>}
        </div>
      </Card>
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{fontWeight:600,fontSize:13}}>반품 사유 내역</div>
          <input value={filterProd} onChange={e=>setFilterProd(e.target.value)}
            style={{...inp,width:160,fontSize:11,padding:"5px 8px"}} placeholder="상품명 검색"/>
        </div>
        <div style={{overflowY:"auto",maxHeight:480}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
              {["날짜","판매처","상품명","반품 사유",""].map(h=>(
                <th key={h} style={{padding:"6px 8px",textAlign:"left",color:D.textMeta,fontWeight:400}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {filtered.slice(0,100).map(r=>(
                <tr key={r.id} style={{borderBottom:`1px solid ${D.border}`}}>
                  <td style={{padding:"5px 8px",color:D.textMeta,whiteSpace:"nowrap"}}>{r.date}</td>
                  <td style={{padding:"5px 8px"}}><span style={{color:chColor(r.channel),fontWeight:600}}>{r.channel}</span></td>
                  <td style={{padding:"5px 8px",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.product_name}</td>
                  <td style={{padding:"5px 8px",color:D.textSub}}>{r.return_reason}</td>
                  <td style={{padding:"5px 8px"}}>
                    <button onClick={()=>del(r.id)}
                      style={{background:"transparent",border:"none",color:D.textMeta,cursor:"pointer",fontSize:10}}>✕</button>
                  </td>
                </tr>
              ))}
              {filtered.length===0&&<tr><td colSpan={5} style={{padding:24,textAlign:"center",color:D.textMeta}}>데이터 없음</td></tr>}
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

  const loadHistory=useCallback(async()=>{
    const db=await getSupabase();
    const {data}=await db.from("revenues").select("*").order("date",{ascending:false}).limit(50);
    setHistory(data||[]); setHistTs(nowStr());
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
    for(const d of dates){
      const {error}=await db.from("revenues").upsert({
        date:d,channel:ch,
        amount:Math.round(num/dates.length),
        order_count:Math.round((Number(orderCnt)||0)/dates.length),
        refund_amount:Math.round((Number(refundAmt.replace(/,/g,""))||0)/dates.length),
        refund_count:Math.round((Number(refundCnt)||0)/dates.length),
      },{onConflict:"date,channel"});
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

  const handleCsvFile=useCallback(file=>{
    Papa.parse(file,{header:true,skipEmptyLines:true,complete:async({data})=>{
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
        channel:chCol?normChannel(r[chCol]):"자사몰",
        amount:Number(String(r[amtCol]||"0").replace(/[^0-9.-]/g,""))||0,
        order_count:ordCol?Number(r[ordCol]||0):0,
        refund_amount:refAmtCol?Number(String(r[refAmtCol]||"0").replace(/[^0-9.-]/g,"")):0,
        refund_count:refCntCol?Number(r[refCntCol]||0):0,
      }));
      // 기존 데이터와 겹치는 (date, channel) 쌍 확인
      const db=await getSupabase();
      const dates=[...new Set(rows.map(r=>r.date))];
      const {data:existing}=await db.from("revenues").select("date,channel").in("date",dates);
      const existSet=new Set((existing||[]).map(r=>`${r.date}__${r.channel}`));
      const overlaps=rows.filter(r=>existSet.has(`${r.date}__${r.channel}`));
      setCsvPreview({rows,overlaps});
      setCsvConflictChoice(null);
    }});
  },[]);

  const handleCsvUpload=async(choice)=>{
    if(!csvPreview?.rows) return;
    const db=await getSupabase();
    let toUpload=csvPreview.rows;
    if(choice==="keep"){
      const overlapKeys=new Set(csvPreview.overlaps.map(r=>`${r.date}__${r.channel}`));
      toUpload=toUpload.filter(r=>!overlapKeys.has(`${r.date}__${r.channel}`));
    }
    for(let i=0;i<toUpload.length;i+=200){
      const {error}=await db.from("revenues").upsert(toUpload.slice(i,i+200),{onConflict:"date,channel"});
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
          <DropZone onFile={handleCsvFile} label="매출 CSV 업로드 (날짜·판매처·매출금액 컬럼 필요)"/>
          {csvPreview?.error&&<div style={{color:D.red,fontSize:10,marginTop:4}}>{csvPreview.error}</div>}
          {csvPreview&&!csvPreview.error&&(csvPreview.overlaps?.length===0||csvConflictChoice)&&(
            <div style={{marginTop:6,display:"flex",gap:8,alignItems:"center"}}>
              <span style={{fontSize:11,color:D.textSub}}>{csvPreview.rows.length}건 파싱됨</span>
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
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div style={{display:"flex",alignItems:"baseline",gap:6}}>
            <span style={{fontWeight:600,fontSize:13}}>최근 입력 내역</span>
            <UpdatedAt ts={histTs}/>
          </div>
          <Btn onClick={loadHistory} variant="ghost" style={{padding:"4px 11px",fontSize:11}}>불러오기</Btn>
        </div>
        {history.length>0?(
          <div style={{overflowY:"auto",maxHeight:480}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
                {["날짜","판매처","매출","주문","환불금","환불수",""].map(h=>(
                  <th key={h} style={{padding:"5px 7px",textAlign:h===""?"center":"left",color:D.textMeta,fontWeight:400}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {history.map(r=>(
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
        ):<div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>불러오기를 눌러주세요</div>}
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
  const dateValid=startDate&&endDate&&startDate<=endDate;

  const confirmDate=async()=>{
    setLoading(true);
    const db=await getSupabase();
    const {data}=await db.from("stock_uploads").select("*").gte("upload_date",startDate).lte("upload_date",endDate).order("upload_date");
    setExisting(data||[]); setStep(1); setLoading(false);
  };
  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    Papa.parse(file,{header:true,skipEmptyLines:true,
      complete:({data})=>{
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
      },
      error:e=>setResult({type:"error",msg:e.message}),
    });
  },[]);
  const handleUpload=async()=>{
    if(!preview?.length||!dateValid) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    const {error:delErr}=await db.from("stock_uploads").delete().gte("upload_date",startDate).lte("upload_date",endDate);
    if(delErr){setResult({type:"error",msg:"삭제 실패: "+delErr.message});setLoading(false);return;}
    const rows=preview.map(r=>({...r,upload_date:startDate}));
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
            <DropZone onFile={handleFile} fileName={fileName} label="입고 CSV 업로드"/>
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
          <div style={{fontWeight:500,fontSize:12,marginBottom:20}}>
            {step<2?`기존 DB — ${startDate}~${endDate}`:`새 파일 — ${fileName}`}
          </div>
          {step===0&&<div style={{color:D.textMeta,textAlign:"center",padding:60,fontSize:12}}>기간 선택 후 기존 데이터 표시</div>}
          {step>=1&&step<2&&(existing?.length?
            <PreviewTable rows={existing} cols={[
              {key:"upload_date",label:"업로드일",color:D.textMeta},
              {key:"product_name",label:"상품명",maxW:150},
              {key:"option_name",label:"옵션",color:D.textMeta},
              {key:"qty",label:"수량",bold:true},
              {key:"memo",label:"메모",color:D.textMeta},
            ]}/>:
            <div style={{color:D.green,textAlign:"center",padding:60,fontSize:12}}>해당 기간 기존 데이터 없음</div>)}
          {step>=2&&preview&&<PreviewTable rows={preview} cols={[
            {key:"product_name",label:"상품명",maxW:180},
            {key:"option_name",label:"옵션",color:D.textMeta},
            {key:"qty",label:"수량",bold:true},
            {key:"memo",label:"메모",color:D.textMeta},
          ]}/>}
        </Card>
      </div>
    </div>
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
    Papa.parse(file,{header:true,skipEmptyLines:true,
      complete:({data})=>{
        try{
          if(!data.length) throw new Error("CSV 데이터가 없습니다");
          const f=detectFields(Object.keys(data[0]));
          if(!f.orderId) throw new Error("관리번호 컬럼을 찾을 수 없습니다");

          // 관리번호+상품명+옵션 기준 중복 합산
          const grouped={};
          data.filter(r=>r[f.orderId]).forEach(r=>{
            const oid=String(r[f.orderId]).trim();
            const prod=String(r[f.product]||"").trim();
            const opt=String(r[f.option]||"").trim();
            const ch=normChannel(r[f.channel]);
            const rawDate=r[f.date];
            const dateVal=toDate(rawDate);
            const csRaw=f.cs?String(r[f.cs]||"").trim():"";
            const statusRaw=f.status?String(r[f.status]||"").trim():"";
            const status=csRaw?normCS(csRaw):(statusRaw?"배송":"배송");
            const qty=toNum(r[f.qty])||1;
            // 관리번호+상품명+옵션 조합을 DB key로 사용 → 같은 관리번호 내 여러 상품 허용
            const dbKey=`${oid}||${prod}||${opt}`;
            if(!grouped[dbKey]){
              grouped[dbKey]={order_id:dbKey,order_date:dateVal,channel:ch,
                product_name:prod,option_name:opt,qty:0,status,raw_status:csRaw||statusRaw};
            }
            grouped[dbKey].qty+=qty;
            grouped[dbKey].status=status;
          });
          const parsed=Object.values(grouped);
          setParsedFile(parsed);
          setResult(null);
        }catch(e){setResult({type:"error",msg:e.message});}
      },
      error:e=>setResult({type:"error",msg:e.message}),
    });
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
      const {error}=await db.from("orders").insert(inRange.slice(i,i+500));
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
            <DropZone onFile={handleFile} fileName={fileName} label="이지어드민 CSV 선택"/>
            {parsedFile&&(
              <div style={{color:D.green,fontSize:11,marginTop:8}}>
                ✓ {parsedFile.length}건 파싱 완료
              </div>
            )}
            {result?.type==="error"&&<Alert type="error" msg={result.msg}/>}
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
    </div>
  );
}

// ─────────────────────────────────────────────
// DATA INPUT (탭 컨테이너)
// ─────────────────────────────────────────────
function DataInput({ onUpdate, onDataChange, orders=[], stocks=[], revenues=[] }) {
  const [tab,setTab]=useState("revenue");
  const [stockInfoOpen,setStockInfoOpen]=useState(false);
  const [orderInfoOpen,setOrderInfoOpen]=useState(false);

  const InfoBtn=({onClick})=>(
    <button onClick={onClick}
      style={{background:"transparent",border:`1px solid ${D.border}`,borderRadius:"50%",
        width:20,height:20,display:"inline-flex",alignItems:"center",justifyContent:"center",
        fontSize:10,cursor:"pointer",color:D.textSub,marginLeft:6,verticalAlign:"middle"}}>
      i
    </button>
  );

  const lastDate=(arr,field)=>{
    const d=arr.map(r=>r[field]).filter(Boolean).sort().at(-1);
    return d?<span style={{fontSize:10,color:D.textMeta,fontWeight:400,marginLeft:6}}>({d})</span>:null;
  };

  const tabs=[
    {key:"revenue",label:<span>매출 입력{lastDate(revenues,"date")}</span>},
    {key:"stock",label:<span>입고 CSV{lastDate(stocks,"upload_date")} <InfoBtn onClick={()=>setStockInfoOpen(true)}/></span>},
    {key:"orders",label:<span>이지어드민 CSV(배송일 기준){lastDate(orders,"order_date")} <InfoBtn onClick={()=>setOrderInfoOpen(true)}/></span>},
    {key:"cs",label:"CS 데이터"},
    {key:"delete",label:"데이터 삭제"},
  ];

  return (
    <div style={{padding:"20px 24px",maxWidth:1400,margin:"0 auto"}}>
      <div style={{display:"flex",borderBottom:`1px solid ${D.border}`,marginBottom:18}}>
        {tabs.map(t=>(
          <button key={t.key} onClick={()=>setTab(t.key)}
            style={{background:"transparent",border:"none",
              borderBottom:tab===t.key?`2px solid ${D.black}`:"2px solid transparent",
              color:tab===t.key?D.black:D.textSub,
              padding:"9px 16px",fontWeight:tab===t.key?600:400,
              fontSize:13,cursor:"pointer",marginBottom:-1,transition:"all 0.12s",
              display:"flex",alignItems:"center"}}>
            {t.label}
          </button>
        ))}
      </div>

      {tab==="revenue"&&<RevenueForm onUpdate={ts=>onUpdate("revenue",ts)}/>}
      {tab==="stock"&&<StockUploader onUpdate={ts=>onUpdate("stock",ts)}/>}
      {tab==="orders"&&<EasyAdminUploader onUpdate={ts=>{onUpdate("orders",ts);onDataChange?.();}}/>}
      {tab==="cs"&&<CSDataInput/>}
      {tab==="delete"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14}}>
          <DataDeleteSection table="revenues" dateField="date" label="매출 입력" onDone={()=>onDataChange?.()}/>
          <DataDeleteSection table="stock_uploads" dateField="upload_date" label="입고 CSV" onDone={()=>onDataChange?.()}/>
          <DataDeleteSection table="orders" dateField="order_date" label="이지어드민 CSV" onDone={()=>onDataChange?.()}/>
        </div>
      )}

      {/* 입고 CSV 가이드 */}
      <InfoModal show={stockInfoOpen} onClose={()=>setStockInfoOpen(false)} title="입고 CSV 업로드 가이드">
        이지어드민 재고 관리에서 기간 선택 후 작업 <strong>[배송]</strong> 선택 후 수량에 <strong>1</strong> 입력 후 <strong>애널리틱스용 양식</strong>으로 다운로드 → CSV로 변환 후 업로드
      </InfoModal>

      {/* 이지어드민 CSV 가이드 */}
      <InfoModal show={orderInfoOpen} onClose={()=>setOrderInfoOpen(false)} title="이지어드민 CSV 업로드 가이드">
        이지어드민 주문 관리 <strong>확장주문검색2</strong>에서 기간을 <strong>배송일</strong>로 설정 후 원하는 기간 선택 → 검색 후 다운로드 → CSV로 변환 후 업로드
        <br/><br/>
        <strong>CS 컬럼 상태 매핑:</strong><br/>
        • 정상 → 배송<br/>
        • 배송후 전체 교환 → 교환<br/>
        • 배송후 전체 취소 → 반품
      </InfoModal>
    </div>
  );
}

// ─────────────────────────────────────────────
// APP ROOT
// ─────────────────────────────────────────────
export default function App() {
  const [page,setPage]=useState("dashboard");
  const [orders,setOrders]=useState([]);
  const [stocks,setStocks]=useState([]);
  const [revenues,setRevenues]=useState([]);
  const [ts,setTs]=useState(()=>{
    try{return JSON.parse(localStorage.getItem("merryon_ts")||"null")||{orders:null,stock:null,revenue:null};}
    catch{return{orders:null,stock:null,revenue:null};}
  });

  const loadData=useCallback(async()=>{
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
    const [s,r]=await Promise.all([
      db.from("stock_uploads").select("*"),
      db.from("revenues").select("*").order("date",{ascending:false}),
    ]);
    setOrders(allOrders.map(r=>({...r,channel:normChannel(r.channel)})));
    setStocks(s.data||[]);
    setRevenues(r.data||[]);
  },[]);

  useEffect(()=>{ loadData(); },[loadData]);

  const updateTs=useCallback((key,val)=>setTs(prev=>{
    const next={...prev,[key]:val};
    try{localStorage.setItem("merryon_ts",JSON.stringify(next));}catch{}
    return next;
  }),[]);

  const nav=[
    {key:"dashboard",label:"대시보드"},
    {key:"flow",label:"물류 플로우"},
    {key:"promo",label:"프로모션"},
    {key:"input",label:"데이터 입력"},
  ];

  return (
    <div style={{ display:"flex", flexDirection:"column", minHeight:"100vh", background:D.bg,
      fontFamily:"'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      color:D.text, fontSize:14 }}>

      {/* top bar */}
      <div style={{ background:D.surface, borderBottom:`1px solid ${D.border}`,
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
        <div style={{ color:D.textMeta, fontSize:11, flexShrink:0 }}>
          {new Date().toLocaleDateString("ko-KR",{year:"numeric",month:"long",day:"numeric"})}
        </div>
      </div>

      {/* main content */}
      <div style={{ flex:1, overflowY:"auto" }}>
        {page==="dashboard"&&(
          <Dashboard orders={orders} stocks={stocks} revenues={revenues} ts={ts}
            onRefresh={loadData}/>
        )}
        {page==="flow"&&<LogisticsFlow orders={orders} stocks={stocks} ts={ts}/>}
        {page==="promo"&&<PromoFlow revenues={revenues}/>}
        {page==="input"&&(
          <DataInput
            onUpdate={updateTs}
            onDataChange={loadData}
            orders={orders}
            stocks={stocks}
            revenues={revenues}
          />
        )}
      </div>
    </div>
  );
}
