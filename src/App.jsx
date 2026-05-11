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
  textSub:    "#666666",
  textMeta:   "#aaaaaa",
  black:      "#111111",
  green:      "#1a7a4f",
  red:        "#c0392b",
  amber:      "#b07d00",
  blue:       "#1a4fa5",
  SANKEY: [
    "#e05a3a","#4a7fc1","#5aab6e","#d4a017","#9b59b6",
    "#1abc9c","#e67e22","#2980b9","#e91e63","#00bcd4",
    "#ff5722","#607d8b","#8bc34a","#ff9800","#673ab7",
    "#f06292","#26a69a","#ffa726","#42a5f5","#ab47bc",
  ],
};

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
  if (period === "1m") {
    const c = new Date(); c.setMonth(c.getMonth()-1);
    return rows.filter(r => r[dateField] >= c.toISOString().slice(0,10));
  }
  if (period === "3m") {
    const c = new Date(); c.setMonth(c.getMonth()-3);
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
      <span style={{ color:D.textSub, fontSize:10, letterSpacing:"0.1em", textTransform:"uppercase" }}>{children}</span>
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
// MULTI-COLUMN SANKEY  (입고 → 상품명 → 판매처 → 반품)
// ─────────────────────────────────────────────
function ProductSankey({ stockRows, orderRows, period="3m", customStart, customEnd }) {
  // period filter for orders
  const filteredOrders = useMemo(() => {
    return filterByDate(orderRows, "order_date", period, customStart, customEnd);
  }, [orderRows, period, customStart, customEnd]);

  const data = useMemo(() => {
    const prodMap = {};
    stockRows.forEach(r => {
      const key = r.product_name || "미분류";
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, shipped:0, returned:0, byChannel:{} };
      prodMap[key].stock += (r.qty||0);
    });
    filteredOrders.forEach(r => {
      const key = r.product_name || "미분류";
      const ch = r.channel || "미분류";
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, shipped:0, returned:0, byChannel:{} };
      if (!prodMap[key].byChannel[ch]) prodMap[key].byChannel[ch] = { shipped:0, returned:0 };
      if (r.status==="배송") {
        prodMap[key].shipped++;
        prodMap[key].byChannel[ch].shipped++;
      }
      if (["반품","교환"].includes(r.status)) {
        prodMap[key].returned++;
        prodMap[key].byChannel[ch].returned++;
      }
    });

    const prods = Object.values(prodMap)
      .filter(p => p.shipped>0||p.stock>0)
      .sort((a,b)=>b.shipped-a.shipped); // 수량 많은 순 내림차순

    const chanMap = {};
    filteredOrders.forEach(r => {
      const ch = r.channel||"미분류";
      if (!chanMap[ch]) chanMap[ch] = { name:ch, shipped:0, returned:0 };
      if (r.status==="배송") chanMap[ch].shipped++;
      if (["반품","교환"].includes(r.status)) chanMap[ch].returned++;
    });
    const channels = Object.values(chanMap).sort((a,b)=>b.shipped-a.shipped);
    const totalReturned = filteredOrders.filter(r=>["반품","교환"].includes(r.status)).length;

    return { prods, channels, totalReturned };
  }, [stockRows, filteredOrders]);

  if (!data.prods.length) return (
    <div style={{ textAlign:"center", padding:80, color:D.textMeta, fontSize:13 }}>
      입고 CSV 또는 이지어드민 CSV를 업로드하면<br/>상품별 물류 흐름이 표시됩니다
    </div>
  );

  const { prods, channels, totalReturned } = data;
  const n = prods.length;

  const PAD_L=20, PAD_R=20, PAD_T=36;
  const COL_W=140, NODE_H=24, ROW_GAP=5;
  const totalH = PAD_T + n*(NODE_H+ROW_GAP) + 60;
  const COLS_X = [PAD_L, PAD_L+COL_W*1.4, PAD_L+COL_W*2.8, PAD_L+COL_W*4.4];
  const SVG_W  = PAD_L + COL_W*5.8 + PAD_R;

  const headers = ["입고","상품명","판매처별 배송","반품"];
  const yOf = i => PAD_T + 20 + i*(NODE_H+ROW_GAP);

  const totalStock = prods.reduce((s,p)=>s+p.stock,0)||1;
  const totalShipped = prods.reduce((s,p)=>s+p.shipped,0)||1;
  const chanTotal = channels.reduce((s,c)=>s+c.shipped,0)||1;

  // channel y positions
  let cy = PAD_T+20;
  const chanYOf = {};
  channels.forEach(ch=>{
    const h = Math.max(NODE_H,(ch.shipped/chanTotal)*(n*(NODE_H+ROW_GAP))-ROW_GAP);
    chanYOf[ch.name] = cy + h/2;
    cy += h + ROW_GAP;
  });

  return (
    <div style={{ overflowX:"auto", overflowY:"auto", maxHeight:600 }}>
      <svg width={SVG_W} height={totalH} style={{ display:"block", minWidth:700 }}>
        {headers.map((h,ci)=>(
          <text key={h} x={COLS_X[ci]+48} y={PAD_T}
            textAnchor="middle" fill={D.textSub} fontSize="11" fontWeight="600">{h}</text>
        ))}

        {/* 입고 → 상품 연결선 */}
        {prods.map((p,i)=>{
          if (!p.stock) return null;
          const x1=COLS_X[0]+96, y1=PAD_T+20+(n*(NODE_H+ROW_GAP))/2;
          const x2=COLS_X[1],    y2=yOf(i)+NODE_H/2;
          const thick=Math.max(1,(p.stock/totalStock)*16);
          const mx=(x1+x2)/2;
          return <path key={`s${i}`} d={`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`}
            fill="none" stroke={D.SANKEY[i%D.SANKEY.length]} strokeWidth={thick} opacity={0.18}/>;
        })}

        {/* 상품 → 판매처 연결선 */}
        {prods.map((p,i)=>{
          if (!p.shipped) return null;
          const x1=COLS_X[1]+96, y1=yOf(i)+NODE_H/2;
          return Object.entries(p.byChannel).map(([ch,v])=>{
            if (!v.shipped) return null;
            const x2=COLS_X[2], y2=chanYOf[ch]||PAD_T+30;
            const thick=Math.max(1,(v.shipped/totalShipped)*14);
            const mx=(x1+x2)/2;
            return <path key={`p${i}c${ch}`} d={`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`}
              fill="none" stroke={D.SANKEY[i%D.SANKEY.length]} strokeWidth={thick} opacity={0.15}/>;
          });
        })}

        {/* 판매처 → 반품 연결선 */}
        {channels.map((ch,ci)=>{
          if (!ch.returned) return null;
          const x1=COLS_X[2]+96, y1=chanYOf[ch.name]||PAD_T+30;
          const x2=COLS_X[3],    y2=PAD_T+20+(n*(NODE_H+ROW_GAP))/2;
          const thick=Math.max(1,(ch.returned/(totalReturned||1))*18);
          const mx=(x1+x2)/2;
          return <path key={`r${ci}`} d={`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`}
            fill="none" stroke={D.red} strokeWidth={thick} opacity={0.2}/>;
        })}

        {/* 컬럼0: 입고 블록 */}
        <rect x={COLS_X[0]} y={PAD_T+20} width={96} height={n*(NODE_H+ROW_GAP)-ROW_GAP} rx={4} fill={D.SANKEY[0]} opacity={0.12}/>
        <rect x={COLS_X[0]} y={PAD_T+20} width={4} height={n*(NODE_H+ROW_GAP)-ROW_GAP} rx={2} fill={D.SANKEY[0]}/>
        <text x={COLS_X[0]+8} y={PAD_T+20+n*(NODE_H+ROW_GAP)/2} dominantBaseline="middle"
          fill={D.SANKEY[0]} fontSize="11" fontWeight="600">
          입고 {prods.reduce((s,p)=>s+p.stock,0).toLocaleString()}개
        </text>

        {/* 컬럼1: 상품 블록 */}
        {prods.map((p,i)=>{
          const y=yOf(i); const col=D.SANKEY[i%D.SANKEY.length];
          const barW=Math.max(3,Math.min(92,(p.shipped/totalShipped)*92));
          return (
            <g key={p.name}>
              <rect x={COLS_X[1]} y={y} width={96} height={NODE_H} rx={3} fill={col} opacity={0.1}/>
              <rect x={COLS_X[1]} y={y} width={barW} height={NODE_H} rx={3} fill={col} opacity={0.28}/>
              <rect x={COLS_X[1]} y={y} width={3} height={NODE_H} rx={1} fill={col}/>
              <text x={COLS_X[1]+7} y={y+NODE_H/2-2} dominantBaseline="middle"
                fill={D.black} fontSize="9" fontWeight="600">
                {p.name.length>18?p.name.slice(0,18)+"…":p.name}
              </text>
              <text x={COLS_X[1]+7} y={y+NODE_H/2+8} dominantBaseline="middle"
                fill={D.textMeta} fontSize="8">
                배송 {p.shipped} · 입고 {p.stock}
              </text>
            </g>
          );
        })}

        {/* 컬럼2: 판매처 블록 */}
        {(()=>{
          let ry=PAD_T+20;
          return channels.map((ch,ci)=>{
            const h=Math.max(NODE_H,(ch.shipped/chanTotal)*(n*(NODE_H+ROW_GAP))-ROW_GAP);
            const y=ry; ry+=h+ROW_GAP;
            const col=D.SANKEY[(ci+5)%D.SANKEY.length];
            chanYOf[ch.name]=y+h/2;
            return (
              <g key={ch.name}>
                <rect x={COLS_X[2]} y={y} width={96} height={h} rx={4} fill={col} opacity={0.12}/>
                <rect x={COLS_X[2]} y={y} width={4} height={h} rx={2} fill={col}/>
                <text x={COLS_X[2]+8} y={y+h/2-4} dominantBaseline="middle"
                  fill={col} fontSize="10" fontWeight="600">{ch.name}</text>
                <text x={COLS_X[2]+8} y={y+h/2+8} dominantBaseline="middle"
                  fill={D.textMeta} fontSize="9">{ch.shipped.toLocaleString()}건</text>
              </g>
            );
          });
        })()}

        {/* 컬럼3: 반품 블록 */}
        {totalReturned>0&&(
          <g>
            <rect x={COLS_X[3]} y={PAD_T+20} width={96} height={n*(NODE_H+ROW_GAP)-ROW_GAP}
              rx={4} fill={D.red} opacity={0.1}/>
            <rect x={COLS_X[3]} y={PAD_T+20} width={4} height={n*(NODE_H+ROW_GAP)-ROW_GAP}
              rx={2} fill={D.red}/>
            <text x={COLS_X[3]+8} y={PAD_T+20+n*(NODE_H+ROW_GAP)/2-4}
              dominantBaseline="middle" fill={D.red} fontSize="11" fontWeight="600">
              반품 {totalReturned.toLocaleString()}건
            </text>
          </g>
        )}
      </svg>
    </div>
  );
}

// ─────────────────────────────────────────────
// ANALYTICS ENGINE
// ─────────────────────────────────────────────
function analyze(orderRows, stockRows, revenueRows) {
  // 매출 입력 데이터 기반 KPI
  const totalRevenue    = revenueRows.reduce((s,r)=>s+(r.amount||0),0);
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
    byChannel[ch].revenue+=(r.amount||0);
    byChannel[ch].orderCount+=(r.order_count||0);
    byChannel[ch].refundCount+=(r.refund_count||0);
  });
  orderRows.forEach(r=>{
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,revenue:0,orderCount:0,refundCount:0,shipped:0,returned:0};
    if(r.status==="배송") byChannel[ch].shipped++;
    if(["반품","교환"].includes(r.status)) byChannel[ch].returned++;
  });
  const channelList=Object.values(byChannel).sort((a,b)=>b.revenue-a.revenue||b.shipped-a.shipped);
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
  {key:"week",label:"이번 주"},
  {key:"1m",label:"1개월"},
  {key:"3m",label:"3개월"},
  {key:"custom",label:"기간 선택"},
];

function Dashboard({ orders, stocks, revenues, ts, onRefresh }) {
  const [period,setPeriod]=useState("week");
  const [customStart,setCustomStart]=useState("");
  const [customEnd,setCustomEnd]=useState("");
  const [weekTab,setWeekTab]=useState("best");
  const [rankChannel,setRankChannel]=useState("전체");
  const [shippingPeriod,setShippingPeriod]=useState("thisMonth");
  const [returnPeriod,setReturnPeriod]=useState("1m");

  const axTick={fill:D.textMeta,fontSize:10};

  const filteredOrders=useMemo(()=>filterByDate(orders,"order_date",period,customStart,customEnd),[orders,period,customStart,customEnd]);
  const filteredRevenues=useMemo(()=>filterByDate(revenues,"date",period,customStart,customEnd),[revenues,period,customStart,customEnd]);
  const stats=useMemo(()=>analyze(filteredOrders,stocks,filteredRevenues),[filteredOrders,stocks,filteredRevenues]);

  // 판매처 채널 목록
  const activeChannels=useMemo(()=>[...new Set(filteredOrders.map(r=>r.channel||"미분류"))]
    .filter(Boolean).slice(0,6),[filteredOrders]);

  // 채널 필터된 랭킹 데이터
  const rankRows=useMemo(()=>{
    const rows=rankChannel==="전체"?stats.weekRows:stats.weekRows.filter(r=>r.channel===rankChannel);
    const byProd={};
    rows.forEach(r=>{
      const key=r.product_name||"미분류";
      if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
      byProd[key].qty+=(r.qty||0); byProd[key].orders++;
      if(["반품","교환"].includes(r.status)) byProd[key].returned++;
    });
    const list=Object.values(byProd);
    const best=[...list].sort((a,b)=>b.qty-a.qty).slice(0,20)
      .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));
    const worst=[...list].filter(p=>p.returned>0).sort((a,b)=>b.returned-a.returned).slice(0,20)
      .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));
    return {best,worst};
  },[stats.weekRows,rankChannel]);

  // 월별 배송량 차트 데이터
  const shippingChartData=useMemo(()=>{
    const today=new Date().toISOString().slice(0,10);
    if(shippingPeriod==="thisMonth"){
      const now=new Date();
      const ms=`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-01`;
      const byDay={};
      orders.filter(r=>r.order_date>=ms&&r.order_date<=today).forEach(r=>{
        const d=r.order_date;
        if(!byDay[d]) byDay[d]={date:d.slice(5),shipped:0};
        if(r.status==="배송") byDay[d].shipped++;
      });
      return Object.values(byDay).sort((a,b)=>a.date>b.date?1:-1);
    }
    const months=shippingPeriod==="3m"?3:shippingPeriod==="6m"?6:12;
    const c=new Date(); c.setMonth(c.getMonth()-months);
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
    const c=new Date();
    let start;
    if(returnPeriod==="thisMonth"){
      start=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}-01`;
    } else {
      const d=new Date(); d.setMonth(d.getMonth()-(returnPeriod==="1m"?1:3));
      start=d.toISOString().slice(0,10);
    }
    const chs=[...new Set(orders.filter(r=>r.order_date>=start).map(r=>r.channel||"미분류"))].slice(0,5);
    const byDate={};
    orders.filter(r=>r.order_date>=start).forEach(r=>{
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
      <div style={{display:"flex",gap:9,marginBottom:16,flexWrap:"wrap"}}>
        <KPI label="총 매출" value={fmtWon(stats.totalRevenue)} accent={D.black}/>
        <KPI label="총 주문" value={stats.totalOrderCount>0?stats.totalOrderCount.toLocaleString()+"건":"—"} accent={D.black}/>
        <KPI label="배송" value={stats.totalShipped.toLocaleString()+"건"} accent={D.green}/>
        <KPI label="반품·취소" value={stats.totalRefundCount>0?stats.totalRefundCount.toLocaleString()+"건":stats.totalReturned.toLocaleString()+"건"}
          sub={stats.returnRate+"%"} accent={parseFloat(stats.returnRate)>10?D.red:D.textSub}/>
        <KPI label="입고 수량" value={stats.totalStock.toLocaleString()+"개"} accent={D.blue}/>
      </div>

      {/* 판매처 점유율 + 판매처별 매출 */}
      <div style={{display:"grid",gridTemplateColumns:"280px 1fr",gap:10,marginBottom:12}}>
        <Card>
          <SecTitle ts={ts.orders}>판매처 점유율</SecTitle>
          <ResponsiveContainer width="100%" height={160}>
            <PieChart>
              <Pie data={stats.channelList.slice(0,6).map(c=>({name:c.name,value:parseFloat(c.share)}))}
                dataKey="value" nameKey="name" cx="50%" cy="50%"
                innerRadius={38} outerRadius={60} paddingAngle={2}>
                {stats.channelList.slice(0,6).map((_,i)=>(
                  <Cell key={i} fill={i===0?"#111":i===1?"#444":i===2?"#777":i===3?"#999":i===4?"#bbb":"#ddd"}/>
                ))}
              </Pie>
              <Tooltip formatter={v=>`${v}%`} contentStyle={{background:D.surface,border:`1px solid ${D.border}`,borderRadius:7,fontSize:11}}/>
              <Legend iconSize={8} iconType="circle" wrapperStyle={{fontSize:10,paddingTop:6}}/>
            </PieChart>
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
                {stats.channelList.slice(0,7).map((_,i)=>(
                  <Cell key={i} fill={i===0?"#111":i===1?"#444":i===2?"#777":"#aaa"}/>
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* 판매처 상세 */}
      <Card style={{marginBottom:12}}>
        <SecTitle ts={ts.orders}>판매처 상세</SecTitle>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
            {["판매처","점유율","주문","매출","배송","반품·취소","반품률"].map(h=>(
              <th key={h} style={{padding:"7px 9px",textAlign:h==="판매처"?"left":"right",
                color:D.textMeta,fontWeight:400}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {stats.channelList.map((c)=>(
              <tr key={c.name} style={{borderBottom:`1px solid ${D.border}`}}>
                <td style={{padding:"7px 9px",fontWeight:600}}>{c.name}</td>
                <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>{c.share}%</td>
                <td style={{textAlign:"right",padding:"7px 9px"}}>{c.orderCount>0?c.orderCount.toLocaleString():"—"}</td>
                <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600}}>{c.revenue>0?fmtWon(c.revenue):"—"}</td>
                <td style={{textAlign:"right",padding:"7px 9px",color:D.green}}>{c.shipped.toLocaleString()}</td>
                <td style={{textAlign:"right",padding:"7px 9px",color:D.red}}>{c.refundCount>0?c.refundCount.toLocaleString():c.returned.toLocaleString()}</td>
                <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600,
                  color:parseFloat(c.returnRate)>10?D.red:D.textSub}}>{c.returnRate}%</td>
              </tr>
            ))}
            <tr style={{borderTop:`1px solid ${D.borderMid}`}}>
              <td style={{padding:"7px 9px",fontWeight:700}}>합계</td>
              <td style={{textAlign:"right",padding:"7px 9px",color:D.textSub}}>100%</td>
              <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600}}>{stats.totalOrderCount>0?stats.totalOrderCount.toLocaleString():"—"}</td>
              <td style={{textAlign:"right",padding:"7px 9px",fontWeight:700}}>{fmtWon(stats.totalRevenue)}</td>
              <td style={{textAlign:"right",padding:"7px 9px",color:D.green,fontWeight:600}}>{stats.totalShipped.toLocaleString()}</td>
              <td style={{textAlign:"right",padding:"7px 9px",color:D.red,fontWeight:600}}>{stats.totalRefundCount>0?stats.totalRefundCount.toLocaleString():stats.totalReturned.toLocaleString()}</td>
              <td style={{textAlign:"right",padding:"7px 9px",fontWeight:600,
                color:parseFloat(stats.returnRate)>10?D.red:D.textSub}}>{stats.returnRate}%</td>
            </tr>
          </tbody>
        </table>
      </Card>

      {/* 월별 배송량 (독립 기간) */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <SecTitle ts={ts.orders}>월별 배송량</SecTitle>
            <div style={{display:"flex",gap:4}}>
              {[["thisMonth","이번달"],["3m","3개월"],["6m","6개월"]].map(([v,l])=>(
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
              <Bar dataKey="shipped" name="배송" fill="#111" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        {/* 판매처별 일자 반품 */}
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <SecTitle ts={ts.orders}>판매처별 반품 추이</SecTitle>
            <div style={{display:"flex",gap:4}}>
              {[["thisMonth","이번달"],["1m","1개월"],["3m","3개월"]].map(([v,l])=>(
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

      {/* 주간 상품 랭킹 */}
      <Card>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <SecTitle ts={ts.orders}>주간 상품 랭킹{stats.latestWeek?` — ${stats.latestWeek}`:""}</SecTitle>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {/* 채널 필터 */}
            {["전체",...activeChannels].map(ch=>(
              <button key={ch} onClick={()=>setRankChannel(ch)}
                style={{background:rankChannel===ch?D.black:"transparent",
                  color:rankChannel===ch?"#fff":D.textSub,
                  border:`1px solid ${rankChannel===ch?D.black:D.border}`,
                  borderRadius:5,padding:"3px 9px",fontSize:10,cursor:"pointer"}}>
                {ch}
              </button>
            ))}
            <div style={{width:1,background:D.border,margin:"0 2px"}}/>
            {[["best","판매 Top"],["worst","반품 Top"]].map(([k,l])=>(
              <button key={k} onClick={()=>setWeekTab(k)}
                style={{background:"transparent",border:"none",
                  borderBottom:weekTab===k?`2px solid ${D.black}`:"2px solid transparent",
                  color:weekTab===k?D.black:D.textSub,padding:"3px 10px",
                  fontWeight:weekTab===k?600:400,fontSize:11,cursor:"pointer"}}>
                {l}
              </button>
            ))}
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <RankTable
            data={weekTab==="best"?rankRows.best:rankRows.worst}
            cols={weekTab==="best"?[
              {key:"name",label:"상품명",maxW:190},
              {key:"qty",label:"배송량",right:true,bold:true,fmt:v=>v.toLocaleString()},
              {key:"returnRate",label:"반품률",right:true,color:D.textMeta,fmt:v=>v+"%"},
            ]:[
              {key:"name",label:"상품명",maxW:190},
              {key:"returned",label:"반품",right:true,bold:true,color:D.red,fmt:v=>v.toLocaleString()},
              {key:"returnRate",label:"반품률",right:true,color:D.red,fmt:v=>v+"%"},
              {key:"qty",label:"배송량",right:true,color:D.textSub,fmt:v=>v.toLocaleString()},
            ]}
          />
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={(weekTab==="best"?rankRows.best:rankRows.worst).slice(0,12)} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
              <XAxis type="number" tick={axTick}/>
              <YAxis type="category" dataKey="name" width={130} tick={{...axTick,fontSize:9}}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey={weekTab==="best"?"qty":"returned"} name={weekTab==="best"?"배송량":"반품"} radius={[0,3,3,0]}>
                {(weekTab==="best"?rankRows.best:rankRows.worst).slice(0,12).map((_,i)=>(
                  <Cell key={i} fill={weekTab==="best"?(i===0?"#111":i===1?"#444":i===2?"#777":"#aaa"):D.red}/>
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>
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
          <div style={{color:D.black,fontWeight:600,fontSize:15,marginBottom:3}}>물류 플로우</div>
          <div style={{color:D.textMeta,fontSize:12}}>
            입고 → 상품명 → 판매처별 배송 → 반품 · 전체 상품 표시
            <UpdatedAt ts={ts.orders||ts.stock}/>
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

      <Card style={{overflowX:"auto",marginBottom:12}}>
        <ProductSankey stockRows={stocks} orderRows={orders} period={period} customStart={customStart} customEnd={customEnd}/>
      </Card>

      {/* 상품별 흐름 요약 — 배송/반품으로 통일, 옵션 없이 상품명 기준 */}
      {filteredOrders.length>0&&(
        <Card>
          <SecTitle ts={ts.orders}>상품별 흐름 요약</SecTitle>
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
                  return Object.values(prodMap)
                    .filter(p=>p.shipped>0||p.stock>0)
                    .sort((a,b)=>b.shipped-a.shipped)
                    .map((p,i)=>{
                      const total=p.shipped+p.returned;
                      const rr=total>0?(p.returned/total*100).toFixed(1):"0.0";
                      return(
                        <tr key={p.name} style={{borderBottom:`1px solid ${D.border}`}}>
                          <td style={{padding:"6px 9px",color:D.textMeta,textAlign:"right"}}>{i+1}</td>
                          <td style={{padding:"6px 9px",fontWeight:i<3?700:400,maxWidth:220,
                            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",color:D.blue}}>{p.stock.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",fontWeight:600,color:D.green}}>{p.shipped.toLocaleString()}</td>
                          <td style={{textAlign:"right",padding:"6px 9px",color:D.red}}>{p.returned.toLocaleString()}</td>
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
// DATA INPUT — 매출 입력
// ─────────────────────────────────────────────
const REVENUE_CHANNELS = ["자사몰","29CM","무신사"];

function RevenueForm({ onUpdate }) {
  const today=new Date().toISOString().slice(0,10);
  const [date,setDate]=useState(today);
  const [ch,setCh]=useState(REVENUE_CHANNELS[0]);
  const [amt,setAmt]=useState("");
  const [orderCnt,setOrderCnt]=useState("");
  const [refundAmt,setRefundAmt]=useState("");
  const [refundCnt,setRefundCnt]=useState("");
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
    const {error}=await db.from("revenues").upsert({
      date,channel:ch,
      amount:num,
      order_count:Number(orderCnt)||0,
      refund_amount:Number(refundAmt.replace(/,/g,""))||0,
      refund_count:Number(refundCnt)||0,
    },{onConflict:"date,channel"});
    const ts2=nowStr();
    if(error) setResult({type:"error",msg:error.message});
    else {
      setResult({type:"success",msg:`저장 완료`,ts:ts2});
      setAmt("");setOrderCnt("");setRefundAmt("");setRefundCnt("");
      onUpdate(ts2); if(history.length) loadHistory();
    }
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

  const inp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,
    padding:"7px 10px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box"};
  const numInp=(v,fn)=>(
    <input type="text" value={v} onChange={e=>fn(e.target.value.replace(/[^0-9,]/g,""))} style={inp}/>
  );

  return (
    <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
      <Card>
        <div style={{fontWeight:600,marginBottom:14,fontSize:13}}>매출 입력</div>

        <div style={{marginBottom:10}}>
          <div style={{color:D.textMeta,fontSize:10,marginBottom:6}}>날짜</div>
          <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={inp}/>
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
      </Card>

      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
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
            <div style={{color:D.red,fontSize:10,marginBottom:12}}>⚠ 확정 시 해당 기간 DB 데이터 전체 교체</div>
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
          <div style={{fontWeight:500,fontSize:12,marginBottom:12}}>
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

  // Step 2→3: 확정 업로드 (order_id=관리번호||상품||옵션 복합키 → 같은 관리번호 내 여러 상품 허용)
  const handleUpload=async()=>{
    if(!inRange.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    for(let i=0;i<inRange.length;i+=500){
      const {error}=await db.from("orders").upsert(inRange.slice(i,i+500),{onConflict:"order_id"});
      if(error){setResult({type:"error",msg:"업로드 실패: "+error.message});setLoading(false);return;}
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
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
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
function DataInput({ onUpdate, onDataChange }) {
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

  const tabs=[
    {key:"revenue",label:"매출 입력"},
    {key:"stock",label:<span>입고 CSV <InfoBtn onClick={()=>setStockInfoOpen(true)}/></span>},
    {key:"orders",label:<span>이지어드민 CSV(배송일 기준) <InfoBtn onClick={()=>setOrderInfoOpen(true)}/></span>},
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
  const [ts,setTs]=useState({orders:null,stock:null,revenue:null});

  const loadData=useCallback(async()=>{
    const db=await getSupabase();
    const [o,s,r]=await Promise.all([
      db.from("orders").select("*").order("order_date",{ascending:false}),
      db.from("stock_uploads").select("*"),
      db.from("revenues").select("*").order("date",{ascending:false}),
    ]);
    setOrders((o.data||[]).map(r=>({...r,channel:normChannel(r.channel)})));
    setStocks(s.data||[]);
    setRevenues(r.data||[]);
  },[]);

  useEffect(()=>{ loadData(); },[loadData]);

  const updateTs=useCallback((key,val)=>setTs(prev=>({...prev,[key]:val})),[]);

  const nav=[
    {key:"dashboard",label:"대시보드",icon:"▦"},
    {key:"flow",label:"물류 플로우",icon:"⟶"},
    {key:"input",label:"데이터 입력",icon:"↑"},
  ];

  return (
    <div style={{ display:"flex", minHeight:"100vh", background:D.bg,
      fontFamily:"'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      color:D.text, fontSize:14 }}>

      {/* sidebar */}
      <div style={{ width:180, background:D.surface, borderRight:`1px solid ${D.border}`,
        padding:"18px 10px", display:"flex", flexDirection:"column", flexShrink:0 }}>
        <div style={{ color:D.black, fontWeight:800, fontSize:12, marginBottom:1, paddingLeft:4, letterSpacing:"0.05em" }}>MERRYON</div>
        <div style={{ color:D.black, fontWeight:700, fontSize:11, marginBottom:1, paddingLeft:4 }}>COMMERCE</div>
        <div style={{ color:D.textMeta, fontSize:9, letterSpacing:"0.1em", marginBottom:22, paddingLeft:4 }}>WORK FLOW</div>
        <nav style={{ display:"flex", flexDirection:"column", gap:2 }}>
          {nav.map(n=>(
            <button key={n.key} onClick={()=>setPage(n.key)}
              style={{ background:page===n.key?D.surfaceAlt:"transparent",
                color:page===n.key?D.black:D.textSub,
                border:"none", borderRadius:6, padding:"8px 10px",
                textAlign:"left", cursor:"pointer", fontSize:12,
                fontWeight:page===n.key?600:400, transition:"all 0.1s",
                display:"flex", alignItems:"center", gap:8 }}>
              <span style={{ fontSize:11 }}>{n.icon}</span>{n.label}
            </button>
          ))}
        </nav>
        <div style={{ marginTop:"auto", paddingLeft:4, color:D.textMeta, fontSize:10 }}>
          make by jihoon
        </div>
      </div>

      {/* main */}
      <div style={{ flex:1, overflowY:"auto" }}>
        <div style={{ background:D.surface, borderBottom:`1px solid ${D.border}`,
          padding:"11px 24px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <div style={{ color:D.black, fontWeight:600, fontSize:13 }}>
            {nav.find(n=>n.key===page)?.label}
          </div>
          <div style={{ color:D.textMeta, fontSize:11 }}>
            {new Date().toLocaleDateString("ko-KR",{year:"numeric",month:"long",day:"numeric"})}
          </div>
        </div>
        {page==="dashboard"&&(
          <Dashboard orders={orders} stocks={stocks} revenues={revenues} ts={ts}
            onRefresh={loadData}/>
        )}
        {page==="flow"&&<LogisticsFlow orders={orders} stocks={stocks} ts={ts}/>}
        {page==="input"&&(
          <DataInput
            onUpdate={updateTs}
            onDataChange={loadData}
          />
        )}
      </div>
    </div>
  );
}
