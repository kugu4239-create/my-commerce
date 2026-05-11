import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import Papa from "papaparse";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, PieChart, Pie,
} from "recharts";

// ─────────────────────────────────────────────
// DESIGN TOKENS  (White & Black system)
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
  offBlack:   "#333333",
  green:      "#1a7a4f",
  red:        "#c0392b",
  amber:      "#b07d00",
  blue:       "#1a4fa5",
  // sankey palette — stays colorful for readability
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
const toNum  = v => parseFloat(String(v||"0").replace(/[^0-9.-]/g,""))||0;
const toDate = raw => {
  if (!raw) return null;
  const s = String(raw).trim();
  const m1 = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m1) return `${m1[1]}-${m1[2].padStart(2,"0")}-${m1[3].padStart(2,"0")}`;
  const m2 = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if (m2) return `${m2[3]}-${m2[2].padStart(2,"0")}-${m2[1].padStart(2,"0")}`;
  return null;
};
const normStatus = raw => {
  if (!raw) return "기타";
  const v = String(raw).toLowerCase().replace(/\s/g,"");
  if (v.includes("반품")||v.includes("return"))        return "반품";
  if (v.includes("취소")||v.includes("cancel")||v.includes("환불")) return "취소";
  if (v.includes("배송완료")||v.includes("완료"))       return "배송완료";
  if (v.includes("배송중")||v.includes("발송"))         return "배송중";
  if (v.includes("출고")||v.includes("ship"))           return "출고";
  if (v.includes("입금")||v.includes("결제"))           return "결제완료";
  if (v.includes("입고"))                               return "입고";
  return raw;
};
const fmtWon = n => {
  if (!n) return "—";
  if (n>=1e8) return "₩"+(n/1e8).toFixed(1)+"억";
  if (n>=1e4) return "₩"+(n/1e4).toFixed(0)+"만";
  return "₩"+n.toLocaleString();
};
function detectFields(columns) {
  const lc = columns.map(c=>c.toLowerCase().replace(/\s/g,""));
  const f  = (...kws) => { const i=lc.findIndex(c=>kws.some(k=>c.includes(k))); return i>=0?columns[i]:null; };
  return {
    channel:  f("판매처","channel","플랫폼","채널","mall","store","platform"),
    product:  f("상품명","product","품명","item","name"),
    option:   f("옵션","option","size","color","사이즈","색상"),
    qty:      f("수량","qty","quantity","개수","판매수량"),
    status:   f("상태","status","주문상태","처리상태","배송상태"),
    date:     f("주문일","날짜","date","order_date","주문날짜","reg_date"),
    revenue:  f("금액","revenue","sales","매출","price","가격","결제금액","주문금액"),
    orderId:  f("관리번호","order_id","주문번호","orderid"),
    memo:     f("메모","memo","비고","note"),
  };
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
    primary:  { bg:D.black,   cl:"#fff",      bd:"none" },
    ghost:    { bg:"transparent", cl:D.textSub, bd:`1px solid ${D.border}` },
    danger:   { bg:D.red,     cl:"#fff",      bd:"none" },
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
function PreviewTable({ rows, cols, outIdx=new Set(), maxRows=60 }) {
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

// ─────────────────────────────────────────────
// MULTI-COLUMN SANKEY  (상품별 흐름)
// ─────────────────────────────────────────────
// columns: 입고 | 판매처 | 상품명 | 상태
// flows:   stockRows → orderRows matched by product+option
function ProductSankey({ stockRows, orderRows }) {
  const svgRef = useRef(null);

  const data = useMemo(() => {
    // ── 상품별 집계 ─────────────────────────────
    const prodMap = {};

    // 입고 데이터
    stockRows.forEach(r => {
      const key = [r.product_name, r.option_name].filter(Boolean).join(" / ");
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, ordered:0, shipped:0, done:0, inDel:0, returned:0, cancelled:0 };
      prodMap[key].stock += (r.qty||0);
    });

    // 주문 데이터
    orderRows.forEach(r => {
      const key = [r.product_name, r.option_name].filter(Boolean).join(" / ");
      if (!prodMap[key]) prodMap[key] = { name:key, stock:0, ordered:0, shipped:0, done:0, inDel:0, returned:0, cancelled:0 };
      prodMap[key].ordered++;
      if (["배송완료","배송중","출고"].includes(r.status)) prodMap[key].shipped++;
      if (r.status==="배송완료") prodMap[key].done++;
      if (r.status==="배송중")   prodMap[key].inDel++;
      if (r.status==="반품")     prodMap[key].returned++;
      if (r.status==="취소")     prodMap[key].cancelled++;
    });

    // 상품 필터 (주문 또는 입고 있는 것만), 주문 기준 정렬
    const prods = Object.values(prodMap)
      .filter(p => p.ordered > 0 || p.stock > 0)
      .sort((a,b) => b.ordered - a.ordered)
      .slice(0, 20); // 최대 20개

    return prods;
  }, [stockRows, orderRows]);

  if (!data.length) return (
    <div style={{ textAlign:"center", padding:80, color:D.textMeta, fontSize:13 }}>
      입고 CSV 또는 카페24 주문 CSV를 업로드하면<br/>상품별 물류 흐름이 표시됩니다
    </div>
  );

  // ── SVG 레이아웃 계산 ──────────────────────
  const PAD_L  = 20;
  const PAD_R  = 20;
  const PAD_T  = 30;
  const COL_W  = 130;   // 각 컬럼 너비
  const NODE_H = 22;    // 각 노드 높이
  const ROW_GAP= 6;     // 노드 사이 갭
  const COLS   = 4;     // 입고 | 판매채널 | 상품 | 배송상태
  const n      = data.length;
  const totalH = PAD_T + n * (NODE_H + ROW_GAP) + 60;
  const COLS_X = [PAD_L, PAD_L + COL_W*1.5, PAD_L + COL_W*3, PAD_L + COL_W*4.5];
  const SVG_W  = PAD_L + COL_W*6 + PAD_R;

  // 컬럼 헤더
  const headers = ["입고", "판매처", "상품명", "배송 현황"];

  // 상품별 y위치
  const yOf = i => PAD_T + 20 + i*(NODE_H+ROW_GAP);

  // 배송상태별 집계 (4번째 컬럼용)
  const statusTotals = { 배송완료:0, 배송중:0, 반품:0, 취소:0, 기타:0 };
  data.forEach(p => {
    statusTotals["배송완료"] += p.done;
    statusTotals["배송중"]   += p.inDel;
    statusTotals["반품"]     += p.returned;
    statusTotals["취소"]     += p.cancelled;
    statusTotals["기타"]     += Math.max(0, p.ordered - p.shipped - p.returned - p.cancelled);
  });
  const statusList = Object.entries(statusTotals).filter(([,v])=>v>0);
  const statusColors = { 배송완료:D.green, 배송중:D.blue, 반품:D.red, 취소:D.red, 기타:D.textMeta };
  // y positions for status nodes
  const statusTotalCount = statusList.reduce((s,[,v])=>s+v,0)||1;
  let statusY = PAD_T + 20;
  const statusYOf = {};
  statusList.forEach(([k,v])=>{
    statusYOf[k] = statusY + (v/statusTotalCount)*(n*(NODE_H+ROW_GAP))/2;
    statusY += (v/statusTotalCount)*(n*(NODE_H+ROW_GAP)) + ROW_GAP;
  });

  // 판매처 집계
  const chanMap = {};
  orderRows.forEach(r=>{
    const ch = r.channel||"미분류";
    if (!chanMap[ch]) chanMap[ch]=0;
    chanMap[ch]++;
  });
  const chanList = Object.entries(chanMap).sort((a,b)=>b[1]-a[1]).slice(0,8);
  const chanTotal = chanList.reduce((s,[,v])=>s+v,0)||1;
  let chanY = PAD_T + 20;
  const chanYOf = {};
  chanList.forEach(([k,v])=>{
    chanYOf[k] = chanY + (v/chanTotal)*(n*(NODE_H+ROW_GAP))/2;
    chanY += (v/chanTotal)*(n*(NODE_H+ROW_GAP)) + ROW_GAP;
  });

  // 입고 집계
  const totalStock = data.reduce((s,p)=>s+p.stock,0)||1;
  const totalOrdered = data.reduce((s,p)=>s+p.ordered,0)||1;

  return (
    <div style={{ overflowX:"auto", overflowY:"hidden" }}>
      <svg width={SVG_W} height={totalH} style={{ display:"block", minWidth:600 }}>
        {/* 컬럼 헤더 */}
        {headers.map((h,ci)=>(
          <text key={h} x={COLS_X[ci]+48} y={PAD_T} textAnchor="middle"
            fill={D.textSub} fontSize="11" fontWeight="600">{h}</text>
        ))}

        {/* 연결선 — 입고 → 상품 */}
        {data.map((p,i)=>{
          const x1=COLS_X[0]+96, y1=PAD_T+20+i*(NODE_H+ROW_GAP)+NODE_H/2;
          const x2=COLS_X[2],    y2=yOf(i)+NODE_H/2;
          const hasStock = p.stock>0;
          if (!hasStock) return null;
          const thick=Math.max(1,(p.stock/totalStock)*18);
          const mx=(x1+x2)/2;
          return (
            <path key={`stock-${i}`}
              d={`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`}
              fill="none" stroke={D.SANKEY[i%D.SANKEY.length]} strokeWidth={thick} opacity={0.18}/>
          );
        })}

        {/* 연결선 — 판매처 → 상품 */}
        {data.map((p,i)=>{
          const x2=COLS_X[2], y2=yOf(i)+NODE_H/2;
          if (!p.ordered) return null;
          // 주 판매처 찾기 (단순화)
          const mainCh=chanList[0]?.[0]||"미분류";
          const cy=chanYOf[mainCh]||PAD_T+30;
          const x1=COLS_X[1]+96;
          const thick=Math.max(1,(p.ordered/totalOrdered)*14);
          const mx=(x1+x2)/2;
          return (
            <path key={`chan-${i}`}
              d={`M${x1},${cy} C${mx},${cy} ${mx},${y2} ${x2},${y2}`}
              fill="none" stroke={D.SANKEY[(i+5)%D.SANKEY.length]} strokeWidth={thick} opacity={0.13}/>
          );
        })}

        {/* 연결선 — 상품 → 배송상태 */}
        {data.map((p,i)=>{
          const x1=COLS_X[2]+96, y1=yOf(i)+NODE_H/2;
          const links=[
            {k:"배송완료",v:p.done},
            {k:"배송중",  v:p.inDel},
            {k:"반품",    v:p.returned},
            {k:"취소",    v:p.cancelled},
          ].filter(l=>l.v>0);
          return links.map(({k,v})=>{
            const x2=COLS_X[3], y2=statusYOf[k]||PAD_T+30;
            const thick=Math.max(1,(v/totalOrdered)*18);
            const mx=(x1+x2)/2;
            return (
              <path key={`stat-${i}-${k}`}
                d={`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`}
                fill="none" stroke={statusColors[k]||D.textMeta} strokeWidth={thick} opacity={0.22}/>
            );
          });
        })}

        {/* ── 컬럼0: 입고 노드 (하나로 합산) ── */}
        <rect x={COLS_X[0]} y={PAD_T+20} width={96} height={n*(NODE_H+ROW_GAP)-ROW_GAP}
          rx={4} fill={D.SANKEY[0]} opacity={0.15}/>
        <rect x={COLS_X[0]} y={PAD_T+20} width={4} height={n*(NODE_H+ROW_GAP)-ROW_GAP}
          rx={2} fill={D.SANKEY[0]}/>
        <text x={COLS_X[0]+8} y={PAD_T+20+n*(NODE_H+ROW_GAP)/2}
          dominantBaseline="middle" fill={D.SANKEY[0]} fontSize="11" fontWeight="600">
          입고 {data.reduce((s,p)=>s+p.stock,0).toLocaleString()}개
        </text>

        {/* ── 컬럼1: 판매처 노드 ── */}
        {(() => {
          let cy = PAD_T+20;
          return chanList.map(([ch,cnt],ci)=>{
            const h=Math.max(NODE_H,(cnt/chanTotal)*(n*(NODE_H+ROW_GAP))-ROW_GAP);
            const y=cy; cy+=h+ROW_GAP;
            const col=D.SANKEY[(ci+5)%D.SANKEY.length];
            return (
              <g key={ch}>
                <rect x={COLS_X[1]} y={y} width={96} height={h} rx={4} fill={col} opacity={0.12}/>
                <rect x={COLS_X[1]} y={y} width={4} height={h} rx={2} fill={col}/>
                <text x={COLS_X[1]+8} y={y+h/2}
                  dominantBaseline="middle" fill={col} fontSize="10" fontWeight="600">
                  {ch}
                </text>
                <text x={COLS_X[1]+8} y={y+h/2+12}
                  dominantBaseline="middle" fill={D.textMeta} fontSize="9">
                  {cnt.toLocaleString()}건
                </text>
              </g>
            );
          });
        })()}

        {/* ── 컬럼2: 상품 노드 ── */}
        {data.map((p,i)=>{
          const y=yOf(i);
          const col=D.SANKEY[i%D.SANKEY.length];
          const barW=Math.max(4,Math.min(92,(p.ordered/totalOrdered)*92));
          return (
            <g key={p.name}>
              <rect x={COLS_X[2]} y={y} width={96} height={NODE_H} rx={3} fill={col} opacity={0.1}/>
              <rect x={COLS_X[2]} y={y} width={barW} height={NODE_H} rx={3} fill={col} opacity={0.25}/>
              <rect x={COLS_X[2]} y={y} width={3} height={NODE_H} rx={1} fill={col}/>
              <text x={COLS_X[2]+7} y={y+NODE_H/2-3}
                dominantBaseline="middle" fill={D.black} fontSize="9" fontWeight="500">
                {p.name.length>18?p.name.slice(0,18)+"…":p.name}
              </text>
              <text x={COLS_X[2]+7} y={y+NODE_H/2+6}
                dominantBaseline="middle" fill={D.textMeta} fontSize="8">
                주문 {p.ordered} · 입고 {p.stock}
              </text>
            </g>
          );
        })}

        {/* ── 컬럼3: 배송상태 노드 ── */}
        {(() => {
          let sy=PAD_T+20;
          return statusList.map(([k,v])=>{
            const h=Math.max(NODE_H,(v/statusTotalCount)*(n*(NODE_H+ROW_GAP))-ROW_GAP);
            const y=sy; sy+=h+ROW_GAP;
            const col=statusColors[k]||D.textMeta;
            statusYOf[k]=y+h/2;
            return (
              <g key={k}>
                <rect x={COLS_X[3]} y={y} width={96} height={h} rx={3} fill={col} opacity={0.1}/>
                <rect x={COLS_X[3]} y={y} width={4} height={h} rx={2} fill={col}/>
                <text x={COLS_X[3]+8} y={y+h/2-4}
                  dominantBaseline="middle" fill={col} fontSize="10" fontWeight="600">
                  {k}
                </text>
                <text x={COLS_X[3]+8} y={y+h/2+7}
                  dominantBaseline="middle" fill={D.textMeta} fontSize="9">
                  {v.toLocaleString()}건
                </text>
              </g>
            );
          });
        })()}
      </svg>
    </div>
  );
}

// ─────────────────────────────────────────────
// ANALYTICS ENGINE
// ─────────────────────────────────────────────
function analyze(orderRows, stockRows, revenueRows) {
  const totalRevenue  = revenueRows.reduce((s,r)=>s+(r.amount||0),0);
  const totalOrders   = orderRows.length;
  const totalShipped  = orderRows.filter(r=>["배송완료","배송중","출고"].includes(r.status)).length;
  const totalReturned = orderRows.filter(r=>["반품","취소"].includes(r.status)).length;
  const returnRate    = totalOrders>0?(totalReturned/totalOrders*100).toFixed(1):0;
  const totalStock    = stockRows.reduce((s,r)=>s+(r.qty||0),0);

  const byChannel={};
  orderRows.forEach(r=>{
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,orders:0,qty:0,shipped:0,returned:0,revenue:0};
    byChannel[ch].orders++;
    byChannel[ch].qty+=(r.qty||0);
    if(["배송완료","배송중","출고"].includes(r.status)) byChannel[ch].shipped++;
    if(["반품","취소"].includes(r.status)) byChannel[ch].returned++;
  });
  revenueRows.forEach(r=>{
    const ch=r.channel||"미분류";
    if(!byChannel[ch]) byChannel[ch]={name:ch,orders:0,qty:0,shipped:0,returned:0,revenue:0};
    byChannel[ch].revenue+=(r.amount||0);
  });
  const channelList=Object.values(byChannel).sort((a,b)=>b.revenue-a.revenue);
  const totalRev=channelList.reduce((s,c)=>s+(c.revenue||0),0)||1;
  channelList.forEach(c=>{
    c.share=((c.revenue||0)/totalRev*100).toFixed(1);
    c.returnRate=c.orders>0?(c.returned/c.orders*100).toFixed(1):"0.0";
  });

  const byMonth={};
  orderRows.forEach(r=>{
    const ym=r.order_date?r.order_date.slice(0,7):null;
    if(!ym) return;
    if(!byMonth[ym]) byMonth[ym]={month:ym,total:0,shipped:0,returned:0};
    byMonth[ym].total++;
    if(["배송완료","배송중","출고"].includes(r.status)) byMonth[ym].shipped++;
    if(["반품","취소"].includes(r.status)) byMonth[ym].returned++;
  });
  const monthlyData=Object.values(byMonth)
    .sort((a,b)=>a.month>b.month?1:-1)
    .map(m=>({...m,returnRate:m.total>0?(m.returned/m.total*100).toFixed(1):0}));

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
    const key=[r.product_name,r.option_name].filter(Boolean).join(" / ");
    if(!byProd[key]) byProd[key]={name:key,qty:0,orders:0,returned:0};
    byProd[key].qty+=(r.qty||0); byProd[key].orders++;
    if(["반품","취소"].includes(r.status)) byProd[key].returned++;
  });
  const prodList=Object.values(byProd);
  const weekBest=[...prodList].sort((a,b)=>b.qty-a.qty).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));
  const weekWorst=[...prodList].filter(p=>p.returned>0).sort((a,b)=>b.returned-a.returned).slice(0,20)
    .map(p=>({...p,returnRate:p.orders>0?(p.returned/p.orders*100).toFixed(1):"0.0"}));

  return {
    totalRevenue,totalOrders,totalShipped,totalReturned,returnRate,totalStock,
    channelList,monthlyData,weekBest,weekWorst,latestWeek,
  };
}

// ─────────────────────────────────────────────
// SUPABASE CLIENT  (실제 배포 시 교체)
// ─────────────────────────────────────────────
const SUPA_URL = typeof import.meta!=="undefined"&&import.meta.env?.VITE_SUPABASE_URL || "";
const SUPA_KEY = typeof import.meta!=="undefined"&&import.meta.env?.VITE_SUPABASE_ANON_KEY || "";
let _supabase = null;
async function getSupabase() {
  if (_supabase) return _supabase;
  if (!SUPA_URL||!SUPA_KEY) {
    // stub for local preview
    _supabase = {
      from: () => ({
        select:()=>({ order:()=>({ limit:()=>Promise.resolve({data:[],error:null}), ascending:false }),
          gte:()=>({ lte:()=>({ order:()=>Promise.resolve({data:[],error:null}) }) }),
          in:()=>Promise.resolve({data:[],error:null}),
        }),
        insert:rows=>Promise.resolve({data:rows,error:null}),
        upsert:(rows,o)=>Promise.resolve({data:rows,error:null}),
        delete:()=>({ gte:()=>({ lte:()=>Promise.resolve({error:null}) }), eq:()=>Promise.resolve({error:null}) }),
      }),
    };
    return _supabase;
  }
  const { createClient } = await import("@supabase/supabase-js");
  _supabase = createClient(SUPA_URL, SUPA_KEY);
  return _supabase;
}

// ─────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────
function Dashboard({ orders, stocks, revenues, ts }) {
  const [weekTab, setWeekTab] = useState("best");
  const [period,  setPeriod]  = useState("all");
  const axTick = { fill:D.textMeta, fontSize:10 };

  const filtered = useMemo(()=>{
    if (period==="all") return orders;
    const months=period==="3m"?3:1;
    const cutoff=new Date(); cutoff.setMonth(cutoff.getMonth()-months);
    const cutStr=cutoff.toISOString().slice(0,10);
    return orders.filter(r=>r.order_date>=cutStr);
  },[orders,period]);

  const stats = useMemo(()=>analyze(filtered,stocks,revenues),[filtered,stocks,revenues]);

  function RankTable({ data, cols }) {
    return (
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
        <thead><tr>
          <th style={{ padding:"5px 7px",color:D.textMeta,fontWeight:400,borderBottom:`1px solid ${D.border}`,width:22 }}>#</th>
          {cols.map(c=><th key={c.key} style={{ padding:"5px 7px",textAlign:c.right?"right":"left",color:D.textMeta,fontWeight:400,borderBottom:`1px solid ${D.border}` }}>{c.label}</th>)}
        </tr></thead>
        <tbody>
          {data.map((row,i)=>(
            <tr key={i} style={{ borderBottom:`1px solid ${D.border}` }}>
              <td style={{ padding:"5px 7px",color:i<3?D.black:D.textMeta,fontWeight:i<3?600:400 }}>{i+1}</td>
              {cols.map(c=><td key={c.key} style={{ padding:"5px 7px",textAlign:c.right?"right":"left",
                color:c.color||D.text,fontWeight:c.bold?600:400,maxWidth:c.maxW,
                overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>
                {c.fmt?c.fmt(row[c.key],row):row[c.key]}
              </td>)}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  return (
    <div style={{ padding:"20px 24px", maxWidth:1400, margin:"0 auto" }}>
      <div style={{ display:"flex", justifyContent:"flex-end", gap:5, marginBottom:16 }}>
        {[["all","전체"],["3m","3개월"],["1m","1개월"]].map(([k,l])=>(
          <button key={k} onClick={()=>setPeriod(k)}
            style={{ background:"transparent", border:`1px solid ${period===k?D.black:D.border}`,
              color:period===k?D.black:D.textSub, borderRadius:6,
              padding:"4px 11px", fontSize:11, cursor:"pointer",
              fontWeight:period===k?600:400 }}>
            {l}
          </button>
        ))}
      </div>

      <div style={{ display:"flex", gap:9, marginBottom:16, flexWrap:"wrap" }}>
        <KPI label="총 매출" value={fmtWon(stats.totalRevenue)} accent={D.black}/>
        <KPI label="총 주문" value={stats.totalOrders.toLocaleString()+"건"} accent={D.black}/>
        <KPI label="배송 완료" value={stats.totalShipped.toLocaleString()+"건"} accent={D.green}/>
        <KPI label="반품·취소" value={stats.totalReturned.toLocaleString()+"건"}
          sub={stats.returnRate+"%"} accent={parseFloat(stats.returnRate)>10?D.red:D.textSub}/>
        <KPI label="입고 수량" value={stats.totalStock.toLocaleString()+"개"} accent={D.blue}/>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"240px 1fr", gap:10, marginBottom:12 }}>
        <Card>
          <SecTitle ts={ts.orders}>판매처 점유율</SecTitle>
          <ResponsiveContainer width="100%" height={190}>
            <PieChart>
              <Pie data={stats.channelList.slice(0,6).map(c=>({name:c.name,value:parseFloat(c.share)}))}
                dataKey="value" nameKey="name" cx="50%" cy="50%"
                innerRadius={44} outerRadius={72} paddingAngle={2}
                label={({name,value})=>`${name} ${value}%`} labelLine={false}>
                {stats.channelList.slice(0,6).map((_,i)=>(
                  <Cell key={i} fill={i===0?"#111":i===1?"#444":i===2?"#777":"#aaa"}/>
                ))}
              </Pie>
              <Tooltip formatter={v=>`${v}%`} contentStyle={{ background:D.surface, border:`1px solid ${D.border}`, borderRadius:7, fontSize:11 }}/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SecTitle ts={ts.orders}>판매처별 매출</SecTitle>
          <ResponsiveContainer width="100%" height={190}>
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

      <Card style={{ marginBottom:12 }}>
        <SecTitle ts={ts.orders}>판매처 상세</SecTitle>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead><tr style={{ borderBottom:`1px solid ${D.border}` }}>
            {["판매처","점유율","주문","매출","배송완료","반품·취소","반품률"].map(h=>(
              <th key={h} style={{ padding:"7px 9px", textAlign:h==="판매처"?"left":"right",
                color:D.textMeta, fontWeight:400 }}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {stats.channelList.map((c,i)=>(
              <tr key={c.name} style={{ borderBottom:`1px solid ${D.border}` }}>
                <td style={{ padding:"7px 9px",fontWeight:600 }}>{c.name}</td>
                <td style={{ textAlign:"right",padding:"7px 9px",color:D.textSub }}>{c.share}%</td>
                <td style={{ textAlign:"right",padding:"7px 9px" }}>{c.orders.toLocaleString()}</td>
                <td style={{ textAlign:"right",padding:"7px 9px",fontWeight:600 }}>{c.revenue>0?fmtWon(c.revenue):"—"}</td>
                <td style={{ textAlign:"right",padding:"7px 9px",color:D.green }}>{c.shipped.toLocaleString()}</td>
                <td style={{ textAlign:"right",padding:"7px 9px",color:D.red }}>{c.returned.toLocaleString()}</td>
                <td style={{ textAlign:"right",padding:"7px 9px",fontWeight:600,
                  color:parseFloat(c.returnRate)>10?D.red:D.textSub }}>{c.returnRate}%</td>
              </tr>
            ))}
            <tr style={{ borderTop:`1px solid ${D.borderMid}` }}>
              <td style={{ padding:"7px 9px",fontWeight:700 }}>합계</td>
              <td style={{ textAlign:"right",padding:"7px 9px",color:D.textSub }}>100%</td>
              <td style={{ textAlign:"right",padding:"7px 9px",fontWeight:600 }}>{stats.totalOrders.toLocaleString()}</td>
              <td style={{ textAlign:"right",padding:"7px 9px",fontWeight:700 }}>{fmtWon(stats.totalRevenue)}</td>
              <td style={{ textAlign:"right",padding:"7px 9px",color:D.green,fontWeight:600 }}>{stats.totalShipped.toLocaleString()}</td>
              <td style={{ textAlign:"right",padding:"7px 9px",color:D.red,fontWeight:600 }}>{stats.totalReturned.toLocaleString()}</td>
              <td style={{ textAlign:"right",padding:"7px 9px",fontWeight:600,
                color:parseFloat(stats.returnRate)>10?D.red:D.textSub }}>{stats.returnRate}%</td>
            </tr>
          </tbody>
        </table>
      </Card>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
        <Card>
          <SecTitle ts={ts.orders}>월별 배송량</SecTitle>
          <ResponsiveContainer width="100%" height={170}>
            <BarChart data={stats.monthlyData}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="month" tick={axTick}/><YAxis tick={axTick}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey="shipped" name="배송" fill="#111" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SecTitle ts={ts.orders}>월별 반품수 · 반품률</SecTitle>
          <ResponsiveContainer width="100%" height={170}>
            <BarChart data={stats.monthlyData}>
              <CartesianGrid strokeDasharray="3 3" stroke={D.border}/>
              <XAxis dataKey="month" tick={axTick}/>
              <YAxis yAxisId="l" tick={axTick}/>
              <YAxis yAxisId="r" orientation="right" unit="%" tick={axTick}/>
              <Tooltip content={<Tip/>}/>
              <Bar yAxisId="l" dataKey="returned" name="반품수" fill={D.red} radius={[3,3,0,0]}/>
              <Line yAxisId="r" type="monotone" dataKey="returnRate" name="반품률(%)"
                stroke={D.amber} strokeWidth={1.5} dot={{fill:D.amber,r:2}}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
          <SecTitle ts={ts.orders}>주간 상품 랭킹{stats.latestWeek?` — ${stats.latestWeek}`:""}</SecTitle>
          <div style={{ display:"flex", borderBottom:`1px solid ${D.border}` }}>
            {[["best","판매 Top 20"],["worst","반품 Top 20"]].map(([k,l])=>(
              <button key={k} onClick={()=>setWeekTab(k)}
                style={{ background:"transparent",border:"none",
                  borderBottom:weekTab===k?`2px solid ${D.black}`:"2px solid transparent",
                  color:weekTab===k?D.black:D.textSub,padding:"6px 14px",
                  fontWeight:weekTab===k?600:400,fontSize:12,cursor:"pointer",marginBottom:-1 }}>{l}</button>
            ))}
          </div>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
          <RankTable
            data={weekTab==="best"?stats.weekBest:stats.weekWorst}
            cols={weekTab==="best"?[
              {key:"name",label:"상품명/옵션",maxW:190},
              {key:"qty",label:"수량",right:true,bold:true,fmt:v=>v.toLocaleString()},
              {key:"returnRate",label:"반품률",right:true,color:D.textMeta,fmt:v=>v+"%"},
            ]:[
              {key:"name",label:"상품명/옵션",maxW:190},
              {key:"returned",label:"반품",right:true,bold:true,color:D.red,fmt:v=>v.toLocaleString()},
              {key:"returnRate",label:"반품률",right:true,color:D.red,fmt:v=>v+"%"},
              {key:"qty",label:"수량",right:true,color:D.textSub,fmt:v=>v.toLocaleString()},
            ]}
          />
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={(weekTab==="best"?stats.weekBest:stats.weekWorst).slice(0,12)} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" stroke={D.border} horizontal={false}/>
              <XAxis type="number" tick={axTick}/><YAxis type="category" dataKey="name" width={130} tick={{...axTick,fontSize:9}}/>
              <Tooltip content={<Tip/>}/>
              <Bar dataKey={weekTab==="best"?"qty":"returned"} name={weekTab==="best"?"수량":"반품"} radius={[0,3,3,0]}>
                {(weekTab==="best"?stats.weekBest:stats.weekWorst).slice(0,12).map((_,i)=>(
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
  return (
    <div style={{ padding:"20px 24px", maxWidth:1400, margin:"0 auto" }}>
      <div style={{ marginBottom:16 }}>
        <div style={{ color:D.black, fontWeight:600, fontSize:15, marginBottom:4 }}>
          물류 플로우
        </div>
        <div style={{ color:D.textMeta, fontSize:12 }}>
          상품별 입고 → 판매처 → 배송 흐름 · 최대 20개 상품
          <UpdatedAt ts={ts.orders||ts.stock}/>
        </div>
      </div>

      {/* 범례 */}
      <div style={{ display:"flex", gap:12, marginBottom:16, flexWrap:"wrap" }}>
        {[
          {label:"배송완료", color:D.green},
          {label:"배송중",   color:D.blue},
          {label:"반품",     color:D.red},
          {label:"취소",     color:D.red},
        ].map(({label,color})=>(
          <div key={label} style={{ display:"flex", alignItems:"center", gap:5, fontSize:11 }}>
            <div style={{ width:8, height:8, borderRadius:"50%", background:color }}/>
            <span style={{ color:D.textSub }}>{label}</span>
          </div>
        ))}
      </div>

      <Card style={{ overflowX:"auto" }}>
        <ProductSankey stockRows={stocks} orderRows={orders}/>
      </Card>

      {/* 상품별 요약 테이블 */}
      {orders.length>0 && (
        <Card style={{ marginTop:12 }}>
          <SecTitle ts={ts.orders}>상품별 흐름 요약</SecTitle>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
              <thead><tr style={{ borderBottom:`1px solid ${D.border}` }}>
                {["#","상품명/옵션","입고","주문","배송완료","배송중","반품","취소","반품률"].map(h=>(
                  <th key={h} style={{ padding:"7px 9px", textAlign:h==="상품명/옵션"?"left":"right",
                    color:D.textMeta, fontWeight:400, whiteSpace:"nowrap" }}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {(() => {
                  const prodMap={};
                  stocks.forEach(r=>{
                    const k=[r.product_name,r.option_name].filter(Boolean).join(" / ");
                    if(!prodMap[k]) prodMap[k]={name:k,stock:0,ordered:0,done:0,inDel:0,returned:0,cancelled:0};
                    prodMap[k].stock+=(r.qty||0);
                  });
                  orders.forEach(r=>{
                    const k=[r.product_name,r.option_name].filter(Boolean).join(" / ");
                    if(!prodMap[k]) prodMap[k]={name:k,stock:0,ordered:0,done:0,inDel:0,returned:0,cancelled:0};
                    prodMap[k].ordered++;
                    if(r.status==="배송완료") prodMap[k].done++;
                    if(r.status==="배송중")   prodMap[k].inDel++;
                    if(r.status==="반품")     prodMap[k].returned++;
                    if(r.status==="취소")     prodMap[k].cancelled++;
                  });
                  return Object.values(prodMap)
                    .sort((a,b)=>b.ordered-a.ordered).slice(0,30)
                    .map((p,i)=>{
                      const rr=p.ordered>0?((p.returned+p.cancelled)/p.ordered*100).toFixed(1):"0.0";
                      return (
                        <tr key={p.name} style={{ borderBottom:`1px solid ${D.border}` }}>
                          <td style={{ padding:"6px 9px",color:D.textMeta,textAlign:"right" }}>{i+1}</td>
                          <td style={{ padding:"6px 9px",fontWeight:i<3?600:400,maxWidth:200,
                            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{p.name}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",color:D.blue }}>{p.stock.toLocaleString()}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",fontWeight:600 }}>{p.ordered.toLocaleString()}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",color:D.green }}>{p.done.toLocaleString()}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",color:D.textSub }}>{p.inDel.toLocaleString()}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",color:D.red }}>{p.returned.toLocaleString()}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",color:D.red }}>{p.cancelled.toLocaleString()}</td>
                          <td style={{ textAlign:"right",padding:"6px 9px",fontWeight:600,
                            color:parseFloat(rr)>10?D.red:D.textSub }}>{rr}%</td>
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
// DATA INPUT (매출 / 입고 / 카페24)
// ─────────────────────────────────────────────
const CHANNELS_LIST = ["스마트스토어","쿠팡","카페24","29CM","무신사","자사몰","기타"];

function RevenueForm({ onUpdate }) {
  const today=new Date().toISOString().slice(0,10);
  const [date,setDate]=useState(today);
  const [ch,setCh]=useState(CHANNELS_LIST[0]);
  const [amt,setAmt]=useState("");
  const [memo,setMemo]=useState("");
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [history,setHistory]=useState([]);
  const [histTs,setHistTs]=useState(null);

  const loadHistory=useCallback(async()=>{
    const db=await getSupabase();
    const { data }=await db.from("revenues").select("*").order("date",{ascending:false}).limit(30);
    setHistory(data||[]); setHistTs(nowStr());
  },[]);

  const handleSave=async()=>{
    const num=Number(amt.replace(/,/g,""));
    if(!amt||isNaN(num)){setResult({type:"error",msg:"금액을 입력해주세요."});return;}
    setLoading(true);setResult(null);
    const db=await getSupabase();
    const { error }=await db.from("revenues").upsert({date,channel:ch,amount:num,memo},{onConflict:"date,channel"});
    const ts=nowStr();
    if(error) setResult({type:"error",msg:error.message});
    else { setResult({type:"success",msg:`₩${num.toLocaleString()} 저장 완료`,ts}); setAmt("");setMemo(""); onUpdate(ts); if(history.length) loadHistory(); }
    setLoading(false);
  };

  const inp={background:"transparent",border:`1px solid ${D.border}`,borderRadius:6,padding:"7px 10px",fontSize:12,color:D.text,width:"100%",boxSizing:"border-box"};
  return (
    <div style={{ display:"grid", gridTemplateColumns:"320px 1fr", gap:14 }}>
      <Card>
        <div style={{ fontWeight:600, marginBottom:14, fontSize:13 }}>매출 입력</div>
        {[
          {label:"날짜",el:<input type="date" value={date} onChange={e=>setDate(e.target.value)} style={inp}/>},
          {label:"판매처",el:<select value={ch} onChange={e=>setCh(e.target.value)} style={{...inp,background:D.surface}}>{CHANNELS_LIST.map(c=><option key={c}>{c}</option>)}</select>},
          {label:"매출 금액",el:<input type="text" value={amt} placeholder="1500000" onChange={e=>setAmt(e.target.value.replace(/[^0-9,]/g,""))} style={inp}/>},
          {label:"메모",el:<input type="text" value={memo} placeholder="선택사항" onChange={e=>setMemo(e.target.value)} style={inp}/>},
        ].map(({label,el})=>(
          <div key={label} style={{ marginBottom:10 }}>
            <div style={{ color:D.textMeta,fontSize:10,marginBottom:4 }}>{label}</div>{el}
          </div>
        ))}
        {amt&&<div style={{color:D.textSub,fontSize:11,marginBottom:10}}>₩{Number(amt.replace(/,/g,"")||0).toLocaleString()}</div>}
        <Btn onClick={handleSave} disabled={loading} style={{width:"100%"}}>{loading?"저장 중...":"저장"}</Btn>
        {result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}
      </Card>
      <Card>
        <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12 }}>
          <div style={{ display:"flex",alignItems:"baseline",gap:6 }}>
            <span style={{fontWeight:600,fontSize:13}}>최근 입력 내역</span>
            <UpdatedAt ts={histTs}/>
          </div>
          <Btn onClick={loadHistory} variant="ghost" style={{padding:"4px 11px",fontSize:11}}>불러오기</Btn>
        </div>
        {history.length>0?(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{borderBottom:`1px solid ${D.border}`}}>
              {["날짜","판매처","매출","메모"].map(h=><th key={h} style={{padding:"5px 7px",textAlign:"left",color:D.textMeta,fontWeight:400}}>{h}</th>)}
            </tr></thead>
            <tbody>{history.map((r,i)=>(
              <tr key={r.id} style={{borderBottom:`1px solid ${D.border}`}}>
                <td style={{padding:"5px 7px",color:D.textMeta}}>{r.date}</td>
                <td style={{padding:"5px 7px"}}>{r.channel}</td>
                <td style={{padding:"5px 7px",fontWeight:600}}>₩{(r.amount||0).toLocaleString()}</td>
                <td style={{padding:"5px 7px",color:D.textMeta}}>{r.memo||"—"}</td>
              </tr>
            ))}</tbody>
          </table>
        ):<div style={{color:D.textMeta,textAlign:"center",padding:40,fontSize:12}}>불러오기를 눌러주세요</div>}
      </Card>
    </div>
  );
}

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
    const { data }=await db.from("stock_uploads").select("*").gte("upload_date",startDate).lte("upload_date",endDate).order("upload_date");
    setExisting(data||[]); setStep(1); setLoading(false);
  };
  const handleFile=useCallback(file=>{
    if(!file) return;
    setFileName(file.name); setResult(null);
    Papa.parse(file,{header:true,skipEmptyLines:true,
      complete:({data})=>{
        try {
          const f=detectFields(Object.keys(data[0]||{}));
          const rows=data.filter(r=>f.product&&r[f.product]).map(r=>({
            product_name:String(r[f.product]||"").trim(),
            option_name:String(r[f.option]||"").trim(),
            qty:toNum(r[f.qty]),
            memo:String(r[f.memo]||"").trim(),
          }));
          setPreview(rows); setStep(2);
        } catch(e){setResult({type:"error",msg:e.message});}
      },
      error:e=>setResult({type:"error",msg:e.message}),
    });
  },[]);
  const handleUpload=async()=>{
    if(!preview?.length||!dateValid) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    const { error:delErr }=await db.from("stock_uploads").delete().gte("upload_date",startDate).lte("upload_date",endDate);
    if(delErr){setResult({type:"error",msg:"삭제 실패: "+delErr.message});setLoading(false);return;}
    const rows=preview.map(r=>({...r,upload_date:startDate}));
    for(let i=0;i<rows.length;i+=500){
      const { error }=await db.from("stock_uploads").insert(rows.slice(i,i+500));
      if(error){setResult({type:"error",msg:"삽입 실패: "+error.message});setLoading(false);return;}
    }
    const ts=nowStr();
    await db.from("upload_logs").insert({upload_type:"stock",file_name:fileName,row_count:preview.length,inserted:preview.length,deleted:existing?.length||0,date_start:startDate,date_end:endDate});
    setStep(3); setResult({type:"success",msg:`기존 ${existing?.length||0}건 삭제 → 새 ${preview.length}건 등록`,ts});
    onUpdate(ts); setLoading(false);
  };
  const reset=()=>{setStep(0);setPreview(null);setExisting(null);setFileName("");setResult(null);};
  return (
    <div>
      <Steps current={step} steps={["기간 선택","파일 업로드","미리보기 확인","완료"]}/>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<><div style={{fontWeight:600,marginBottom:12,fontSize:13}}>입고 기간 선택</div><DateRange start={startDate} end={endDate} onStart={setStartDate} onEnd={setEndDate}/><div style={{color:D.red,fontSize:10,marginBottom:12}}>⚠ 확정 시 해당 기간 DB 데이터 전체 교체</div><Btn onClick={confirmDate} disabled={!dateValid||loading} style={{width:"100%"}}>{loading?"조회 중...":"기간 확정"}</Btn></>}
          {step===1&&<><div style={{fontWeight:600,marginBottom:12,fontSize:13}}>파일 업로드</div><StatRow items={[{label:"삭제 예정",value:`${existing?.length||0}건`,color:D.red}]}/><DropZone onFile={handleFile} fileName={fileName} label="입고 CSV 업로드"/><button onClick={()=>{setStep(0);setExisting(null);}} style={{width:"100%",background:"transparent",border:"none",color:D.textMeta,fontSize:11,cursor:"pointer",marginTop:8,padding:"5px"}}>← 기간 다시 선택</button></>}
          {step===2&&<><div style={{fontWeight:600,marginBottom:12,fontSize:13}}>미리보기 확인</div><StatRow items={[{label:"삭제 예정",value:`${existing?.length||0}건`,color:D.red},{label:"새 등록",value:`${preview?.length||0}건`,color:D.green}]}/><div style={{color:D.amber,fontSize:10,marginBottom:12}}>기존 {existing?.length||0}건 삭제 후 새 {preview?.length||0}건으로 교체</div><Btn onClick={handleUpload} disabled={loading} variant="danger" style={{width:"100%",marginBottom:7}}>{loading?"처리 중...":"확정 교체"}</Btn><button onClick={()=>setStep(1)} style={{width:"100%",background:"transparent",border:"none",color:D.textMeta,fontSize:11,cursor:"pointer",padding:"5px"}}>← 파일 다시 선택</button>{result?.type==="error"&&<Alert type="error" msg={result.msg}/>}</>}
          {step===3&&<div style={{textAlign:"center"}}><div style={{fontSize:36,marginBottom:8}}>✓</div><div style={{color:D.green,fontWeight:600,marginBottom:10}}>교체 완료</div>{result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}<Btn onClick={reset} variant="ghost" style={{width:"100%",marginTop:12}}>새 업로드</Btn></div>}
          {result?.type==="error"&&step!==2&&<Alert type="error" msg={result.msg}/>}
        </Card>
        <Card>
          <div style={{fontWeight:500,fontSize:12,marginBottom:12}}>{step<2?`기존 DB — ${startDate}~${endDate}`:`새 파일 — ${fileName}`}</div>
          {step===0&&<div style={{color:D.textMeta,textAlign:"center",padding:60,fontSize:12}}>기간 선택 후 기존 데이터 표시</div>}
          {step>=1&&step<2&&(existing?.length?<PreviewTable rows={existing} cols={[{key:"upload_date",label:"업로드일",color:D.textMeta},{key:"product_name",label:"상품명",maxW:150},{key:"option_name",label:"옵션",color:D.textMeta},{key:"qty",label:"수량",bold:true},{key:"memo",label:"메모",color:D.textMeta}]}/>:<div style={{color:D.green,textAlign:"center",padding:60,fontSize:12}}>해당 기간 기존 데이터 없음</div>)}
          {step>=2&&preview&&<PreviewTable rows={preview} cols={[{key:"product_name",label:"상품명",maxW:180},{key:"option_name",label:"옵션",color:D.textMeta},{key:"qty",label:"수량",bold:true},{key:"memo",label:"메모",color:D.textMeta}]}/>}
        </Card>
      </div>
    </div>
  );
}

function OrderUploader({ onUpdate }) {
  const today=new Date().toISOString().slice(0,10);
  const [startDate,setStartDate]=useState(today);
  const [endDate,setEndDate]=useState(today);
  const [step,setStep]=useState(0);
  const [fileName,setFileName]=useState("");
  const [allRows,setAllRows]=useState([]);
  const [inRange,setInRange]=useState([]);
  const [outRows,setOutRows]=useState([]);
  const [dupInfo,setDupInfo]=useState(null);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const dateValid=startDate&&endDate&&startDate<=endDate;

  const handleFile=useCallback(async file=>{
    if(!file||!dateValid) return;
    setFileName(file.name); setResult(null);
    Papa.parse(file,{header:true,skipEmptyLines:true,
      complete:async({data})=>{
        try {
          const f=detectFields(Object.keys(data[0]||{}));
          if(!f.orderId) throw new Error("관리번호 컬럼을 찾을 수 없습니다");
          const parsed=data.filter(r=>r[f.orderId]).map(r=>({
            order_id:String(r[f.orderId]).trim(),
            order_date:toDate(r[f.date]),
            channel:String(r[f.channel]||"미분류").trim(),
            product_name:String(r[f.product]||"").trim(),
            option_name:String(r[f.option]||"").trim(),
            qty:toNum(r[f.qty])||1,
            status:normStatus(r[f.status]),
            raw_status:String(r[f.status]||"").trim(),
          }));
          setAllRows(parsed);
          const inR=parsed.filter(r=>r.order_date&&r.order_date>=startDate&&r.order_date<=endDate);
          const outR=parsed.filter(r=>!r.order_date||r.order_date<startDate||r.order_date>endDate);
          setInRange(inR); setOutRows(outR);
          if(inR.length>0){
            const db=await getSupabase();
            const { data:existing }=await db.from("orders").select("order_id,status").in("order_id",inR.map(r=>r.order_id).slice(0,1000));
            const existMap=new Map((existing||[]).map(r=>[r.order_id,r.status]));
            setDupInfo({total:inR.length,newCount:inR.filter(r=>!existMap.has(r.order_id)).length,updateCount:inR.filter(r=>existMap.has(r.order_id)&&existMap.get(r.order_id)!==r.status).length,sameCount:inR.filter(r=>existMap.has(r.order_id)&&existMap.get(r.order_id)===r.status).length});
          } else setDupInfo({total:0,newCount:0,updateCount:0,sameCount:0});
          setStep(2);
        } catch(e){setResult({type:"error",msg:e.message});}
      },
      error:e=>setResult({type:"error",msg:e.message}),
    });
  },[dateValid,startDate,endDate]);

  const handleUpload=async()=>{
    if(!inRange.length) return;
    setLoading(true); setResult(null);
    const db=await getSupabase();
    for(let i=0;i<inRange.length;i+=500){
      const { error }=await db.from("orders").upsert(inRange.slice(i,i+500),{onConflict:"order_id"});
      if(error){setResult({type:"error",msg:error.message});setLoading(false);return;}
    }
    const ts=nowStr();
    await db.from("upload_logs").insert({upload_type:"orders",file_name:fileName,row_count:allRows.length,inserted:dupInfo?.newCount||0,updated:dupInfo?.updateCount||0,skipped:outRows.length,date_start:startDate,date_end:endDate});
    setStep(3); setResult({type:"success",msg:`신규 ${dupInfo?.newCount}건 추가 / ${dupInfo?.updateCount}건 업데이트 / 기간 외 ${outRows.length}건 제외`,ts});
    onUpdate(ts); setLoading(false);
  };

  const reset=()=>{setStep(0);setAllRows([]);setInRange([]);setOutRows([]);setDupInfo(null);setFileName("");setResult(null);};
  const outIdx=new Set(allRows.reduce((acc,r,i)=>{ if(!r.order_date||r.order_date<startDate||r.order_date>endDate) acc.push(i); return acc; },[]));

  return (
    <div>
      <Steps current={step} steps={["기간 선택","파일 업로드","미리보기 확인","완료"]}/>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:14}}>
        <Card>
          {step===0&&<><div style={{fontWeight:600,marginBottom:12,fontSize:13}}>주문 기간 선택</div><DateRange start={startDate} end={endDate} onStart={setStartDate} onEnd={setEndDate}/><div style={{color:D.blue,fontSize:10,marginBottom:12,lineHeight:1.7}}>관리번호 기준 처리 · 신규→추가 / 기존→상태업데이트<br/>기간 밖 데이터 → 자동 제외</div>{dateValid&&<DropZone onFile={handleFile} fileName={fileName} label="카페24 주문 CSV"/>}{result?.type==="error"&&<Alert type="error" msg={result.msg}/>}</>}
          {step===2&&<><div style={{fontWeight:600,marginBottom:12,fontSize:13}}>미리보기 확인</div>{dupInfo&&<StatRow items={[{label:"기간 내",value:dupInfo.total},{label:"신규",value:dupInfo.newCount,color:D.green},{label:"업데이트",value:dupInfo.updateCount,color:D.amber},{label:"변동없음",value:dupInfo.sameCount}]}/>}{outRows.length>0&&<Alert type="warn" msg={`기간 밖 ${outRows.length}건 제외`}/>}<div style={{display:"flex",flexDirection:"column",gap:7,marginTop:12}}><Btn onClick={handleUpload} disabled={loading||!inRange.length} style={{width:"100%"}}>{loading?"처리 중...":`확정 업로드 (${dupInfo?.newCount||0}추가/${dupInfo?.updateCount||0}업데이트)`}</Btn><button onClick={reset} style={{width:"100%",background:"transparent",border:"none",color:D.textMeta,fontSize:11,cursor:"pointer",padding:"5px"}}>← 처음으로</button></div>{result?.type==="error"&&<Alert type="error" msg={result.msg}/>}</>}
          {step===3&&<div style={{textAlign:"center"}}><div style={{fontSize:36,marginBottom:8}}>✓</div><div style={{color:D.green,fontWeight:600,marginBottom:10}}>업로드 완료</div>{result&&<Alert type={result.type} msg={result.msg} ts={result.ts}/>}<Btn onClick={reset} variant="ghost" style={{width:"100%",marginTop:12}}>새 업로드</Btn></div>}
        </Card>
        <Card>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <span style={{fontWeight:500,fontSize:12}}>파일 미리보기</span>
            {allRows.length>0&&<div style={{display:"flex",gap:7}}>
              <span style={{background:`${D.green}12`,color:D.green,fontSize:10,padding:"2px 9px",borderRadius:20}}>기간 내 {inRange.length}건</span>
              {outRows.length>0&&<span style={{background:`${D.red}12`,color:D.red,fontSize:10,padding:"2px 9px",borderRadius:20}}>기간 외 {outRows.length}건</span>}
            </div>}
          </div>
          {allRows.length>0?<PreviewTable rows={allRows} outIdx={outIdx} cols={[{key:"order_id",label:"관리번호",color:D.textMeta,maxW:90},{key:"order_date",label:"주문일",color:D.textMeta},{key:"channel",label:"판매처",bold:true},{key:"product_name",label:"상품명",maxW:140},{key:"option_name",label:"옵션",color:D.textMeta},{key:"qty",label:"수량",bold:true},{key:"status",label:"상태",fmt:v=><span style={{color:v==="반품"||v==="취소"?D.red:v==="배송완료"?D.green:D.text,fontWeight:500}}>{v}</span>}]}/>:<div style={{color:D.textMeta,textAlign:"center",padding:80,fontSize:12}}>기간 선택 후 CSV 파일을 업로드하면 미리보기가 표시됩니다</div>}
        </Card>
      </div>
    </div>
  );
}

function DataInput({ onUpdate }) {
  const [tab,setTab]=useState("revenue");
  return (
    <div style={{padding:"20px 24px",maxWidth:1400,margin:"0 auto"}}>
      <TabBar tabs={[{key:"revenue",label:"매출 입력"},{key:"stock",label:"입고 CSV"},{key:"orders",label:"카페24 주문 CSV"}]} active={tab} onChange={setTab}/>
      {tab==="revenue"&&<RevenueForm onUpdate={ts=>onUpdate("revenue",ts)}/>}
      {tab==="stock"&&<StockUploader onUpdate={ts=>onUpdate("stock",ts)}/>}
      {tab==="orders"&&<OrderUploader onUpdate={ts=>onUpdate("orders",ts)}/>}
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

  const updateTs=useCallback((key,val)=>setTs(prev=>({...prev,[key]:val})),[]);

  useEffect(()=>{
    (async()=>{
      const db=await getSupabase();
      const [o,s,r]=await Promise.all([
        db.from("orders").select("*").order("order_date",{ascending:false}),
        db.from("stock_uploads").select("*"),
        db.from("revenues").select("*").order("date",{ascending:false}),
      ]);
      setOrders(o.data||[]);
      setStocks(s.data||[]);
      setRevenues(r.data||[]);
    })();
  },[]);

  const nav=[
    {key:"dashboard", label:"대시보드",    icon:"▦"},
    {key:"flow",      label:"물류 플로우",  icon:"⟶"},
    {key:"input",     label:"데이터 입력",  icon:"↑"},
  ];

  return (
    <div style={{ display:"flex", minHeight:"100vh", background:D.bg,
      fontFamily:"'Pretendard','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      color:D.text, fontSize:14 }}>

      {/* sidebar */}
      <div style={{ width:172, background:D.surface, borderRight:`1px solid ${D.border}`,
        padding:"18px 10px", display:"flex", flexDirection:"column", flexShrink:0 }}>
        <div style={{ color:D.black, fontWeight:700, fontSize:13, marginBottom:2, paddingLeft:4 }}>COMMERCE</div>
        <div style={{ color:D.textMeta, fontSize:9, letterSpacing:"0.1em", marginBottom:22, paddingLeft:4 }}>ANALYTICS</div>
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
        {page==="dashboard" && <Dashboard orders={orders} stocks={stocks} revenues={revenues} ts={ts}/>}
        {page==="flow"      && <LogisticsFlow orders={orders} stocks={stocks} ts={ts}/>}
        {page==="input"     && <DataInput onUpdate={updateTs}/>}
      </div>
    </div>
  );
}
