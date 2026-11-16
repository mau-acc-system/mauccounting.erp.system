import { validateLicense, readExcel, getSheet } from "../excel.js";

export const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
export const num = v => { const n=Number(String(v ?? 0).replace(/[,₱$ ]/g,"")); return Number.isFinite(n)?n:0; };
export const money = (v,cur="PHP") => new Intl.NumberFormat("en-PH",{style:"currency",currency:cur==="PHP"?"PHP":cur,minimumFractionDigits:2}).format(num(v));
export const dateVal = v => { if(!v) return ""; if(v instanceof Date) return v.toISOString().slice(0,10); const s=String(v); if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10); const d=new Date(v); return isNaN(d)?s.slice(0,10):d.toISOString().slice(0,10); };
export const rows = name => { try { const x=getSheet(name); return Array.isArray(x)?x:[]; } catch { return []; } };
export const byId = (arr,key) => Object.fromEntries(arr.map(x=>[String(x[key]??""),x]));
export const active = v => v===true || String(v).toLowerCase()==="true" || String(v)==="1" || String(v).toLowerCase()==="active";
export async function init(title, build) {
  try {
    const ok=await validateLicense(); if(ok===false) throw new Error("License validation failed.");
    await readExcel();
    const company=rows("CompanySetup")[0]||{};
    const cur=company.Currency||"PHP";
    document.getElementById("company").textContent=company.CompanyName||"ACC System";
    document.getElementById("reportTitle").textContent=title;
    document.getElementById("currency").textContent=cur;
    const data=build({rows,cur,company});
    render(data,cur);
  } catch(e) { document.getElementById("reportBody").innerHTML=`<div class="error">Unable to load report: ${esc(e.message||e)}<br><small>Check license, Excel database connection, and required worksheet names.</small></div>`; }
}
function render(data,cur) {
 const stats=data.stats||[];
 document.getElementById("stats").innerHTML=stats.map(s=>`<div class="stat"><span>${esc(s.label)}</span><b>${s.money?money(s.value,cur):esc(s.value)}</b></div>`).join("");
 const cols=data.columns||[];
 document.getElementById("reportBody").innerHTML=`<div class="table-wrap"><table><thead><tr>${cols.map(c=>`<th>${esc(c.label)}</th>`).join("")}</tr></thead><tbody>${(data.rows||[]).map(r=>`<tr>${cols.map(c=>`<td class="${c.right?'right':''}">${c.money?money(r[c.key],cur):esc(r[c.key])}</td>`).join("")}</tr>`).join("")||`<tr><td colspan="${cols.length||1}" class="empty">No records found for the selected filters.</td></tr>`}</tbody></table></div>`;
 document.getElementById("printBtn").onclick=()=>window.print();
 document.getElementById("csvBtn").onclick=()=>{const lines=[cols.map(c=>c.label),...(data.rows||[]).map(r=>cols.map(c=>r[c.key]??""))];const csv=lines.map(a=>a.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download=(data.filename||"acc-report")+".csv";a.click();URL.revokeObjectURL(a.href);};
}
