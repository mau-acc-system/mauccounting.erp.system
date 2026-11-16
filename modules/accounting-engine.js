// Mauccounting ERP integrated posting engine.
// Uses the existing .dat-backed excel.js compatibility API.
import {getSheet, addRecord, updateRecord} from "../excel.js";

const num = v => Number(v || 0);
const now = () => new Date().toISOString();
const id = (prefix, rows, key) => prefix + String(Math.max(0,...rows.map(r=>parseInt(String(r[key]||"").replace(/\D/g,""))||0))+1).padStart(6,"0");
const rows = name => getSheet(name);
const find = (name,key,value) => rows(name).find(r=>String(r[key])===String(value));
const active = x => !["false","0",false,0].includes(x);
function accountByCode(code) {
  const a=rows("ChartOfAccounts").find(x=>String(x.AccountCode)===String(code));
  if(!a) throw Error(`Required account code ${code} is missing from Chart of Accounts.`);
  return a;
}
function accountByType(type, hints=[]) {
  const all=rows("ChartOfAccounts").filter(x=>active(x.IsActive));
  const normalizeType = v => {
    const x=String(v||"").toLowerCase();
    return ({assets:"asset",liabilities:"liability",expenses:"expense",revenues:"revenue",equities:"equity"})[x]||x;
  };
  const wanted = normalizeType(type);
  const a=all.find(x=>normalizeType(x.AccountType)===wanted &&
    hints.some(h=>(String(x.AccountName||"")+" "+String(x.AccountSubtype||"")).toLowerCase().includes(h)));
  if(!a) throw Error(`No ${type} account found for ${hints.join("/")}. Configure Chart of Accounts first.`);
  return a;
}
function addJournal(sourceType, sourceId, date, description, lines) {
  const existing=rows("JournalEntries").find(j=>j.SourceType===sourceType && String(j.SourceID)===String(sourceId));
  if(existing) return existing;
  const debit=lines.reduce((s,l)=>s+num(l.Debit),0), credit=lines.reduce((s,l)=>s+num(l.Credit),0);
  if(Math.abs(debit-credit)>0.005) throw Error(`Unbalanced journal: debit ${debit.toFixed(2)} / credit ${credit.toFixed(2)}.`);
  if(debit<=0) throw Error("Journal must contain a positive amount.");
  const js=rows("JournalEntries"), journalId=id("JE",js,"JournalID"), journalNo=id("JE-",js,"JournalNo"), t=now();
  addRecord("JournalEntries",{JournalID:journalId,JournalNo:journalNo,JournalDate:date||t.slice(0,10),ReferenceNo:String(sourceId),Description:description,Status:"Posted",TotalDebit:debit,TotalCredit:credit,CreatedBy:"System",CreatedAt:t,PostedBy:"System",PostedAt:t,UpdatedAt:t,SourceType:sourceType,SourceID:String(sourceId)});
  let lineNo=0;
  for(const l of lines){const a=typeof l.account==="string"?accountByCode(l.account):l.account;lineNo++;
    addRecord("JournalEntryLines",{LineID:id("JEL",rows("JournalEntryLines"),"LineID"),JournalID:journalId,LineNo:lineNo,AccountID:a.AccountID,AccountCode:a.AccountCode,AccountName:a.AccountName,Description:l.description||description,Debit:num(l.Debit),Credit:num(l.Credit),UpdatedAt:t,SourceType:sourceType,SourceID:String(sourceId)});
  }
  return {JournalID:journalId,JournalNo:journalNo};
}
function itemStock(itemId) {
  return rows("InventoryTransactions").filter(x=>String(x.ItemID)===String(itemId))
    .reduce((s,x)=>s+num(x.QuantityIn)-num(x.QuantityOut),0);
}
function costingMethod() {
  const c=rows("CompanySetup")[0]||{};
  return String(c.InventoryCostingMethod||"WA").toUpperCase().replace(/[^A-Z]/g,"");
}
function costForIssue(itemId, quantity) {
  const method=costingMethod();
  const movements=rows("InventoryTransactions").filter(x=>String(x.ItemID)===String(itemId));
  if(method.includes("FIFO")||method.includes("LIFO")) {
    const layers=[];
    for(const t of movements) {
      const qi=num(t.QuantityIn), qo=num(t.QuantityOut), cost=num(t.UnitCost);
      if(qi>0) layers.push({qty:qi,cost,when:String(t.TransactionDate||""),seq:layers.length});
      if(qo>0) {
        let remaining=qo;
        while(remaining>0.000001 && layers.length) {
          const ix=method.includes("LIFO")?layers.length-1:0;
          const take=Math.min(remaining,layers[ix].qty);
          layers[ix].qty-=take;remaining-=take;
          if(layers[ix].qty<=0.000001) layers.splice(ix,1);
        }
      }
    }
    let need=quantity,total=0;
    const work=layers.map(x=>({...x}));
    while(need>0.000001 && work.length) {
      const ix=method.includes("LIFO")?work.length-1:0;
      const take=Math.min(need,work[ix].qty);
      total+=take*work[ix].cost;need-=take;work[ix].qty-=take;
      if(work[ix].qty<=0.000001)work.splice(ix,1);
    }
    if(need>0.000001)throw Error("Insufficient inventory costing layers.");
    return total/quantity;
  }
  return averageCost(itemId);
}
function averageCost(itemId) {
  const tx=rows("InventoryTransactions").filter(x=>String(x.ItemID)===String(itemId));
  let qty=0,value=0;
  for(const t of tx){const qi=num(t.QuantityIn), qo=num(t.QuantityOut), cost=num(t.UnitCost);
    if(qi>0){qty+=qi;value+=qi*cost;}
    if(qo>0){const avg=qty>0?value/qty:cost;qty-=qo;value-=qo*avg;}
  }
  return qty>0?value/qty:0;
}
function stockTx(itemId, type, qtyIn, qtyOut, cost, refType, refId, memo) {
  const all=rows("InventoryTransactions"), t=now();
  addRecord("InventoryTransactions",{TransactionID:id("IT",all,"TransactionID"),TransactionNo:id("STK-",all,"TransactionNo"),TransactionDate:t.slice(0,10),ItemID:itemId,TransactionType:type,QuantityIn:qtyIn,QuantityOut:qtyOut,UnitCost:cost,ReferenceType:refType,ReferenceID:String(refId),Warehouse:"Default",Memo:memo,CreatedAt:t,CreatedBy:"System",UpdatedAt:t});
}
function itemForLine(line) {
  const item=find("Items","ItemID",line.ItemID);
  if(!item) throw Error(`Item ${line.ItemID} not found.`);
  return item;
}
function invoiceLines(invoiceId) {return rows("InvoiceLines").filter(x=>String(x.InvoiceID)===String(invoiceId));}
function billLines(billId) {return rows("BillLines").filter(x=>String(x.BillID)===String(billId));}

function clearDerivedPosting(sourceType, sourceId) {
  const journals=rows("JournalEntries");
  const matches=journals.filter(j=>j.SourceType===sourceType && String(j.SourceID)===String(sourceId));
  const journalIds=new Set(matches.map(j=>String(j.JournalID)));
  const lines=rows("JournalEntryLines");
  for(let i=lines.length-1;i>=0;i--)if(journalIds.has(String(lines[i].JournalID)))lines.splice(i,1);
  for(let i=journals.length-1;i>=0;i--)if(journalIds.has(String(journals[i].JournalID)))journals.splice(i,1);
  const tx=rows("InventoryTransactions");
  for(let i=tx.length-1;i>=0;i--)if(String(tx[i].ReferenceID)===String(sourceId)&&String(tx[i].ReferenceType)===sourceType)tx.splice(i,1);
}

function postInvoice(rec) {
  if(rec.Status==="Draft"||rec.Status==="Cancelled") return;
  const lines=invoiceLines(rec.InvoiceID);
  if(!lines.length) return;
  clearDerivedPosting("SalesInvoice",rec.InvoiceID); // Header is saved first; posting occurs when the first detail line is saved.
  const total=num(rec.TotalAmount)||lines.reduce((s,l)=>s+num(l.LineTotal),0)+num(rec.TaxAmount);
  const paid=Math.min(total,Math.max(0,num(rec.AmountPaid))), ar=total-paid;
  const cash=accountByType("Asset",["cash","bank"]);
  const receivable=accountByType("Asset",["receivable","accounts receivable"]);
  const revenue=accountByType("Revenue",["sales","service","revenue"]);
  const taxPayable=rows("ChartOfAccounts").find(a=>String(a.AccountType).toLowerCase()==="liability" && /tax payable|output tax|vat payable/i.test(a.AccountName||""));
  const jl=[];
  if(paid>0) jl.push({account:cash,Debit:paid});
  if(ar>0) jl.push({account:receivable,Debit:ar});
  jl.push({account:revenue,Credit:total-num(rec.TaxAmount)});
  if(num(rec.TaxAmount)>0){if(!taxPayable)throw Error("Tax amount entered but no Output Tax/VAT Payable liability account exists.");jl.push({account:taxPayable,Credit:num(rec.TaxAmount)});}
  let cogs=0;
  for(const l of lines){const item=itemForLine(l);if(String(item.ItemType||"").toLowerCase()==="service"||String(item.ItemType||"").toLowerCase()==="non-inventory")continue;
    const q=num(l.Quantity);const stock=itemStock(item.ItemID);if(q>stock)throw Error(`Insufficient stock for ${item.ItemName}. Available ${stock}, required ${q}.`);
    const cost=costForIssue(item.ItemID,q)||num(item.PurchaseCost);cogs+=q*cost;stockTx(item.ItemID,"Sales",0,q,cost,"SalesInvoice",rec.InvoiceID,`Sales ${rec.InvoiceNo}`);
  }
  if(cogs>0){jl.push({account:accountByType("Expense",["cost of goods sold","cogs"]),Debit:cogs});jl.push({account:accountByType("Asset",["inventory","stock"]),Credit:cogs});}
  addJournal("SalesInvoice",rec.InvoiceID,rec.InvoiceDate,`Sales invoice ${rec.InvoiceNo}`,jl);
  updateRecord("Invoices","InvoiceID",rec.InvoiceID,{Subtotal:total-num(rec.TaxAmount),TotalAmount:total,AmountPaid:paid,BalanceDue:ar,Status:ar<=0.005?"Paid":paid>0?"Partially Paid":"Posted"});
}
function postBill(rec) {
  if(rec.Status==="Draft"||rec.Status==="Cancelled")return;
  const lines=billLines(rec.BillID);if(!lines.length)return;
  clearDerivedPosting("PurchaseBill",rec.BillID); // Header is saved first; posting occurs when detail is saved.
  const total=num(rec.TotalAmount)||lines.reduce((s,l)=>s+num(l.LineTotal),0)+num(rec.TaxAmount);
  const paid=Math.min(total,Math.max(0,num(rec.AmountPaid))), ap=total-paid;
  const inventory=accountByType("Asset",["inventory","stock"]);
  const payable=accountByType("Liability",["payable","accounts payable"]);
  const cash=accountByType("Asset",["cash","bank"]);
  let inventoryValue=0;
  for(const l of lines){const item=itemForLine(l);const q=num(l.Quantity),cost=num(l.UnitCost)||num(item.PurchaseCost);
    if(String(item.ItemType||"").toLowerCase()!=="service"&&String(item.ItemType||"").toLowerCase()!=="non-inventory"){stockTx(item.ItemID,"Purchase",q,0,cost,"PurchaseBill",rec.BillID,`Purchase ${rec.BillNo}`);inventoryValue+=q*cost;}
  }
  const expense=total-num(rec.TaxAmount)-inventoryValue;
  const jl=[];if(inventoryValue>0)jl.push({account:inventory,Debit:inventoryValue});
  if(expense>0)jl.push({account:accountByType("Expense",["purchase","supplies","expense"]),Debit:expense});
  if(num(rec.TaxAmount)>0){const inputTax=rows("ChartOfAccounts").find(a=>/input tax|vat input|tax receivable/i.test(a.AccountName||""));if(!inputTax)throw Error("Tax amount entered but no Input Tax/VAT Receivable account exists.");jl.push({account:inputTax,Debit:num(rec.TaxAmount)});}
  if(ap>0)jl.push({account:payable,Credit:ap});if(paid>0)jl.push({account:cash,Credit:paid});
  addJournal("PurchaseBill",rec.BillID,rec.BillDate,`Purchase bill ${rec.BillNo}`,jl);
  updateRecord("Bills","BillID",rec.BillID,{TotalAmount:total,AmountPaid:paid,BalanceDue:ap,Status:ap<=0.005?"Paid":paid>0?"Partially Paid":"Posted"});
}
function postCustomerPayment(rec) {
  if(rows("JournalEntries").some(j=>j.SourceType==="CustomerPayment"&&String(j.SourceID)===String(rec.PaymentID)))return;
  const candidates=rows("Invoices").filter(x=>String(x.CustomerID)===String(rec.CustomerID) && num(x.BalanceDue)>0);
  const invoice=rec.InvoiceID?find("Invoices","InvoiceID",rec.InvoiceID):(candidates.length===1?candidates[0]:null);
  const amount=num(rec.Amount);if(amount<=0)throw Error("Payment amount must be greater than zero.");
  if(!invoice)throw Error(candidates.length>1?"Select the invoice to apply this payment to.":"No uniquely identifiable open invoice. Link an Invoice ID before posting.");
  const balance=num(invoice.BalanceDue);if(amount>balance+0.005)throw Error("Payment exceeds the invoice balance.");
  const cash=rec.BankID?find("Banks","BankID",rec.BankID):null;
  const cashAccount=rec.CashAccountID?find("CashAccounts","CashAccountID",rec.CashAccountID):null;
  const cashCoa=accountByType("Asset",["cash","bank"]);
  const ar=accountByType("Asset",["receivable","accounts receivable"]);
  addJournal("CustomerPayment",rec.PaymentID,rec.PaymentDate,`Customer receipt ${rec.PaymentNo}`,[{account:cashCoa,Debit:amount},{account:ar,Credit:amount}]);
  if(invoice)updateRecord("Invoices","InvoiceID",invoice.InvoiceID,{AmountPaid:num(invoice.AmountPaid)+amount,BalanceDue:Math.max(0,balance-amount),Status:balance-amount<=0.005?"Paid":"Partially Paid"});
}
function postVendorPayment(rec) {
  if(rows("JournalEntries").some(j=>j.SourceType==="VendorPayment"&&String(j.SourceID)===String(rec.PaymentID)))return;
  const candidates=rows("Bills").filter(x=>String(x.VendorID)===String(rec.VendorID) && num(x.BalanceDue)>0);
  const bill=rec.BillID?find("Bills","BillID",rec.BillID):(candidates.length===1?candidates[0]:null);
  const amount=num(rec.Amount);if(amount<=0)throw Error("Payment amount must be greater than zero.");
  if(!bill)throw Error(candidates.length>1?"Select the bill to apply this payment to.":"No uniquely identifiable open bill. Link a Bill ID before posting.");
  const balance=num(bill.BalanceDue);if(amount>balance+0.005)throw Error("Payment exceeds the bill balance.");
  const cash=accountByType("Asset",["cash","bank"]);const ap=accountByType("Liability",["payable","accounts payable"]);
  addJournal("VendorPayment",rec.PaymentID,rec.PaymentDate,`Vendor payment ${rec.PaymentNo}`,[{account:ap,Debit:amount},{account:cash,Credit:amount}]);
  if(bill)updateRecord("Bills","BillID",bill.BillID,{AmountPaid:num(bill.AmountPaid)+amount,BalanceDue:Math.max(0,balance-amount),Status:balance-amount<=0.005?"Paid":"Partially Paid"});
}
function postSalesReturn(rec) {
  if(rows("JournalEntries").some(j=>j.SourceType==="SalesReturn"&&String(j.SourceID)===String(rec.ReturnID)))return;
  const qty=num(rec.Quantity), amount=num(rec.Amount||rec.TotalAmount);
  if(qty<=0||amount<=0)throw Error("Sales return requires positive quantity and amount.");
  const item=itemForLine(rec);
  const cost=costForIssue(item.ItemID,qty);
  const revenue=accountByType("Revenue",["sales","service","revenue"]);
  const ar=accountByType("Asset",["receivable","accounts receivable"]);
  const inventory=accountByType("Asset",["inventory","stock"]);
  const cogs=accountByType("Expense",["cost of goods sold","cogs"]);
  stockTx(item.ItemID,"Sales Return",qty,0,cost,"SalesReturn",rec.ReturnID,"Sales return");
  addJournal("SalesReturn",rec.ReturnID,rec.ReturnDate,`Sales return ${rec.ReturnNo}`,[{account:revenue,Debit:amount},{account:ar,Credit:amount},{account:inventory,Debit:qty*cost},{account:cogs,Credit:qty*cost}]);
}
function postPurchaseReturn(rec) {
  if(rows("JournalEntries").some(j=>j.SourceType==="PurchaseReturn"&&String(j.SourceID)===String(rec.ReturnID)))return;
  const qty=num(rec.Quantity), amount=num(rec.Amount||rec.TotalAmount);
  if(qty<=0||amount<=0)throw Error("Purchase return requires positive quantity and amount.");
  const item=itemForLine(rec), stock=itemStock(item.ItemID);
  if(qty>stock)throw Error(`Cannot return ${qty}; only ${stock} units are on hand.`);
  const cost=costForIssue(item.ItemID,qty);
  const inventory=accountByType("Asset",["inventory","stock"]);
  const ap=accountByType("Liability",["payable","accounts payable"]);
  stockTx(item.ItemID,"Purchase Return",0,qty,cost,"PurchaseReturn",rec.ReturnID,"Purchase return");
  addJournal("PurchaseReturn",rec.ReturnID,rec.ReturnDate,`Purchase return ${rec.ReturnNo}`,[{account:ap,Debit:amount},{account:inventory,Credit:qty*cost}]);
}
function postStockAdjustment(rec) {
  if(rows("JournalEntries").some(j=>j.SourceType==="StockAdjustment"&&String(j.SourceID)===String(rec.TransactionID)))return;
  const item=itemForLine(rec), qi=num(rec.QuantityIn), qo=num(rec.QuantityOut);
  if(qi>0&&qo>0)throw Error("Stock adjustment cannot increase and decrease stock in one line.");
  if(qi<=0&&qo<=0)throw Error("Enter a positive quantity in or out.");
  const cost=num(rec.UnitCost)||averageCost(item.ItemID)||num(item.PurchaseCost);
  if(qo>itemStock(item.ItemID))throw Error("Stock adjustment would create negative inventory.");
  const inv=accountByType("Asset",["inventory","stock"]);
  const offset=accountByType(qi>0?"Revenue":"Expense",qi>0?["inventory gain","adjustment gain","other income"]:["inventory adjustment","shrinkage","stock loss"]);
  if(qi>0){stockTx(item.ItemID,"Adjustment In",qi,0,cost,"StockAdjustment",rec.TransactionID,rec.Memo);addJournal("StockAdjustment",rec.TransactionID,rec.TransactionDate,"Inventory adjustment in",[{account:inv,Debit:qi*cost},{account:offset,Credit:qi*cost}]);}
  else{stockTx(item.ItemID,"Adjustment Out",0,qo,cost,"StockAdjustment",rec.TransactionID,rec.Memo);addJournal("StockAdjustment",rec.TransactionID,rec.TransactionDate,"Inventory adjustment out",[{account:offset,Debit:qo*cost},{account:inv,Credit:qo*cost}]);}
}
export async function postAccountingTransaction(sheet,rec) {
  if(sheet==="Invoices")postInvoice(rec);
  else if(sheet==="Bills")postBill(rec);
  else if(sheet==="InvoiceLines"){const h=find("Invoices","InvoiceID",rec.InvoiceID);if(h)postInvoice(h);}
  else if(sheet==="BillLines"){const h=find("Bills","BillID",rec.BillID);if(h)postBill(h);}
  else if(sheet==="CustomerPayments")postCustomerPayment(rec);
  else if(sheet==="VendorPayments")postVendorPayment(rec);
  else if(sheet==="SalesReturns")postSalesReturn(rec);
  else if(sheet==="PurchaseReturns")postPurchaseReturn(rec);
  else if(sheet==="InventoryTransactions" && /adjustment/i.test(String(rec.TransactionType||"")))postStockAdjustment(rec);
}
export function getInventoryOnHand(itemId){return itemStock(itemId);}
export function getInventoryAverageCost(itemId){return averageCost(itemId);}
