const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto').webcrypto;
const ctx=vm.createContext({console,crypto,TextEncoder,Uint8Array,Blob,Intl,setTimeout,document:{addEventListener(){}}});
vm.runInContext(fs.readFileSync('app.js','utf8')+'\n'+fs.readFileSync('insurance.js','utf8'),ctx);
(async()=>{
const result=await vm.runInContext(`(async()=>{
 const s={code:'TEST00001',name:'Học sinh kiểm thử',className:'6A1',hasFeeBreakdown:true,dueItems:[{id:'h',catalogId:'h',category:'insurance',name:'BHYT',amount:631800,paymentCode:'TEST00001YT'},{id:'b',catalogId:'b',category:'mandatory',name:'BHTT',amount:150000,paymentCode:'TEST00001TT'}]};
 const config={bin:'970418',accountNumber:'0000000000',accountName:'TAI KHOAN KIEM THU'};
 const groups=insuranceChoices(s,[],'catalog:h','catalog:b'),h=await insurancePlan(s,groups.options[0],config,'2026-2027'),both=await insurancePlan(s,groups.options[1],config,'2026-2027');
 setNoticeBundleCache({items:[h,both]});
 const tx=(id,code,amount,date='2026-10-05')=>({id,ref:id,content:'THANH TOAN '+code,amount,date,bankStatus:'Thành công'});
 const match=arr=>reconcileTransactions([s],arr);
 const combo=match([tx('1',both.remark,781800)]);
 const legacy=match([{...tx('1','',631800),reportedPaymentCode:'TEST00001YT'}]);
 const hp=insuranceChoices(s,legacy,'catalog:h','catalog:b');
 const body=match([{...tx('1','',150000),reportedPaymentCode:'TEST00001TT'}]);
 const bp=insuranceChoices(s,body,'catalog:h','catalog:b');
 const same=await insurancePlan(s,groups.options[1],config,'2026-2027');
 const changed=await insurancePlan(s,groups.options[1],config,'2027-2028');
 const long=await insurancePlan({...s,code:'TEST00001XXXXXXXXXXXXXXXXXXXXX'},groups.options[1],config,'2026-2027');
 const results={initial:groups.options.map(o=>o.amount),combo:combo.map(t=>[t.paymentStatus,t.matchedDueItemIds]),totals:totals([s],combo),receipts:receiptCandidates([s],combo).map(x=>x.item.amount),healthPaid:hp.options.map(o=>o.label),bodyPaid:bp.options.map(o=>o.label),complete:insuranceChoices(s,combo,'catalog:h','catalog:b').complete,missing:insuranceChoices({...s,dueItems:[s.dueItems[0]]},[],'catalog:h','catalog:b').error,same:h.remark!==both.remark&&same.remark===both.remark,changed:changed.remark!==both.remark,long:long.remark!==both.remark,wrong:match([tx('1',both.remark,631800)])[0].paymentStatus,unknown:match([tx('1','IB00000000000000000000',781800)])[0].paymentStatus,duplicate:match([tx('1',h.remark,631800,'2026-10-04'),tx('2',both.remark,781800) ]).map(t=>t.paymentStatus),reverse:match([tx('1',both.remark,781800,'2026-10-04'),tx('2',h.remark,631800)]).map(t=>t.paymentStatus),failed:match([{...tx('1',both.remark,781800),bankStatus:'Thất bại'}])[0].paymentStatus,substring:match([tx('1',both.remark+'ABC',781800)])[0].paymentStatus,ambiguous:match([tx('1',h.remark+' '+both.remark,781800)])[0].paymentStatus};
 const modified={...s,dueItems:s.dueItems.map(x=>({...x,amount:x.amount+1}))};results.changedDue=reconcileTransactions([modified],[tx('1',both.remark,781800)])[0].paymentStatus;
 results.removed=reconcileTransactions([{...s,dueItems:[s.dueItems[0]]}],[tx('1',both.remark,781800)])[0].paymentStatus;
 return JSON.stringify(results);
})()`,ctx);
const r=JSON.parse(result);
assert.deepEqual(r.initial,[631800,781800]);assert.deepEqual(r.combo,[['valid',['b','h']]]);assert.equal(r.totals.paid,781800);assert.equal(r.totals.paidItems,2);assert.deepEqual(r.receipts,[150000,631800]);assert.deepEqual(r.healthPaid,['Chỉ nộp BHTT']);assert.deepEqual(r.bodyPaid,['Chỉ nộp BHYT']);assert.equal(r.complete,true);assert.ok(r.missing);assert.ok(r.same&&r.changed&&r.long);assert.equal(r.wrong,'amount_mismatch');assert.equal(r.unknown,'unmatched');assert.deepEqual(r.duplicate,['valid','duplicate']);assert.deepEqual(r.reverse,['valid','duplicate']);assert.equal(r.failed,'bank_not_successful');assert.notEqual(r.substring,'valid');assert.equal(r.ambiguous,'ambiguous');assert.equal(r.changedDue,'amount_mismatch');assert.equal(r.removed,'no_due');console.log('PASS: 21 insurance reconciliation and allocation checks');
})().catch(e=>{console.error(e);process.exitCode=1;});
