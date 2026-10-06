/* Two alternative payment plans; all data, QR generation and exports stay local. */
let insuranceView={students:[],transactions:[],entries:[],index:0,busy:false,cancel:false,revision:0};
const IB_CODE=/\bIB[A-F0-9]{20}\b/g;

// Issued plans take priority over heuristics. Never infer a different plan from its amount.
function reconcileInsuranceTransaction(t,byCode,claimed){
  if(t.sourceType==='cash')return null;
  const text=[t.reportedPaymentCode,t.reportedStudentCode,t.content,t.ref].filter(Boolean).join(' ').toUpperCase();
  const codes=[...new Set(text.match(IB_CODE)||[])];
  const possibleCodes=text.match(/\bIB[A-F0-9]{8,}\b/g)||[];
  const plans=codes.map(code=>noticeBundleCache.get(slug(code))).filter(b=>b?.kind==='insurance-choice');
  if(!codes.length&& !possibleCodes.length)return null;
  const base={...t,matched:false,matchedDueItemId:'',matchedDueItemIds:[],matchedDueItem:'',studentCode:'',studentName:'',feeCategory:'other',matchedBy:''};
  if(!isSuccessfulBankStatus(t.bankStatus))return {...base,paymentStatus:'bank_not_successful'};
  if(!codes.length)return {...base,paymentStatus:'unmatched'};
  if(codes.length!==1||possibleCodes.some(code=>!codes.includes(code)))return {...base,paymentStatus:'ambiguous'};
  if(plans.length!==1)return {...base,paymentStatus:'unmatched'};
  const plan=plans[0],student=byCode.get(slug(plan.studentCode));
  if(!student)return {...base,paymentStatus:'unmatched'};
  Object.assign(base,{studentCode:student.code,studentName:student.name});
  const reportedStudent=byCode.get(slug(t.reportedStudentCode||''));
  if(reportedStudent&&reportedStudent.code!==student.code)return {...base,paymentStatus:'ambiguous'};
  const allItems=studentDueItems(student),items=plan.allocations.map(a=>allItems.find(x=>x.id===a.itemId));
  if(items.some(x=>!x)||!items.length||new Set(plan.itemIds).size!==items.length)return {...base,paymentStatus:'no_due'};
  if(num(t.amount)!==plan.amount||items.some((x,i)=>x.amount!==plan.allocations[i].amount)||items.reduce((s,x)=>s+x.amount,0)!==plan.amount)return {...base,paymentStatus:'amount_mismatch'};
  const keys=items.map(x=>`${student.code}|${x.id}`);
  if(keys.some(k=>claimed.has(k)))return {...base,paymentStatus:'duplicate'};
  keys.forEach(k=>claimed.add(k));
  return {...base,matched:true,paymentStatus:'valid',matchedBy:'insurance_plan',paymentChannel:'transfer',feeCategory:items.length===1?items[0].category:'other',feeDetail:items.map(x=>x.name).join(' + '),matchedDueItem:items.map(x=>x.name).join(' + '),matchedDueItemId:items[0].id,matchedDueItemIds:items.map(x=>x.id),noticeBundleRemark:plan.remark};
}

function insuranceChoices(student,transactions,healthKey,bodyKey){
  const items=studentDueItems(student),health=items.filter(x=>noticeItemKey(x)===healthKey),body=items.filter(x=>noticeItemKey(x)===bodyKey);
  if(health.length!==1||body.length!==1||health[0].id===body[0].id||health[0].legacy||body[0].legacy)return {student,options:[],error:'Thiếu hoặc trùng khoản đã chọn'};
  const paid=new Set(transactions.filter(t=>t.paymentStatus==='valid'&&t.studentCode===student.code).flatMap(transactionMatchedItemIds));
  const h=health[0],b=body[0],hp=paid.has(h.id),bp=paid.has(b.id);
  let options=[];
  if(!hp&&!bp)options=[{label:'Chỉ nộp BHYT',items:[h]},{label:'BHYT + BHTT',items:[h,b]}];
  else if(!hp)options=[{label:'Chỉ nộp BHYT',items:[h]}];
  else if(!bp)options=[{label:'Chỉ nộp BHTT',items:[b]}];
  return {student,health:h,body:b,healthPaid:hp,bodyPaid:bp,complete:hp&&bp,options:options.map(o=>({...o,amount:o.items.reduce((sum,x)=>sum+x.amount,0)}))};
}
function insuranceCurrentGroups(){
  const classes=new Set(selectedValues($('#ibClasses'))),allClasses=classes.has('all')||!classes.size;
  return insuranceView.students.filter(s=>allClasses||classes.has(s.className||'Chưa xếp lớp')).map(s=>insuranceChoices(s,insuranceView.transactions,$('#ibHealth').value,$('#ibBody').value)).sort((a,b)=>(a.student.className||'').localeCompare(b.student.className||'','vi',{numeric:true})||a.student.name.localeCompare(b.student.name,'vi'));
}
function invalidateInsurancePreview(){
  insuranceView.revision++;insuranceView.entries=[];insuranceView.index=0;
  $('#ibPreviewImage').innerHTML='<div class="empty-inline">Bấm Xem trước để cập nhật thông báo theo dữ liệu hiện tại.</div>';
  $('#ibPrev').disabled=$('#ibNext').disabled=true;
  $('#ibPreviewLabel').textContent='Mẫu A5 dọc · Ảnh PNG rõ nét';
  const groups=insuranceCurrentGroups(),valid=groups.filter(g=>!g.error),complete=valid.filter(g=>g.complete).length;
  $('#ibSummary').textContent=`${valid.length} học sinh đủ dữ liệu · ${complete} đã hoàn thành · ${groups.length-valid.length} thiếu/trùng khoản (không xuất).`;
  if(groups.some(g=>g.error)){
    const details=document.createElement('details'),summary=document.createElement('summary'),list=document.createElement('p');
    summary.textContent='Xem học sinh cần kiểm tra';list.textContent=groups.filter(g=>g.error).map(g=>`${g.student.className||'—'} · ${g.student.name} (${g.student.code})`).join('; ');
    details.append(summary,list);$('#ibSummary').append(details);
  }
}
async function renderInsuranceTool(students,transactions,config){
  insuranceView.students=students;insuranceView.transactions=transactions;
  const saved=(await request('meta','get','insuranceConfig'))||{};
  for(const [id,category,key] of [['ibHealth','insurance','healthKey'],['ibBody','mandatory','bodyKey']]){
    const el=$('#'+id),previous=el.value||saved[key],options=new Map();
    students.forEach(s=>studentDueItems(s).filter(x=>x.category===category).forEach(x=>options.set(noticeItemKey(x),x.name)));
    el.innerHTML='<option value="">Chọn khoản thu…</option>'+[...options].map(([v,label])=>`<option value="${escapeHTML(v)}">${escapeHTML(label)}</option>`).join('');
    if(options.has(previous))el.value=previous;else if(options.size===1)el.value=[...options.keys()][0];
  }
  const classes=[...new Set(students.map(s=>s.className||'Chưa xếp lớp'))].sort((a,b)=>a.localeCompare(b,'vi',{numeric:true})),el=$('#ibClasses'),selected=new Set(selectedValues(el));
  el.innerHTML='<option value="all">Tất cả lớp</option>'+classes.map(c=>`<option value="${escapeHTML(c)}">${escapeHTML(c)}</option>`).join('');
  [...el.options].forEach(o=>o.selected=selected.has(o.value));if(!el.selectedOptions.length)el.options[0].selected=true;
  if(!$('#ibYear').dataset.loaded){for(const [id,key] of [['ibYear','year'],['ibDeadline','deadline'],['ibContact','contact']])if(saved[key]!==undefined)$('#'+id).value=saved[key];$('#ibYear').dataset.loaded='1';}
  $('#ibBin').value=config?.bin||'970418';$('#ibAccount').value=config?.accountNumber||'';$('#ibAccountName').value=config?.accountName||'';
  invalidateInsurancePreview();
}
async function saveInsuranceAccount(){
  const config={key:'qrAccount',bin:$('#ibBin').value.trim(),accountNumber:$('#ibAccount').value.trim(),accountName:qrText($('#ibAccountName').value,25)};
  if(!/^\d{6}$/.test(config.bin)||!/^\d{1,19}$/.test(config.accountNumber)||!config.accountName)throw new Error('Nhập BIN 6 chữ số, số tài khoản 1–19 chữ số và tên chủ tài khoản.');
  await putMany('meta',[config]);invalidateInsurancePreview();toast('Đã lưu tài khoản nhận tiền.');
}
async function insurancePlan(student,option,config,year){
  const allocations=option.items.map(x=>({itemId:x.id,amount:x.amount})).sort((a,b)=>a.itemId.localeCompare(b.itemId));
  const signature=JSON.stringify([student.code,allocations,config.bin,config.accountNumber,year]);
  const digest=new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(signature)));
  const remark='IB'+[...digest].slice(0,10).map(x=>x.toString(16).padStart(2,'0')).join('').toUpperCase();
  return {kind:'insurance-choice',remark,signature,studentCode:student.code,itemIds:allocations.map(x=>x.itemId),allocations,amount:option.amount,year,account:{bin:config.bin,accountNumber:config.accountNumber},createdAt:new Date().toISOString()};
}
async function buildInsuranceEntries(){
  const [students,stored,config,school,bundles]=await Promise.all([all('students'),all('transactions'),request('meta','get','qrAccount'),getReceiptConfig(),getNoticeBundles()]);
  if(!$('#ibHealth').value||!$('#ibBody').value)throw new Error('Chọn đủ khoản BHYT và BHTT đã phân giao.');
  if(!config||!/^\d{6}$/.test(config.bin)||!/^\d{1,19}$/.test(config.accountNumber)||!config.accountName)throw new Error('Mở mục Tài khoản nhận tiền, nhập và lưu tài khoản của trường trước.');
  if($('#ibBin').value.trim()!==config.bin||$('#ibAccount').value.trim()!==config.accountNumber||qrText($('#ibAccountName').value,25)!==config.accountName)throw new Error('Thông tin tài khoản đang sửa chưa được lưu. Bấm Lưu tài khoản trước.');
  setNoticeBundleCache(bundles);insuranceView.students=students;insuranceView.transactions=reconcileTransactions(students,stored);
  const groups=insuranceCurrentGroups().filter(g=>!g.error&&($('#ibStatus').value==='all'||!g.complete));
  if(!groups.length)throw new Error('Không có học sinh phù hợp. Kiểm tra lớp, khoản phân giao và trạng thái đã nộp.');
  const settings={key:'insuranceConfig',healthKey:$('#ibHealth').value,bodyKey:$('#ibBody').value,year:$('#ibYear').value.trim(),deadline:$('#ibDeadline').value,contact:$('#ibContact').value.trim()};
  if(!settings.year)throw new Error('Nhập năm học để phân biệt thông báo của từng năm.');
  const mapping=new Map((bundles.items||[]).map(b=>[slug(b.remark),b])),aliases=new Set(students.flatMap(s=>studentDueItems(s).map(x=>slug(x.paymentCode||'')))),entries=[];
  for(const group of groups){
    const options=[];
    for(const option of group.options){
      const plan=await insurancePlan(group.student,option,config,settings.year),key=slug(plan.remark),old=mapping.get(key);
      if((old&&old.signature!==plan.signature)||aliases.has(key))throw new Error('Phát hiện trùng mã thanh toán. Không xuất; cần kiểm tra danh sách mã.');
      mapping.set(key,old||plan);options.push({...option,plan:old||plan,payload:buildVietQrPayload(config,option.amount,plan.remark)});
    }
    entries.push({...group,options,settings,school,config});
  }
  // Keep every previously issued code: old QR images may still be paid later.
  await persistInsurancePlans([...mapping.values()],settings);
  return entries;
}

// Read and merge within one IndexedDB transaction so simultaneous tabs keep all issued codes.
async function persistInsurancePlans(plans,settings){
  let record;
  await new Promise((resolve,reject)=>{
    const tx=db.transaction('meta','readwrite'),store=tx.objectStore('meta'),read=store.get('noticeBundles');
    let failure;
    read.onsuccess=()=>{
      const merged=new Map((read.result?.items||[]).map(b=>[slug(b.remark),b]));
      for(const plan of plans){
        const key=slug(plan.remark),old=merged.get(key);
        if(old&&plan.kind==='insurance-choice'&&old.signature!==plan.signature){failure=new Error('Trùng mã thanh toán; không thể lưu thông báo.');tx.abort();return;}
        if(!old)merged.set(key,plan);
      }
      record={key:'noticeBundles',items:[...merged.values()]};store.put(record);store.put(settings);
    };
    tx.oncomplete=resolve;tx.onerror=()=>reject(failure||tx.error);tx.onabort=()=>reject(failure||tx.error||new Error('Không lưu được bảng mã.'));
  });
  setNoticeBundleCache(record);
}

// Render the encoded matrix at integer pixels with a four-module white quiet zone.
function insuranceQrCanvas(payload,maxSize=402){
  const holder=document.createElement('div'),qr=new QRCode(holder,{text:payload,width:256,height:256,correctLevel:QRCode.CorrectLevel.M});
  const matrix=qr._oQRCode,n=matrix.getModuleCount(),scale=Math.floor(maxSize/(n+8)),size=(n+8)*scale;
  if(scale<3)throw new Error('Nội dung QR quá dài để in rõ.');
  const canvas=document.createElement('canvas');canvas.width=canvas.height=size;
  const ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,size,size);ctx.fillStyle='#000';
  for(let row=0;row<n;row++)for(let col=0;col<n;col++)if(matrix.isDark(row,col))ctx.fillRect((col+4)*scale,(row+4)*scale,scale,scale);
  return canvas;
}
function insuranceCanvas(entry){
  const canvas=document.createElement('canvas');canvas.width=1240;canvas.height=1754;
  const c=canvas.getContext('2d'),teal='#006b68',ink='#173b3b',muted='#536b6b',gold='#e3bb53';
  c.fillStyle='#fff';c.fillRect(0,0,1240,1754);c.textBaseline='top';
  const text=(str,x,y,size=28,color=ink,weight='400',align='left',max=1100)=>{
    c.fillStyle=color;c.textAlign=align;let fs=size;c.font=`${weight} ${fs}px Arial`;
    while(c.measureText(String(str)).width>max&&fs>16){fs--;c.font=`${weight} ${fs}px Arial`;}
    c.fillText(str,x,y);return fs;
  };
  const wrap=(str,x,y,width,size=26,line=36,color=muted)=>{
    c.font=`400 ${size}px Arial`;let row='',dy=y;
    for(const word of String(str).split(/\s+/)){const next=row?row+' '+word:word;if(c.measureText(next).width>width&&row){text(row,x,dy,size,color);row=word;dy+=line;}else row=next;}
    if(row)text(row,x,dy,size,color);return dy+line;
  };
  c.fillStyle=teal;c.fillRect(0,0,1240,18);
  text(entry.school.schoolName,62,54,31,teal,'700','left',875);text('BIDV',1178,51,47,teal,'700','right',220);
  text('THÔNG BÁO NỘP BẢO HIỂM',620,139,49,teal,'700','center');
  text('Năm học '+entry.settings.year,620,206,28,muted,'400','center');
  c.fillStyle='#eef7f5';c.fillRect(62,272,1116,160);
  text('HỌC SINH',86,291,21,muted,'700');text(entry.student.name,86,325,39,ink,'700','left',1060);
  text('Lớp: '+(entry.student.className||'—'),86,386,26,ink,'700');text('Mã HS: '+entry.student.code,1154,386,26,ink,'400','right',830);
  text('CÁC KHOẢN THU',62,477,23,muted,'700');
  text('Bảo hiểm y tế (BHYT)',62,522,28,ink,'700');text(entry.healthPaid?'ĐÃ NỘP':money(entry.health.amount),1178,522,28,teal,'700','right');
  text('Bắt buộc theo đối tượng tham gia · Mức thu đã phân giao',62,562,23,muted);
  text('Bảo hiểm thân thể (BHTT)',62,613,28,ink,'700');text(entry.bodyPaid?'ĐÃ NỘP':money(entry.body.amount),1178,613,28,teal,'700','right');
  text('Tự nguyện · Phụ huynh lựa chọn tham gia',62,653,23,muted);
  if(entry.complete){
    c.fillStyle='#eef7f5';c.fillRect(62,737,1116,520);
    text('ĐÃ HOÀN THÀNH',620,906,52,teal,'700','center');
    text('CÁC KHOẢN BẢO HIỂM ĐÃ CHỌN',620,986,32,ink,'700','center');
    text('Không cần thanh toán thêm.',620,1050,30,muted,'400','center');
  }else{
    const width=entry.options.length===2?540:720,start=entry.options.length===2?62:260;
    entry.options.forEach((o,i)=>{
      const x=start+i*576,y=734;
      c.strokeStyle='#b6d7d0';c.lineWidth=2;c.strokeRect(x,y,width,646);c.fillStyle='#eef7f5';c.fillRect(x+1,y+1,width-2,171);
      text(entry.options.length===2?'PHƯƠNG ÁN '+(i+1):'KHOẢN CÒN LẠI',x+width/2,y+24,22,muted,'700','center');
      text(o.label,x+width/2,y+65,34,teal,'700','center',width-30);
      text(money(o.amount),x+width/2,y+116,40,ink,'700','center',width-30);
      const qr=insuranceQrCanvas(o.payload);c.drawImage(qr,Math.round(x+(width-qr.width)/2),y+190);
      text(o.plan.remark,x+width/2,y+601,23,muted,'400','center',width-30);
    });
  }
  const warning=entry.complete?'Cảm ơn quý phụ huynh đã hoàn thành.':entry.options.length===2?'CHỈ THANH TOÁN 01 TRONG 02 MÃ QR':'CHỈ THANH TOÁN KHOẢN CÒN LẠI';
  c.fillStyle='#fff5d9';c.fillRect(62,1410,1116,92);c.fillStyle=gold;c.fillRect(62,1410,7,92);
  text(warning,620,1427,31,ink,'700','center',1070);
  text(entry.complete?'Trạng thái theo báo cáo đã cập nhật tại trường.':'Không thanh toán lại nếu đã nộp. Không quét cả hai mã.',620,1470,23,ink,'400','center',1070);
  if(!entry.complete){
    text('TK nhận: '+entry.config.accountNumber+' · BIN '+entry.config.bin,62,1523,24,ink);
    text(entry.config.accountName,62,1558,25,ink,'700');
    text('Kiểm tra người nhận, giữ nguyên số tiền và mã nội dung khi chuyển.',62,1595,23,muted);
  }
  text('Hạn nộp: '+(entry.settings.deadline?noticeDateVi(entry.settings.deadline):'Theo thông báo của trường'),62,1642,24,ink,'700','left',1090);
  if(entry.settings.contact)text('Liên hệ: '+entry.settings.contact,62,1677,23,muted,'400','left',1090);
  c.fillStyle=teal;c.fillRect(0,1722,1240,32);text('BIDV – Đồng hành chuyển đổi số cùng ngành Giáo dục',620,1728,20,'#fff','400','center');
  return canvas;
}

// Minimal PDF writer: one high-resolution JPEG per A5 page, embedded without dependencies.
function insurancePdf(pages){
  const encoder=new TextEncoder(),parts=[],offsets=[0];let size=0;
  const append=value=>{const bytes=typeof value==='string'?encoder.encode(value):value;parts.push(bytes);size+=bytes.length;};
  const object=(id,body)=>{offsets[id]=size;append(`${id} 0 obj\n${body}\nendobj\n`);};
  append('%PDF-1.4\n');object(1,'<< /Type /Catalog /Pages 2 0 R >>');
  object(2,`<< /Type /Pages /Count ${pages.length} /Kids [${pages.map((_,i)=>`${3+i*3} 0 R`).join(' ')}] >>`);
  pages.forEach((jpeg,i)=>{
    const id=3+i*3;
    object(id,`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 419.53 595.28] /Resources << /XObject << /Im0 ${id+1} 0 R >> >> /Contents ${id+2} 0 R >>`);
    offsets[id+1]=size;append(`${id+1} 0 obj\n<< /Type /XObject /Subtype /Image /Width 1240 /Height 1754 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpeg.length} >>\nstream\n`);append(jpeg);append('\nendstream\nendobj\n');
    const stream='q\n419.53 0 0 595.28 0 0 cm\n/Im0 Do\nQ\n';object(id+2,`<< /Length ${encoder.encode(stream).length} >>\nstream\n${stream}endstream`);
  });
  const count=pages.length*3+3,xref=size;append(`xref\n0 ${count}\n0000000000 65535 f \n`);
  for(let i=1;i<count;i++)append(`${String(offsets[i]).padStart(10,'0')} 00000 n \n`);
  append(`trailer\n<< /Size ${count} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`);
  return new Blob(parts,{type:'application/pdf'});
}
function insuranceSafeName(value){return qrText(value,100).replace(/[^A-Z0-9._-]/g,'_')||'CHUA_XEP_LOP';}
async function showInsurancePreview(){
  const entry=insuranceView.entries[insuranceView.index];if(!entry)return;
  const canvas=insuranceCanvas(entry);canvas.setAttribute('role','img');canvas.setAttribute('aria-label',`Thông báo ${entry.student.name}, lớp ${entry.student.className}. ${entry.options.map(o=>o.label+': '+money(o.amount)).join('. ')}`);
  $('#ibPreviewImage').replaceChildren(canvas);$('#ibPreviewLabel').textContent=`${insuranceView.index+1}/${insuranceView.entries.length} · ${entry.student.name} · ${entry.student.className||'—'}`;
  $('#ibPrev').disabled=insuranceView.index===0;$('#ibNext').disabled=insuranceView.index===insuranceView.entries.length-1;
}
async function exportInsurance(){
  const entries=await buildInsuranceEntries(),zip=new JSZip(),classes=new Map();let done=0;
  entries.forEach(e=>{const name=e.student.className||'Chưa xếp lớp';if(!classes.has(name))classes.set(name,[]);classes.get(name).push(e);});
  const usedFolders=new Set();
  for(const [className,group] of classes){
    let folderName=insuranceSafeName(className);while(usedFolders.has(folderName))folderName+='_';usedFolders.add(folderName);
    const folder=zip.folder(folderName),pages=[],usedNames=new Set();
    for(const entry of group){
      if(insuranceView.cancel)throw new Error('Đã hủy xuất. Chưa tải file; mã đã tạo được giữ để đối soát.');
      const canvas=insuranceCanvas(entry);let stem=`${folderName}_${insuranceSafeName(entry.student.code)}_${insuranceSafeName(entry.student.name)}_BHYT_BHTT`;
      while(usedNames.has(stem))stem+='_';usedNames.add(stem);
      const png=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));if(!png)throw new Error('Không tạo được ảnh PNG.');
      folder.file(stem+'.png',await png.arrayBuffer());
      const jpeg=await new Promise(resolve=>canvas.toBlob(resolve,'image/jpeg',0.96));if(!jpeg)throw new Error('Không tạo được trang PDF.');
      pages.push(new Uint8Array(await jpeg.arrayBuffer()));
      $('#ibProgress').textContent=`Đang tạo ${++done}/${entries.length} thông báo · Lớp ${className}`;
      await new Promise(resolve=>setTimeout(resolve,0));
    }
    folder.file(`THONG_BAO_BHYT_BHTT_${folderName}.pdf`,await insurancePdf(pages).arrayBuffer());
  }
  // A readable mapping supports review and retains the exact codes issued with these notices.
  const csvCell=v=>'"'+String(v??'').replace(/^[=+@-]/,"'$&").replace(/"/g,'""')+'"';
  const rows=[['Lớp','Mã HS','Họ tên','Phương án','Mã tham chiếu','Tổng tiền','BHYT','BHTT']];
  entries.forEach(e=>e.options.forEach(o=>rows.push([e.student.className,e.student.code,e.student.name,o.label,o.plan.remark,o.amount,o.items.filter(x=>x.id===e.health.id).reduce((s,x)=>s+x.amount,0),o.items.filter(x=>x.id===e.body.id).reduce((s,x)=>s+x.amount,0)])));
  zip.file('BANG_MA_THANH_TOAN_BHYT_BHTT.csv','\uFEFF'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n'));
  zip.file('HUONG_DAN.txt','Mỗi thư mục lớp có ảnh PNG từng học sinh và PDF A5 cả lớp.\nPhụ huynh chỉ thanh toán một phương án. QR ảnh không tự vô hiệu hóa.\nĐối soát bằng báo cáo có mã IB... trong nội dung/mã tham chiếu. Mã này không tự đăng ký tại ngân hàng.\nSao lưu dữ liệu bằng chức năng Sao lưu của web sau khi phát hành; bảng CSV chỉ dùng tra cứu, không thay thế bản sao lưu.\n');
  const blob=await zip.generateAsync({type:'blob',compression:'STORE'},p=>{if(insuranceView.cancel)throw new Error('Đã hủy xuất.');$('#ibProgress').textContent=`Đóng gói ${Math.round(p.percent)}%`;});
  const scope=classes.size===1?insuranceSafeName([...classes.keys()][0]):`${classes.size}_LOP`;
  download(`QR_BAO_HIEM_BHYT_BHTT_${scope}_${insuranceSafeName(entries[0].settings.year)}.zip`,blob,'application/zip');
  $('#ibProgress').textContent=`Đã xuất ${entries.length} thông báo, ${classes.size} PDF lớp. Hãy sao lưu dữ liệu để giữ bảng mã đối soát.`;
  insuranceView.entries=entries;insuranceView.index=0;await showInsurancePreview();
}
async function insuranceRun(action){
  if(insuranceView.busy)return;insuranceView.busy=true;insuranceView.cancel=false;
  const controls=$$('#page-insurance input, #page-insurance select, #page-insurance button').filter(el=>el.id!=='ibCancel'),prior=controls.map(el=>el.disabled);
  controls.forEach(el=>el.disabled=true);$('#ibCancel').hidden=action!=='export';$('#ibProgress').textContent='Đang kiểm tra dữ liệu và tạo mã…';
  try{if(action==='export')await exportInsurance();else{insuranceView.entries=await buildInsuranceEntries();insuranceView.index=0;await showInsurancePreview();$('#ibProgress').textContent='Bản xem trước đã cập nhật. Mỗi QR có mã thanh toán riêng.';}}
  catch(error){console.error(error);$('#ibProgress').textContent=error.message;toast(error.message,true);}
  finally{insuranceView.busy=false;controls.forEach((el,i)=>el.disabled=prior[i]);$('#ibCancel').hidden=true;if(insuranceView.entries.length)await showInsurancePreview();}
}
function wireInsuranceTool(){
  ['ibHealth','ibBody','ibYear','ibDeadline','ibContact','ibStatus'].forEach(id=>$('#'+id).addEventListener('change',invalidateInsurancePreview));
  $('#ibClasses').addEventListener('change',()=>{const el=$('#ibClasses');if(el.selectedOptions.length>1)el.options[0].selected=false;invalidateInsurancePreview();});
  $('#ibSaveAccount').onclick=()=>saveInsuranceAccount().catch(e=>toast(e.message,true));
  $('#ibPreview').onclick=()=>insuranceRun('preview');$('#ibExport').onclick=()=>insuranceRun('export');$('#ibCancel').onclick=()=>{insuranceView.cancel=true;};
  $('#ibPrev').onclick=async()=>{if(insuranceView.index>0){insuranceView.index--;await showInsurancePreview();}};
  $('#ibNext').onclick=async()=>{if(insuranceView.index<insuranceView.entries.length-1){insuranceView.index++;await showInsurancePreview();}};
}
