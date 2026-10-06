const demoStudent={code:'TEST00001',name:'NGUYỄN MINH AN · DỮ LIỆU THỬ',className:'6A1',dueItems:[{id:'h',catalogId:'h',name:'BHYT',category:'insurance',amount:631800},{id:'b',catalogId:'b',name:'BHTT',category:'mandatory',amount:150000}]};
const demoConfig={bin:'970418',accountNumber:'0000000000',accountName:'TAI KHOAN KIEM THU'};
let demoCanvas;
async function drawDemo(){
  try{
    const state=$('#state').value,ids=state==='health'?['h']:state==='body'?['b']:state==='both'?['h','b']:[];
    const entry=insuranceChoices(demoStudent,ids.length?[{studentCode:demoStudent.code,paymentStatus:'valid',matchedDueItemIds:ids}]:[],'catalog:h','catalog:b');
    entry.settings={year:'2026–2027',deadline:'2026-10-31',contact:'Văn phòng trường · DỮ LIỆU KIỂM THỬ'};
    entry.school={schoolName:'TRƯỜNG THCS SỐ 2 HIẾU GIANG'};entry.config=demoConfig;
    for(const option of entry.options){option.plan=await insurancePlan(demoStudent,option,demoConfig,entry.settings.year);option.payload=buildVietQrPayload(demoConfig,option.amount,option.plan.remark);}
    demoCanvas=insuranceCanvas(entry);$('#preview').replaceChildren(demoCanvas);$('#result').textContent=`${entry.options.length} QR · ${entry.options.map(o=>o.label+' '+money(o.amount)).join(' / ')||'Đã hoàn thành'}`;
  }catch(e){$('#result').textContent=e.message;console.error(e);}
}
$('#state').onchange=drawDemo;
$('#png').onclick=()=>demoCanvas.toBlob(blob=>download('KIEM_THU_BHYT_BHTT_6A1.png',blob,'image/png'));
$('#pdf').onclick=()=>demoCanvas.toBlob(async blob=>download('KIEM_THU_BHYT_BHTT_6A1.pdf',insurancePdf([new Uint8Array(await blob.arrayBuffer())]),'application/pdf'),'image/jpeg',.96);
drawDemo();
