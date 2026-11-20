'use strict';
const $=id=>document.getElementById(id);
let uploads=[],result=null,busy=false;
function status(message,error=false){$('status').textContent=message;$('status').classList.toggle('error',error);}
function invalidate(){result=null;$('download').disabled=true;$('badge').textContent='대기 중';$('row-count').textContent='—';$('col-count').textContent='—';$('preview-empty').hidden=false;$('preview-wrap').hidden=true;}
function renderFiles(){
  $('file-list').replaceChildren();
  uploads.forEach((file,i)=>{
    const row=document.createElement('div');row.className='file-row';
    const icon=document.createElement('span');icon.className='file-icon';icon.textContent=file.name.split('.').pop().toUpperCase();
    const info=document.createElement('div');info.className='file-info';
    const name=document.createElement('div');name.className='file-name';name.textContent=file.name;name.title=file.name;
    const size=document.createElement('div');size.className='file-size';size.textContent=`${(file.size/1024).toLocaleString('ko-KR',{maximumFractionDigits:1})} KB`;
    info.append(name,size);const remove=document.createElement('button');remove.textContent='×';remove.disabled=busy;remove.setAttribute('aria-label',`${file.name} 삭제`);remove.onclick=()=>{uploads.splice(i,1);invalidate();renderFiles();status('파일 목록이 변경되었습니다. 다시 병합해 주세요.');};row.append(icon,info,remove);$('file-list').append(row);
  });
  $('empty').hidden=uploads.length>0;$('file-count').replaceChildren(document.createTextNode(`${uploads.length} `),Object.assign(document.createElement('small'),{textContent:'개'}));
  $('merge').disabled=busy||!uploads.length;$('clear').disabled=busy||!uploads.length;
}
function addFiles(files){if(busy)return;let rejected=0;for(const file of files){if(!/\.(csv|xlsx)$/i.test(file.name)){rejected++;continue;}if(!uploads.some(f=>f.name===file.name&&f.size===file.size&&f.lastModified===file.lastModified))uploads.push(file);}invalidate();renderFiles();status(rejected?`지원하지 않는 파일 ${rejected}개를 제외했습니다. CSV 또는 XLSX만 추가해 주세요.`:`${uploads.length}개 파일을 추가했습니다. 병합 설정을 확인해 주세요.`,rejected>0);}
$('files').onchange=e=>{addFiles(e.target.files);e.target.value='';};
$('drop').onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();if(!busy)$('files').click();}};
for(const event of ['dragenter','dragover'])$('drop').addEventListener(event,e=>{e.preventDefault();$('drop').classList.add('over');});
for(const event of ['dragleave','drop'])$('drop').addEventListener(event,e=>{e.preventDefault();$('drop').classList.remove('over');});
$('drop').addEventListener('drop',e=>addFiles(e.dataTransfer.files));
window.addEventListener('dragover',e=>e.preventDefault());window.addEventListener('drop',e=>e.preventDefault());
$('clear').onclick=()=>{uploads=[];invalidate();renderFiles();status('파일을 추가하면 병합할 수 있습니다.');};
for(const id of ['header','source','sheets','encoding'])$(id).onchange=()=>{invalidate();status('설정이 변경되었습니다. 다시 병합해 주세요.');};
function setBusy(value){busy=value;for(const id of ['header','source','sheets','encoding','files'])$(id).disabled=value;renderFiles();$('merge').textContent=value?'병합하는 중…':'파일 병합하기';}
function preview(){
  const table=$('preview');table.replaceChildren();
  const head=table.createTHead().insertRow();result.headers.slice(0,30).forEach(title=>{const th=document.createElement('th');th.scope='col';th.textContent=title;head.append(th);});
  const body=table.createTBody();result.rows.slice(0,100).forEach(row=>{const tr=body.insertRow();row.slice(0,30).forEach(value=>{const td=tr.insertCell();td.textContent=String(value);td.title=String(value);});});
  $('row-count').textContent=result.rows.length.toLocaleString();$('col-count').textContent=result.headers.length.toLocaleString();$('badge').textContent='병합 완료';$('preview-empty').hidden=true;$('preview-wrap').hidden=false;$('preview-note').textContent=`전체 ${result.rows.length.toLocaleString()}행 · ${result.headers.length}열이 다운로드 파일에 포함됩니다.`;$('download').disabled=false;
}
$('merge').onclick=async()=>{
  if(busy||!uploads.length)return;invalidate();setBusy(true);
  try{
    if(!window.XLSX)throw new Error('Excel 처리 라이브러리를 불러오지 못했습니다. 페이지를 다시 열어 주세요.');
    const tables=[];let current='';
    for(let i=0;i<uploads.length;i++){
      const file=uploads[i];current=file.name;status(`${i+1} / ${uploads.length} · ${current} 읽는 중…`);await new Promise(resolve=>setTimeout(resolve,0));
      try{
        const bytes=await file.arrayBuffer();
        const csv=/\.csv$/i.test(file.name);
        const book=csv?XLSX.read(new TextDecoder($('encoding').value).decode(bytes),{type:'string',raw:true}):XLSX.read(bytes,{type:'array',cellDates:false});
        const names=$('sheets').value==='all'?book.SheetNames:book.SheetNames.slice(0,1);
        for(const name of names){const sheet=book.Sheets[name];if(!sheet||!sheet['!ref'])continue;
          const range=XLSX.utils.decode_range(sheet['!ref']);if(range.e.c>=16384||range.e.r>=1048576)throw new Error('Excel의 행 또는 열 한도를 초과합니다.');
          tables.push({file:file.name,sheet:csv?'CSV':name,rows:XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:false})});
        }
      }catch(error){throw new Error(`${current}: 파일을 읽지 못했습니다. ${error.message||'손상 또는 암호화 여부를 확인해 주세요.'}`);}
    }
    result=MergeCore.mergeTables(tables,{header:$('header').checked,source:$('source').checked});
    if(!result.headers.length){invalidate();throw new Error('병합할 데이터가 없습니다. 파일 내용을 확인해 주세요.');}
    preview();status(`${uploads.length}개 파일 병합 완료. ${result.rows.length.toLocaleString()}행을 저장할 수 있습니다.`);
  }catch(error){invalidate();status(error.message||'병합 중 오류가 발생했습니다.',true);}finally{setBusy(false);}
};
$('download').onclick=()=>{
  if(!result||busy)return;
  try{
    const format=document.querySelector('input[name="format"]:checked').value;
    const base=($('filename').value.trim().replace(/\.(csv|xlsx)$/i,'').replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_')||'병합된_데이터').slice(0,120);
    let blob;
    if(format==='csv')blob=new Blob([MergeCore.csvText(result)],{type:'text/csv;charset=utf-8;'});
    else{
      if(result.rows.length+1>1048576||result.headers.length>16384)throw new Error('XLSX 저장 한도(1,048,576행 / 16,384열)를 초과했습니다. CSV를 선택해 주세요.');
      const book=XLSX.utils.book_new();const sheet=XLSX.utils.aoa_to_sheet([result.headers,...result.rows]);
      sheet['!cols']=result.headers.map(()=>({wch:20}));XLSX.utils.book_append_sheet(book,sheet,'병합 데이터');blob=new Blob([XLSX.write(book,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    }
    const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=`${base}.${format}`;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);status(`${base}.${format} 다운로드를 시작했습니다.`);
  }catch(error){status(error.message||'다운로드 중 오류가 발생했습니다.',true);}
};
