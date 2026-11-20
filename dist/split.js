'use strict';
const $=id=>document.getElementById(id);
let file=null,book=null,rows=[],rowNumbers=[],result=null,loading=false,version=0,selected=0;
function status(text,error=false){$('status').textContent=text;$('status').classList.toggle('error',error);}
function invalidate(){result=null;$('results').hidden=true;$('preview-empty').hidden=false;$('download-all').disabled=true;$('open-map').disabled=true;$('badge').textContent='대기 중';for(const id of ['row-count','group-count','col-count'])$(id).textContent='—';}
function busy(value){loading=value;for(const id of ['file','header','encoding','filename','split-mode','date-order'])$(id).disabled=value;$('sheet').disabled=value||!book||/\.csv$/i.test(file?.name||'');$('column').disabled=value||!rows.length;$('reset').disabled=value||!file;$('split').disabled=value||!rows.length;$('download-all').disabled=value||!result;$('open-map').disabled=value||!result;document.querySelectorAll('#groups button,input[name="format"]').forEach(el=>el.disabled=value);}
function readSheet(){
  invalidate();const sheet=book?.Sheets[$('sheet').value];rows=[];rowNumbers=[];
  if(sheet?.['!ref']){
    const range=XLSX.utils.decode_range(sheet['!ref']);range.s.c=0;
    const all=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:true,range});
    all.forEach((row,i)=>{if(row.some(value=>value!==''&&value!=null)){rows.push(row);rowNumbers.push(range.s.r+i);}});
  }
  const width=rows.reduce((max,row)=>Math.max(max,row.length),0);$('column').replaceChildren();for(let i=0;i<width;i++){const name=$('header').checked?String(rows[0]?.[i]??''):`열 ${i+1}`;$('column').add(new Option(`${name||'(제목 없음)'} · ${i+1}열`,String(i)));}
  busy(false);status(width?'분리 기준 열을 선택한 뒤 파일 분리를 눌러주세요.':'선택한 시트에 데이터가 없습니다.',!width);
}
async function loadFile(next=file){
  if(!next)return;
  if(!/\.(csv|xlsx)$/i.test(next.name)){status('CSV 또는 XLSX 파일을 선택하세요.',true);return;}
  file=next;const token=++version;invalidate();book=null;rows=[];busy(true);$('file-name').textContent=file.name;status('파일을 읽고 있습니다…');
  try{
    const bytes=await next.arrayBuffer();if(token!==version)return;
    if(!window.XLSX)throw Error('Excel 처리 라이브러리를 불러오지 못했습니다.');
    const csv=/\.csv$/i.test(next.name);book=csv?XLSX.read(new TextDecoder($('encoding').value).decode(bytes),{type:'string',raw:true}):XLSX.read(bytes,{type:'array',cellNF:true});
    $('sheet').replaceChildren();book.SheetNames.forEach(name=>$('sheet').add(new Option(name,name)));readSheet();
  }catch(error){book=null;rows=[];$('sheet').replaceChildren(new Option('파일을 선택하세요'));$('column').replaceChildren(new Option('파일을 선택하세요'));status(`파일을 읽지 못했습니다. ${error.message}`,true);}finally{if(token===version)busy(false);}
}
function accept(files){if(loading)return;if(files.length!==1){status('분리할 파일을 한 개만 선택해 주세요.',true);return;}loadFile(files[0]);}
$('file').onchange=e=>{accept(e.target.files);e.target.value='';};
$('drop').onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();if(!loading)$('file').click();}};
for(const type of ['dragenter','dragover'])$('drop').addEventListener(type,e=>{e.preventDefault();$('drop').classList.add('over');});
for(const type of ['drop','dragleave'])$('drop').addEventListener(type,e=>{e.preventDefault();$('drop').classList.remove('over');});
$('drop').addEventListener('drop',e=>accept(e.dataTransfer.files));window.addEventListener('dragover',e=>e.preventDefault());window.addEventListener('drop',e=>e.preventDefault());
$('encoding').onchange=()=>loadFile();$('header').onchange=()=>{if(book)readSheet();};$('sheet').onchange=readSheet;$('column').onchange=()=>{invalidate();status('기준 열이 바뀌었습니다. 다시 분리해 주세요.');};
function updateDateSettings(){const date=$('split-mode').value!=='value';$('date-settings').hidden=!date;$('date-hint').hidden=!date;invalidate();status('분리 방식이 바뀌었습니다. 다시 분리해 주세요.');}
$('split-mode').onchange=updateDateSettings;$('date-order').onchange=()=>{invalidate();status('날짜 순서가 바뀌었습니다. 다시 분리해 주세요.');};
function dateValues(column){
  const sheet=book.Sheets[$('sheet').value],date1904=!!book.Workbook?.WBProps?.date1904;
  return rows.map((row,i)=>{
    const cell=sheet[XLSX.utils.encode_cell({r:rowNumbers[i],c:column})];
    if(cell?.t==='n'&&cell.z&&XLSX.SSF.is_date(cell.z)){
      const parsed=XLSX.SSF.parse_date_code(cell.v,{date1904});
      return parsed?`${String(parsed.y).padStart(4,'0')}-${String(parsed.m).padStart(2,'0')}-${String(parsed.d).padStart(2,'0')}`:String(row[column]??'');
    }
    return row[column]??'';
  });
}
$('reset').onclick=()=>{++version;file=book=null;rows=[];invalidate();$('sheet').replaceChildren(new Option('파일을 선택하세요'));$('column').replaceChildren(new Option('파일을 선택하세요'));$('file-name').textContent='선택한 파일이 없습니다.';busy(false);status('파일을 추가하고 분리할 기준 열을 선택하세요.');};
function preview(i){selected=i;const group=result.groups[i];$('preview-title').textContent=`${group.value===''?'(빈 값)':JSON.stringify(group.value)} 미리보기`;
  const table=$('preview');table.replaceChildren();const head=table.createTHead().insertRow();result.headers.slice(0,30).forEach(name=>{const th=document.createElement('th');th.scope='col';th.textContent=name;head.append(th);});
  const body=table.createTBody();group.rows.slice(0,100).forEach(row=>{const tr=body.insertRow();row.slice(0,30).forEach(value=>{const td=tr.insertCell();td.textContent=String(value);});});
  document.querySelectorAll('.group-preview').forEach((button,index)=>button.setAttribute('aria-pressed',String(index===i)));
}
$('split').onclick=()=>{if(loading||!rows.length)return;invalidate();try{const column=Number($('column').value),mode=$('split-mode').value;result=SplitCore.splitRows(rows,column,$('header').checked,{mode,dateOrder:$('date-order').value,dateValues:mode==='value'?undefined:dateValues(column)});if(!result.groups.length)throw Error('제목 행 외에 분리할 데이터가 없습니다.');$('groups').replaceChildren();
  result.groups.forEach((group,i)=>{const row=document.createElement('div');row.className='group-row';const button=document.createElement('button');button.className='group-preview';button.textContent=`${group.value===''?'(빈 값)':JSON.stringify(group.value)} · ${group.rows.length.toLocaleString()}행`;button.onclick=()=>preview(i);const download=document.createElement('button');download.className='text-button';download.textContent='다운로드';download.setAttribute('aria-label',`${group.value||'빈 값'} 그룹 다운로드`);download.onclick=()=>downloadOne(i);row.append(button,download);$('groups').append(row);});
  $('row-count').textContent=result.groups.reduce((n,g)=>n+g.rows.length,0).toLocaleString();$('group-count').textContent=result.groups.length.toLocaleString();$('col-count').textContent=result.headers.length;$('badge').textContent='분리 완료';$('results').hidden=false;$('preview-empty').hidden=true;$('download-all').disabled=false;$('open-map').disabled=false;preview(0);status(`${result.groups.length.toLocaleString()}개 그룹으로 분리했습니다. 개별 또는 ZIP으로 저장하세요.`+(result.invalidCount?` 날짜를 확인할 수 없는 ${result.invalidCount}행은 ‘날짜 확인 필요’ 그룹에 포함했습니다.`:''));
}catch(error){invalidate();status(error.message,true);}};
function extension(){return document.querySelector('input[name="format"]:checked').value;}
function groupBytes(group,format){const data={headers:result.headers,rows:group.rows};if(format==='csv')return new TextEncoder().encode(MergeCore.csvText(data));if(data.rows.length+1>1048576||data.headers.length>16384)throw Error('XLSX 저장 한도를 초과했습니다. CSV를 선택해 주세요.');const workbook=XLSX.utils.book_new();XLSX.utils.book_append_sheet(workbook,XLSX.utils.aoa_to_sheet([data.headers,...data.rows]),'분리 데이터');return new Uint8Array(XLSX.write(workbook,{type:'array',bookType:'xlsx'}));}
function save(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
function downloadOne(i){if(!result||loading)return;try{const format=extension(),names=SplitCore.namesFor(result.groups,$('filename').value,format);save(new Blob([groupBytes(result.groups[i],format)],{type:format==='csv'?'text/csv;charset=utf-8':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),names[i]);status(`${names[i]} 다운로드를 시작했습니다.`);}catch(error){status(error.message,true);}}
$('download-all').onclick=async()=>{if(!result||loading)return;const format=extension(),prefix=$('filename').value,names=SplitCore.namesFor(result.groups,prefix,format);busy(true);try{const files=[];for(let i=0;i<result.groups.length;i++){status(`${i+1} / ${result.groups.length} 파일 준비 중…`);await new Promise(resolve=>setTimeout(resolve,0));files.push({name:names[i],data:groupBytes(result.groups[i],format)});}save(SplitCore.zip(files),SplitCore.safeName(prefix||'분리된_데이터')+'.zip');status(`${files.length}개 파일이 들어 있는 ZIP 다운로드를 시작했습니다.`);}catch(error){status(error.message,true);}finally{busy(false);}};
$('open-map').onclick=async()=>{if(!result||loading)return;busy(true);try{const token=await MapBridge.put({name:file?.name||'분리 결과',headers:result.headers,groups:result.groups});location.href='map.html?split='+encodeURIComponent(token);}catch(error){status(error.message+' 분리 파일을 다운로드한 뒤 위치 지도 페이지에서 직접 추가할 수도 있습니다.',true);busy(false);}};
