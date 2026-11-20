(function(root){
  'use strict';
  function dateParts(value,order='ymd'){
    const text=String(value??'').trim();let match;
    if((match=/^(\d{4})(\d{2})(\d{2})$/.exec(text)))return validDate(+match[1],+match[2],+match[3]);
    if((match=/^(\d{4})(\d{2})$/.exec(text)))return validDate(+match[1],+match[2],1);
    if((match=/^(\d{4})$/.exec(text)))return +match[1]>0?{year:+match[1],month:null}:null;
    const normalized=text.replace(/년\s*/g,'-').replace(/월\s*/g,'-').replace(/일/g,'').replace(/\s*([./-])\s*/g,'$1').replace(/\.$/,'');
    match=/^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?(.*)$/.exec(normalized);
    if(match&&validTime(match[4]))return validDate(+match[1],+match[2],+(match[3]||1));
    if(order!=='ymd'){
      match=/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(.*)$/.exec(normalized);
      if(match&&validTime(match[4]))return validDate(+match[3],+(order==='mdy'?match[1]:match[2]),+(order==='mdy'?match[2]:match[1]));
    }
    return null;
  }
  function validTime(suffix){return suffix===''||/^[T ](?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d(?:\.\d+)?)?(?:Z|[+-](?:[01]\d|2[0-3]):?[0-5]\d)?$/.test(suffix);}
  function validDate(year,month,day){
    const leap=year%4===0&&(year%100!==0||year%400===0),days=[31,leap?29:28,31,30,31,30,31,31,30,31,30,31];
    return year>=1&&year<=9999&&month>=1&&month<=12&&day>=1&&day<=days[month-1]?{year,month}:null;
  }
  function splitRows(rows,column,header=true,options={}){
    const data=rows.filter(row=>row.some(v=>v!==''&&v!=null));
    const width=data.reduce((n,row)=>Math.max(n,row.length),0);
    if(!Number.isInteger(column)||column<0||column>=width)throw Error('분리 기준 열을 선택하세요.');
    const headers=Array.from({length:width},(_,i)=>header?String(data[0]?.[i]??''):`열 ${i+1}`);
    const groups=new Map(),mode=options.mode||'value';let invalidCount=0;
    if(!['value','year','month'].includes(mode))throw Error('분리 방식을 선택하세요.');
    for(const [index,row] of data.slice(header?1:0).entries()){
      let key=String(row[column]??''),kind='value';
      if(mode!=='value'){
        const value=options.dateValues?.[index+(header?1:0)]??key;
        if(String(value).trim()===''){key='';kind='empty';}
        else{
          const parts=dateParts(value,options.dateOrder);
          if(!parts||(mode==='month'&&!parts.month)){key='날짜 확인 필요';kind='invalid';invalidCount++;}
          else{key=String(parts.year).padStart(4,'0')+(mode==='month'?'-'+String(parts.month).padStart(2,'0'):'');kind='date';}
        }
      }
      const mapKey=JSON.stringify([kind,key]);
      if(!groups.has(mapKey))groups.set(mapKey,{value:key,kind,rows:[]});
      groups.get(mapKey).rows.push(Array.from({length:width},(_,i)=>row[i]??''));
    }
    const output=[...groups.values()];if(mode!=='value')output.sort((a,b)=>((a.kind==='date'?0:a.kind==='empty'?1:2)-(b.kind==='date'?0:b.kind==='empty'?1:2))||a.value.localeCompare(b.value));
    return {headers,groups:output,invalidCount};
  }
  function safeName(value){const name=String(value).replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').replace(/[. ]+$/g,'').slice(0,80);return name||'빈 값';}
  function namesFor(groups,prefix,extension){
    const used=new Set();return groups.map(group=>{const base=safeName(prefix||'분리된_데이터')+'_'+safeName(group.value);let name=base+'.'+extension,n=2;while(used.has(name.toLowerCase()))name=base+' ('+(n++)+').'+extension;used.add(name.toLowerCase());return name;});
  }
  const crcTable=Array.from({length:256},(_,n)=>{for(let k=0;k<8;k++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
  function crc32(bytes){let crc=0xffffffff;for(const b of bytes)crc=crcTable[(crc^b)&255]^(crc>>>8);return (crc^0xffffffff)>>>0;}
  // ZIP store method: UTF-8 filenames, CRC checksums and central directory.
  function zip(files){
    if(files.length>65535)throw Error('ZIP 하나에 넣을 수 있는 파일 수(65,535개)를 초과했습니다.');
    const parts=[],central=[];let offset=0;
    const now=new Date(),date=((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate(),time=(now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1);
    for(const file of files){
      const name=new TextEncoder().encode(file.name),data=file.data,crc=crc32(data);
      if(offset+data.length+name.length+30>0xffffffff)throw Error('ZIP 크기가 4GB 한도를 초과합니다.');
      const local=new Uint8Array(30+name.length),v=new DataView(local.buffer);
      v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint16(10,time,true);v.setUint16(12,date,true);v.setUint32(14,crc,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,name.length,true);local.set(name,30);parts.push(local,data);
      const c=new Uint8Array(46+name.length),cv=new DataView(c.buffer);cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(8,0x800,true);cv.setUint16(12,time,true);cv.setUint16(14,date,true);cv.setUint32(16,crc,true);cv.setUint32(20,data.length,true);cv.setUint32(24,data.length,true);cv.setUint16(28,name.length,true);cv.setUint32(42,offset,true);c.set(name,46);central.push(c);offset+=local.length+data.length;
    }
    const size=central.reduce((n,p)=>n+p.length,0);if(offset+size+22>0xffffffff)throw Error('ZIP 크기가 4GB 한도를 초과합니다.');
    const end=new Uint8Array(22),ev=new DataView(end.buffer);ev.setUint32(0,0x06054b50,true);ev.setUint16(8,files.length,true);ev.setUint16(10,files.length,true);ev.setUint32(12,size,true);ev.setUint32(16,offset,true);
    return new Blob([...parts,...central,end],{type:'application/zip'});
  }
  root.SplitCore={splitRows,dateParts,safeName,namesFor,zip};if(typeof module!=='undefined')module.exports=root.SplitCore;
})(typeof globalThis!=='undefined'?globalThis:this);
