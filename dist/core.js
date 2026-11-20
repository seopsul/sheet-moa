(function(root){
  'use strict';
  // Headers are matched by their displayed text and occurrence, including duplicates.
  function mergeTables(tables,options={}){
    const header=options.header!==false, source=!!options.source;
    const columns=[], index=new Map(), rows=[];
    for(const table of tables){
      const data=table.rows.filter(row=>row.some(v=>v!==null&&v!==undefined&&v!==''));
      if(!data.length)continue;
      const width=data.reduce((max,row)=>Math.max(max,row.length),0);
      const first=header?data[0]:[];const seen=new Map();
      const mapping=Array.from({length:width},(_,i)=>{
        const title=header?String(first[i]??'').trim():`열 ${i+1}`;
        const base=title||`열 ${i+1}`;
        const occurrence=(seen.get(base)||0)+1;seen.set(base,occurrence);
        const key=JSON.stringify([base,occurrence]);
        if(!index.has(key)){index.set(key,columns.length);columns.push(base);}
        return index.get(key);
      });
      for(const input of data.slice(header?1:0)){
        const values=[];mapping.forEach((target,i)=>{values[target]=input[i]??'';});
        rows.push({values,file:table.file,sheet:table.sheet});
      }
    }
    const used=new Set();const unique=base=>{let name=base,n=2;while(used.has(name))name=`${base} (${n++})`;used.add(name);return name;};
    const headers=columns.map(unique);
    if(source)headers.push(unique('원본 파일'),unique('원본 시트'));
    return {headers,rows:rows.map(row=>{
      const values=Array.from({length:columns.length},(_,i)=>row.values[i]??'');
      if(source)values.push(row.file,row.sheet);return values;
    })};
  }
  function csvText(result){
    const cell=v=>{let s=v==null?'':String(v);if(typeof v==='string'&&/^[\s\u0000-\u001f]*[=+@-]/.test(s))s="'"+s;return /[",\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;};
    return '\uFEFF'+[result.headers,...result.rows].map(row=>row.map(cell).join(',')).join('\r\n');
  }
  root.MergeCore={mergeTables,csvText};
  if(typeof module!=='undefined')module.exports=root.MergeCore;
})(typeof globalThis!=='undefined'?globalThis:this);
