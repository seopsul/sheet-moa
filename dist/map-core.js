(function(root){
  'use strict';
  function number(value){if(value===null||value===undefined||String(value).trim()==='')return null;const text=String(value).trim();if(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text))return null;const n=Number(text);return Number.isFinite(n)?n:null;}
  function coordinates(headers,groups,latCol,lonCol,labelCol=-1){
    if(!Number.isInteger(latCol)||!Number.isInteger(lonCol)||latCol<0||lonCol<0||latCol>=headers.length||lonCol>=headers.length||latCol===lonCol)throw Error('서로 다른 위도 열과 경도 열을 선택하세요.');
    const points=[],invalid=[];
    groups.forEach((group,groupIndex)=>group.rows.forEach((row,rowIndex)=>{
      const lat=number(row[latCol]),lon=number(row[lonCol]);let reason='';
      if(lat===null||lon===null)reason='위도 또는 경도가 비어 있거나 숫자가 아닙니다.';
      else if(lat < -90||lat > 90||lon < -180||lon > 180)reason='좌표 범위 오류 (위도 ±90, 경도 ±180). 열 순서를 확인하세요.';
      const record={groupIndex,rowIndex,group:group.value,lat,lon,row,label:labelCol>=0?String(row[labelCol]??''):`지점 ${rowIndex+1}`};
      if(reason)invalid.push({...record,reason});else points.push(record);
    }));return {points,invalid};
  }
  function guess(headers,kind){const aliases=kind==='lat'?['latitude','lat','위도','y','y좌표']:kind==='lon'?['longitude','lon','lng','long','경도','x','x좌표']:['name','label','title','id','이름','명칭','지점명','주소'];return headers.findIndex(h=>aliases.includes(String(h).trim().toLowerCase().replace(/[ _-]/g,'')));}
  root.MapCore={number,coordinates,guess};if(typeof module!=='undefined')module.exports=root.MapCore;
})(typeof globalThis!=='undefined'?globalThis:this);
