const {chromium}=require('playwright');
const fs=require('node:fs'),assert=require('node:assert/strict'),XLSX=require('./dist/vendor/xlsx.full.min.js');
const core=require('./dist/split-core.js');
const data=[['이름','지역'],['홍길동','서울'],['김민수','부산'],['이영희','서울'],['미정',''],['원본빈값','빈 값'],['슬래시','A/B'],['역슬래시','A\\B'],['대문자','TEAM'],['소문자','team']];
assert.equal(core.splitRows(data,1).groups.length,8);
assert.equal(core.splitRows(data,0,false).groups.length,10);
const groups=core.splitRows(data,1).groups,names=core.namesFor(groups,'분리','csv');assert.equal(new Set(names.map(n=>n.toLowerCase())).size,names.length);
(async()=>{fs.mkdirSync('test-results',{recursive:true});const browser=await chromium.launch({headless:true,channel:'msedge'});try{
  const p=await browser.newPage({viewport:{width:1360,height:1100}}),errors=[];p.on('pageerror',e=>errors.push(e.message));
  await p.goto(require('node:url').pathToFileURL(require('node:path').resolve('dist/split.html')).href);
  const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(data),'지역별');XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['부서','수량'],['영업',1],['기술',2]]),'부서별');
  await p.locator('#file').setInputFiles({name:'원본.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from(XLSX.write(book,{type:'array',bookType:'xlsx'}))});await p.waitForFunction(()=>!document.getElementById('split').disabled);
  await p.locator('#column').selectOption('1');await p.locator('#split').click();assert.equal(await p.locator('#group-count').textContent(),'8');assert.equal(await p.locator('#row-count').textContent(),'9');
  const one=p.waitForEvent('download');await p.locator('.group-row').first().getByRole('button',{name:'다운로드',exact:false}).click();const single=await one;const b=XLSX.read(fs.readFileSync(await single.path()),{type:'buffer'});assert.equal(b.Sheets[b.SheetNames[0]]['!ref'],'A1:B3');assert.equal(b.Sheets[b.SheetNames[0]].A3.v,'이영희');
  fs.mkdirSync('test-results/split-tests',{recursive:true});
  for(const format of ['xlsx','csv']){await p.locator(`input[value="${format}"]`).check();const pending=p.waitForEvent('download');await p.locator('#download-all').click();const download=await pending;await download.saveAs(`test-results/split-tests/${format}.zip`);await p.waitForFunction(()=>!document.getElementById('download-all').disabled);}
  await p.screenshot({path:'test-results/split-preview.png',fullPage:true});await p.setViewportSize({width:390,height:844});assert.equal(await p.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);
  await p.locator('#sheet').selectOption('부서별');assert.equal(await p.locator('#download-all').isDisabled(),true);await p.locator('#split').click();assert.equal(await p.locator('#group-count').textContent(),'2');
  await p.locator('#header').uncheck();await p.locator('#split').click();assert.equal(await p.locator('#row-count').textContent(),'3');
  await p.locator('#reset').click();assert.equal(await p.locator('#split').isDisabled(),true);
  await p.locator('#file').setInputFiles({name:'한글.csv',mimeType:'text/csv',buffer:Buffer.from('\uFEFF지역,코드\n서울,001\n부산,002\n서울,003')});await p.waitForFunction(()=>!document.getElementById('split').disabled);await p.locator('#header').check();await p.locator('#split').click();assert.equal(await p.locator('#group-count').textContent(),'2');assert.ok((await p.locator('#preview').textContent()).includes('001'));
  await p.getByRole('link',{name:'파일 병합',exact:true}).click();assert.equal(await p.title(),'시트모아 · Excel 파일 병합');assert.equal(await p.getByRole('link',{name:'파일 분리',exact:true}).count(),1);
  assert.deepEqual(errors,[]);console.log('PASS: split groups, empty and special values, name collisions, XLSX individual output, CSV/XLSX ZIP, sheet/header changes, reset, CSV leading zeros, mobile layout and navigation');
}finally{await browser.close();}})().catch(error=>{console.error(error);process.exit(1)});
