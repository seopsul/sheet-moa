const {chromium}=require('playwright');
const fs=require('node:fs');const assert=require('node:assert/strict');const XLSX=require('./dist/vendor/xlsx.full.min.js');
(async()=>{fs.mkdirSync('test-results',{recursive:true});
  const browser=await chromium.launch({headless:true,channel:'msedge'});
  try{
    const page=await browser.newPage({viewport:{width:1360,height:1100}});const errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.goto(require('node:url').pathToFileURL(require('node:path').resolve('dist/index.html')).href);
    const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['금액','이름','지역'],[2000,'김민수','서울']]),'첫 시트');XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['이름','금액'],['이영희',3000]]),'둘째 시트');
    await page.locator('#files').setInputFiles([{name:'첫파일.csv',mimeType:'text/csv',buffer:Buffer.from('\uFEFF이름,금액\r\n홍길동,1000')},{name:'둘째.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(XLSX.write(book,{type:'array',bookType:'xlsx'}))}]);
    await page.locator('#sheets').selectOption('all');await page.locator('#source').check();await page.locator('#merge').click();await page.waitForFunction(()=>document.getElementById('badge').textContent==='병합 완료');
    assert.equal(await page.locator('#row-count').textContent(),'3');assert.equal(await page.locator('#col-count').textContent(),'5');
    for(const format of ['xlsx','csv']){
      await page.locator(`input[value="${format}"]`).check();const promise=page.waitForEvent('download');await page.locator('#download').click();const download=await promise;const bytes=fs.readFileSync(await download.path());const merged=XLSX.read(bytes,{type:'buffer'});assert.equal(merged.Sheets[merged.SheetNames[0]].A2.v,'홍길동');assert.equal(merged.Sheets[merged.SheetNames[0]].A4.v,'이영희');
    }
    await page.screenshot({path:'test-results/merge-preview.png',fullPage:true});
    await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);
    await page.locator('#header').uncheck();assert.equal(await page.locator('#download').isDisabled(),true);
    await page.locator('#clear').click();assert.equal(await page.locator('#merge').isDisabled(),true);
    await page.locator('#files').setInputFiles({name:'깨진.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from([0x50,0x4b,0x03,0x04,0x00])});await page.locator('#merge').click();await page.waitForFunction(()=>!document.getElementById('merge').disabled);assert.equal(await page.locator('#download').isDisabled(),true);
    assert.deepEqual(errors,[]);console.log('PASS: browser upload, multiple sheets, merged preview, CSV/XLSX downloads reopened, settings invalidate, reset, invalid file, mobile overflow, no runtime errors');
  }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exit(1)});

