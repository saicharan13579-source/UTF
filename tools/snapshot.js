// Record today's full snapshot from the built site: node tools/snapshot.js [YYYY-MM-DD]
const {chromium}=require('playwright');const fs=require('fs');const path=require('path');
const R=path.resolve(__dirname,'..');
(async()=>{
  const now=new Date(Date.now()+5.5*3600e3);const date=process.argv[2]||now.toISOString().slice(0,10); // IST date
  const b=await chromium.launch();const p=await b.newPage();
  await p.goto('file://'+path.join(R,'index.html'));await p.waitForTimeout(1500);
  const snap=await p.evaluate(()=>({counts:{done:VC.done,wip:VC.wip,sold:VC.sold,nyt:VC.nyt,listed:VC.listed,unk:VC.unk},
    needs:PLOTS.filter(p=>p.group==='unk'||p.group==='listed').length,
    plots:PLOTS.map(p=>({plot:p.id,villa:p.entry||p.id,series:p.series,status:p.fromDrawing?'nyt':p.group,stage:p.group==='wip'?p.stage:'',customer:p.customer||'',contractor:p.contractor||'',area:p.area??'',sales:p.sales||'',remarks:p.remarks||'',edited:!!p.ov}))}));
  await b.close();
  fs.mkdirSync(path.join(R,'data/snapshots'),{recursive:true});
  fs.writeFileSync(path.join(R,'data/snapshots',date+'.json'),JSON.stringify({date,...snap},null,1));
  const edits=JSON.parse(fs.readFileSync(path.join(R,'data/edits.json')));
  const changes=edits.filter(e=>(e.at||'').slice(0,10)===date).length;
  const dfile=path.join(R,'data/daily.json');const daily=JSON.parse(fs.readFileSync(dfile)).filter(r=>r.date!==date);
  const row={date,done:snap.counts.done,wip:snap.counts.wip,sold:snap.counts.sold,nyt:snap.counts.nyt,needs:snap.needs,changes};
  daily.push(row);daily.sort((a,b)=>a.date.localeCompare(b.date));fs.writeFileSync(dfile,JSON.stringify(daily,null,1));
  console.log(JSON.stringify(row));
})();
