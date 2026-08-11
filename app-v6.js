const BASE_MONTHS=[8,9,10,11,12,1,2];
const MONTHS=[...BASE_MONTHS];
const DEFAULTS=[
 {id:'major',name:'国有及主要股份行',keywords:'工商银行,建设银行,交通银行,邮储银行,邮政储蓄银行,国家开发银行,进出口银行,农业发展银行,广发银行,民生银行,兴业银行,光大银行,华夏银行,招商银行,浦发银行,上海浦东发展银行,平安银行',rates:{8:2.10,9:.60,10:.65,11:.65,12:.60,1:.57,2:.57},smallMin:1.90},
 {id:'big2',name:'中国银行、农业银行',keywords:'中国银行,农业银行',rates:{8:2.15,9:.65,10:.70,11:.70,12:.65,1:.60,2:.60}},
 {id:'city',name:'重点城商行',keywords:'江苏银行,北京银行,上海银行,宁波银行,浙商银行,南京银行',rates:{8:0,9:0,10:0,11:0,12:0,1:.60,2:.60}},
 {id:'other',name:'其他类银票',keywords:'',rates:{8:2.25,9:.75,10:.80,11:.80,12:.75,1:.70,2:.70}}
];
let rates=loadRates(),smallRate=loadSmallRate(),rows=[];
const $=id=>document.getElementById(id);
const money=new Intl.NumberFormat('zh-CN',{maximumFractionDigits:2});
function today(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function loadRates(){try{return JSON.parse(localStorage.getItem('quote-rate-table-v1'))||structuredClone(DEFAULTS)}catch{return structuredClone(DEFAULTS)}}
function saveRates(){localStorage.setItem('quote-rate-table-v1',JSON.stringify(rates))}
function loadSmallRate(){try{const saved=localStorage.getItem('quote-small-rate-v1');return saved===null?1.90:JSON.parse(saved)}catch{return 1.90}}
function saveSmallRate(){localStorage.setItem('quote-small-rate-v1',JSON.stringify(Number.isFinite(smallRate)?smallRate:null))}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function formatDate(d){return d&&!isNaN(d)?d.toLocaleDateString('zh-CN',{year:'numeric',month:'2-digit',day:'2-digit'}):'—'}
function excelDate(v){if(v instanceof Date)return new Date(v.getFullYear(),v.getMonth(),v.getDate());if(typeof v==='number'){const d=new Date(Math.round((v-25569)*86400000));return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())}if(typeof v==='string'&&v.trim()){const d=new Date(v);if(!isNaN(d))return new Date(d.getFullYear(),d.getMonth(),d.getDate())}return null}
function cellText(cell){const v=cell?.value;if(v==null)return '';try{return String(cell.text??'')}catch{if(typeof v==='object'&&v.result!=null)return String(v.result);if(typeof v==='object'&&Array.isArray(v.richText))return v.richText.map(x=>x.text||'').join('');return typeof v==='object'?'':String(v)}}
function isBlankValue(v){return v==null||(typeof v==='string'&&v.trim()==='')||(typeof v==='object'&&v.result==null&&!Array.isArray(v.richText))}
function numericCellValue(value){
 const raw=typeof value==='object'&&value?.result!=null?value.result:value;
 if(typeof raw==='number')return raw;
 if(raw==null||String(raw).trim()==='')return NaN;
 return Number(String(raw).replace(/[,，￥¥\s]/g,''));
}
function isSummaryLabel(value){
 const label=String(value??'').replace(/\s/g,'').replace(/[.．。、:：]/g,'');
 return /^(合计|总计|小计|汇总|总持有)/.test(label)||/^(持有银行承兑汇票合计)$/.test(label);
}
function isInputSummary(row){return isSummaryLabel(row?.bank)}
function isSummarySheetRow(row){
 let found=false;
 row.eachCell({includeEmpty:false},cell=>{if(isSummaryLabel(cellText(cell)))found=true});
 return found;
}
function normalizedHeader(value){return String(value??'').replace(/[\s（）()【】\[\]：:]/g,'')}
function headerColumnType(value){
 const text=normalizedHeader(value);
 if(/^(金额|票面金额|票据金额|出票金额|汇票金额|票据面额|票面余额)$/.test(text))return 'amount';
 if(/到期日|到期日期|票据到期|汇票到期|到期时间/.test(text))return 'maturity';
 if(/银行名称|承兑行|付款行|付款银行|承兑银行|出票行|承兑人|付款人|金融机构/.test(text))return 'bank';
 return null;
}
function findHeaderColumns(ws){
 for(let r=1;r<=Math.min(20,ws.rowCount);r++){
  const cols={};
  ws.getRow(r).eachCell((cell,c)=>{const type=headerColumnType(cellText(cell));if(type&&!cols[type])cols[type]=c});
  if(cols.amount&&cols.maturity)return{header:r,cols};
 }
 return null;
}
function bankContentScore(value){
 const text=String(value??'').replace(/\s/g,'');
 if(!text)return 0;
 if(/银行|农商行|信用社|农信|财务公司|村镇银行|金融服务中心|结算中心/.test(text))return 3;
 if(/^(中信|招行|工行|建行|农行|中行|交行|邮储|浦发|光大|民生|兴业|华夏|广发|平安|浙商)/.test(text))return 2;
 return 0;
}
function inferBankColumn(ws,header,excluded=[]){
 const excludedSet=new Set(excluded.filter(Boolean)),end=Math.min(ws.rowCount,header+80),maxCol=Math.min(ws.columnCount||60,80);
 let best=null;
 for(let c=1;c<=maxCol;c++){
  if(excludedSet.has(c))continue;
  let nonBlank=0,hits=0,companyOnly=0;
  for(let r=header+1;r<=end;r++){
   const text=cellText(ws.getCell(r,c)).trim();
   if(!text||isSummaryLabel(text))continue;
   nonBlank++;
   const contentScore=bankContentScore(text);
   if(contentScore)hits+=contentScore;
   else if(/有限公司|有限责任公司|公司$/.test(text))companyOnly++;
  }
  if(!nonBlank)continue;
  const headerText=cellText(ws.getCell(header,c)),headerHint=/银行|承兑|付款|金融|机构/.test(headerText)?6:0;
  const hitRows=Math.ceil(hits/3),ratio=hitRows/nonBlank,score=hits*8+ratio*20+headerHint-companyOnly*2;
  if(hitRows>=Math.min(2,nonBlank)&&ratio>=.35&&(!best||score>best.score))best={column:c,score};
 }
 return best?.column||null;
}
function syncDetectedMonths(source){
 const ordered=[...source].filter(r=>r.maturity instanceof Date&&!isNaN(r.maturity)).sort((a,b)=>a.maturity-b.maturity).map(r=>r.maturity.getMonth()+1);
 const detected=[...new Set(ordered)],next=[...BASE_MONTHS];
 detected.forEach(month=>{if(!next.includes(month))next.push(month)});
 const before=MONTHS.join(','),added=next.filter(month=>!MONTHS.includes(month));
 MONTHS.splice(0,MONTHS.length,...next);
 return{changed:before!==MONTHS.join(','),added};
}
function classify(bank){for(const r of rates.filter(x=>x.id!=='other')){if(r.keywords.split(/[,，、\n]/).map(x=>x.trim()).filter(Boolean).some(k=>bank.includes(k)))return r}return rates.find(x=>x.id==='other')||rates[rates.length-1]}
function renderRates(){
 $('rateHead').innerHTML=`<tr><th>承兑行分类</th>${MONTHS.map(m=>`<th>${m}月</th>`).join('')}<th>银行匹配关键词</th></tr>`;
 $('rateBody').innerHTML=rates.map(r=>`<tr data-id="${r.id}"><td><strong>${esc(r.name)}</strong></td>${MONTHS.map(m=>`<td><input data-month="${m}" type="number" step="0.01" value="${r.rates[m]??''}" aria-label="${esc(r.name)}${m}月大票利率"></td>`).join('')}<td><textarea data-keywords aria-label="${esc(r.name)}匹配关键词" placeholder="${r.id==='other'?'默认分类，无需关键词':'用逗号分隔'}">${esc(r.keywords)}</textarea></td></tr>`).join('');
 document.querySelectorAll('[data-month]').forEach(el=>el.addEventListener('change',e=>{const id=e.target.closest('tr').dataset.id;rates.find(r=>r.id===id).rates[e.target.dataset.month]=Number(e.target.value);saveRates();recalc()}));
 document.querySelectorAll('[data-keywords]').forEach(el=>el.addEventListener('change',e=>{const id=e.target.closest('tr').dataset.id;rates.find(r=>r.id===id).keywords=e.target.value;saveRates();recalc()}));
}
function renderSmallRate(){$('smallRate').value=Number.isFinite(smallRate)?smallRate:''}
function calculate(source){const discount=new Date($('discountDate').value+'T00:00:00');const cleanSource=source.filter(row=>!isInputSummary(row)).map((row,index)=>({...row,index:index+1}));return cleanSource.map(row=>{const cat=classify(row.bank);const days=row.maturity?Math.round((row.maturity-discount)/86400000):null;const isSmall=Number.isFinite(row.amount)&&row.amount<1000000;const configured=isSmall?smallRate:(row.maturity?cat.rates[row.maturity.getMonth()+1]:undefined);const rate=typeof configured==='number'&&Number.isFinite(configured)?configured:null;const status=!row.maturity||!row.bank||!Number.isFinite(row.amount)||row.amount<=0||days===null||days<0?'error':rate===null||rate===0?'warn':'ok';return {...row,discount,days,rate,category:cat.name,sizeTier:isSmall?'100万以下':'100万以上',status}})}
function recalc(){if(!rows.length)return;rows=calculate(rows);renderResults()}
function summary(){const valid=rows.filter(r=>r.status==='ok'&&r.days!==null&&r.rate!==null);const amount=valid.reduce((s,r)=>s+r.amount,0);const amountDays=valid.reduce((s,r)=>s+r.amount*r.days,0);return{valid:valid.length,issues:rows.length-valid.length,amount,days:amount?amountDays/amount:0,rate:amountDays?valid.reduce((s,r)=>s+r.amount*r.days*r.rate,0)/amountDays:0}}
function renderResults(){const s=summary();$('kpiDays').textContent=Math.round(s.days).toLocaleString('zh-CN');$('kpiAmount').textContent=money.format(s.amount);$('kpiRate').textContent=s.rate.toFixed(2);$('resultMeta').innerHTML=`<span>${s.valid} 条参与汇总</span>${s.issues?`<span class="issue">● ${s.issues} 条待检查</span>`:''}`;$('exportBtn').disabled=false;
 $('resultBody').innerHTML=rows.slice(0,150).map(r=>`<tr class="${r.status!=='ok'?'row-issue':''}"><td>${r.index}</td><td>${formatDate(r.maturity)}</td><td>${formatDate(r.discount)}</td><td>${r.days??'—'}</td><td>${r.rate===null?'未配置':r.rate.toFixed(2)+'%'}</td><td><span>${esc(r.bank||'未识别')}</span><small>${esc(r.category)} · ${esc(r.sizeTier)}</small></td><td>${Number.isFinite(r.amount)?money.format(r.amount):'—'}</td><td><b class="badge ${r.status}">${r.status==='ok'?'正常':r.status==='warn'?'待确认':'异常'}</b></td></tr>`).join('');
 $('tableNote').textContent=rows.length>150?`网页仅预览前 150 条，导出的 Excel 包含全部 ${rows.length} 条。`:'';
}
async function readFile(file){
 if(!/\.xlsx?$/i.test(file.name)){setMessage('文件格式不支持，请上传 .xlsx 或 .xls 文件');return}
 setBusy(true,'正在读取并核对列名…');
 try{
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  const ws=wb.worksheets.find(s=>!/^sheet[12]$/i.test(s.name))||wb.worksheets[0];
  if(!ws)throw Error('工作簿中没有可读取的工作表');
  const detected=findHeaderColumns(ws);
  if(!detected)throw Error('没有自动识别出金额列或到期日列，请检查表头和数据格式');
  const header=detected.header,cols=detected.cols;
  if(!cols.bank)cols.bank=inferBankColumn(ws,header,[cols.amount,cols.maturity]);
  if(!cols.bank)throw Error('没有自动识别出银行列，请确保该列下方填写了银行、农商行、信用社或财务公司名称');
  const parsed=[];
  for(let r=header+1;r<=ws.rowCount;r++){
   const sourceRow=ws.getRow(r);
   if(isSummarySheetRow(sourceRow))continue;
   const av=ws.getCell(r,cols.amount).value,mv=ws.getCell(r,cols.maturity).value,bank=cellText(ws.getCell(r,cols.bank)).trim();
   const requiredFieldCount=[!isBlankValue(av),!isBlankValue(mv),Boolean(bank)].filter(Boolean).length;
   if(requiredFieldCount<=1)continue;
   const amount=numericCellValue(av);
   parsed.push({index:parsed.length+1,maturity:excelDate(mv),bank,amount});
  }
  rows=calculate(parsed);
  $('fileTitle').textContent=file.name;
  const issues=rows.filter(r=>r.status!=='ok').length;
  const bankHeader=cellText(ws.getCell(header,cols.bank)).trim()||'内容识别列';
  setMessage(`已读取 ${rows.length} 条票据，银行列识别为“${bankHeader}”${issues?`；其中 ${issues} 条需要检查`:'；全部匹配成功'}`);
  renderResults();
 }catch(e){
  rows=[];
  $('exportBtn').disabled=true;
  setMessage(e.message||'文件读取失败，请检查格式');
 }finally{setBusy(false)}
}
function setMessage(t){$('message').textContent=t}function setBusy(b,t){$('dropzone').classList.toggle('busy',b);if(t)setMessage(t)}
async function exportExcel(){if(!rows.length)return;const s=summary(),wb=new ExcelJS.Workbook();wb.creator='票据报价测算工具';const rs=wb.addWorksheet('利率维护');rs.addRow(['承兑行分类',...MONTHS.map(m=>m+'月'),'匹配关键词','100万以下参考下限']);rates.forEach(r=>rs.addRow([r.name,...MONTHS.map(m=>r.rates[m]??null),r.keywords,r.smallMin??null]));rs.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};rs.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFC8102E'}};rs.columns=[{width:25},...MONTHS.map(()=>({width:11})),{width:70},{width:20}];for(let c=2;c<=MONTHS.length+1;c++)rs.getColumn(c).numFmt='0.00';const ws=wb.addWorksheet('报价结果',{views:[{state:'frozen',ySplit:1}]});ws.addRow(['到期日','贴现日','计息天数','分行报价贴现利率','承兑行','金额','匹配分类','检查状态']);rows.forEach(r=>ws.addRow([r.maturity,r.discount,r.days,r.rate,r.bank,r.amount,r.category,r.status==='ok'?'正常':r.status==='warn'?'待确认':'异常']));ws.columns=[{width:14},{width:14},{width:12},{width:20},{width:48},{width:18},{width:24},{width:12}];ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFC8102E'}};ws.getColumn(1).numFmt='yyyy/m/d';ws.getColumn(2).numFmt='yyyy/m/d';ws.getColumn(4).numFmt='0.00';ws.getColumn(6).numFmt='#,##0.00';ws.autoFilter={from:'A1',to:'H1'};ws.addRow([]);const start=ws.rowCount+1;[['加权平均天数',s.days],['合计金额',s.amount],['加权利率',s.rate]].forEach(x=>ws.addRow([null,null,null,null,null,x[0],x[1]]));for(let r=start;r<start+3;r++)for(let c=6;c<=7;c++){ws.getCell(r,c).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFE36E'}};ws.getCell(r,c).font={bold:true}}ws.getCell(start,7).numFmt='0';ws.getCell(start+1,7).numFmt='#,##0.00';ws.getCell(start+2,7).numFmt='0.00';const blob=new Blob([await wb.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`报价测算结果_${$('discountDate').value}.xlsx`;a.click();URL.revokeObjectURL(url)}
$('discountDate').value=today();renderRates();$('discountDate').addEventListener('change',recalc);$('resetRates').addEventListener('click',()=>{rates=structuredClone(DEFAULTS);saveRates();renderRates();recalc()});$('exportBtn').addEventListener('click',exportExcel);const dz=$('dropzone'),fi=$('fileInput');dz.addEventListener('click',()=>fi.click());dz.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' ')fi.click()});dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('dragging')});dz.addEventListener('dragleave',()=>dz.classList.remove('dragging'));dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('dragging');if(e.dataTransfer.files[0])readFile(e.dataTransfer.files[0])});fi.addEventListener('change',()=>{if(fi.files[0])readFile(fi.files[0]);fi.value=''})

let lastMaturityData=[];
function getMonthlyMaturityData(){
 const grouped=new Map();
 rows.forEach(row=>{
  if(!(row.maturity instanceof Date)||isNaN(row.maturity)||!Number.isFinite(row.amount)||row.amount<=0)return;
  const year=row.maturity.getFullYear(),month=row.maturity.getMonth()+1;
  const key=`${year}-${String(month).padStart(2,'0')}`;
  if(!grouped.has(key))grouped.set(key,{key,label:`${year}年${month}月`,large:0,small:0,total:0,amount:0});
  const item=grouped.get(key);
  if(row.amount>=1000000)item.large++;else item.small++;
  item.total++;
  item.amount+=row.amount;
 });
 return [...grouped.values()].sort((a,b)=>a.key.localeCompare(b.key));
}
function renderMaturityAnalysis(){
 lastMaturityData=getMonthlyMaturityData();
 const total=lastMaturityData.reduce((sum,item)=>sum+item.total,0),totalAmount=lastMaturityData.reduce((sum,item)=>sum+item.amount,0);
 $('maturityTotal').textContent=`共 ${total} 张 · ${money.format(totalAmount)} 元`;
 $('maturityBody').innerHTML=lastMaturityData.length?lastMaturityData.map(item=>{const share=totalAmount?item.amount/totalAmount:0;return `<tr><td>${item.label}</td><td class="large-count">${item.large}</td><td class="small-count">${item.small}</td><td><strong>${item.total}</strong></td><td class="amount-count">${money.format(item.amount)}</td><td class="share-count">${(share*100).toFixed(2)}%</td></tr>`}).join(''):'<tr><td colspan="6" class="empty">上传文件后自动生成月度统计</td></tr>';
 drawMaturityChart(lastMaturityData);
 drawAmountShareChart(lastMaturityData);
}
function drawMaturityChart(data){
 const canvas=$('maturityChart');
 if(!canvas)return;
 const width=Math.max(320,canvas.parentElement.clientWidth),height=360,dpr=Math.max(1,window.devicePixelRatio||1);
 canvas.width=Math.round(width*dpr);canvas.height=Math.round(height*dpr);canvas.style.height=height+'px';
 const ctx=canvas.getContext('2d');ctx.scale(dpr,dpr);ctx.clearRect(0,0,width,height);
 ctx.font='12px "Microsoft YaHei",sans-serif';ctx.fillStyle='#746f64';
 if(!data.length){ctx.textAlign='center';ctx.fillText('上传文件后自动生成柱形图',width/2,height/2);return}
 const margin={top:54,right:20,bottom:72,left:48},chartW=width-margin.left-margin.right,chartH=height-margin.top-margin.bottom;
 const maxValue=Math.max(...data.flatMap(item=>[item.large,item.small]),1),axisMax=Math.max(5,Math.ceil(maxValue/5)*5),ticks=5;
 ctx.strokeStyle='#ded8cb';ctx.lineWidth=1;ctx.textAlign='right';ctx.textBaseline='middle';
 for(let i=0;i<=ticks;i++){const value=Math.round(axisMax*i/ticks),y=margin.top+chartH-chartH*i/ticks;ctx.beginPath();ctx.moveTo(margin.left,y);ctx.lineTo(width-margin.right,y);ctx.stroke();ctx.fillStyle='#746f64';ctx.fillText(String(value),margin.left-9,y)}
 const groupW=chartW/data.length,barW=Math.max(8,Math.min(30,(groupW-14)/2));
 data.forEach((item,index)=>{
  const center=margin.left+groupW*(index+.5),bars=[{value:item.large,x:center-barW-2,color:'#173f35'},{value:item.small,x:center+2,color:'#e5683a'}];
  bars.forEach(bar=>{const h=bar.value/axisMax*chartH,y=margin.top+chartH-h;ctx.fillStyle=bar.color;ctx.fillRect(bar.x,y,barW,h);if(bar.value){ctx.fillStyle='#39443f';ctx.textAlign='center';ctx.textBaseline='bottom';ctx.fillText(String(bar.value),bar.x+barW/2,y-4)}});
  ctx.save();ctx.translate(center,margin.top+chartH+12);if(data.length>8)ctx.rotate(-Math.PI/4);ctx.fillStyle='#5f5a50';ctx.textAlign=data.length>8?'right':'center';ctx.textBaseline='top';ctx.fillText(item.label,0,0);ctx.restore();
 });
 ctx.textBaseline='middle';ctx.textAlign='left';ctx.fillStyle='#173f35';ctx.fillRect(margin.left,18,12,12);ctx.fillStyle='#39443f';ctx.fillText('100万以上',margin.left+18,24);ctx.fillStyle='#e5683a';ctx.fillRect(margin.left+112,18,12,12);ctx.fillStyle='#39443f';ctx.fillText('100万以下',margin.left+130,24);
}
const AMOUNT_SHARE_COLORS=['#173f35','#e5683a','#e1b44b','#244f43','#305f50','#d98a58','#8fa99d','#b78b32','#746f64','#5d806f','#ef9a73','#c7aa62'];
function drawAmountShareChart(data){
 const canvas=$('amountShareChart');
 if(!canvas)return;
 const width=Math.max(280,canvas.parentElement.clientWidth),columns=width<360?1:2,height=Math.max(360,280+Math.ceil(data.length/columns)*23),dpr=Math.max(1,window.devicePixelRatio||1);
 canvas.width=Math.round(width*dpr);canvas.height=Math.round(height*dpr);canvas.style.height=height+'px';
 const ctx=canvas.getContext('2d');ctx.scale(dpr,dpr);ctx.clearRect(0,0,width,height);ctx.font='11px "Microsoft YaHei",sans-serif';
 const totalAmount=data.reduce((sum,item)=>sum+item.amount,0);
 if(!data.length||!totalAmount){ctx.fillStyle='#7b6f70';ctx.textAlign='center';ctx.fillText('上传文件后自动生成金额占比饼图',width/2,height/2);return}
 const radius=Math.min(108,width*.3),cx=width/2,cy=132;
 let angle=-Math.PI/2;
 data.forEach((item,index)=>{
  const share=item.amount/totalAmount,next=angle+share*Math.PI*2,color=AMOUNT_SHARE_COLORS[index%AMOUNT_SHARE_COLORS.length];
  ctx.beginPath();ctx.moveTo(cx,cy);ctx.arc(cx,cy,radius,angle,next);ctx.closePath();ctx.fillStyle=color;ctx.fill();ctx.strokeStyle='#fffaf8';ctx.lineWidth=2;ctx.stroke();
  if(share>=.075){const mid=(angle+next)/2;ctx.fillStyle='#fff';ctx.font='700 11px "Microsoft YaHei",sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(`${(share*100).toFixed(1)}%`,cx+Math.cos(mid)*radius*.62,cy+Math.sin(mid)*radius*.62)}
  angle=next;
 });
 const columnWidth=width/columns,legendTop=cy+radius+24;
 data.forEach((item,index)=>{
  const column=index%columns,row=Math.floor(index/columns),x=column*columnWidth+18,y=legendTop+row*23,share=item.amount/totalAmount;
  ctx.fillStyle=AMOUNT_SHARE_COLORS[index%AMOUNT_SHARE_COLORS.length];ctx.fillRect(x,y-6,10,10);ctx.fillStyle='#4c3b3e';ctx.font='11px "Microsoft YaHei",sans-serif';ctx.textAlign='left';ctx.textBaseline='middle';ctx.fillText(`${item.label}  ${(share*100).toFixed(2)}%`,x+16,y-1);
 });
}
const baseRenderResults=renderResults;
renderResults=function(){baseRenderResults();renderMaturityAnalysis()};
const baseReadFile=readFile;
readFile=async function(file){
 await baseReadFile(file);
 if(!rows.length){renderMaturityAnalysis();showPanel('upload');return}
 const monthSync=syncDetectedMonths(rows);
 if(monthSync.changed){renderRates();recalc()}
 if(monthSync.added.length){setMessage(`${$('message').textContent}；已新增 ${monthSync.added.map(m=>m+'月').join('、')} 利率列，请填写后确认`);showPanel('rates')}
 else showPanel('results');
};
let chartResizeTimer;
window.addEventListener('resize',()=>{clearTimeout(chartResizeTimer);chartResizeTimer=setTimeout(()=>{drawMaturityChart(lastMaturityData);drawAmountShareChart(lastMaturityData)},120)});
renderMaturityAnalysis();

async function exportExcelV4(){
 if(!rows.length)return;
 const s=summary(),monthly=getMonthlyMaturityData(),wb=new ExcelJS.Workbook();
 wb.creator='票据报价测算工具';

 const rs=wb.addWorksheet('利率维护',{views:[{state:'frozen',ySplit:3}]});
 rs.addRow(['单张票面100万以下统一利率',smallRate]);
 rs.addRow([]);
 rs.addRow(['承兑行分类（单张100万以上）',...MONTHS.map(m=>m+'月'),'匹配关键词']);
 rates.forEach(r=>rs.addRow([r.name,...MONTHS.map(m=>r.rates[m]??null),r.keywords]));
 rs.columns=[{width:30},...MONTHS.map(()=>({width:11})),{width:70}];
 rs.getRow(3).font={bold:true,color:{argb:'FFFFFFFF'}};
 rs.getRow(3).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFC8102E'}};
 rs.getCell(1,1).font={bold:true};rs.getCell(1,1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFE36E'}};
 rs.getCell(1,2).font={bold:true};rs.getCell(1,2).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFE36E'}};
 for(let c=2;c<=MONTHS.length+1;c++)rs.getColumn(c).numFmt='0.00';

 const ws=wb.addWorksheet('报价结果',{views:[{state:'frozen',ySplit:1}]});
 ws.addRow(['到期日','贴现日','计息天数','分行报价贴现利率','承兑行','金额','金额档位','匹配分类','检查状态']);
 rows.forEach(r=>ws.addRow([r.maturity,r.discount,r.days,r.rate,r.bank,r.amount,r.sizeTier,r.category,r.status==='ok'?'正常':r.status==='warn'?'待确认':'异常']));
 ws.columns=[{width:14},{width:14},{width:12},{width:20},{width:48},{width:18},{width:14},{width:24},{width:12}];
 ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFC8102E'}};
 ws.getColumn(1).numFmt='yyyy/m/d';ws.getColumn(2).numFmt='yyyy/m/d';ws.getColumn(4).numFmt='0.00';ws.getColumn(6).numFmt='#,##0.00';ws.autoFilter={from:'A1',to:'I1'};
 ws.addRow([]);
 const start=ws.rowCount+1;
 [['加权平均天数',s.days],['合计金额',s.amount],['加权利率',s.rate]].forEach(x=>ws.addRow([null,null,null,null,null,x[0],x[1]]));
 for(let r=start;r<start+3;r++)for(let c=6;c<=7;c++){ws.getCell(r,c).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFE36E'}};ws.getCell(r,c).font={bold:true}}
 ws.getCell(start,7).numFmt='0';ws.getCell(start+1,7).numFmt='#,##0.00';ws.getCell(start+2,7).numFmt='0.00';

 const ms=wb.addWorksheet('月度到期统计',{views:[{state:'frozen',ySplit:1}]});
 const monthlyAmountTotal=monthly.reduce((n,x)=>n+x.amount,0);
 ms.addRow(['到期月份','100万以上张数','100万以下张数','合计张数','到期金额总计','各月合计金额占比']);
 monthly.forEach(item=>ms.addRow([item.label,item.large,item.small,item.total,item.amount,monthlyAmountTotal?item.amount/monthlyAmountTotal:0]));
 const totalRow=ms.addRow(['合计',monthly.reduce((n,x)=>n+x.large,0),monthly.reduce((n,x)=>n+x.small,0),monthly.reduce((n,x)=>n+x.total,0),monthlyAmountTotal,monthlyAmountTotal?1:0]);
 ms.columns=[{width:18},{width:18},{width:18},{width:14},{width:22},{width:22}];
 ms.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};ms.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFC8102E'}};
 totalRow.font={bold:true};totalRow.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFE36E'}};
 ms.getColumn(5).numFmt='#,##0.00';ms.getColumn(6).numFmt='0.00%';ms.autoFilter={from:'A1',to:'F1'};

 const blob=new Blob([await wb.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download=`报价测算结果_${$('discountDate').value}.xlsx`;a.click();URL.revokeObjectURL(url);
}

renderSmallRate();
$('smallRate').addEventListener('change',e=>{const value=e.target.value.trim();smallRate=value===''?null:Number(value);if(!Number.isFinite(smallRate))smallRate=null;saveSmallRate();recalc()});
$('resetRates').addEventListener('click',()=>{smallRate=1.90;saveSmallRate();renderSmallRate();recalc()});
const currentExportButton=$('exportBtn'),cleanExportButton=currentExportButton.cloneNode(true);
currentExportButton.replaceWith(cleanExportButton);
cleanExportButton.addEventListener('click',exportExcelV4);

const PANEL_INFO={
 upload:{title:'上传票据数据',hint:'选择Excel文件开始测算'},
 rates:{title:'利率维护',hint:'维护大票月度利率和100万以下小票利率'},
 results:{title:'测算结果',hint:'查看加权指标、报价明细并导出Excel'},
 maturity:{title:'月度到期统计',hint:'查看各月到期张数、金额总计与占比'}
};
function showPanel(name,shouldScroll=true){
 const target=PANEL_INFO[name]?name:'upload';
 document.querySelectorAll('.tool-panel').forEach(panel=>{const active=panel.dataset.panel===target;panel.classList.toggle('is-active',active);panel.hidden=!active});
 $('workspaceSelect').value=target;$('panelTitle').textContent=PANEL_INFO[target].title;$('panelHint').textContent=PANEL_INFO[target].hint;
 if(target==='maturity')requestAnimationFrame(()=>{drawMaturityChart(lastMaturityData);drawAmountShareChart(lastMaturityData)});
 if(shouldScroll)document.querySelector('.workspace-switcher').scrollIntoView({block:'start'});
}
$('workspaceSelect').addEventListener('change',e=>showPanel(e.target.value));
showPanel('upload',false);
