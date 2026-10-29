/* Editable Excel project format. All processing happens in the browser. */
(function (root) {
'use strict';
const graphHeaders=['Graph ID','Data sheet','Name','Title','Categories and series','Pivoted','X-axis name','Other X-axis name','Y-axis name','Y minimum','Y maximum','Lower limit','Upper limit','Standard deviation','Value labels','Overall average','Individual measurements','Width','Palette','Custom appearance','Colors JSON','Pivot colors JSON'];
const paletteNames={palette1:'NVIDIA',palette2:'Matplotlib default',palette3:'MATLAB default'};
const header=['Category','Series','Input type','Mean','SD','n'];
const limitCells=2000000;
// Small grid adapter over the bundled ExcelJS browser build.
function encode_col(n){let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s}
function encode_cell({r,c}){return encode_col(c)+(r+1)}
function decode_cell(address){const m=/^([A-Z]+)(\d+)$/.exec(address);if(!m)throw Error('Invalid cell address.');let c=0;for(const ch of m[1])c=c*26+ch.charCodeAt(0)-64;return {r:Number(m[2])-1,c:c-1}}
function decode_range(ref){const [a,b=a]=ref.split(':');return {s:decode_cell(a),e:decode_cell(b)}}
const grid={utils:{encode_col,encode_cell,decode_range,
 book_new:()=>({SheetNames:[],Sheets:{}}),
 book_append_sheet:(wb,ws,name)=>{if(wb.Sheets[name])throw Error('Duplicate worksheet name.');wb.SheetNames.push(name);wb.Sheets[name]=ws},
 aoa_to_sheet:rows=>{const ws={};let max=0;rows.forEach((row,r)=>{max=Math.max(max,row.length);row.forEach((v,c)=>{if(v!==null&&v!==undefined)ws[encode_cell({r,c})]={t:typeof v==='number'?'n':typeof v==='boolean'?'b':'s',v}})});ws['!ref']=`A1:${encode_cell({r:Math.max(0,rows.length-1),c:Math.max(0,max-1)})}`;return ws},
 sheet_to_json:ws=>{const range=decode_range(ws['!ref']||'A1'),rows=[];for(let r=0;r<=range.e.r;r++){const row=[];for(let c=0;c<=range.e.c;c++)row.push(ws[encode_cell({r,c})]?.v??'');rows.push(row)}return rows}
},
 async write(book){
  if(!root.ExcelJS)throw Error('The Excel library did not load. Refresh and try again.');
  const wb=new root.ExcelJS.Workbook();wb.creator='Bar-Tender';wb.title='Bar-Tender project';wb.calcProperties.fullCalcOnLoad=true;
  for(const name of book.SheetNames){
   const source=book.Sheets[name],ws=wb.addWorksheet(name,{views:[{state:'frozen',ySplit:1}]});
   (source['!cols']||[]).forEach((col,i)=>{ws.getColumn(i+1).width=col.wch});
   for(const [address,cell] of Object.entries(source)){
    if(address.startsWith('!'))continue;
    if(typeof cell.v==='string'&&cell.v.length>32767)throw Error(`Cell ${name}!${address} exceeds Excel's text limit.`);
    const target=ws.getCell(address);target.value=cell.f?{formula:cell.f,result:cell.v}:cell.v;
    if(cell.z)target.numFmt=cell.z;
    if(cell.f){target.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF1F3F5'}};target.font={color:{argb:'FF56616D'}}}
   }
   ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF16324F'}};ws.getRow(1).height=24;
   if(source['!autofilter'])ws.autoFilter=source['!autofilter'].ref;
   if(name==='Appearance'){ws.getColumn(3).alignment={wrapText:true,vertical:'top'};ws.getColumn(4).alignment={wrapText:true,vertical:'top'}}
  }
  return wb.xlsx.writeBuffer();
 },
 async read(bytes){
  if(!root.ExcelJS)throw Error('The Excel library did not load. Refresh and try again.');
  const input=new root.ExcelJS.Workbook();await input.xlsx.load(bytes);const book={SheetNames:[],Sheets:{}};let total=0;
  for(const ws of input.worksheets){
   if(ws.rowCount>5003||ws.columnCount>10007)throw Error(`Sheet "${ws.name}" is too large.`);
   total+=ws.rowCount*ws.columnCount;if(total>limitCells)throw Error('The workbook is too large to open safely in this browser.');
   const output={'!ref':`A1:${encode_cell({r:Math.max(0,ws.rowCount-1),c:Math.max(0,ws.columnCount-1)})}`};
   ws.eachRow((row,r)=>row.eachCell((cell,c)=>{
    let v=cell.value;if(v===null)return;const entry={};
    if(v&&typeof v==='object'&&('formula' in v||'sharedFormula' in v)){entry.f=cell.formula;v=v.result}
    if(v&&typeof v==='object'){
     if(v.richText)v=v.richText.map(part=>part.text).join('');else if(v.error){entry.t='e';v=v.error}else if(v.text!==undefined)v=v.text;else throw Error(`Unsupported cell value at ${ws.name}!${cell.address}.`);
    }
    entry.v=v;entry.t=entry.t|| (typeof v==='number'?'n':typeof v==='boolean'?'b':'s');output[encode_cell({r:r-1,c:c-1})]=entry;
   }));book.SheetNames.push(ws.name);book.Sheets[ws.name]=output;
  }
  return book;
 }
};

const X=()=>grid;
function text(v){return v===null||v===undefined?'':String(v)}
function numberCell(v){return text(v).trim()!==''&&Number.isFinite(Number(v))?Number(v):text(v)}
function bool(v,label){if(typeof v==='boolean')return v;const s=text(v).trim().toLowerCase();if(['true','yes','1'].includes(s))return true;if(['false','no','0',''].includes(s))return false;throw Error(`${label}: enter TRUE or FALSE.`)}
function chunks(s){const out=[];for(let i=0;i<s.length;){let end=Math.min(i+15000,s.length);if(end<s.length&&/[\uD800-\uDBFF]/.test(s[end-1]))end--;out.push(s.slice(i,end));i=end}return out}
function sheet(rows,widths){const ws=X().utils.aoa_to_sheet(rows);ws['!cols']=widths.map(wch=>({wch}));return ws}
async function write(project,active=0,combined=false){
 const lib=X(),wb=lib.utils.book_new();wb.Props={Title:'Bar-Tender project',Application:'Bar-Tender'};
 wb.Workbook={CalcPr:{calcMode:'auto'}};
 lib.utils.book_append_sheet(wb,sheet([
 ['Bar-Tender — editable Excel project'],
 ['Open this workbook in Bar-Tender using Open Excel to continue working.'],
 ['Each numbered data sheet holds one graph. Edit Category, Series and Input type as needed.'],
 ['Mean + SD rows: enter Mean and optional SD. Raw data rows: enter samples under Measurement 1, 2, …'],
 ['For raw data, Mean, SD and n are formulas for viewing in Excel. Bar-Tender recalculates from the measurements on import.'],
 ['You may add rows and Measurement N columns. Keep the column headers and named sheets.'],
 ['Graphs holds editable titles, axes, palettes, limits and view settings.'],
 ['Appearance holds applied and draft Python code, split into numbered parts. Keep Graph ID and Part intact.'],
 ['Workbook includes all open graphs. SVG export and Copy PNG remain separate image actions.'],
 ['Formula inputs: recalculate and save in Excel before importing; otherwise the cached value may be unavailable.'],
 ['For Mean + SD rows, stored measurement columns are kept for switching back to Raw data; they are not plotted in summary mode.']
 ],[115]),'Read me');
 lib.utils.book_append_sheet(wb,sheet([['Bar-Tender workbook',1],['Active graph',project.graphs.length?`g${active+1}`:''],['Combined view',!!combined]],[25,30]),'Workspace');
 const configs=[graphHeaders],scripts=[['Graph ID','Part','Applied script','Draft script']];
 project.graphs.forEach((g,i)=>{
  const id=`g${i+1}`,dataSheet=`${String(i+1).padStart(2,'0')} ${g.name.replace(/[\\/?*\[\]:]/g,' ').replace(/^'+|'+$/g,'').trim()||'Graph'}`.slice(0,31);
  configs.push([id,dataSheet,g.name,g.title,!!g.grouped,!!g.pivoted,g.xlabel,g.otherXLabel||'Series',g.ylabel,numberCell(g.ymin),numberCell(g.ymax),numberCell(g.lower),numberCell(g.upper),g.sd,g.labels,g.average,!!g.points,g.width,paletteNames[g.palette]||'NVIDIA',!!g.visualEnabled,JSON.stringify(g.colors||{}),JSON.stringify(g.pivotColors||{})]);
  const applied=chunks(g.visualScript||''),draft=chunks(g.visualDraft||'');
  for(let part=0;part<Math.max(1,applied.length,draft.length);part++)scripts.push([id,part+1,applied[part]||'',draft[part]||'']);
  const maxSamples=Math.max(1,...g.rows.map(r=>r.samples?.length||0));
  const rows=[[...header,...Array.from({length:maxSamples},(_,j)=>`Measurement ${j+1}`)]];
  for(const r of g.rows)rows.push([r.category,r.series,r.type==='raw'?'Raw data':'Mean + SD',r.type==='raw'?'':numberCell(r.value),r.type==='raw'?'':numberCell(r.sd),'',...(r.samples||[]).map(numberCell)]);
  const ws=sheet(rows,[23,23,17,15,15,10,...Array(maxSamples).fill(17)]);
  for(let row=0;row<g.rows.length;row++){
   const r=g.rows[row];if(r.type!=='raw')continue;
   const nrow=row+2,range=`G${nrow}:${lib.utils.encode_col(5+maxSamples)}${nrow}`;
   const samples=(r.samples||[]).filter(v=>text(v).trim()!==''&&Number.isFinite(Number(v))).map(Number);
   const n=samples.length,mean=n?samples.reduce((a,b)=>a+b,0)/n:0,sd=n>1?Math.sqrt(samples.reduce((sum,v)=>sum+(v-mean)**2,0)/(n-1)):0;
   ws[`D${nrow}`]={t:n?'n':'s',v:n?mean:'',f:`IF(COUNT(${range})>0,AVERAGE(${range}),"")`,z:'0.###'};
   ws[`E${nrow}`]={t:n>1?'n':'s',v:n>1?sd:'',f:`IF(COUNT(${range})>1,STDEV(${range}),"")`,z:'0.###'};
   ws[`F${nrow}`]={t:'n',v:n,f:`COUNT(${range})`};
  }
  ws['!autofilter']={ref:`A1:${lib.utils.encode_col(5+maxSamples)}${rows.length}`};
  lib.utils.book_append_sheet(wb,ws,dataSheet);
 });
 lib.utils.book_append_sheet(wb,sheet(configs,graphHeaders.map((h,i)=>i<4?25:20)),'Graphs');
 lib.utils.book_append_sheet(wb,sheet(scripts,[16,10,80,80]),'Appearance');
 return lib.write(wb,{bookType:'xlsx',type:'array',compression:true});
}
async function read(bytes){
 const lib=X(),wb=await lib.read(bytes,{type:'array',cellFormula:true,sheetRows:5003});let totalCells=0;
 function table(name,maxRows=5001){
  const ws=wb.Sheets[name];if(!ws)throw Error(`Missing sheet "${name}". Open a workbook exported by Bar-Tender.`);
  const ref=ws['!fullref']||ws['!ref']||'A1',range=lib.utils.decode_range(ref);
  if(range.e.r+1>maxRows||range.e.c>10006)throw Error(`Sheet "${name}" is too large.`);
  totalCells+=(range.e.r+1)*(range.e.c+1);if(totalCells>limitCells)throw Error('The workbook is too large to open safely in this browser.');
  // Reject formula cells without cached results when they supply source inputs.
  const values=lib.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:true});
  return {ws,values};
 }
 const workspace=table('Workspace',20).values;
 if(workspace[0]?.[0]!=='Bar-Tender workbook'||Number(workspace[0]?.[1])!==1)throw Error('Choose an Excel project exported by Bar-Tender.');
 const config=table('Graphs',51).values,heads=config[0]||[];
 for(const h of graphHeaders)if(!heads.includes(h))throw Error(`Graphs sheet is missing the "${h}" column.`);
 const get=(row,key)=>row[heads.indexOf(key)]??'';
 const configs=config.slice(1).filter(row=>row.some(v=>v!=='')),seen=new Set(),usedSheets=new Set();
 if(configs.length>50)throw Error('A project can contain at most 50 graphs.');
 const scripts=new Map(),appearance=table('Appearance',1001).values,scriptHeaders=appearance[0]||[];
 for(const h of ['Graph ID','Part','Applied script','Draft script'])if(!scriptHeaders.includes(h))throw Error(`Appearance sheet is missing "${h}".`);
 for(const line of appearance.slice(1)){
  if(!line.some(v=>v!==''))continue;
  const getScript=k=>line[scriptHeaders.indexOf(k)]??'',id=text(getScript('Graph ID')),part=Number(getScript('Part'));
  if(!Number.isInteger(part)||part<1||part>10)throw Error('Appearance script parts must be numbered 1–10.');
  if(!scripts.has(id))scripts.set(id,new Map());if(scripts.get(id).has(part))throw Error('Duplicate appearance script part.');
  scripts.get(id).set(part,[text(getScript('Applied script')),text(getScript('Draft script'))]);
 }
 const graphs=configs.map(row=>{
  const id=text(get(row,'Graph ID')),name=text(get(row,'Data sheet'));
  if(!id||seen.has(id)||usedSheets.has(name))throw Error('Graph IDs and data sheet names must be unique.');seen.add(id);usedSheets.add(name);
  const g={name:text(get(row,'Name')),title:text(get(row,'Title')),grouped:bool(get(row,'Categories and series'),'Categories and series'),pivoted:bool(get(row,'Pivoted'),'Pivoted'),xlabel:text(get(row,'X-axis name')),otherXLabel:text(get(row,'Other X-axis name')),ylabel:text(get(row,'Y-axis name')),ymin:text(get(row,'Y minimum')),ymax:text(get(row,'Y maximum')),lower:text(get(row,'Lower limit')),upper:text(get(row,'Upper limit')),sd:bool(get(row,'Standard deviation'),'Standard deviation'),labels:bool(get(row,'Value labels'),'Value labels'),average:bool(get(row,'Overall average'),'Overall average'),points:bool(get(row,'Individual measurements'),'Individual measurements'),width:text(get(row,'Width')),palette:Object.keys(paletteNames).find(k=>paletteNames[k]===get(row,'Palette'))||text(get(row,'Palette')),visualEnabled:bool(get(row,'Custom appearance'),'Custom appearance'),colors:JSON.parse(text(get(row,'Colors JSON'))||'{}'),pivotColors:JSON.parse(text(get(row,'Pivot colors JSON'))||'{}'),visualScript:'',visualDraft:'',rows:[]};
  if(!['wide','narrow'].includes(g.width)||!paletteNames[g.palette])throw Error(`${g.name}: invalid width or palette in Graphs sheet.`);
  const parts=scripts.get(id)||new Map(),ordered=[...parts.keys()].sort((a,b)=>a-b);
  if(ordered.some((part,i)=>part!==i+1))throw Error(`${g.name}: appearance script parts are missing.`);
  g.visualScript=ordered.map(part=>parts.get(part)[0]).join('');g.visualDraft=ordered.map(part=>parts.get(part)[1]).join('');
  if(g.visualEnabled&&!parts.size)throw Error(`${g.name}: missing appearance script.`);
  const {ws,values}=table(name),titles=values[0]||[];
  for(const h of header)if(!titles.includes(h))throw Error(`Sheet "${name}" is missing the "${h}" column.`);
  const sampleColumns=titles.map((h,i)=>({match:/^Measurement (\d+)$/.exec(text(h)),i})).filter(c=>c.match).sort((a,b)=>Number(a.match[1])-Number(b.match[1]));
  if(sampleColumns.some((c,i)=>Number(c.match[1])!==i+1))throw Error(`Sheet "${name}": measurement columns must be numbered from 1 without gaps.`);
  function source(line,ri,ci){const cell=ws[lib.utils.encode_cell({r:ri,c:ci})];if(cell?.f&&(cell.v===undefined||cell.v===null||cell.t==='e'))throw Error(`Sheet "${name}", row ${ri+1}: recalculate formulas and save in Excel before importing.`);if(cell?.t==='e')throw Error(`Sheet "${name}", row ${ri+1}: correct the Excel cell error first.`);return text(line[ci]??'')}
  values.slice(1).forEach((line,index)=>{
   if(!line.some(v=>v!==''))return;
   const ri=index+1,cell=h=>source(line,ri,titles.indexOf(h)),samples=sampleColumns.map(c=>source(line,ri,c.i));
   while(samples.length&&samples.at(-1)==='')samples.pop();samples.push('');
   const type=cell('Input type').trim().toLowerCase(),raw=type==='raw data'||type==='raw'||(!type&&samples.some(v=>v.trim()));
   if(type&&!['raw data','raw','mean + sd','summary'].includes(type))throw Error(`Sheet "${name}", row ${ri+1}: Input type must be Raw data or Mean + SD.`);
   const value=raw?samples.filter(v=>v.trim()).join(', '):cell('Mean'),sd=raw?'':cell('SD');
   g.rows.push({category:cell('Category'),series:cell('Series'),type:raw?'raw':'summary',samples,value,sd});
  });
  return g;
 });
 for(const id of scripts.keys())if(!seen.has(id))throw Error('Appearance sheet references an unknown Graph ID.');
 const activeId=text(workspace.find(row=>row[0]==='Active graph')?.[1]);
 return {project:{version:2,graphs},active:Math.max(0,configs.findIndex(row=>text(get(row,'Graph ID'))===activeId)),combined:bool(workspace.find(row=>row[0]==='Combined view')?.[1],'Combined view')};
}
root.BarTenderExcel={write,read};
})(typeof window!=='undefined'?window:globalThis);
