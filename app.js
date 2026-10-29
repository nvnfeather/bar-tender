const $=id=>document.getElementById(id);
const palettes={"palette1":["#77B900","#18BBA4","#0074DD","#962EC6","#EF9205","#D94F59","#294A75","#737373"],"palette2":["#1F77B4","#FF7F0E","#2CA02C","#D62728","#9467BD","#8C564B","#E377C2","#7F7F7F","#BCBD22","#17BECF"],"palette3":["#1171BE","#DD5400","#EDB120","#8516D1","#3BAA32","#2FBEEF","#D1048B"]};
let project={version:2,graphs:[]};
let active=0,combined=false; const fields=['title','xlabel','ylabel','ymin','ymax','lower','upper','sd','labels','average','points'];
const WORKSPACE_KEY='bar-tender.workspace.v1';
let lastSavedWorkspace=null,autosaveBlocked=false,autosaveWarningShown=false;
function saveWorkspace(){
 if(autosaveBlocked)return;
 try{
  const snapshot=JSON.stringify({version:1,project,active,combined});
  if(snapshot!==lastSavedWorkspace){localStorage.setItem(WORKSPACE_KEY,snapshot);lastSavedWorkspace=snapshot}
  $('autosaveStatus').textContent='Saved in this browser';
  $('autosaveStatus').title='Graphs, data and settings restore after refresh on this browser and website. Use Save Excel for a portable backup.';
 }catch(error){$('autosaveStatus').textContent='Autosave unavailable — use Save project';if(!autosaveWarningShown){toast('Browser storage is unavailable or full. Use Save Excel to keep your work.');autosaveWarningShown=true}}
}
function restoreWorkspace(){
 try{
  const saved=localStorage.getItem(WORKSPACE_KEY);if(!saved)return;
  const state=JSON.parse(saved),p=state.project;
  if(state.version!==1||!p||![1,2].includes(p.version)||!Array.isArray(p.graphs)||p.graphs.length>50)throw Error('Invalid saved workspace');
  // Restore unfinished rows as well as complete graphs; plotting validation happens later.
  for(const g of p.graphs){
   if(!g||!Array.isArray(g.rows)||g.rows.length>5000||!g.colors||typeof g.colors!=='object')throw Error('Invalid saved graph');
   for(const k of ['name',...fields.filter(k=>!['sd','labels','average','points'].includes(k))])if(typeof g[k]!=='string')throw Error('Invalid saved settings');
   for(const k of ['sd','labels','average'])if(typeof g[k]!=='boolean')throw Error('Invalid saved display option');
   for(const r of g.rows){if(!r)throw Error('Invalid saved row');for(const k of ['category','series','value','sd'])if(typeof r[k]!=='string')throw Error('Invalid saved row')}
   normalizeGraph(g);
  }
  p.version=2;project=p;active=Number.isInteger(state.active)?Math.max(0,Math.min(state.active,Math.max(0,p.graphs.length-1))):0;combined=state.combined===true;
 }catch(error){autosaveBlocked=true;$('autosaveStatus').textContent='Saved workspace could not be restored';$('autosaveStatus').title='The previous saved copy has been preserved. Use Open Excel to load a backup.';toast('Could not restore the saved workspace. Your previous saved copy has been preserved.')}
}
window.addEventListener('pagehide',saveWorkspace);
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='hidden')saveWorkspace()});
const current=()=>project.graphs[active];
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function toast(message){$('toast').textContent=message;$('toast').style.display='block';setTimeout(()=>$('toast').style.display='none',2500)}
function mean(a){return a.reduce((s,x)=>s+x,0)/a.length}function deviation(a){return a.length>1?Math.sqrt(a.reduce((s,x)=>s+(x-mean(a))**2,0)/(a.length-1)):null}
function numbers(text){const parts=String(text).trim().replace(/^\[/,'').replace(/\]$/,'').split(/[,;\s]+/).filter(Boolean);if(!parts.length)throw Error('Enter a measurement.');const a=parts.map(Number);if(a.some(x=>!Number.isFinite(x)))throw Error('Measurements must be finite numbers.');return a}
function rowUsed(r){return !!(r.category.trim()||r.series.trim()||r.value.trim()||r.sd.trim()||r.samples?.some(v=>v.trim()))}
function blankRow(){return {category:'',series:'',type:'summary',samples:[''],value:'',sd:''}}
function normalizeGraph(g){normalizeAppearance(g);if(g.points===undefined)g.points=false;if(typeof g.points!=='boolean')throw Error('Invalid individual measurements option.');if(g.grouped===undefined)g.grouped=true;if(typeof g.pivoted!=='boolean')g.pivoted=false;if(!palettes[g.palette])g.palette='palette1';if(!g.pivotColors)g.pivotColors={};if(typeof g.otherXLabel!=='string')g.otherXLabel='Series';if(!['wide','narrow'].includes(g.width))g.width='wide';for(const r of g.rows){if(!r.type)r.type=g.mode==='summary'||r.sd.trim()!==''?'summary':'raw';if(!['raw','summary'].includes(r.type))throw Error('Invalid row input type.');if(r.samples!==undefined&&(!Array.isArray(r.samples)||r.samples.length>10000||r.samples.some(v=>typeof v!=='string')))throw Error('Invalid measurement array.');if(!r.samples)r.samples=r.type==='raw'?r.value.replace(/^\[/,'').replace(/\]$/,'').split(/[,;\s]+/).filter(Boolean):[];if(!r.samples.length||r.samples.at(-1).trim())r.samples.push('')}delete g.mode;return g}
function rowStats(r){if(r.type==='summary'){const a=numbers(r.value);if(a.length!==1)throw Error('Enter one mean, or choose Raw data for this row.');const sd=r.sd.trim()===''?null:Number(r.sd);if(sd!==null&&(!Number.isFinite(sd)||sd<0))throw Error('SD must be a nonnegative number.');return {mean:a[0],sd,raw:null,n:null}}const a=r.samples.filter(v=>v.trim()).map(v=>{const parsed=numbers(v);if(parsed.length!==1)throw Error('One measurement per field. Paste an Excel row or column to fill several fields.');return parsed[0]});if(!a.length)throw Error('Enter at least one measurement.');return {mean:mean(a),sd:deviation(a),raw:a,n:a.length}}
function compile(g){const keys=new Set();return g.rows.filter(rowUsed).map((r,i)=>{if(!r.category.trim())throw Error(`Row ${i+1}: enter a category.`);const pivoted=g.grouped&&g.pivoted;const group=g.grouped?(pivoted?r.category.trim():(r.series.trim()||'Series 1')):'Bars';const category=pivoted?(r.series.trim()||'Series 1'):r.category.trim();const key=JSON.stringify([category,group]);if(keys.has(key))throw Error(`Duplicate bar: ${category}${g.grouped?' / '+group:''}. Use one row per bar.`);keys.add(key);try{return {category,series:group,...rowStats(r)}}catch(e){throw Error(`Row ${i+1}: ${e.message}`)}})}
function parseExcelArray(text){const lines=text.trim().split(/\r?\n/).filter(l=>l.trim());if(lines.length>1&&lines.some(l=>l.split('\t').filter(v=>v.trim()).length>1))throw Error('Paste a single Excel row or column for this array.');const vals=text.trim().replace(/^\[/,'').replace(/\]$/,'').split(/[,;\s]+/).filter(Boolean);if(!vals.length||vals.some(v=>!Number.isFinite(Number(v))))throw Error('Paste numeric measurements only, without a header.');return vals}


const appearanceCandidates=new Map();
let advancedOpen=false,lastGoodPlot=null;
function normalizeAppearance(g){
 for(const key of ['visualScript','visualDraft']){if(g[key]===undefined)g[key]='';if(typeof g[key]!=='string'||g[key].length>100000)throw Error('Invalid appearance script')}
 if(g.visualEnabled===undefined)g.visualEnabled=false;
 if(typeof g.visualEnabled!=='boolean')throw Error('Invalid appearance setting');
}
function starterAppearance(g){return `# Runs after Bar-Tender draws this graph using the current table data.
# Edit these Matplotlib commands, then click Run.
# ax = this graph's axes; fig = the figure; bars = bar artists.
# data is a read-only tuple of records; categories and series are tuples.

ax.set_title(${JSON.stringify(g.title)})
ax.set_xlabel(${JSON.stringify(g.xlabel)})
ax.set_ylabel(${JSON.stringify(g.ylabel)})
ax.tick_params(axis="both", labelsize=10)

# Examples — remove the # to enable:
# ax.tick_params(axis="x", labelrotation=30)
# ax.grid(axis="y", linestyle="--", alpha=0.25)
# ax.set_axisbelow(True)
# for bar in bars:
#     bar.set_edgecolor("black")
#     bar.set_linewidth(0.8)
# ax.legend(loc="upper left", frameon=False)
`}
function syncAdvanced(reloadEditor=false){
 const g=current();$('previewTab').disabled=!g;$('advancedTab').disabled=!g;
 $('advancedPanel').hidden=!g||!advancedOpen;
 $('previewTab').setAttribute('aria-selected',String(!advancedOpen));$('advancedTab').setAttribute('aria-selected',String(advancedOpen));
 if(!g){$('appearanceBadge').textContent='';return}
 normalizeAppearance(g);
 if(reloadEditor)$('appearanceScript').value=g.visualDraft||g.visualScript||starterAppearance(g);
 const locked=g.visualEnabled||appearanceCandidates.has(g);
 for(const id of fields)$(id).disabled=locked;
 if(!locked)syncAverage();
 $('wide').disabled=locked;$('narrow').disabled=locked;$('palette').disabled=locked;
 $('appearanceLock').hidden=!locked;
 $('appearanceBadge').textContent=g.visualEnabled?'Custom appearance':'';
 $('scriptState').textContent=appearanceCandidates.has(g)?'Running…':g.visualDraft&&g.visualDraft!==g.visualScript?'Draft — click Run':g.visualEnabled?'Custom appearance active':'Standard appearance';
}
$('previewTab').onclick=()=>{advancedOpen=false;syncAdvanced()};
$('advancedTab').onclick=()=>{advancedOpen=true;syncAdvanced(true)};
$('appearanceScript').oninput=()=>{if(!current())return;current().visualDraft=$('appearanceScript').value;saveWorkspace();syncAdvanced()};
$('appearanceScript').onkeydown=e=>{if((e.ctrlKey||e.metaKey)&&e.key==='Enter'){e.preventDefault();$('runAppearance').click()}else if(e.key==='Tab'){e.preventDefault();const t=e.target,start=t.selectionStart,end=t.selectionEnd;t.setRangeText('    ',start,end,'end');t.dispatchEvent(new Event('input',{bubbles:true}))}};
$('runAppearance').onclick=()=>{const g=current();if(!g)return;g.visualDraft=$('appearanceScript').value;if(g.visualDraft.length>100000){toast('Appearance script is too long.');return}appearanceCandidates.set(g,g.visualDraft);saveWorkspace();render();syncAdvanced()};
$('resetAppearance').onclick=()=>{const g=current();if(!g)return;appearanceCandidates.delete(g);g.visualEnabled=false;g.visualScript='';g.visualDraft='';saveWorkspace();syncAdvanced(true);render()};
$('stopPython').onclick=()=>{
 invalidatePlot();appearanceCandidates.clear();
 for(const g of project.graphs)g.visualEnabled=false;
 const oldEngine=plotEngine;plotEngine=null;if(oldEngine)oldEngine.destroy();
 saveWorkspace();syncAdvanced();engineStatus('Restarting Python…');warmEngine();render();
 toast('Python restarted. Custom scripts are paused; drafts and data are preserved.');
};
function samePlotView(graphs){return lastGoodPlot&&lastGoodPlot.combined===combined&&lastGoodPlot.graphs.length===graphs.length&&lastGoodPlot.graphs.every((g,i)=>g===graphs[i])}
function recoverLastPlot(graphs){if(!samePlotView(graphs))return false;$('chart').innerHTML=lastGoodPlot.html;exportSvg=lastGoodPlot.svg;exportPng=lastGoodPlot.png;$('export').disabled=false;$('copyPng').disabled=false;return true}
let plotEngine=null,renderRevision=0,renderTimer=null,pendingRender=null,rendering=false,exportSvg='',exportPng=null;
function engineStatus(message){$('engineStatus').textContent=message}
async function warmEngine(){
 if(!plotEngine)plotEngine=new MatplotlibEngine(engineStatus);
 const engine=plotEngine;
 try{await engine.warmup();if(plotEngine===engine)$('retryEngine').hidden=true}
 catch(error){if(plotEngine!==engine)return;engineStatus('Could not load the chart engine. Check your connection, then retry.');$('retryEngine').hidden=false}
}
function invalidatePlot(){renderRevision++;clearTimeout(renderTimer);pendingRender=null;exportSvg='';exportPng=null;$('copyPng').disabled=true;$('export').disabled=true;$('chart').removeAttribute('aria-busy')}
function graphPayload(g){
 const data=compile(g);if(!data.length)throw Error('Add at least one measurement.');
 for(const k of ['lower','upper','ymin','ymax'])if(g[k]!==''&&!Number.isFinite(Number(g[k])))throw Error('Axis bounds and limits must be finite.');
 if(g.lower!==''&&g.upper!==''&&Number(g.lower)>Number(g.upper))throw Error('Lower limit must not exceed upper limit.');
 if(g.ymin!==''&&g.ymax!==''&&Number(g.ymax)<=Number(g.ymin))throw Error('Y maximum must be greater than Y minimum.');
 if(g.average&&data.some(r=>r.raw===null))throw Error('Overall average needs raw data in every row.');
 return {...g,appearanceCode:appearanceCandidates.has(g)?appearanceCandidates.get(g):(g.visualEnabled?g.visualScript:''),rows:undefined,colors:undefined,pivotColors:undefined,data,paletteColors:palettes[g.palette]||palettes.palette1};
}
async function drainRenderer(){
 if(rendering)return;rendering=true;
 try{while(pendingRender){const job=pendingRender;pendingRender=null;
  try{
   if(!plotEngine)plotEngine=new MatplotlibEngine(engineStatus);
   const result=await plotEngine.render(job.payload);
   if(job.revision!==renderRevision)continue;
   exportSvg=result.svg;
   exportPng=new Blob([Uint8Array.from(atob(result.png),c=>c.charCodeAt(0))],{type:'image/png'});
   // Matplotlib generates this SVG locally, with escaped labels and glyph paths.
   const svg=result.svg.slice(result.svg.indexOf('<svg'));
   $('chart').innerHTML=svg;$('chart').querySelector('svg').setAttribute('aria-label',job.label);$('chart').querySelector('svg').setAttribute('role','img');
   lastGoodPlot={graphs:job.graphs,combined:job.payload.combined,html:$('chart').innerHTML,svg:exportSvg,png:exportPng};
   for(const [i,g] of job.graphs.entries())if(appearanceCandidates.has(g)&&appearanceCandidates.get(g)===job.payload.graphs[i].appearanceCode){g.visualScript=appearanceCandidates.get(g);g.visualEnabled=true;appearanceCandidates.delete(g)}
   saveWorkspace();syncAdvanced();
   $('chartError').textContent='';$('chart').removeAttribute('aria-busy');$('export').disabled=false;$('copyPng').disabled=false;
   const stats=result.summaries;
   $('summary').textContent=`${stats.reduce((n,p)=>n+p.bars,0)} bars · ${job.payload.combined?`${stats.length} graphs`:`${stats[0].cats} categories · ${stats[0].groups} series`}`;
   engineStatus(`Matplotlib ${result.matplotlibVersion} · Running in your browser`);$('retryEngine').hidden=true;
  }catch(error){if(job.revision!==renderRevision)continue;
   for(const g of job.graphs)appearanceCandidates.delete(g);
   exportSvg='';exportPng=null;$('copyPng').disabled=true;$('export').disabled=true;$('chart').removeAttribute('aria-busy');
   const recovered=recoverLastPlot(job.graphs);if(!recovered)$('chart').innerHTML='';
   $('chartError').textContent=error.message.slice(-1500)+(recovered?'\nLast successful graph shown; these changes were not applied.':'');
   $('summary').textContent=recovered?'Last successful graph':'Chart unavailable';
   engineStatus(error.kind==='appearance'?'Appearance script needs a correction.':'Unable to draw the chart. Retry or restart Python.');$('retryEngine').hidden=error.kind==='appearance';
   saveWorkspace();syncAdvanced();
  }
 }}finally{rendering=false}
}
function render(){
 invalidatePlot();if(!current()){load();return}
 syncAverage();saveWorkspace();syncGraphControls();syncAdvanced();$('wide').classList.toggle('active',current().width==='wide');$('narrow').classList.toggle('active',current().width==='narrow');$('chart').classList.toggle('narrow',!combined&&current().width==='narrow');$('previewtitle').textContent=combined?'Combined figure':current().title;updateComputed();
 try{
  const selected=combined?project.graphs.filter(g=>g.rows.some(rowUsed)):[current()];
  if(!selected.length||(!combined&&!current().rows.some(rowUsed))){$('chart').innerHTML='<div class="emptychart"><strong>Your graph starts here</strong><span>Name a category, then enter samples or a mean and SD.</span></div>';$('chartError').textContent='';$('summary').textContent='No measurements yet';return}
  const payload={combined,graphs:selected.map(graphPayload)},revision=renderRevision,label=combined?'Combined figure':current().title;
  $('chartError').textContent='';$('chart').setAttribute('aria-busy','true');$('summary').textContent='Updating chart…';
  if(!$('chart').querySelector('svg'))$('chart').innerHTML='<div class="emptychart"><strong>Preparing your chart…</strong><span>Python and Matplotlib are loading in the background. You can keep entering data.</span></div>';
  renderTimer=setTimeout(()=>{pendingRender={revision,payload,label,graphs:selected};drainRenderer()},180);
 }catch(error){$('chartError').textContent=error.message;$('chart').innerHTML='';$('summary').textContent='Check the data and settings'}
}
$('retryEngine').onclick=()=>{if(plotEngine)plotEngine.destroy();plotEngine=null;$('retryEngine').hidden=true;warmEngine();render()};
function syncAverage(){const hasRaw=current().rows.some(r=>r.type==='raw'&&r.samples.some(v=>v.trim()));$('points').disabled=!hasRaw;$('points').title=hasRaw?'Show raw measurements over each bar; can be combined with SD':'Enter raw measurements to show individual points';const unavailable=current().rows.some(r=>rowUsed(r)&&r.type==='summary');$('average').disabled=unavailable;if(unavailable&&current().average){current().average=false;$('average').checked=false}}
function updateComputed(){current().rows.forEach((r,i)=>{const cells=$('rows').children[i]?.querySelectorAll('[data-computed]');if(!cells||r.type!=='raw')return;try{const stat=rowStats(r);cells[0].textContent=Number(stat.mean.toFixed(3));cells[1].textContent=stat.sd===null?'n/a':Number(stat.sd.toFixed(3));cells[2].textContent=stat.n}catch{cells.forEach(c=>c.textContent='—')}})}
function makeInput(value,label,oninput,type='text',placeholder=''){const input=document.createElement('input');input.value=value;input.type=type;input.setAttribute('aria-label',label);input.placeholder=placeholder;if(type==='number')input.step='any';input.oninput=()=>oninput(input.value);return input}
function setRowArray(row,values){row.type='raw';row.samples=[...values,''];row.value=values.join(', ');row.sd=''}
function syncPasteTargets(){const select=$('pasteTarget'),previous=select.value;select.innerHTML='';current().rows.forEach((r,i)=>{const option=document.createElement('option');option.value=String(i);option.textContent=`Row ${i+1}: ${r.category.trim()||'(unnamed category)'}${current().grouped?' / '+(r.series.trim()||'(unnamed series)'):''}`;select.append(option)});if([...select.options].some(o=>o.value===previous))select.value=previous;$('apply').disabled=!current().rows.length}
function paintSamples(wrap,r,i){wrap.innerHTML='';function addSample(j){const input=makeInput(r.samples[j],`Measurement ${j+1}, row ${i+1}`,v=>{r.samples[j]=v;r.value=r.samples.filter(x=>x.trim()).join(', ');if(j===r.samples.length-1&&v.trim()){r.samples.push('');addSample(j+1)}render()});input.onkeydown=e=>{if((e.key!=='Enter'&&e.key!=='ArrowRight')||e.altKey||e.ctrlKey||e.metaKey||e.shiftKey)return;e.preventDefault();if(j===r.samples.length-1){r.samples.push('');addSample(j+1)}const next=wrap.querySelectorAll('input')[j+1];next.focus();next.select();next.scrollIntoView({block:'nearest',inline:'nearest'})};input.inputMode='decimal';input.placeholder='+';input.className='sampleinput';input.onpaste=e=>{const text=e.clipboardData.getData('text');if(!/[\t\n,;]/.test(text))return;e.preventDefault();try{const vals=parseExcelArray(text);r.samples.splice(j,r.samples[j].trim()?1:0,...vals);r.samples=r.samples.filter(v=>v.trim());r.samples.push('');r.value=r.samples.filter(v=>v.trim()).join(', ');paintSamples(wrap,r,i);render();toast(`${vals.length} measurements pasted`)}catch(err){toast(err.message)}};wrap.append(input);if(j===r.samples.length-1)wrap.scrollLeft=wrap.scrollWidth}r.samples.forEach((_,j)=>addSample(j))}
function paintRows(){const g=current();$('rows').innerHTML='';$('tableHead').innerHTML='<tr>'+[g.grouped?(g.pivoted?'Category · color':'Category · axis'):'Category',...(g.grouped?[g.pivoted?'Series · axis':'Series · color']:[]),'Mean','SD','n','Input type','Add data',''].map(h=>`<th>${h}</th>`).join('')+'</tr>';g.rows.forEach((r,i)=>{const tr=document.createElement('tr');for(const key of ['category',...(g.grouped?['series']:[])]){const td=document.createElement('td');const input=makeInput(r[key],`${key}, row ${i+1}`,v=>{r[key]=v;syncPasteTargets();render()},'text',key==='category'?'Category A':'Series 1');td.append(input);tr.append(td)}const typecell=document.createElement('td'),select=document.createElement('select');select.setAttribute('aria-label',`Input type, row ${i+1}`);for(const [value,label]of [['raw','Raw data'],['summary','Mean + SD']]){const o=document.createElement('option');o.value=value;o.textContent=label;select.append(o)}select.value=r.type;select.onchange=()=>{if(select.value==='summary'){try{const stat=rowStats(r);r.value=String(stat.mean);r.sd=stat.sd===null?'':String(stat.sd)}catch{r.value='';r.sd=''}}else{r.value=r.samples.filter(v=>v.trim()).join(', ');r.sd=''}r.type=select.value;paintRows();render()};typecell.append(select);const measurement=document.createElement('td');measurement.className='measurecell';if(r.type==='raw'){const wrap=document.createElement('div');wrap.className='samples';paintSamples(wrap,r,i);measurement.append(wrap)}else{const hint=document.createElement('span');hint.className='providedhint';hint.textContent='← Enter mean + SD';measurement.append(hint)}for(const [key,label]of [['value','Mean'],['sd','SD']]){const td=document.createElement('td');if(r.type==='raw'){td.className='computed';td.dataset.computed=key;td.textContent='—'}else{const input=makeInput(r[key],`${label}, row ${i+1}`,v=>{r[key]=v;render()},'number',key==='value'?'Mean':'Optional');if(key==='sd')input.min=0;if(key==='value')input.onpaste=e=>{const text=e.clipboardData.getData('text');if(!/[\t\n,;]/.test(text))return;e.preventDefault();try{const values=parseExcelArray(text);setRowArray(r,values);paintRows();render();toast(`${values.length} measurements added to this category`)}catch(error){toast(error.message)}};td.append(input)}tr.append(td)}const n=document.createElement('td');n.className='computed';if(r.type==='raw')n.dataset.computed='n';n.textContent='—';tr.append(n);tr.append(typecell);tr.append(measurement);const td=document.createElement('td'),del=document.createElement('button');del.textContent='×';del.setAttribute('aria-label',`Delete row ${i+1}`);del.onclick=()=>{g.rows.splice(i,1);paintRows();render()};td.append(del);tr.append(td);$('rows').append(tr)});updateComputed();syncPasteTargets()}

function load(){syncAdvanced(true);saveWorkspace();invalidatePlot();const g=current();const hasGraph=!!g;document.querySelector('aside').hidden=!hasGraph;document.querySelector('.data').hidden=!hasGraph;document.querySelector('.view-options').hidden=!hasGraph;for(const id of ['pivot','wide','narrow','single','combined','palette'])$(id).disabled=!hasGraph;$('save').disabled=!hasGraph;if(!hasGraph){$('tabs').innerHTML='';$('rows').innerHTML='';$('chart').classList.remove('narrow');$('chart').innerHTML='<div class="emptychart"><strong>Create your first graph</strong><span>Start with categories, or compare categories and series.</span><button class="primary" id="startGraph">+ New graph</button></div>';$('startGraph').onclick=()=>$('newGraph').click();$('previewtitle').textContent='Your workspace';$('dimensionRoles').textContent='';$('chartError').textContent='';$('summary').textContent='No graphs yet';$('dataStatus').textContent='Your measurements · browser only';$('export').disabled=true;return}$('tabs').innerHTML='';project.graphs.forEach((gr,i)=>{const b=document.createElement('button');b.textContent=gr.name;b.className=i===active?'active':'';b.onclick=()=>{active=i;load()};$('tabs').append(b);if(i===active){const close=document.createElement('button');close.textContent='×';close.className='closegraph';close.setAttribute('aria-label','Delete '+gr.name);close.onclick=()=>{project.graphs.splice(i,1);active=Math.min(i,project.graphs.length-1);load()};$('tabs').append(close)}});for(const id of fields){if($(id).type==='checkbox')$(id).checked=g[id];else $(id).value=g[id]}$('pasteError').textContent='';$('paste').value='';$('single').classList.toggle('active',!combined);$('combined').classList.toggle('active',combined);$('modeBadge').textContent=g.grouped?'CATEGORIES + SERIES':'CATEGORIES';$('dataHelp').textContent='Rows start with Mean + SD. Paste an Excel number row or column into Mean to switch that category to raw data, or choose Raw data and paste into a measurement field.';$('pasteHelp').textContent='Paste numbers only: one Excel row or one Excel column. All numbers become measurements for the selected category'+(g.grouped?' and series.':'.')+' Its mean and sample SD are calculated automatically.';document.querySelector('aside .small').textContent='Raw data uses sample SD (n − 1). At least two samples are needed for SD. Overall averages require all rows to contain raw data.';$('dataStatus').textContent='Your measurements · browser only';paintRows();render()}

for(const id of fields)$(id).oninput=()=>{current()[id]=$(id).type==='checkbox'?$(id).checked:$(id).value;render()};
$('single').onclick=()=>{combined=false;$('single').classList.add('active');$('combined').classList.remove('active');render()};$('combined').onclick=()=>{combined=true;$('combined').classList.add('active');$('single').classList.remove('active');render()};
$('add').onclick=()=>{current().rows.push(blankRow());paintRows();saveWorkspace();$('rows').lastElementChild.querySelector('input').focus()};
$('apply').onclick=()=>{try{const values=parseExcelArray($('paste').value),index=Number($('pasteTarget').value),row=current().rows[index];if(!row)throw Error('Add a category row first.');setRowArray(row,values);paintRows();render();$('pasteError').textContent='';$('paste').value='';toast(`${values.length} measurements added to row ${index+1}`)}catch(e){$('pasteError').textContent=e.message}};

function download(text,name,type){const url=URL.createObjectURL(new Blob([text],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.hidden=true;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000)}
$('copyPng').onclick=async()=>{
 if(!exportPng)return;
 if(!navigator.clipboard?.write||typeof ClipboardItem==='undefined'){toast('Image copying requires a supported browser and the HTTPS website link.');return}
 const revision=renderRevision;const button=$('copyPng');button.disabled=true;
 try{
  // PNG is prepared by Matplotlib ahead of the click, preserving user activation.
  await navigator.clipboard.write([new ClipboardItem({'image/png':exportPng})]);
  toast('Graph copied as PNG — paste it into your document or slide.');
 }catch(error){toast('Could not copy. Allow clipboard access for this website, then try again.')}
 finally{if(revision===renderRevision)button.disabled=!exportPng}
};
$('export').onclick=()=>{if(!exportSvg)return;download(exportSvg,combined?'combined-figure.svg':`${current().name}.svg`,'image/svg+xml')};
function validateImportedProject(p){
 if(!p||typeof p!=='object')throw Error('Invalid project file.');
 if(![1,2].includes(p.version)||!Array.isArray(p.graphs)||p.graphs.length>50)throw Error('Choose a Plot Studio project with up to 50 graphs.');for(const g of p.graphs){if(g&&g.points===undefined)g.points=false;if(!g||!Array.isArray(g.rows)||!g.colors||g.rows.length>5000)throw Error('Invalid project data.');for(const k of ['name',...fields.filter(x=>!['sd','labels','average','points'].includes(x))])if(typeof g[k]!=='string')throw Error(`Invalid project setting: ${k}`);for(const k of ['sd','labels','average','points'])if(typeof g[k]!=='boolean')throw Error('Invalid display options.');for(const v of Object.values(g.colors))if(!/^#[0-9a-f]{6}$/i.test(v))throw Error('Invalid series color.');for(const r of g.rows)for(const k of ['category','series','value','sd'])if(typeof r[k]!=='string')throw Error('Invalid measurement row.');for(const value of Object.values(g.pivotColors||{}))if(typeof value!=='string'||!/^#[0-9a-f]{6}$/i.test(value))throw Error('Invalid pivot color.');if(g.grouped!==undefined&&typeof g.grouped!=='boolean')throw Error('Invalid graph structure.');normalizeGraph(g)}
 p.version=2;return p;
}
$('save').onclick=async()=>{try{
 const bytes=await BarTenderExcel.write(project,active,combined);
 download(bytes,'bar-tender-project.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
 toast('Excel project saved — reopen it with Open Excel.');
}catch(error){toast('Could not save Excel: '+error.message)}};
$('open').onclick=()=>$('file').click();
$('file').onchange=async()=>{try{
 const file=$('file').files[0];if(!file)return;
 if(file.size>25000000)throw Error('Choose a project workbook smaller than 25 MB.');
 const state=/\.json$/i.test(file.name)?{project:JSON.parse(await file.text()),active:0,combined:false}:await BarTenderExcel.read(await file.arrayBuffer());
 const restored=validateImportedProject(state.project);
 invalidatePlot();appearanceCandidates.clear();lastGoodPlot=null;
 project=restored;active=Math.max(0,Math.min(state.active||0,Math.max(0,project.graphs.length-1)));combined=!!state.combined;
 autosaveBlocked=false;load();toast('Project opened with data and graph settings.');
}catch(error){toast(error.message)}finally{$('file').value=''}};
project.graphs.forEach(normalizeGraph);

$('newGraph').onclick=()=>{$('newGraphName').value='';document.querySelector('input[name=graphStructure][value=categories]').checked=true;$('newGraphError').textContent='';$('newGraphDialog').showModal();$('newGraphName').focus()};
$('closeDialog').onclick=$('cancelDialog').onclick=()=>$('newGraphDialog').close();
$('newGraphForm').onsubmit=e=>{e.preventDefault();const name=$('newGraphName').value.trim();if(!name){$('newGraphError').textContent='Enter a graph name.';return}if(project.graphs.some(g=>g.name.toLowerCase()===name.toLowerCase())){$('newGraphError').textContent='Choose a name that is not already in use.';return}if(project.graphs.length>=50){$('newGraphError').textContent='This project already contains 50 graphs.';return}const grouped=document.querySelector('input[name="graphStructure"]:checked').value==='series';project.graphs.push({name,title:name,grouped,pivoted:false,palette:'palette1',pivotColors:{},otherXLabel:'Series',width:'wide',xlabel:'Category',ylabel:'Value',ymin:'',ymax:'',lower:'',upper:'',sd:true,labels:true,average:false,points:false,colors:{},rows:[blankRow()]});active=project.graphs.length-1;combined=false;$('newGraphDialog').close();load();$('rows').querySelector('input').focus();toast(grouped?'Categories and series graph created':'Categories graph created')};
$('wide').onclick=()=>{current().width='wide';render()};$('narrow').onclick=()=>{current().width='narrow';render()};

function syncGraphControls(){const g=current();$('pivot').disabled=!g.grouped;$('pivot').setAttribute('aria-pressed',String(!!g.pivoted));$('pivot').classList.toggle('selected',!!g.pivoted);$('pivot').title=g.grouped?'Swap the axis and color dimensions':'Pivot is available for graphs with categories and series';$('palette').value=g.palette||'palette1';$('dimensionRoles').textContent=g.grouped?(g.pivoted?'Axis: Series  ·  Color: Category':'Axis: Category  ·  Color: Series'):'Axis: Category';}
$('pivot').onclick=()=>{const g=current();if(!g.grouped)return;g.pivoted=!g.pivoted;[g.xlabel,g.otherXLabel]=[g.otherXLabel,g.xlabel];$('xlabel').value=g.xlabel;paintRows();render()};
$('palette').onchange=()=>{const g=current();g.palette=$('palette').value;g.colors={};g.pivotColors={};render()};

restoreWorkspace();
load();
warmEngine();
