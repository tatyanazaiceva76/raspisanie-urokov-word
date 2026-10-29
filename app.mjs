import {DAYS,buildDocx,validateModel} from './core.mjs';
const createDraft=()=>Array.from({length:6},()=>Array(8).fill(''));
function snapshotDraft(draft,inputs){for(const input of inputs){const day=Number(input.dataset.day),period=Number(input.dataset.period);if(Number.isInteger(day)&&day>=0&&day<6&&Number.isInteger(period)&&period>=0&&period<8)draft[day][period]=input.value;}return draft;}
const visibleCells=(draft,days,periods)=>Array.from({length:days},(_,day)=>draft[day].slice(0,periods));
function setDownloadState(anchor,{enabled,href=null}={}){if(enabled){if(!href)throw new Error('DOWNLOAD_HREF_REQUIRED');anchor.href=href;anchor.removeAttribute('aria-disabled');}else{anchor.setAttribute('aria-disabled','true');anchor.removeAttribute('href');}return anchor;}
const canDownload=(anchor,objectUrl,generated)=>Boolean(objectUrl&&generated&&anchor.hasAttribute('href')&&!anchor.hasAttribute('aria-disabled'));
function containTable(document,table){const wrapper=document.createElement('div');wrapper.className='table-scroll';wrapper.append(table);return wrapper;}


const $=id=>document.getElementById(id),form=$('planner'),grid=$('grid'),preview=$('preview'),errors=$('errors'),download=$('download'),build=$('build');
let generated=null,objectUrl=null;
let draft=createDraft();
const revoke=()=>{if(objectUrl)URL.revokeObjectURL(objectUrl);objectUrl=null;generated=null;setDownloadState(download,{enabled:false});preview.replaceChildren();$('result').hidden=true;};
const invalidate=()=>{revoke();errors.textContent='';};

function settings(){return{days:Number($('days').value),periods:Number($('periods').value),paper:$('paper').value,title:$('title').value};}
function snapshotVisible(){snapshotDraft(draft,grid.querySelectorAll('input[data-day]'));}
function values(){snapshotVisible();const s=settings(),cells=visibleCells(draft,s.days,s.periods);return{...s,cells};}
function make(tag,text,className){const node=document.createElement(tag);if(text!==undefined)node.textContent=text;if(className)node.className=className;return node;}

function renderInputs(){const s=settings();grid.replaceChildren();const table=make('table');table.setAttribute('aria-label','Предметы по дням и урокам');const head=document.createElement('thead'),hr=document.createElement('tr');hr.append(make('th','№'));for(const day of DAYS.slice(0,s.days))hr.append(make('th',day));head.append(hr);table.append(head);const body=document.createElement('tbody');for(let period=0;period<s.periods;period++){const row=document.createElement('tr');row.append(make('th',String(period+1)));for(let day=0;day<s.days;day++){const cell=document.createElement('td'),input=document.createElement('input');input.type='text';input.maxLength=120;input.dataset.day=String(day);input.dataset.period=String(period);input.value=draft[day][period];input.setAttribute('aria-label',`${DAYS[day]}, урок ${period+1}`);cell.append(input);row.append(cell);}body.append(row);}table.append(body);grid.append(table);}

function renderPreview(model){const table=make('table');table.setAttribute('aria-label','Предпросмотр готового расписания');const head=document.createElement('thead'),hr=document.createElement('tr');hr.append(make('th','№'));for(const day of DAYS.slice(0,model.days))hr.append(make('th',day));head.append(hr);table.append(head);const body=document.createElement('tbody');for(let p=0;p<model.periods;p++){const row=document.createElement('tr');row.append(make('th',String(p+1)));for(let d=0;d<model.days;d++)row.append(make('td',model.cells[d][p]||' '));body.append(row);}table.append(body);preview.replaceChildren(make('h2',model.title||'Расписание уроков'),containTable(document,table),make('p','В DOCX останется редактируемая таблица. Перед печатью проверьте переносы и страницы в своём редакторе.','fine'));$('result').hidden=false;}

function produce(){invalidate();try{const model=validateModel(values()),bytes=buildDocx(model);generated=bytes;objectUrl=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}));setDownloadState(download,{enabled:true,href:objectUrl});download.download='raspisanie-urokov.docx';renderPreview(model);$('result').scrollIntoView({block:'start',behavior:'smooth'});}catch(error){errors.textContent=error.message;errors.focus();}}

function demo(){const s=settings();$('title').value='Моё расписание';const examples=['Русский язык','Математика','Литература','История','Биология','Иностранный язык','Физкультура','Информатика'];draft=createDraft();for(let day=0;day<s.days;day++)for(let period=0;period<s.periods;period++)draft[day][period]=examples[(day+period)%examples.length];renderInputs();invalidate();}
function reset(){HTMLFormElement.prototype.reset.call(form);$('days').value='5';$('periods').value='6';$('paper').value='letterPortrait';$('title').value='';draft=createDraft();renderInputs();invalidate();$('title').focus();}

form.addEventListener('submit',event=>{event.preventDefault();produce();});
form.addEventListener('input',event=>{if(event.target.matches?.('input[data-day]')){const day=Number(event.target.dataset.day),period=Number(event.target.dataset.period);draft[day][period]=event.target.value;}invalidate();});
form.addEventListener('change',event=>{if(event.target=== $('days')||event.target=== $('periods')){snapshotVisible();renderInputs();}invalidate();});
$('demo').addEventListener('click',demo);
$('reset').addEventListener('click',reset);
download.addEventListener('click',event=>{if(!canDownload(download,objectUrl,generated))event.preventDefault();});
window.addEventListener('beforeunload',()=>{if(objectUrl)URL.revokeObjectURL(objectUrl);});
renderInputs();build.disabled=false;
