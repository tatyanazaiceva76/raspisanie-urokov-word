const DAYS=['Понедельник','Вторник','Среда','Четверг','Пятница','Суббота'];
const PAPER={
  letterPortrait:{label:'Letter, книжная',width:12240,height:15840,orientation:null},
  a4Portrait:{label:'A4, книжная',width:11906,height:16838,orientation:null},
  a4Landscape:{label:'A4, альбомная',width:16838,height:11906,orientation:'landscape'}
};
const encoder=new TextEncoder();
const codepoints=value=>[...value].length;
const assertText=(value,max,label,{empty=true}={})=>{
  if(typeof value!=='string')throw new Error(`${label}: ожидается текст`);
  if(!empty&&value.length===0)throw new Error(`${label}: заполните поле`);
  if(codepoints(value)>max)throw new Error(`${label}: не более ${max} символов`);
  for(let i=0;i<value.length;i++){const c=value.charCodeAt(i);if((c<=0x1f)||(c>=0x7f&&c<=0x9f))throw new Error(`${label}: управляющие символы не поддерживаются`);if(c===0xfffe||c===0xffff)throw new Error(`${label}: символ не поддерживается в XML 1.0`);if(c>=0xd800&&c<=0xdbff){const n=value.charCodeAt(++i);if(!(n>=0xdc00&&n<=0xdfff))throw new Error(`${label}: повреждённая Unicode-пара`);}else if(c>=0xdc00&&c<=0xdfff)throw new Error(`${label}: повреждённая Unicode-пара`);}
  return value;
};
const xml=value=>String(value).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&apos;');
export {DAYS,PAPER,xml};

export function validateModel(raw){
  if(!raw||typeof raw!=='object'||Array.isArray(raw))throw new Error('Нужны параметры расписания');
  const days=Number(raw.days),periods=Number(raw.periods),paper=String(raw.paper??'');
  if(![5,6].includes(days))throw new Error('Учебных дней должно быть 5 или 6');
  if(![4,6,8].includes(periods))throw new Error('Уроков в день должно быть 4, 6 или 8');
  if(!Object.hasOwn(PAPER,paper))throw new Error('Выберите доступный формат бумаги');
  const title=assertText(raw.title??'',80,'Заголовок').trim();
  if(!Array.isArray(raw.cells)||raw.cells.length!==days)throw new Error('Сетка дней не совпадает с настройками');
  const cells=raw.cells.map((column,day)=>{
    if(!Array.isArray(column)||column.length!==periods)throw new Error(`Сетка уроков для дня ${day+1} не совпадает с настройками`);
    return column.map((value,row)=>assertText(value,60,`${DAYS[day]}, урок ${row+1}`));
  });
  return Object.freeze({days,periods,paper,title,cells:Object.freeze(cells.map(row=>Object.freeze([...row]))) });
}

const crcTable=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
export function crc32(bytes){let c=0xffffffff;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0;}
const u16=(view,offset,value)=>view.setUint16(offset,value,true),u32=(view,offset,value)=>view.setUint32(offset,value>>>0,true);

export function zipStore(entries){
  if(!Array.isArray(entries)||entries.length<1||entries.length>32)throw new Error('Недопустимое число частей DOCX');
  const seen=new Set(),locals=[],centrals=[];let offset=0,total=0;
  for(const entry of entries){if(!entry||typeof entry.name!=='string'||seen.has(entry.name)||entry.name.startsWith('/')||entry.name.includes('..'))throw new Error('Недопустимое имя части DOCX');seen.add(entry.name);const name=encoder.encode(entry.name),data=entry.data instanceof Uint8Array?entry.data:encoder.encode(String(entry.data));total+=data.length;if(total>4*1024*1024)throw new Error('DOCX превышает локальный лимит 4 МиБ');const crc=crc32(data),local=new Uint8Array(30+name.length+data.length),lv=new DataView(local.buffer);u32(lv,0,0x04034b50);u16(lv,4,20);u16(lv,6,0x0800);u16(lv,8,0);u16(lv,10,0);u16(lv,12,0x0021);u32(lv,14,crc);u32(lv,18,data.length);u32(lv,22,data.length);u16(lv,26,name.length);u16(lv,28,0);local.set(name,30);local.set(data,30+name.length);locals.push(local);
    const central=new Uint8Array(46+name.length),cv=new DataView(central.buffer);u32(cv,0,0x02014b50);u16(cv,4,20);u16(cv,6,20);u16(cv,8,0x0800);u16(cv,10,0);u16(cv,12,0);u16(cv,14,0x0021);u32(cv,16,crc);u32(cv,20,data.length);u32(cv,24,data.length);u16(cv,28,name.length);u16(cv,30,0);u16(cv,32,0);u16(cv,34,0);u16(cv,36,0);u32(cv,38,0);u32(cv,42,offset);central.set(name,46);centrals.push(central);offset+=local.length;
  }
  const centralOffset=offset,centralBytes=centrals.reduce((n,b)=>n+b.length,0),eocd=new Uint8Array(22),ev=new DataView(eocd.buffer);u32(ev,0,0x06054b50);u16(ev,4,0);u16(ev,6,0);u16(ev,8,entries.length);u16(ev,10,entries.length);u32(ev,12,centralBytes);u32(ev,16,centralOffset);u16(ev,20,0);const out=new Uint8Array(centralOffset+centralBytes+eocd.length);let cursor=0;for(const part of [...locals,...centrals,eocd]){out.set(part,cursor);cursor+=part.length;}return out;
}

const pText=(text,{bold=false,size=21,align='center',style=null}={})=>`<w:p><w:pPr>${style?`<w:pStyle w:val="${style}"/>`:''}<w:jc w:val="${align}"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>${bold?'<w:b/>':''}<w:color w:val="000000"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p>`;
const borders='<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="D9D9D9"/><w:left w:val="single" w:sz="6" w:space="0" w:color="D9D9D9"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="D9D9D9"/><w:right w:val="single" w:sz="6" w:space="0" w:color="D9D9D9"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="D9D9D9"/><w:insideV w:val="single" w:sz="6" w:space="0" w:color="D9D9D9"/></w:tblBorders>';
const cell=(text,width,{header=false}={})=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>${header?'<w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/>':''}</w:tcPr>${pText(text,{bold:header,size:header?20:21})}</w:tc>`;

export function documentXml(model){
  const v=validateModel(model),page=PAPER[v.paper],margin=720,usable=page.width-margin*2,numberWidth=600,dayWidth=Math.floor((usable-numberWidth)/v.days),grid=[numberWidth,...Array(v.days).fill(dayWidth)],header=`<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>${cell('№',numberWidth,{header:true})}${DAYS.slice(0,v.days).map((day,i)=>cell(day,grid[i+1],{header:true})).join('')}</w:tr>`,rows=[];
  for(let lesson=0;lesson<v.periods;lesson++)rows.push(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(String(lesson+1),numberWidth)}${Array.from({length:v.days},(_,day)=>cell(v.cells[day][lesson],dayWidth)).join('')}</w:tr>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${pText(v.title||'Расписание уроков',{size:32,style:'Title'})}${pText('Редактируемая таблица на неделю',{size:21})}<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableNormal"/><w:tblW w:w="${usable}" w:type="dxa"/><w:tblLayout w:type="fixed"/>${borders}<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${grid.map(width=>`<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${header}${rows.join('')}</w:tbl><w:p><w:pPr><w:spacing w:before="160"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/><w:color w:val="000000"/><w:sz w:val="18"/></w:rPr><w:t>Проверьте переносы и страницы в своём редакторе перед печатью.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="${page.width}" w:h="${page.height}"${page.orientation?' w:orient="landscape"':''}/><w:pgMar w:top="720" w:right="${margin}" w:bottom="720" w:left="${margin}" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>`;
}

const contentTypes=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
const rootRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
const docRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:color w:val="000000"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/><w:b/><w:color w:val="000000"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
const core=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Расписание уроков</dc:title><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`;
const app=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Local browser DOCX generator</Application><AppVersion>1.0</AppVersion></Properties>`;

export const DOCX_MEMBERS=Object.freeze(['[Content_Types].xml','_rels/.rels','docProps/core.xml','docProps/app.xml','word/document.xml','word/styles.xml','word/_rels/document.xml.rels']);
export function buildDocx(raw){const model=validateModel(raw),parts={'[Content_Types].xml':contentTypes,'_rels/.rels':rootRels,'docProps/core.xml':core,'docProps/app.xml':app,'word/document.xml':documentXml(model),'word/styles.xml':styles,'word/_rels/document.xml.rels':docRels};return zipStore(DOCX_MEMBERS.map(name=>({name,data:encoder.encode(parts[name])})));}

export function parseStoreZip(bytes){
  const out=new Map();let offset=0;while(offset+4<=bytes.length){const view=new DataView(bytes.buffer,bytes.byteOffset+offset,bytes.byteLength-offset),signature=view.getUint32(0,true);if(signature!==0x04034b50)break;const flags=view.getUint16(6,true),method=view.getUint16(8,true),crc=view.getUint32(14,true),size=view.getUint32(18,true),nameLength=view.getUint16(26,true),extraLength=view.getUint16(28,true),name=new TextDecoder().decode(bytes.subarray(offset+30,offset+30+nameLength)),start=offset+30+nameLength+extraLength,data=bytes.slice(start,start+size);if(method!==0||!(flags&0x0800)||crc32(data)!==crc)throw new Error('Некорректная ZIP-часть');out.set(name,data);offset=start+size;}return out;
}
