/*
  Estandarizador Documental TSF · GitHub Pages (solo Word)
  Todo el procesamiento sucede en el navegador. No existen llamadas POST ni APIs de almacenamiento.
  JSZip: licencia MIT/GPL-3.0 — https://stuk.github.io/jszip/
  PDF.js: licencia Apache-2.0 — https://mozilla.github.io/pdf.js/
*/
'use strict';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture'
};
const MAX_BYTES = 25 * 1024 * 1024;
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const PDF_WORKER = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
const form = document.getElementById('standard-form');
const fileInput = document.getElementById('source-file');
const drop = document.getElementById('dropzone');
const filenameEl = document.getElementById('filename');
const notice = document.getElementById('notice');
const button = document.getElementById('generate');
const buttonLabel = document.getElementById('btn-label');

const esc = (s) => String(s == null ? '' : s).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&apos;');
const xmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const namespaceString = `xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"`;
const W = (el, name) => el.getAttributeNS(NS.w, name) || el.getAttribute('w:' + name) || '';
const by = (root, tag) => Array.from(root.getElementsByTagNameNS(NS.w,tag));

function xmlParse(xml, what='documento') {
  const x = new DOMParser().parseFromString(xml,'application/xml');
  if (x.getElementsByTagName('parsererror').length) throw Error('El XML de ' + what + ' no se pudo interpretar.');
  return x;
}
function xmlSave(dom) { return xmlDeclaration + new XMLSerializer().serializeToString(dom.documentElement); }
function node(doc, tag, attrs={}) {
  const el=doc.createElementNS(NS.w,'w:'+tag);
  for(const [k,v] of Object.entries(attrs)) el.setAttributeNS(NS.w,'w:'+k,String(v));
  return el;
}
function xmlFragment(doc, text) {
  const parsed = xmlParse(`<root ${namespaceString}>${text}</root>`,'fragmento');
  return Array.from(parsed.documentElement.childNodes).filter(n=>n.nodeType===1).map(n=>doc.importNode(n,true));
}
function appendXml(doc,parent,text) { for(const el of xmlFragment(doc,text)) parent.appendChild(el); }
function childrenNS(parent,name) { return Array.from(parent.childNodes).filter(x=>x.nodeType===1 && x.namespaceURI===NS.w && x.localName===name); }
function firstNS(parent,name) { return childrenNS(parent,name)[0] || null; }
function requireChild(doc,parent,name) { let c=firstNS(parent,name);if(!c){c=node(doc,name);parent.insertBefore(c,parent.firstChild);}return c; }
function hex(s){return String(s).replace(/[^A-Fa-f0-9]/g,'').slice(0,6).toUpperCase()||'000000';}

const styleDefinition = [
  ['TSFBody', 'Texto TSF', 20, '242424', false],
  ['TSFH1', 'Título TSF 1', 26, '242424', true],
  ['TSFH2', 'Título TSF 2', 23, '242424', true],
  ['TSFH3', 'Título TSF 3', 21, '343434', true]
].map(([id,label,size,color,bold])=>`<w:style w:type="paragraph" w:styleId="${id}" w:customStyle="1"><w:name w:val="${esc(label)}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="${bold?180:0}" w:after="${bold?100:125}" w:line="270" w:lineRule="auto"/>${bold?'<w:keepNext/>':''}</w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="${color}"/><w:sz w:val="${size}"/>${bold?'<w:b/>':''}</w:rPr></w:style>`).join('');

function normalizeStyles(stylesDoc){
  const root=stylesDoc.documentElement;
  const current=new Map(Array.from(root.getElementsByTagNameNS(NS.w,'style')).map(x=>[W(x,'styleId'),x]));
  for(const id of ['TSFBody','TSFH1','TSFH2','TSFH3']){
    if(current.has(id))root.removeChild(current.get(id));
  }
  appendXml(stylesDoc,root,styleDefinition);
  const normal=current.get('Normal');
  if(normal){
    const rp=requireChild(stylesDoc,normal,'rPr');
    for(const child of Array.from(rp.children)){
      if(['sz','rFonts','color','highlight','shd'].includes(child.localName))rp.removeChild(child);
    }
    appendXml(stylesDoc,rp,'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:color w:val="242424"/>');
  }
}
function stylesBlank(){return `${xmlDeclaration}<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:style>${styleDefinition}</w:styles>`;}

function p(text,style='TSFBody',extra=''){
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/>${extra}</w:pPr><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}
function run(t,{bold=false,color='242424',sz=20}={}){
  return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${sz}"/><w:color w:val="${hex(color)}"/>${bold?'<w:b/>':''}</w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
}
function cp(text, {style='TSFBody',before='',after='',extra=''}={}){
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/>${extra}</w:pPr>${before}${run(text)}${after}</w:p>`;
}
function tableCell(text,{shading='FFFFFF',label=false}={}){
  return `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/><w:shd w:fill="${shading}"/><w:tcMar><w:top w:w="90" w:type="dxa"/><w:start w:w="120" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:end w:w="120" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${run(text,{bold:label,sz:19})}</w:p></w:tc>`;
}
function tableRow(a,b){return `<w:tr>${tableCell(a,{label:true,shading:'FFF3CD'})}${tableCell(b,{shading:'FFFFFF'})}</w:tr>`;}
function infoTable(meta){
  const date = new Intl.DateTimeFormat('es-PE',{day:'2-digit',month:'2-digit',year:'numeric'}).format(new Date());
  const data=[['Tipo documental',meta.type],['Documento',meta.title],['Elaborador',meta.author],['Revisor',meta.reviewer],['Aprobador',meta.approver||'No consignado'],['Fecha de estandarización',date]];
  return `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="6" w:color="D7DCE2"/><w:bottom w:val="single" w:sz="6" w:color="D7DCE2"/><w:insideH w:val="single" w:sz="5" w:color="D7DCE2"/><w:insideV w:val="single" w:sz="5" w:color="D7DCE2"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>${data.map(x=>tableRow(x[0],x[1])).join('')}</w:tbl>`;
}
function imageMarkup(rId, {cx=1790000,cy=406000,id=1205}={}){
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Logo TOTAL"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="tsf_logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${esc(rId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
function cover(meta,logoRid){
  // La portada no altera ni borra texto del cuerpo fuente.
  return `<w:p><w:pPr><w:spacing w:before="250" w:after="380"/></w:pPr>${imageMarkup(logoRid,{cx:2600000,cy:590000,id:1206})}</w:p>`+
    `<w:p><w:pPr><w:spacing w:before="220" w:after="95"/></w:pPr>${run(meta.type.toUpperCase(),{bold:true,color:'7C818B',sz:21})}</w:p>`+
    `<w:p><w:pPr><w:spacing w:before="40" w:after="300"/></w:pPr>${run(meta.title.toUpperCase(),{bold:true,color:'242424',sz:36})}</w:p>`+
    `<w:p><w:pPr><w:spacing w:before="160" w:after="135"/></w:pPr>${run('INFORMACIÓN DEL DOCUMENTO',{bold:true,color:'242424',sz:22})}</w:p>`+
    infoTable(meta)+
    `<w:p><w:pPr><w:spacing w:before="300" w:after="90"/></w:pPr>${run('CONTROL DOCUMENTAL',{bold:true,sz:21})}</w:p>`+
    `<w:p><w:pPr><w:spacing w:before="30" w:after="100"/></w:pPr>${run('Documento generado en formato corporativo TSF. Su revisión y aprobación corresponden al área responsable.',{color:'6B7280',sz:18})}</w:p>`+
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
}
function headerXml(meta){
  return `${xmlDeclaration}<w:hdr ${namespaceString}><w:tbl><w:tblPr><w:tblW w:w="9400" w:type="dxa"/><w:tblBorders><w:bottom w:val="single" w:sz="13" w:color="F5BD00"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2800"/><w:gridCol w:w="6600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/></w:tcPr><w:p>${imageMarkup('rIdTSFLogoHeader',{cx:1770000,cy:402000,id:1299})}</w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="6600" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr>${run(meta.type.toUpperCase(),{bold:true,color:'424751',sz:17})}</w:p></w:tc></w:tr></w:tbl><w:p/></w:hdr>`;
}
function footerXml(){
  return `${xmlDeclaration}<w:ftr ${namespaceString}><w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:top w:val="single" w:sz="6" w:color="E4E6EA" w:space="6"/></w:pBdr></w:pPr>${run('TOTAL SERVICIOS FINANCIEROS  ·  Documento sujeto a revisión  ·  Página ',{color:'747C87',sz:16})}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;
}
function childRel(els,id){return Array.from(els).find(el=>el.getAttribute('Id')===id);}
function uniqueRelId(rels,base){let id=base,i=1;while(childRel(rels.getElementsByTagNameNS(NS.rel,'Relationship'),id))id=base+(i++);return id;}
function addRel(doc,root,id,type,target){const el=doc.createElementNS(NS.rel,'Relationship');el.setAttribute('Id',id);el.setAttribute('Type',`http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}`);el.setAttribute('Target',target);root.appendChild(el);}
function addContentType(doc,root,part,type){const exists=Array.from(root.getElementsByTagNameNS(NS.ct,'Override')).some(x=>x.getAttribute('PartName')===part);if(!exists){const x=doc.createElementNS(NS.ct,'Override');x.setAttribute('PartName',part);x.setAttribute('ContentType',type);root.appendChild(x);}}
function addPngDefault(doc,root){const exists=Array.from(root.getElementsByTagNameNS(NS.ct,'Default')).some(x=>x.getAttribute('Extension')==='png');if(!exists){const x=doc.createElementNS(NS.ct,'Default');x.setAttribute('Extension','png');x.setAttribute('ContentType','image/png');root.appendChild(x);}}
function ensureMargs(doc,sect){let m=firstNS(sect,'pgMar');if(!m){m=node(doc,'pgMar');sect.appendChild(m);}m.setAttributeNS(NS.w,'w:top','1450');m.setAttributeNS(NS.w,'w:bottom','1380');m.setAttributeNS(NS.w,'w:left','1350');m.setAttributeNS(NS.w,'w:right','1350');m.setAttributeNS(NS.w,'w:header','550');m.setAttributeNS(NS.w,'w:footer','560');}
function addSectionHeaderFooter(doc,hId,fId){
  let sects=by(doc,'sectPr');
  if(!sects.length){const body=by(doc,'body')[0];const sect=node(doc,'sectPr');body.appendChild(sect);sects=[sect];}
  for(const sect of sects){
    for(const el of Array.from(sect.childNodes))if(el.nodeType===1 && ['headerReference','footerReference','titlePg'].includes(el.localName))sect.removeChild(el);
    const hr=node(doc,'headerReference',{type:'default'});hr.setAttributeNS(NS.r,'r:id',hId);
    const fr=node(doc,'footerReference',{type:'default'});fr.setAttributeNS(NS.r,'r:id',fId);
    sect.insertBefore(fr,sect.firstChild);sect.insertBefore(hr,sect.firstChild);ensureMargs(doc,sect);
  }
}
function insertBeforeBody(doc,markup){const body=by(doc,'body')[0];if(!body)throw Error('No se encontró el cuerpo del Word.');const ref=body.firstChild;for(const el of xmlFragment(doc,markup))body.insertBefore(el,ref);}
function setPstyle(doc,paragraph,styleId){const pp=requireChild(doc,paragraph,'pPr');let s=firstNS(pp,'pStyle');if(!s){s=node(doc,'pStyle');pp.insertBefore(s,pp.firstChild);}s.setAttributeNS(NS.w,'w:val',styleId);
  for(const old of Array.from(pp.children))if(['shd','pBdr'].includes(old.localName))pp.removeChild(old);
}
function paragraphText(par){return Array.from(par.getElementsByTagNameNS(NS.w,'t')).map(n=>n.textContent).join('');}
function cleanDirectRunStyle(doc){
  for(const rp of by(doc,'rPr')){
    for(const c of Array.from(rp.childNodes)){
      if(c.nodeType===1 && ['color','highlight','shd','sz','szCs','rFonts'].includes(c.localName))rp.removeChild(c);
    }
  }
}
function detectHeading(par){
  const txt=paragraphText(par).trim();if(!txt)return 'TSFBody';
  const pp=firstNS(par,'pPr');const style=pp?firstNS(pp,'pStyle'):null;
  const prior=style?W(style,'val').toLowerCase():'';
  if(/heading\s*1|titulo\s*1|título\s*1|(^|\D)h1($|\D)/.test(prior))return 'TSFH1';
  if(/heading\s*2|titulo\s*2|título\s*2|(^|\D)h2($|\D)/.test(prior))return 'TSFH2';
  if(/heading\s*3|titulo\s*3|título\s*3|(^|\D)h3($|\D)/.test(prior))return 'TSFH3';
  if(txt.length>120)return 'TSFBody';
  // Numeraciones habituales de secciones, evitando confundir ítems numerados muy largos.
  if(/^\d+(?:\.\d+){0,3}\.?\s+[A-ZÁÉÍÓÚÑ][^.!?]{2,86}$/.test(txt))return txt.split('.').length>2?'TSFH2':'TSFH1';
  if(/^[A-ZÁÉÍÓÚÜÑ0-9\s,()/–-]{5,75}$/.test(txt)&&/[A-ZÁÉÍÓÚÑ]/.test(txt)&&txt.split(/\s+/).length<=10)return 'TSFH2';
  return 'TSFBody';
}
function improveTables(doc){
  for(const tbl of by(doc,'tbl')){
    const pr=requireChild(doc,tbl,'tblPr');
    for(const old of Array.from(pr.children)) if(old.localName==='tblBorders')pr.removeChild(old);
    appendXml(doc,pr,'<w:tblBorders><w:top w:val="single" w:sz="5" w:color="D8DDE4"/><w:bottom w:val="single" w:sz="5" w:color="D8DDE4"/><w:left w:val="single" w:sz="5" w:color="D8DDE4"/><w:right w:val="single" w:sz="5" w:color="D8DDE4"/><w:insideH w:val="single" w:sz="4" w:color="E2E5EA"/><w:insideV w:val="single" w:sz="4" w:color="E2E5EA"/></w:tblBorders>');
    const rows=childrenNS(tbl,'tr');
    rows.forEach((row,i)=>{
      if(i===0 && rows.length>1){const trPr=requireChild(doc,row,'trPr');if(!firstNS(trPr,'tblHeader'))trPr.appendChild(node(doc,'tblHeader',{val:'true'}));}
      for(const cell of childrenNS(row,'tc')){
        const tcPr=requireChild(doc,cell,'tcPr');
        for(const old of Array.from(tcPr.children))if(old.localName==='shd')tcPr.removeChild(old);
        // Primera fila en amarillo; las siguientes, blanco o gris claro.
        appendXml(doc,tcPr,`<w:shd w:fill="${(i===0&&rows.length>1)?'F5BD00':(i%2===0?'F9FAFB':'FFFFFF')}"/>`);
        if(i===0&&rows.length>1){for(const paragraph of childrenNS(cell,'p'))for(const run of by(paragraph,'r')){
          const rp=requireChild(doc,run,'rPr');if(!firstNS(rp,'b'))rp.appendChild(node(doc,'b'));
        }}
      }
    });
  }
}
function normalizeBodyContent(doc){
  // Solo el cuerpo: la información de cabeceras, pies y elementos incrustados se conserva en el paquete.
  cleanDirectRunStyle(doc);
  for(const par of by(doc,'p'))setPstyle(doc,par,detectHeading(par));
  improveTables(doc);
}
async function loadAsset(){const r=await fetch('assets/tsf_logo.png',{cache:'force-cache'});if(!r.ok)throw Error('No se pudo cargar el logo TSF.');return r.arrayBuffer();}
function fileExists(zip,name){return !!zip.file(name);}
async function patchPackage(zip,meta,{newDoc=false,bodyXml='' }={}){
  if(!fileExists(zip,'word/document.xml'))throw Error('El archivo Word no contiene el documento principal.');
  const document=xmlParse(await zip.file('word/document.xml').async('string'),'document.xml');
  const styles=fileExists(zip,'word/styles.xml')?xmlParse(await zip.file('word/styles.xml').async('string'),'styles.xml'):xmlParse(stylesBlank());
  normalizeStyles(styles);
  const relsPath='word/_rels/document.xml.rels';
  const rels=fileExists(zip,relsPath)?xmlParse(await zip.file(relsPath).async('string'),'relaciones'):xmlParse(`${xmlDeclaration}<Relationships xmlns="${NS.rel}"/>`);
  const ct=xmlParse(await zip.file('[Content_Types].xml').async('string'),'tipos MIME');
  const idHeader=uniqueRelId(rels,'rIdTSFHdr');
  const idFooter=uniqueRelId(rels,'rIdTSFFtr');
  const idCoverLogo=uniqueRelId(rels,'rIdTSFLogo');
  const root=rels.documentElement;
  addRel(rels,root,idHeader,'header','headerTSF.xml');
  addRel(rels,root,idFooter,'footer','footerTSF.xml');
  addRel(rels,root,idCoverLogo,'image','media/tsf_logo.png');
  if(!Array.from(root.getElementsByTagNameNS(NS.rel,'Relationship')).some(el=>el.getAttribute('Type')?.endsWith('/styles'))){addRel(rels,root,uniqueRelId(rels,'rIdTSFStyles'),'styles','styles.xml');}

  normalizeBodyContent(document);
  // En PDF, el cuerpo se construyó antes de esta llamada.
  insertBeforeBody(document,cover(meta,idCoverLogo));
  addSectionHeaderFooter(document,idHeader,idFooter);
  // Asegurar que el documento tenga namespaces r/drawing para la portada importada.
  for(const [prefix,ns] of Object.entries({r:NS.r,wp:NS.wp,a:NS.a,pic:NS.pic})){
    if(!document.documentElement.lookupNamespaceURI(prefix))document.documentElement.setAttribute('xmlns:'+prefix,ns);
  }

  const image=await loadAsset();
  zip.file('word/media/tsf_logo.png',image);
  zip.file('word/document.xml',xmlSave(document));
  zip.file('word/styles.xml',xmlSave(styles));
  zip.file(relsPath,xmlSave(rels));
  zip.file('word/headerTSF.xml',headerXml(meta));
  zip.file('word/footerTSF.xml',footerXml());
  zip.file('word/_rels/headerTSF.xml.rels',`${xmlDeclaration}<Relationships xmlns="${NS.rel}"><Relationship Id="rIdTSFLogoHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/tsf_logo.png"/></Relationships>`);
  addPngDefault(ct,ct.documentElement);
  addContentType(ct,ct.documentElement,'/word/headerTSF.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  addContentType(ct,ct.documentElement,'/word/footerTSF.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
  addContentType(ct,ct.documentElement,'/word/styles.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
  zip.file('[Content_Types].xml',xmlSave(ct));
  return zip;
}
async function docxToDocx(buffer,meta){
  const zip=await JSZip.loadAsync(buffer);
  const warnings=[];
  if(fileExists(zip,'word/comments.xml'))warnings.push('El documento contiene comentarios de revisión que permanecen en el archivo original.');
  const documentXml=await (zip.file('word/document.xml')?.async('string'));
  if(!documentXml)throw Error('El archivo seleccionado no es un Word DOCX válido.');
  if(/w:altChunk|w:object|w:txbxContent|w:sdt/.test(documentXml))warnings.push('El documento contiene objetos o bloques especiales: revisar su presentación en Word.');
  if(/w:ins|w:del/.test(documentXml))warnings.push('El documento contiene control de cambios: revisar en Word.');
  await patchPackage(zip,meta);
  const out=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:5},mimeType:DOCX_MIME});
  // Confirmamos que el paquete resultante sea legible y contenga sus componentes esenciales.
  const outZip=await JSZip.loadAsync(out);
  if(!outZip.file('word/headerTSF.xml')||!outZip.file('word/document.xml'))throw Error('El Word resultante no pasó la comprobación de integridad.');
  return {out,warnings};
}

function blankDocxZip(body){
  const zip=new JSZip();
  zip.file('[Content_Types].xml',`${xmlDeclaration}<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
  zip.file('_rels/.rels',`${xmlDeclaration}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels',`${xmlDeclaration}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('word/document.xml',`${xmlDeclaration}<w:document ${namespaceString}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`);
  zip.file('word/styles.xml',stylesBlank());
  zip.file('docProps/core.xml',`${xmlDeclaration}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Documento estandarizado TSF</dc:title><dc:creator>Estandarizador Documental TSF</dc:creator></cp:coreProperties>`);
  return zip;
}
function safeSpaces(s){return String(s).replace(/\u0000/g,'').replace(/[\u0001-\u0008\u000b\u000c\u000e-\u001f]/g,'');}
function pdfLines(items){
  const relevant=items.filter(x=>typeof x.str==='string' && x.str.length>0).map(x=>({text:safeSpaces(x.str),x:x.transform?.[4]||0,y:x.transform?.[5]||0,size:Math.abs(x.transform?.[3]||x.height||10),width:Math.max(0,x.width||0),hasEOL:!!x.hasEOL}));
  if(!relevant.length)return [];
  // Se agrupan los fragmentos por coordenada vertical para recuperar el orden de lectura.
  relevant.sort((a,b)=>b.y-a.y || a.x-b.x);
  const rows=[];
  for(const item of relevant){const last=rows[rows.length-1];if(last&&Math.abs(last.y-item.y)<2.7){last.fragments.push(item);last.y=(last.y*(last.fragments.length-1)+item.y)/last.fragments.length;}else{rows.push({y:item.y,fragments:[item]});}}
  return rows.map(row=>{
    row.fragments.sort((a,b)=>a.x-b.x);
    let text='',prev=null;
    for(const frag of row.fragments){
      if(prev){const gap=frag.x-(prev.x+prev.width);if(gap>2) text+=gap>35?'    ':' ';else if(gap>=-1.5 && !text.endsWith(' ') && !frag.text.startsWith(' ') && !prev.text.endsWith(' '))text+=' ';}
      text+=frag.text;prev=frag;
    }
    return {text:text.trimEnd(),y:row.y,size:Math.max(...row.fragments.map(f=>f.size))};
  }).filter(r=>r.text.trim());
}
function headingFromPdf(line,base){
  const t=line.text.trim();if(!t||t.length>145)return 'TSFBody';
  if(line.size>base*1.35 && t.length<115)return 'TSFH1';
  if(line.size>base*1.17 && t.length<110)return 'TSFH2';
  if(/^\d+(?:\.\d+){0,3}\.?\s+[A-ZÁÉÍÓÚÑ]/.test(t)&&t.length<110)return 'TSFH2';
  if(t.length<80 && /^[A-ZÁÉÍÓÚÑ0-9\s,()/–-]{5,80}$/.test(t)&&t.split(/\s+/).length<=9)return 'TSFH2';
  return 'TSFBody';
}
async function pdfToDocx(buffer,meta,onProgress){
  if(typeof pdfjsLib==='undefined')throw Error('No se pudo cargar PDF.js. Verifica tu conexión a Internet o que la red no bloquee cdnjs.');
  pdfjsLib.GlobalWorkerOptions.workerSrc=PDF_WORKER;
  const loadTask=pdfjsLib.getDocument({data:new Uint8Array(buffer)});
  let pdf;
  try{pdf=await loadTask.promise;}catch(e){throw Error('No se pudo abrir el PDF. Puede estar protegido por contraseña o dañado. '+e.message);}
  const num=Math.min(pdf.numPages,120);
  let body='',allChars=0;
  const warnings=['La conversión PDF → Word reconstruye texto: gráficos, fotografías, columnas y tablas complejas pueden requerir revisión manual.'];
  if(pdf.numPages>120)throw Error('El PDF supera 120 páginas y no puede convertirse en esta versión.');
  for(let i=1;i<=num;i++){
    onProgress(`Leyendo página ${i} de ${num}…`);
    const pg=await pdf.getPage(i),txt=await pg.getTextContent();
    const lines=pdfLines(txt.items);
    const nChars=lines.reduce((a,l)=>a+l.text.trim().length,0);
    allChars+=nChars;
    if(nChars<5)throw Error(`La página ${i} no tiene texto seleccionable. El PDF parece escaneado o contiene texto como imagen; esta versión no incluye OCR.`);
    const sizes=lines.map(l=>l.size).filter(x=>x>1).sort((a,b)=>a-b);
    const base=sizes[Math.floor(sizes.length/2)]||10;
    let lastY=null;
    for(const line of lines){
      const style=headingFromPdf(line,base);
      if(lastY!==null&&Math.abs(lastY-line.y)>base*1.8)body+='<w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>';
      body+=p(line.text,style);
      lastY=line.y;
    }
    if(i<num)body+='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    // Recolección de páginas sin subir el documento; se libera el recurso de PDF.js al finalizar.
    pg.cleanup();
  }
  if(allChars<15)throw Error('No se detectó suficiente texto seleccionable para construir un Word editable.');
  let zip=blankDocxZip(body);
  await patchPackage(zip,meta,{newDoc:true});
  const out=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:5},mimeType:DOCX_MIME});
  const validation=await JSZip.loadAsync(out);
  if(!validation.file('word/document.xml'))throw Error('No se pudo validar el Word generado.');
  try{await pdf.destroy();}catch(_){/* PDF.js ya finalizó; no bloquear la descarga */}
  return {out,warnings};
}
function cleanFileName(name){return (name||'Documento').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9 _-]/g,'').trim().replace(/\s+/g,'_').slice(0,85)||'Documento';}
function showNotice(message,kind='success') {notice.hidden=false;notice.className='notice '+kind;notice.textContent=message;}
function showFile(file) {
  if(!file){filenameEl.textContent='Ningún archivo seleccionado';drop.classList.remove('hasfile');return;}
  filenameEl.textContent=`${file.name} · ${(file.size/1024/1024).toFixed(2)} MB`;
  drop.classList.add('hasfile');
}
function saveDoc(blob,name){const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),3000);}
fileInput.addEventListener('change',()=>showFile(fileInput.files[0]));
for(const type of ['dragenter','dragover'])drop.addEventListener(type,e=>{e.preventDefault();drop.classList.add('dragover');});
for(const type of ['dragleave','drop'])drop.addEventListener(type,e=>{e.preventDefault();drop.classList.remove('dragover');});
drop.addEventListener('drop',e=>{if(e.dataTransfer.files.length){const x=new DataTransfer();x.items.add(e.dataTransfer.files[0]);fileInput.files=x.files;showFile(fileInput.files[0]);}});

form.addEventListener('submit',async e=>{
  e.preventDefault();
  const f=fileInput.files[0];
  if(!f){showNotice('Selecciona primero un Word o PDF.','error');return;}
  if(!form.checkValidity()) {form.reportValidity();return;}
  if(f.size>MAX_BYTES){showNotice('El archivo supera el límite de 25 MB.','error');return;}
  const ext=f.name.toLowerCase().split('.').pop();
  if(!['docx','pdf'].includes(ext)){showNotice('Formato no admitido. Utiliza Word (.docx) o PDF (.pdf).','error');return;}
  if(typeof JSZip==='undefined'){showNotice('No se pudo cargar el motor DOCX. Actualiza la página y verifica los archivos publicados.','error');return;}
  const meta={type:document.getElementById('doc-type').value,title:document.getElementById('doc-name').value.trim(),author:document.getElementById('author').value.trim(),reviewer:document.getElementById('reviewer').value.trim(),approver:document.getElementById('approver').value.trim()};
  if(!meta.type||!meta.title||!meta.author||!meta.reviewer){showNotice('Completa el tipo, nombre, elaborador y revisor.','error');return;}
  button.disabled=true;buttonLabel.textContent='Procesando…';
  showNotice('Preparando tu documento; todo el procesamiento ocurre en este navegador.');
  try{
    const bytes=await f.arrayBuffer();
    const result=ext==='pdf'?await pdfToDocx(bytes,meta,msg=>showNotice(msg)):await docxToDocx(bytes,meta);
    const outName=`${cleanFileName(meta.title)}_Estandarizado_TSF.docx`;
    saveDoc(result.out,outName);
    if(result.warnings.length)showNotice('Word generado y descargado. Revisión recomendada: '+result.warnings.join(' '),'warning');
    else showNotice('Word generado y descargado correctamente. El documento conserva su cuerpo original y aplica el formato TSF.');
  }catch(err){console.error('Error de estandarización:',err);showNotice(err.message||'No fue posible generar el documento.','error');}
  finally{button.disabled=false;buttonLabel.textContent='Generar Word';}
});
