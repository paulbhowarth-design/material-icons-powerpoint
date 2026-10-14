/* Material Symbols browser. Catalogue and SVGs provided by Iconify's public API,
   based on Google's open-source Material Symbols. */
const API='https://api.iconify.design';
const $=id=>document.getElementById(id);
let names=[],filtered=[],visible=80,ready=false;
const prefix=()=>`material-symbols`;
const iconName=(name)=>{
  const style=$('style').value;
  // Iconify names: default outlined, -rounded, -sharp. Some names already carry suffixes.
  return name.replace(/-(rounded|sharp|outlined)$/,'')+(style==='outlined'?'':`-${style}`);
};
const iconUrl=(name)=>`${API}/${prefix()}/${encodeURIComponent(iconName(name))}.svg?color=${encodeURIComponent($('color').value)}`;
function render(){const query=$('q').value.trim().toLowerCase().replace(/[ _]+/g,'-');
 filtered=names.filter(n=>n.includes(query)&&!n.endsWith('-rounded')&&!n.endsWith('-sharp'));
 $('status').textContent=`${filtered.length.toLocaleString()} matching icons`;
 const fragment=document.createDocumentFragment();
 for(const n of filtered.slice(0,visible)){const b=document.createElement('button');b.className='tile';b.title=`Insert ${n}`;const im=document.createElement('img');im.src=iconUrl(n);im.loading='lazy';im.alt='';const label=document.createElement('span');label.textContent=n;b.append(im,label);b.addEventListener('click',()=>insert(n));fragment.append(b)}
 $('grid').replaceChildren(fragment);$('more').hidden=visible>=filtered.length;
}
async function load(){try{const r=await fetch(`${API}/collection?prefix=material-symbols`);if(!r.ok)throw Error(`HTTP ${r.status}`);const data=await r.json();names=[...(data.uncategorized||[]),...Object.values(data.categories||{}).flat()].filter((v,i,a)=>a.indexOf(v)===i).sort();if(!names.length)throw Error('Empty icon collection');ready=true;render()}catch(e){$('status').textContent=`Could not load catalogue: ${e.message}. Check your connection.`}}
function toBase64Unicode(str){return btoa(Array.from(new TextEncoder().encode(str),b=>String.fromCharCode(b)).join(''))}
async function insert(name){$('status').textContent=`Inserting ${name}…`;try{
 const r=await fetch(iconUrl(name));if(!r.ok)throw Error('Icon not available in this style');let svg=await r.text();if(!svg.includes('<svg'))throw Error('Invalid SVG');
 // Ensure explicit pixel dimensions for PowerPoint.
 svg=svg.replace(/(<svg\b[^>]*?)\swidth="[^"]*"/,'$1').replace(/(<svg\b[^>]*?)\sheight="[^"]*"/,'$1').replace(/<svg\b/, '<svg width="96" height="96"');
 const encoded=toBase64Unicode(svg);
 if(typeof PowerPoint==='undefined'||!PowerPoint.run)throw Error('Open this add-in inside PowerPoint');
 await PowerPoint.run(async context=>{
   const slides=context.presentation.getSelectedSlides();slides.load('items');await context.sync();
   if(!slides.items.length)throw Error('Select a slide first');
   const shape=slides.items[0].shapes.addImage(encoded);
   shape.left=90;shape.top=90;shape.width=96;shape.height=96;
   await context.sync();
 });
 $('status').textContent=`Inserted ${name}`;
 }catch(e){$('status').textContent=`Insert failed: ${e.message}`;console.error(e)} }
$('q').addEventListener('input',()=>{visible=80;if(ready)render()});
$('style').addEventListener('change',()=>{if(ready)render()});$('color').addEventListener('input',()=>{if(ready)render()});$('more').addEventListener('click',()=>{visible+=80;render()});
Office.onReady(()=>load());
