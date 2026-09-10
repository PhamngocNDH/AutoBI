// ==UserScript==
// @name         AutoBI 16.1.1.17
// @namespace    https://github.com/PhamngocNDH/AutoBI
// @version      16.1.1.17
// @description  AutoBI 16.1.1.17
// @author       38967 - Mr Phạm
// @match        https://crm.thegioididong.com/*
// @match        https://baocao.dienmayxanh.com/*
// @include      https://baocao.dienmayxanh.com/*
// @connect      *
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_16.1.1.16.user.js
// ==/UserScript==

/* AutoBI 16.1.1.17 */
(function(){
'use strict';
if(window.top!==window||!location.pathname.includes('/dashboard/thi-dua'))return;
const DK='tgdd_data_cache_v30',LK='autobi_tra_cham_auto_scan_v17',RK='autobi_tra_cham_reload_once_v17';
const sl=m=>new Promise(r=>setTimeout(r,m));
const nm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const nu=v=>{const n=parseFloat(String(v??'').replace(/,/g,'').replace(/%/g,'').trim());return Number.isFinite(n)?n:0};
const tx=e=>(e?.innerText||e?.textContent||'').trim(),vis=e=>!!e&&!!(e.offsetWidth||e.offsetHeight||e.getClientRects().length);
function cfg(){let c=null;try{c=GM_getValue('tgdd_active_run_config',null)}catch(_){};if(!c||typeof c!=='object')try{c=(typeof unsafeWindow!=='undefined'?unsafeWindow:window).tgdd_active_user_config}catch(_){};return c&&typeof c==='object'?c:{}}
function toast(m){try{const w=typeof unsafeWindow!=='undefined'?unsafeWindow:window;if(w.UI&&typeof w.UI.showToast==='function')return w.UI.showToast(m,2600)}catch(_){}console.log('[AutoBI Trả chậm]',m)}
function toks(c,i){const a=[c['shop'+i],c['shopId'+i],c['warehouse'+i],c['warehouseId'+i]].filter(Boolean).map(String),b=String(c['shop'+i]||'').match(/\b\d{3,6}\b/g)||[];return[...new Set([...a,...b])].map(nm).filter(Boolean)}
function shopKey(s,c){const t=nm(s);for(let i=1;i<=5;i++){if(!c['shop'+i])continue;const f=nm(c['shop'+i]),tail=f.split(' ').slice(-3).join(' ');if((f&&(t.includes(f)||f.includes(t)))||(tail&&t.includes(tail))||toks(c,i).some(x=>/^\d+$/.test(x)&&t.includes(x)))return'shop'+i}return null}
const table=d=>[...d.querySelectorAll('table')].find(t=>{const h=nm(tx(t.querySelector('thead'))||tx(t));return h.includes('dt tra gop')&&h.includes('dt sieu thi')})||null;
async function wait(fn,to=20000,st=300){const e=Date.now()+to;while(Date.now()<e){try{const v=fn();if(v)return v}catch(_){}await sl(st)}return null}
const sig=t=>nm(tx(t)).slice(0,3000);
async function stable(d,old='',to=22000){const e=Date.now()+to;let last='',same=0,t=null;while(Date.now()<e){t=table(d);if(t){const s=sig(t);if(s&&s!==old&&s===last)same++;else same=0;last=s;if(s&&(s!==old||!old)&&same>=2)return t}await sl(500)}return t}
function saveTotal(t,c){const z=GM_getValue(DK,{})||{};z.link7=z.link7||{};z.link2=z.link2||{};let n=0,si=0,ss=0;t.querySelectorAll('tbody tr:not(.ant-table-measure-row)').forEach(r=>{const a=[...r.querySelectorAll('td')];if(a.length<4)return;const k=shopKey(tx(a[0]),c);if(!k)return;const tgop=nu(tx(a[1])),dst=nu(tx(a[2])),tg=nu(tx(a[3])),hc=nu(tx(a[4])),fe=nu(tx(a[6]));z.link7[k]=z.link7[k]||{};Object.assign(z.link7[k],{_SHOP_DT:dst,_SHOP_HC:hc,_SHOP_FE:fe,_SHOP_TOTAL_:{tg,dt_sieu_thi:dst,dt_tragop:tgop}});z.link2[k]=z.link2[k]||{};z.link2[k].tg=tg;si+=tgop;ss+=dst;n++});if(n){z.link2.total=z.link2.total||{};if(ss>0)z.link2.total.tg=si/ss*100;GM_setValue(DK,z)}return n}
function staffKey(z,k,name){const n=nm(name),a=Object.keys(z.link7?.[k]||{}).filter(x=>!x.startsWith('_')),b=Object.keys(z.link6?.[k]?.revenue||{});return[...new Set([...a,...b])].find(x=>{const q=nm(x).replace(/^\d+\s+/,'');return q===n||q.includes(n)||n.includes(q)})||name}
function saveStaff(t,k){const z=GM_getValue(DK,{})||{};z.link7=z.link7||{};z.link7[k]=z.link7[k]||{};let n=0;t.querySelectorAll('tbody tr:not(.ant-table-measure-row)').forEach(r=>{const a=[...r.querySelectorAll('td')];if(a.length<4)return;const name=tx(a[0]);if(!name||nm(name).includes('tong'))return;const sk=staffKey(z,k,name);z.link7[k][sk]=Object.assign({},z.link7[k][sk]||{},{tg:nu(tx(a[3])),dt_sieu_thi:nu(tx(a[2])),dt_tragop:nu(tx(a[1])),hc:nu(tx(a[4])),fe:nu(tx(a[6]))});n++});if(n)GM_setValue(DK,z);return n}
function view(d){return[...d.querySelectorAll('button,a')].filter(vis).find(e=>nm(tx(e))==='xem')||null}
function select(d){const labs=[...d.querySelectorAll('label,span,div')].filter(vis).filter(e=>nm(tx(e))==='sieu thi');for(const l of labs){let p=l.parentElement;for(let n=0;p&&n<5;n++,p=p.parentElement){const s=p.querySelector('.ant-select-selector,[role="combobox"],.ant-select');if(s&&vis(s))return s}}return[...d.querySelectorAll('.ant-select-selector,[role="combobox"]')].filter(vis).find(e=>nm(tx(e)).includes('sieu thi'))||null}
async function open(d){const s=await wait(()=>select(d),12000);if(!s)return null;try{s.click()}catch(_){}return await wait(()=>[...d.querySelectorAll('.ant-select-dropdown,[role="listbox"]')].filter(vis).pop(),6000)}
function action(p,s){const q=nm(s);return[...p.querySelectorAll('button,a,label,span,div')].filter(vis).find(e=>nm(tx(e))===q&&(!e.children.length||['BUTTON','A','LABEL'].includes(e.tagName)))||[...p.querySelectorAll('button,a,label')].filter(vis).find(e=>nm(tx(e)).includes(q))}
async function clear(d){const p=await open(d);if(!p)return false;const b=action(p,'Bỏ chọn tất cả')||action(p,'Bỏ chọn đang lọc');if(b){try{b.click()}catch(_){}await sl(350)}else{for(const c of p.querySelectorAll('input[type="checkbox"]'))if(c.checked)try{c.click()}catch(_){}await sl(250)}try{d.body.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))}catch(_){};await sl(250);return true}
async function choose(d,c,i){const p=await open(d);if(!p)return false;const b=action(p,'Bỏ chọn tất cả')||action(p,'Bỏ chọn đang lọc');if(b){try{b.click()}catch(_){}await sl(300)}const f=nm(c['shop'+i]),ts=toks(c,i),rows=[...p.querySelectorAll('label,li,[role="option"],div')].filter(vis);const r=rows.find(e=>{const t=nm(tx(e));return t&&((f&&t.includes(f))||ts.some(x=>x&&t.includes(x)))});if(!r)return false;try{(r.querySelector('input[type="checkbox"]')||r).click()}catch(_){return false}await sl(350);try{d.body.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))}catch(_){};await sl(250);return true}
async function apply(d,old=''){const b=await wait(()=>view(d),6000);if(!b)return null;try{b.click()}catch(_){return null}await sl(700);return await stable(d,old,22000)}
async function run(){const c=await wait(()=>{const x=cfg();return Object.keys(x).length?x:null},15000);if(!c)return;const shops=[];for(let i=1;i<=5;i++)if(c['shop'+i])shops.push(i);if(!shops.length)return;if(Date.now()-Number(GM_getValue(LK,0)||0)<120000)return;toast('⏳ Đang tự động lấy dữ liệu Trả chậm...');const f=document.createElement('iframe');f.src='/dashboard/tra-cham';f.style.cssText='position:fixed!important;width:1px!important;height:1px!important;opacity:0!important;pointer-events:none!important;left:-9999px!important;top:-9999px!important;border:0!important';document.body.appendChild(f);try{await wait(()=>f.contentDocument?.body,12000);const d=f.contentDocument;let t=await stable(d,'',25000);if(!t)throw Error('Không tải được bảng');let n=saveTotal(t,c);if(n<shops.length){const old=sig(t);await clear(d);t=await apply(d,old);if(t)n=saveTotal(t,c)}if(!n)throw Error('Không đọc được Tổng cụm');for(let p=0;p<shops.length;p++){const i=shops[p],k='shop'+i;toast('🏬 Trả chậm ['+(p+1)+'/'+shops.length+']: '+c[k]);const old=t?sig(t):'';if(!await choose(d,c,i))continue;t=await apply(d,old);if(t&&nm(tx(t.querySelector('thead'))).includes('nhan vien'))saveStaff(t,k)}try{const old=t?sig(t):'';await clear(d);t=await apply(d,old);if(t)saveTotal(t,c)}catch(_){}GM_setValue(LK,Date.now());toast('✅ Đã cập nhật xong dữ liệu Trả chậm');if(!sessionStorage.getItem(RK)){sessionStorage.setItem(RK,'1');await sl(900);location.reload()}}catch(e){console.warn('[AutoBI Trả chậm Auto]',e);toast('⚠️ Chưa tự lấy được dữ liệu Trả chậm')}finally{try{f.remove()}catch(_){}}}
const start=()=>setTimeout(run,4500);if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();
