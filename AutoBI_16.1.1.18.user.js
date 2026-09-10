// ==UserScript==
// @name         AutoBI 16.1.1.18
// @namespace    https://github.com/PhamngocNDH/AutoBI
// @version      16.1.1.18
// @description  AutoBI 16.1.1.18
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
// ==/UserScript==

/* AutoBI 16.1.1.18 */
(function(){
'use strict';
if(window.top!==window)return;
const D='tgdd_data_cache_v30',P='TGDD_BI_STORE_PERMANENT_CONFIG_GM_V1',S='autobi_tc18_state',L='autobi_tc18_last';
const sl=m=>new Promise(r=>setTimeout(r,m)),nm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim(),tx=e=>(e?.innerText||e?.textContent||'').trim(),vis=e=>!!e&&!!(e.offsetWidth||e.offsetHeight||e.getClientRects().length);
const gv=(k,d=null)=>{try{return GM_getValue(k,d)}catch(_){return d}},sv=(k,v)=>{try{GM_setValue(k,v)}catch(_){}},dv=k=>{try{GM_deleteValue(k)}catch(_){}};
function obj(v){if(typeof v==='string')try{v=JSON.parse(v)}catch(_){}return v&&typeof v==='object'?v:{}}
function cfg(){let c=obj(gv('tgdd_active_run_config'));if(Object.keys(c).length)return c;c=obj(gv(P));if(!Object.keys(c).length)try{c=obj(localStorage.getItem(P))}catch(_){};if(Object.keys(c).length)sv('tgdd_active_run_config',c);return c}
function shops(c){let a=[];for(let i=1;i<=5;i++)if(c['shop'+i])a.push({i,key:'shop'+i,name:String(c['shop'+i])});return a}
function toast(m){try{const w=typeof unsafeWindow!=='undefined'?unsafeWindow:window;if(w.UI?.showToast)return w.UI.showToast(m,3200)}catch(_){}console.log('[AutoBI 18]',m)}
async function wait(fn,to=15000){let e=Date.now()+to;while(Date.now()<e){try{let v=fn();if(v)return v}catch(_){}await sl(300)}return null}
function box(c){let a=[...document.querySelectorAll('.ant-select-selector,[role="combobox"],.ant-select')].filter(vis),b=a.find(e=>nm(tx(e)).includes('sieu thi'));if(b)return b;for(const s of shops(c)){let q=nm(s.name).split(' ').slice(-2).join(' ');b=a.find(e=>nm(tx(e)).includes(q));if(b)return b}let labs=[...document.querySelectorAll('label,span,div')].filter(e=>vis(e)&&nm(tx(e))==='sieu thi');for(const l of labs){let p=l;for(let n=0;p&&n<7;n++,p=p.parentElement){let x=p.querySelector?.('.ant-select-selector,[role="combobox"],.ant-select');if(x&&vis(x))return x}}return a[0]}
async function open(c){let b=await wait(()=>box(c),10000);if(!b)return null;(b.closest('.ant-select')||b).click();await sl(350);return wait(()=>[...document.querySelectorAll('.ant-select-dropdown:not(.ant-select-dropdown-hidden),[role="listbox"]')].filter(vis).pop(),5000)}
function clickText(root,s){let q=nm(s),a=[...root.querySelectorAll('button,a,label,span,div,li')].filter(vis),e=a.find(x=>nm(tx(x))===q)||a.find(x=>nm(tx(x)).includes(q));if(!e)return false;try{(e.closest('button,a,label,li,[role="option"]')||e).click();return true}catch(_){return false}}
async function clear(c){let d=await open(c);if(!d)return false;clickText(d,'Bỏ chọn tất cả')||clickText(d,'Bỏ chọn đang lọc');await sl(400);document.body.click();await sl(300);return true}
async function choose(c,s){let d=await open(c);if(!d)return false;clickText(d,'Bỏ chọn tất cả')||clickText(d,'Bỏ chọn đang lọc');await sl(300);let full=nm(s.name),nums=String(s.name).match(/\b\d{3,6}\b/g)||[],ids=[c['shopId'+s.i],c['warehouse'+s.i],c['warehouseId'+s.i],...nums].filter(Boolean).map(x=>nm(x)),rows=[...d.querySelectorAll('label,li,[role="option"],div')].filter(vis),r=rows.find(e=>{let t=nm(tx(e));return t&&(t.includes(full)||ids.some(x=>x&&t.includes(x)))});if(!r){let q=full.split(' ').slice(-2).join(' ');r=rows.find(e=>nm(tx(e)).includes(q))}if(!r){document.body.click();return false}try{(r.querySelector('input[type="checkbox"]')||r.closest('label,li,[role="option"]')||r).click()}catch(_){return false}await sl(400);document.body.click();await sl(300);return true}
async function xem(){let b=await wait(()=>[...document.querySelectorAll('button,a')].filter(e=>vis(e)&&nm(tx(e))==='xem')[0],6000);if(!b)return false;b.click();await sl(4500);return true}
function ready(c){let z=obj(gv(D,{})),ss=shops(c);return ss.length&&ss.every(s=>Number(z.link7?.[s.key]?._SHOP_DT)>0)}
async function scan(){let st=obj(gv(S));if(!st.run)return;let c=await wait(()=>{let x=cfg();return shops(x).length?x:null},12000);if(!c){st.err='Không đọc được cấu hình siêu thị';sv(S,st);location.href=st.back||'/dashboard/thi-dua';return}let ss=shops(c);toast('⏳ Đang lấy Trả chậm...');clickText(document,'Tỷ trọng trả góp');await sl(500);clickText(document,'Lũy kế');await sl(1200);await wait(()=>[...document.querySelectorAll('table')].some(t=>nm(tx(t)).includes('dt sieu thi')),18000);await sl(2200);for(let x=0;x<ss.length;x++){toast('🏬 '+(x+1)+'/'+ss.length+' '+ss[x].name);if(await choose(c,ss[x]))await xem();else console.warn('[AutoBI 18] Không chọn được',ss[x].name)}await clear(c);await xem();sv(L,Date.now());st.run=false;st.done=true;sv(S,st);toast('✅ Đã lấy xong Trả chậm');await sl(900);location.href=st.back||'/dashboard/thi-dua'}
async function start(){let c=await wait(()=>{let x=cfg();return shops(x).length?x:null},12000);if(!c)return;let st=obj(gv(S));if(st.done){dv(S);toast('✅ Đã cập nhật dữ liệu Trả chậm');return}if(st.err){toast('⚠️ '+st.err);dv(S);return}if(ready(c))return;if(Date.now()-Number(gv(L,0)||0)<120000)return;sv(S,{run:true,done:false,back:location.pathname+location.search+location.hash});toast('🔄 Chuyển sang Báo cáo Trả chậm để lấy dữ liệu...');await sl(900);location.href='/dashboard/tra-cham'}
if(location.pathname.includes('/dashboard/tra-cham'))setTimeout(scan,1800);else if(location.pathname.includes('/dashboard/thi-dua'))setTimeout(start,5000);
})();
