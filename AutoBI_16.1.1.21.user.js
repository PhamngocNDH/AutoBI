// ==UserScript==
// @name         AutoBI 16.1.1.21
// @namespace    https://github.com/PhamngocNDH/AutoBI
// @version      16.1.1.21
// @description  AutoBI 16.1.1.21
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

(function(){
'use strict';
if(window.top!==window)return;
const DK='tgdd_data_cache_v30',CK='TGDD_BI_STORE_PERMANENT_CONFIG_GM_V1',ST='autobi_tc21_state',LAST='autobi_tc21_last';
const REGION_CODE='761',AREA_CODE='3866';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const text=e=>(e?.innerText||e?.textContent||'').trim();
const num=v=>{const n=parseFloat(String(v??'').replace(/,/g,'').replace(/%/g,'').trim());return Number.isFinite(n)?n:0};
const vis=e=>!!e&&!!(e.offsetWidth||e.offsetHeight||e.getClientRects().length);
const gv=(k,d=null)=>{try{return GM_getValue(k,d)}catch(_){return d}},sv=(k,v)=>{try{GM_setValue(k,v)}catch(_){}},dv=k=>{try{GM_deleteValue(k)}catch(_){}};
function obj(v){if(typeof v==='string')try{v=JSON.parse(v)}catch(_){}return v&&typeof v==='object'?v:{}}
function cfg(){let c=obj(gv('tgdd_active_run_config'));if(Object.keys(c).length)return c;c=obj(gv(CK));if(!Object.keys(c).length)try{c=obj(localStorage.getItem(CK))}catch(_){}return c}
function shops(c){const a=[];for(let i=1;i<=5;i++)if(c['shop'+i])a.push({key:'shop'+i,name:String(c['shop'+i])});return a}
function toast(m){try{const w=typeof unsafeWindow!=='undefined'?unsafeWindow:window;if(w.UI?.showToast)return w.UI.showToast(m,3500)}catch(_){}console.log('[AutoBI 21]',m)}
async function wait(fn,to=20000,step=250){const end=Date.now()+to;while(Date.now()<end){try{const v=fn();if(v)return v}catch(_){}await sleep(step)}return null}
function findSelect(label){const q=norm(label),sels=[...document.querySelectorAll('.ant-select')].filter(vis);let s=sels.find(e=>norm(text(e)).includes(q));if(s)return s;const labs=[...document.querySelectorAll('span,div,label')].filter(e=>vis(e)&&norm(text(e))===q);for(const l of labs){let p=l;for(let i=0;p&&i<5;i++,p=p.parentElement){s=p.querySelector?.('.ant-select');if(s&&vis(s))return s}}const idx=q==='vung'?0:q==='khu vuc'?1:q==='sieu thi'?2:-1;return idx>=0?sels[idx]||null:null}
async function openSelect(label){const s=await wait(()=>findSelect(label),10000);if(!s)return null;try{(s.querySelector('.ant-select-selector')||s).click()}catch(_){return null}await sleep(350);return wait(()=>[...document.querySelectorAll('.ant-select-dropdown:not(.ant-select-dropdown-hidden)')].filter(vis).pop(),5000)}
async function selectCode(label,code){const d=await openSelect(label);if(!d)return false;const rows=[...d.querySelectorAll('label,li,[role="option"],div')].filter(vis);let r=rows.find(e=>new RegExp('(^|\\s)'+code+'(\\s|$)').test(norm(text(e))));if(!r){document.body.click();return false}const cb=r.querySelector('input[type="checkbox"]');try{(cb||r.closest('label,li,[role="option"]')||r).click()}catch(_){return false}await sleep(450);document.body.click();await sleep(500);return true}
async function clearShop(){const d=await openSelect('Siêu thị');if(!d)return false;const q=[...d.querySelectorAll('button,a,span,div')].filter(vis);const b=q.find(e=>norm(text(e))==='bo chon tat ca')||q.find(e=>norm(text(e)).includes('bo chon tat ca'));if(b)try{(b.closest('button,a')||b).click()}catch(_){}await sleep(400);document.body.click();await sleep(350);return true}
async function clickXem(){const b=await wait(()=>[...document.querySelectorAll('button,a')].find(e=>vis(e)&&norm(text(e))==='xem'),7000);if(!b)return false;b.click();await sleep(3500);return true}
function totalTable(){return[...document.querySelectorAll('table')].find(t=>{const h=norm(text(t.querySelector('thead'))||text(t));return h.includes('sieu thi')&&h.includes('dt tra gop')&&h.includes('dt sieu thi')&&h.includes('homecredit')&&h.includes('fecredit')})||null}
function matchShop(row,c){const r=norm(row);for(let i=1;i<=5;i++){const s=c['shop'+i];if(!s)continue;const f=norm(s),tail=f.split(' ').slice(-2).join(' '),code=(String(s).match(/\b\d{3,6}\b/)||[])[0];if((f&&(r.includes(f)||f.includes(r)))||(tail&&r.includes(tail))||(code&&r.includes(code)))return'shop'+i}if(r.includes('hai anh'))return'shop1';if(r.includes('truc cuong'))return'shop2';return null}
function save(table,c){const z=obj(gv(DK,{}));z.link7=z.link7||{};z.link2=z.link2||{};let n=0,sumTra=0,sumST=0,log=[];for(const r of table.querySelectorAll('tbody tr')){const a=[...r.querySelectorAll('td')];if(a.length<7)continue;const k=matchShop(text(a[0]),c);if(!k)continue;const tra=num(text(a[1])),st=num(text(a[2])),pct=num(text(a[3])),hc=num(text(a[4])),fe=num(text(a[6]));if(!(st>0))continue;z.link7[k]=z.link7[k]||{};Object.assign(z.link7[k],{_SHOP_DT:st,_SHOP_HC:hc,_SHOP_FE:fe,_SHOP_TOTAL_:{tg:pct,dt_sieu_thi:st,dt_tragop:tra}});z.link2[k]=z.link2[k]||{};z.link2[k].tg=pct;sumTra+=tra;sumST+=st;n++;log.push(k+': '+st+'/'+tra+'/'+hc+'/'+fe)}if(n){z.link2.total=z.link2.total||{};z.link2.total.tg=sumST?sumTra/sumST*100:0;sv(DK,z);sv(LAST,Date.now());sv('autobi_tc21_debug',log.join(' | '))}return{n,log}}
async function onTraCham(){const c=await wait(()=>{const x=cfg();return shops(x).length?x:null},12000);if(!c)return;toast('⏳ Chọn Vùng Hà Nội...');if(!await selectCode('Vùng',REGION_CODE)){toast('⚠️ Chưa chọn được Vùng 761');return}await sleep(1000);toast('⏳ Chọn Khu vực Nam Định...');if(!await selectCode('Khu vực',AREA_CODE)){toast('⚠️ Chưa chọn được Khu vực 3866');return}await sleep(1000);await clearShop();await clickXem();const t=await wait(totalTable,22000);if(!t){toast('⚠️ Chưa ra bảng Siêu thị sau khi chọn Vùng/Khu vực');return}await sleep(700);const r=save(t,c);if(!r.n){toast('⚠️ Bảng đã ra nhưng chưa khớp dòng Hải Anh/Trực Cường');return}toast('✅ Đã lấy Trả chậm '+r.n+' siêu thị');const st=obj(gv(ST)),back=st.back||'/dashboard/home';sv(ST,{done:true,back});await sleep(900);location.href=back}
async function main(){const c=await wait(()=>{const x=cfg();return shops(x).length?x:null},12000);if(!c)return;const st=obj(gv(ST));if(st.done){dv(ST);toast('✅ Dữ liệu Trả chậm đã cập nhật');setTimeout(()=>location.reload(),700);return}const z=obj(gv(DK,{}));if(shops(c).every(s=>Number(z.link7?.[s.key]?._SHOP_DT)>0))return;if(Date.now()-Number(gv(LAST,0)||0)<20000)return;sv(ST,{back:location.pathname+location.search+location.hash});toast('🔄 Sang Báo cáo trả chậm: Vùng → Khu vực → Siêu thị...');await sleep(700);location.href='/dashboard/tra-cham'}
if(location.pathname.includes('/dashboard/tra-cham'))setTimeout(onTraCham,1600);else setTimeout(main,3200);
})();
