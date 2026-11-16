// Mauccounting ERP System — GitHub-hosted per-entity .dat storage.
// .dat payloads are Base64 encoded JSON, NOT cryptographic encryption.
// Compatible with legacy getSheet/addRecord/updateRecord/deleteRecord/saveExcel API.
const CONFIG={owner:'mau-acc-system',repo:'mauccounting.erp.database',branch:'main'};
const ROOT=`https://api.github.com/repos/${CONFIG.owner}/${CONFIG.repo}`;
const LICENSE_FILE=new URL('./license.tk',import.meta.url).href;
let githubToken=null, workbook=null, baseHead=null, dirty=new Set(), manifest=null;
function decodeLicense(value){const hex=value.trim();if(!hex||hex.length%2||!/^[a-f0-9]+$/i.test(hex))throw Error('Invalid license file');return new TextDecoder().decode(Uint8Array.from(hex.match(/../g).map(x=>parseInt(x,16))));}
export async function loadLicense(){const r=await fetch(LICENSE_FILE,{cache:'no-store'});if(!r.ok)throw Error('License file not found');githubToken=decodeLicense(await r.text());return true;}
export function setGitHubToken(token){githubToken=token?.trim()||null;}
export function clearGitHubToken(){githubToken=null;workbook=null;baseHead=null;manifest=null;dirty.clear();}
export function hasLicense(){return !!githubToken;}
async function request(url,opts={}){if(!githubToken)await loadLicense();return fetch(url,{...opts,headers:{Authorization:`Bearer ${githubToken}`,Accept:'application/vnd.github+json','X-GitHub-Api-Version':'2022-11-28',...opts.headers}});}
async function jsonGet(url){const r=await request(url);if(!r.ok)throw Error(`GitHub API ${r.status}: ${url.split('/').slice(-2).join('/')}`);return r.json();}
function b64ToText(b64){const s=atob(b64.replace(/\s/g,''));return new TextDecoder().decode(Uint8Array.from(s,c=>c.charCodeAt(0)));}
function textToB64(s){const bytes=new TextEncoder().encode(s);let out='';for(let i=0;i<bytes.length;i+=8192)out+=String.fromCharCode(...bytes.subarray(i,i+8192));return btoa(out);}
async function getContents(path){return jsonGet(`${ROOT}/contents/${path.split('/').map(encodeURIComponent).join('/')}?ref=${encodeURIComponent(CONFIG.branch)}`);}
export async function validateLicense(){try{if(!githubToken)await loadLicense();await getContents('data/manifest.json');return {valid:true,message:'License verified.'};}catch(e){return {valid:false,message:e.message};}}
export async function readExcel(){if(!githubToken)await loadLicense();const m=await getContents('data/manifest.json');manifest=JSON.parse(b64ToText(m.content));if(manifest.format!=='mau-dat-v1')throw Error('Unsupported database manifest');const ref=await jsonGet(`${ROOT}/git/ref/heads/${CONFIG.branch}`);baseHead=ref.object.sha;
 const pairs=await Promise.all(manifest.entities.map(async e=>{const file=await getContents(e.path);const data=JSON.parse(b64ToText(b64ToText(file.content)));if(data.sheet!==e.sheet)throw Error('Invalid sheet: '+e.sheet);return [e.sheet,data];}));
 workbook={SheetNames:pairs.map(x=>x[0]),Sheets:Object.fromEntries(pairs)};dirty.clear();return getSheetNames();}
export function getSheetNames(){if(!workbook)throw Error('Database not loaded');return [...workbook.SheetNames];}
export function getSheet(name){if(!workbook)throw Error('Database not loaded');const s=workbook.Sheets[name];if(!s)throw Error('Sheet not found: '+name);return s.records.map(r=>({...r}));}
export function getSheetHeaders(name){if(!workbook)throw Error('Database not loaded');const s=workbook.Sheets[name];if(!s)throw Error('Sheet not found: '+name);return [...new Set([...(s.headers||[]),...s.records.flatMap(r=>Object.keys(r||{}))])];}
function getData(name){if(!workbook)throw Error('Database not loaded');const s=workbook.Sheets[name];if(!s)throw Error('Sheet not found: '+name);return s;}
export function addRecord(name,record){const s=getData(name);for(const k of Object.keys(record))if(!s.headers.includes(k))s.headers.push(k);s.records.push({...record});dirty.add(name);return true;}
export function updateRecord(name,column,value,updates){const s=getData(name);const index=s.records.findIndex(r=>String(r[column])===String(value));if(index<0)throw Error('Record not found');for(const k of Object.keys(updates))if(!s.headers.includes(k))s.headers.push(k);s.records[index]={...s.records[index],...updates};dirty.add(name);return true;}
export function deleteRecord(name,column,value){const s=getData(name);const before=s.records.length;s.records=s.records.filter(r=>String(r[column])!==String(value));if(before===s.records.length)throw Error('Record not found');dirty.add(name);return true;}
export async function saveExcel(message='Update Mauccounting ERP database'){
 if(!workbook)throw Error('Database not loaded');if(!dirty.size)return true;
 const head=await jsonGet(`${ROOT}/git/ref/heads/${CONFIG.branch}`);if(head.object.sha!==baseHead)throw Error('Database changed in GitHub. Reload before saving to prevent overwriting changes.');
 const commit=await jsonGet(`${ROOT}/git/commits/${baseHead}`);const treeEntries=[];
 for(const name of dirty){const entity=manifest.entities.find(e=>e.sheet===name);if(!entity)throw Error('Unknown sheet '+name);const payload=textToB64(JSON.stringify(workbook.Sheets[name]));treeEntries.push({path:entity.path,mode:'100644',type:'blob',content:payload});}
 const treeResponse=await request(`${ROOT}/git/trees`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({base_tree:commit.tree.sha,tree:treeEntries})});if(!treeResponse.ok)throw Error('Failed to create database tree');const tree=await treeResponse.json();
 const commitResponse=await request(`${ROOT}/git/commits`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message,tree:tree.sha,parents:[baseHead]})});if(!commitResponse.ok)throw Error('Failed to create database commit');const newCommit=await commitResponse.json();
 const update=await request(`${ROOT}/git/refs/heads/${CONFIG.branch}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({sha:newCommit.sha,force:false})});if(!update.ok)throw Error('Database conflict: GitHub branch advanced. Reload.');baseHead=newCommit.sha;dirty.clear();return true;}
export function isDatabaseLoaded(){return !!workbook;}
