// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Detects and writes the "file is in use" markers other office suites use, so a
// colleague who opens a document we are editing (or the other way round) is told.
// Microsoft Office: "~$" owner file, 162 bytes. LibreOffice: ".~lock.<name>#" text file.
const fs=require('node:fs/promises'),sync=require('node:fs'),path=require('node:path'),os=require('node:os');
const OWNER_SIZE=162,OWNER_MAX=52,MARK='elixiroffice:',CELL=new Set(['xlsx','xlsm','xltx','xls','ods','csv']);
const clean=text=>String(text??'').replace(/[\u0000-\u001f\u007f]/g,'').trim();
const sameHost=(a,b)=>{const x=String(a||'').toLowerCase(),y=String(b||'').toLowerCase();return !!x&&(x===y||x.split('.')[0]===y.split('.')[0]);};
// Every spelling Office may use for the owner file of this document: Word and
// PowerPoint drop the first 1 or 2 characters of long names.
function officeOwnerNames(filename){
 const name=path.basename(filename);
 const base=name.length-path.extname(name).length;
 return [...new Set([0,1,2].filter(cut=>base>cut).map(cut=>'~$'+name.slice(cut)))];
}
// The one name we write. Excel keeps the full name. Word and PowerPoint keep the
// "~$" + base name within 8 characters when possible and never drop more than 2
// characters: 6 or fewer base characters drop none, 7 drop one, 8 or more drop two.
// This is the commonly observed rule, not verified against a real Office install.
function officeOwnerName(filename){
 const name=path.basename(filename),ext=path.extname(name),base=name.slice(0,name.length-ext.length);
 if(CELL.has(ext.slice(1).toLowerCase()))return '~$'+name;
 return '~$'+name.slice(Math.min(2,Math.max(0,base.length-6)));
}
const libreLockName=filename=>'.~lock.'+path.basename(filename)+'#';
function parseOwnerFile(buffer){
 try{
  if(!Buffer.isBuffer(buffer)||buffer.length<2)return null;
  if(buffer.length>=58){
   const length=buffer.readUInt16LE(54);
   if(length>0&&length<=OWNER_MAX&&56+length*2<=buffer.length){const wide=clean(buffer.toString('utf16le',56,56+length*2));if(wide)return wide;}
  }
  const length=buffer[0];
  if(length>0&&length<=OWNER_MAX&&1+length<=buffer.length){const narrow=clean(buffer.toString('latin1',1,1+length));if(narrow)return narrow;}
 }catch{}
 return null;
}
function buildOwnerFile(name){
 const text=clean(name).slice(0,OWNER_MAX)||'ElixirOffice',buffer=Buffer.alloc(OWNER_SIZE);
 const narrow=Buffer.from(text.replace(/[^\u0020-\u00ff]/g,'?'),'latin1');
 buffer[0]=narrow.length;buffer.fill(0x20,1,54);narrow.copy(buffer,1);
 const wide=Buffer.from(text,'utf16le');buffer.writeUInt16LE(text.length,54);
 for(let i=56;i<OWNER_SIZE;i+=2){buffer[i]=0x20;buffer[i+1]=0;}
 wide.copy(buffer,56,0,Math.min(wide.length,OWNER_SIZE-56));
 return buffer;
}
const escapeField=value=>String(value??'').replace(/[\\,;]/g,match=>'\\'+match);
function parseLibreLock(text){
 try{
  const fields=[];let current='',escaped=false;
  for(const character of String(text).slice(0,4096)){
   if(escaped){current+=character;escaped=false;}
   else if(character==='\\')escaped=true;
   else if(character===','||character===';'||character==='\n'||character==='\r'){fields.push(current);current='';if(character===';'||character==='\n')break;}
   else current+=character;
  }
  if(current)fields.push(current);
  if(fields.length<3)return null;
  const [display,user,host,time='',url='']=fields.map(clean);
  const pid=url.startsWith(MARK)?Number(url.slice(MARK.length)):NaN;
  return {name:display||user||null,user,host,time,url,ours:Number.isInteger(pid)&&pid>0,pid:Number.isInteger(pid)&&pid>0?pid:null};
 }catch{return null;}
}
function stamp(date){const p=value=>String(value).padStart(2,'0');return p(date.getDate())+'.'+p(date.getMonth()+1)+'.'+date.getFullYear()+' '+p(date.getHours())+':'+p(date.getMinutes());}
// Same field order as LibreOffice: display name, system user, host, time, profile URL; ";" ends the entry.
// The profile URL carries our marker and process id, so a stale lock of ours can be told from a live one.
function buildLibreLock({name,user,host,pid=process.pid,now=new Date()}){
 return [clean(name)||clean(user),clean(user),clean(host),stamp(now),MARK+pid].map(escapeField).join(',')+';\n';
}
const processAlive=pid=>{try{process.kill(pid,0);return true;}catch(error){return error?.code==='EPERM';}};
// locks: {office:[{file,owner}], libre:{file,lock|null}|null}. Returns the owner to show
// (null when unknown) and the stale lock files of this same user that may be replaced.
function decide(locks,me,alive=processAlive){
 const libre=locks.libre?.lock||null;
 const staleLibre=!!libre&&libre.ours&&libre.user===me.user&&sameHost(libre.host,me.host)&&!(libre.pid!==process.pid&&alive(libre.pid)||libre.pid===process.pid&&me.held?.(locks.libre.file));
 const stale=[],active=[];
 if(locks.libre)(staleLibre?stale:active).push({file:locks.libre.file,name:libre?.name||null});
 for(const item of locks.office)(staleLibre&&(!item.owner||item.owner.toLowerCase()===String(me.name).toLowerCase())?stale:active).push({file:item.file,name:item.owner});
 const named=active.find(item=>item.name);
 return {locked:active.length>0,user:named?named.name:null,stale:stale.map(item=>item.file)};
}
async function readSmall(file){
 let stat;
 try{stat=await fs.lstat(file);}catch{return null;}
 if(!stat.isFile()||stat.isSymbolicLink())return null;
 let handle;
 try{handle=await fs.open(file,'r');const buffer=Buffer.alloc(Math.min(stat.size,4096)),{bytesRead}=await handle.read(buffer,0,buffer.length,0);return buffer.subarray(0,bytesRead);}
 catch{return {unreadable:true};}
 finally{await handle?.close().catch(()=>{});}
}
async function inspect(filename,me,alive){
 const dir=path.dirname(filename),locks={office:[],libre:null};
 for(const name of officeOwnerNames(filename)){
  const file=path.join(dir,name),data=await readSmall(file);
  if(data)locks.office.push({file,owner:Buffer.isBuffer(data)?parseOwnerFile(data):null});
 }
 const file=path.join(dir,libreLockName(filename)),data=await readSmall(file);
 if(data)locks.libre={file,lock:Buffer.isBuffer(data)?parseLibreLock(data.toString('utf8')):null};
 return decide(locks,me,alive);
}
async function writable(filename){
 try{await fs.access(filename,sync.constants.W_OK);return true;}catch{return false;}
}
async function hide(file,platform=process.platform){
 if(platform!=='win32')return;
 try{await new Promise(resolve=>require('node:child_process').execFile(path.join(process.env.SystemRoot||'C:\\Windows','System32','attrib.exe'),['+H',file],{windowsHide:true,timeout:5000},()=>resolve()));}catch{}
}
// Writes our lock files next to the document. Files that already exist are never
// touched (except stale ones of ours, removed first). Failures are reported, never thrown.
// Like LibreOffice, the Office "~$" owner file is written only where Microsoft Office runs (Windows, macOS);
// elsewhere it would only clutter shared folders. Detection always checks it.
async function acquire(filename,me,{stale=[],log=()=>{},now=new Date(),platform=process.platform}={}){
 const held=[];
 for(const file of stale)try{await fs.unlink(file);}catch(error){if(error?.code!=='ENOENT')log('lock-stale-remove');}
 const dir=path.dirname(filename);
 const files=[[path.join(dir,libreLockName(filename)),Buffer.from(buildLibreLock({name:me.name,user:me.user,host:me.host,now}),'utf8')],...(['win32','darwin'].includes(platform)?[[path.join(dir,officeOwnerName(filename)),buildOwnerFile(me.name)]]:[])];
 for(const [file,content] of files){
  try{await fs.writeFile(file,content,{flag:'wx',mode:0o644});held.push({file,content});await hide(file,platform);}
  catch(error){log(error?.code==='EEXIST'?'lock-exists':'lock-write');}
 }
 return {filename,held};
}
// Removes only the files we wrote and only while they still hold our bytes.
async function release(handle,log=()=>{}){
 // Office owner file first, LibreOffice lock last: our LibreOffice lock is what
 // marks a leftover ~$ file as ours, so it must outlive it (no foreign-lock window).
 for(const {file,content} of [...(handle?.held||[])].reverse())try{
  const data=await fs.readFile(file);if(data.equals(content))await fs.unlink(file);
 }catch(error){if(error?.code!=='ENOENT')log('lock-release');}
 if(handle)handle.held=[];
}
function releaseSync(handle){
 for(const {file,content} of [...(handle?.held||[])].reverse())try{if(sync.readFileSync(file).equals(content))sync.unlinkSync(file);}catch{}
 if(handle)handle.held=[];
}
function identity(name){
 let user='';try{user=os.userInfo().username;}catch{user=process.env.USER||process.env.USERNAME||'';}
 return {user:user||'user',host:os.hostname(),name:clean(name)||user||'user'};
}
// Save to the original file is refused for a read-only document. Compares paths and file identity.
async function isOriginal(current,target){
 const original=current?.sourceTarget||current?.target;
 if(!original||!target)return false;
 if(path.resolve(original)===path.resolve(target))return true;
 try{const [a,b]=await Promise.all([fs.stat(original),fs.stat(target)]);return a.ino!==0&&a.ino===b.ino&&a.dev===b.dev;}catch{return false;}
}
module.exports={officeOwnerNames,officeOwnerName,libreLockName,parseOwnerFile,buildOwnerFile,parseLibreLock,buildLibreLock,decide,inspect,writable,acquire,release,releaseSync,identity,isOriginal,processAlive,MARK};
