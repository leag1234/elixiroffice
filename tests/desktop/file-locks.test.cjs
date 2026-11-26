// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path'),os=require('node:os');
const locks=require('../../desktop/file-locks.cjs'),catalogs=require('../../tools/desktop/i18n-extra.json'),english=require('../../desktop/lock-strings.cjs');
const me={user:'gael',host:'box.example.org',name:'Gael D'};
const own=(pid,extra={})=>locks.buildLibreLock({name:me.name,user:me.user,host:me.host,pid,...extra});
test('owner file names cover the stripped spellings and the written name follows the Office rule',()=>{
 assert.deepEqual(locks.officeOwnerNames('Document1.docx'),['~$Document1.docx','~$ocument1.docx','~$cument1.docx']);
 assert.deepEqual(locks.officeOwnerNames('a.docx'),['~$a.docx']);
 assert.equal(locks.officeOwnerName('Book1.xlsx'),'~$Book1.xlsx');
 assert.equal(locks.officeOwnerName('Document1.docx'),'~$cument1.docx');
 assert.equal(locks.officeOwnerName('Report.docx'),'~$Report.docx');
 assert.equal(locks.officeOwnerName('abcdefg.pptx'),'~$bcdefg.pptx');
 assert.equal(locks.libreLockName('My file.odt'),'.~lock.My file.odt#');
});
test('owner file round trip, foreign layouts and garbage',()=>{
 const buffer=locks.buildOwnerFile('Gaël Duval');
 assert.equal(buffer.length,162);assert.equal(buffer[0],10);assert.equal(locks.parseOwnerFile(buffer),'Gaël Duval');
 const narrow=Buffer.alloc(162,0x20);narrow[0]=4;narrow.write('Anna',1,'latin1');narrow.writeUInt16LE(0,54);
 assert.equal(locks.parseOwnerFile(narrow),'Anna');
 const wide=Buffer.alloc(162,0);wide.writeUInt16LE(3,54);wide.write('Zoë',56,'utf16le');
 assert.equal(locks.parseOwnerFile(wide),'Zoë');
 for(const bad of [null,undefined,'x',Buffer.alloc(0),Buffer.alloc(1),Buffer.alloc(162),Buffer.alloc(162,0xff),Buffer.from([200,65,66]),Buffer.concat([Buffer.alloc(54),Buffer.from([0xff,0xff])])])assert.equal(locks.parseOwnerFile(bad),null);
 const control=Buffer.alloc(162,0x20);control[0]=3;control.write('a\u0001b',1,'latin1');control.writeUInt16LE(0,54);assert.equal(locks.parseOwnerFile(control),'ab');
});
test('LibreOffice lock parsing and escaping',()=>{
 const real=locks.parseLibreLock(',gael,host,28.03.2013 12:44,file:///home/user/.config/libreoffice/4;');
 assert.equal(real.name,'gael');assert.equal(real.user,'gael');assert.equal(real.host,'host');assert.equal(real.ours,false);
 const named=locks.parseLibreLock('Ann Lee,ann,pc1,01.02.2026 09:05,file:///x;\n');assert.equal(named.name,'Ann Lee');
 const text=locks.buildLibreLock({name:'Doe, John; Jr',user:'jd',host:'h',pid:42,now:new Date(2026,0,2,3,4)});
 assert.equal(text,'Doe\\, John\\; Jr,jd,h,02.01.2026 03:04,elixiroffice:42;\n');
 const back=locks.parseLibreLock(text);assert.equal(back.name,'Doe, John; Jr');assert.equal(back.pid,42);assert.equal(back.ours,true);
 for(const bad of ['','garbage','a,b','\u0000\u0001',null,undefined])assert.equal(locks.parseLibreLock(bad),null);
});
test('stale lock decision',()=>{
 const dead=()=>false,live=()=>true,libre=(text,file='/d/.~lock.a.docx#')=>({file,lock:locks.parseLibreLock(text)});
 // our own dead lock pair: replaced silently
 let result=locks.decide({office:[{file:'/d/~$a.docx',owner:'Gael D'}],libre:libre(own(999999))},me,dead);
 assert.equal(result.locked,false);assert.deepEqual(result.stale.sort(),['/d/.~lock.a.docx#','/d/~$a.docx']);
 // live other instance of ours: locked, named
 result=locks.decide({office:[],libre:libre(own(999999))},me,live);assert.equal(result.locked,true);assert.equal(result.user,'Gael D');assert.deepEqual(result.stale,[]);
 // our pid but not held by any document: stale; held: locked
 assert.equal(locks.decide({office:[],libre:libre(own(process.pid))},{...me,held:()=>false},live).locked,false);
 assert.equal(locks.decide({office:[],libre:libre(own(process.pid))},{...me,held:()=>true},dead).locked,true);
 // same user and host but written by LibreOffice itself (cannot know it is dead): locked
 assert.equal(locks.decide({office:[],libre:libre('G,gael,box,01.01.2026 10:00,file:///p;')},me,dead).locked,true);
 // other user or other host: locked even with our marker
 assert.equal(locks.decide({office:[],libre:libre(locks.buildLibreLock({name:'x',user:'other',host:me.host,pid:1}))},me,dead).locked,true);
 assert.equal(locks.decide({office:[],libre:libre(locks.buildLibreLock({name:'x',user:me.user,host:'elsewhere',pid:1}))},me,dead).locked,true);
 // a Word owner file with another name is never taken over, even next to our stale lock
 result=locks.decide({office:[{file:'/d/~$a.docx',owner:'Colleague'}],libre:libre(own(999999))},me,dead);
 assert.equal(result.locked,true);assert.equal(result.user,'Colleague');assert.deepEqual(result.stale,['/d/.~lock.a.docx#']);
 // owner file alone, or unparsable lock, is locked with unknown user
 result=locks.decide({office:[{file:'/d/~$a.docx',owner:null}],libre:null},me,dead);assert.equal(result.locked,true);assert.equal(result.user,null);
 assert.equal(locks.decide({office:[],libre:{file:'/d/l',lock:null}},me,dead).locked,true);
 assert.equal(locks.decide({office:[],libre:null},me,dead).locked,false);
});
async function folder(t){const dir=await fs.mkdtemp(path.join(os.tmpdir(),'eo-locks-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));return dir;}
test('inspect, acquire and release on disk',async t=>{
 const dir=await folder(t),file=path.join(dir,'Report.docx');await fs.writeFile(file,'x');
 assert.equal((await locks.inspect(file,me)).locked,false);
 // Word spelling with the first character stripped, and a symlink that must be ignored
 await fs.writeFile(path.join(dir,'~$eport.docx'),locks.buildOwnerFile('Colleague'));
 let found=await locks.inspect(file,me);assert.equal(found.locked,true);assert.equal(found.user,'Colleague');
 await fs.rm(path.join(dir,'~$eport.docx'));await fs.symlink(file,path.join(dir,'~$Report.docx'));
 assert.equal((await locks.inspect(file,me)).locked,false);await fs.rm(path.join(dir,'~$Report.docx'));
 const handle=await locks.acquire(file,me,{platform:'darwin'});
 assert.equal(handle.held.length,2);
 assert.equal((await fs.readFile(path.join(dir,'.~lock.Report.docx#'),'utf8')).split(',')[1],'gael');
 assert.equal(locks.parseOwnerFile(await fs.readFile(path.join(dir,'~$Report.docx'))),'Gael D');
 // another process alive: locked; this process holding nothing: stale and replaced
 found=await locks.inspect(file,me,()=>true);assert.equal(found.locked,false);assert.equal(found.stale.length,2);
 await locks.release(handle);assert.deepEqual(await fs.readdir(dir),['Report.docx']);
});
test('never overwrite or delete someone else lock files',async t=>{
 const dir=await folder(t),file=path.join(dir,'a.xlsx');await fs.writeFile(file,'x');
 const theirs='Other,other,pc,01.01.2026 10:00,file:///p;';await fs.writeFile(path.join(dir,'.~lock.a.xlsx#'),theirs);
 const handle=await locks.acquire(file,me,{platform:'darwin'});
 assert.equal(await fs.readFile(path.join(dir,'.~lock.a.xlsx#'),'utf8'),theirs);assert.equal(handle.held.length,1);
 await fs.writeFile(path.join(dir,'~$a.xlsx'),'changed by someone');
 await locks.release(handle);assert.equal(await fs.readFile(path.join(dir,'~$a.xlsx'),'utf8'),'changed by someone');
 assert.equal(await fs.readFile(path.join(dir,'.~lock.a.xlsx#'),'utf8'),theirs);
});
test('stale own lock is replaced silently and a read-only folder does not throw',async t=>{
 const dir=await folder(t),file=path.join(dir,'a.docx');await fs.writeFile(file,'x');
 await fs.writeFile(path.join(dir,'.~lock.a.docx#'),own(999999));await fs.writeFile(path.join(dir,'~$a.docx'),locks.buildOwnerFile(me.name));
 const found=await locks.inspect(file,me,()=>false);assert.equal(found.locked,false);assert.equal(found.stale.length,2);
 const handle=await locks.acquire(file,me,{stale:found.stale,platform:'darwin'});assert.equal(handle.held.length,2);
 assert.equal(locks.parseLibreLock(await fs.readFile(path.join(dir,'.~lock.a.docx#'),'utf8')).pid,process.pid);
 await locks.release(handle);
 const missing=await locks.acquire(path.join(dir,'nope','a.docx'),me,{log:code=>assert(code)});assert.deepEqual(missing.held,[]);
});
test('writable and read-only save refusal',async t=>{
 const dir=await folder(t),file=path.join(dir,'a.docx'),alias=path.join(dir,'alias.docx'),other=path.join(dir,'copy.docx');
 await fs.writeFile(file,'x');await fs.link(file,alias);
 assert.equal(await locks.writable(file),true);assert.equal(await locks.writable(path.join(dir,'none.docx')),false);
 if(process.getuid?.()!==0){await fs.chmod(file,0o444);assert.equal(await locks.writable(file),false);}
 const current={target:file,sourceTarget:file,readOnly:true};
 assert.equal(await locks.isOriginal(current,file),true);assert.equal(await locks.isOriginal(current,alias),true);
 assert.equal(await locks.isOriginal(current,other),false);assert.equal(await locks.isOriginal({target:null,sourceTarget:null},file),false);
});
test('lock and read-only strings exist in all 45 catalogs with placeholders',()=>{
 assert.equal(Object.keys(catalogs).length,45);
 for(const [language,strings] of Object.entries(catalogs))for(const key of Object.keys(english)){
  assert.equal(typeof strings[key],'string',language+':'+key);assert(strings[key].trim());
 }
 for(const [language,strings] of Object.entries(catalogs)){assert(strings.lockedMessage.includes('{file}')&&strings.lockedMessage.includes('{user}'),language);}
 assert.equal(catalogs.en.lockedMessage,english.lockedMessage);
});
test("Linux writes only the LibreOffice lock, not the Office owner file",async()=>{
 const fs=require("node:fs/promises"),os=require("node:os"),path=require("node:path"),locks=require("../../desktop/file-locks.cjs");
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),"eo-lock-")),file=path.join(dir,"Report.docx");await fs.writeFile(file,"x");
 const handle=await locks.acquire(file,locks.identity("Test"),{platform:"linux"});
 require("node:assert").deepEqual(handle.held.map(item=>path.basename(item.file)),[".~lock.Report.docx#"]);
 await locks.release(handle);await fs.rm(dir,{recursive:true});
});
