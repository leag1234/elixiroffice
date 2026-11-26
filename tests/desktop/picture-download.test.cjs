// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),path=require('node:path');
const pictureDownload=require('../../desktop/picture-download.cjs');
const ORIGIN='eoapp://desktop.elixiroffice.invalid';
const DIR='/home/user/Pictures';

test('allows a blob download from the app origin with an image name',()=>{
 const options=pictureDownload.decide({url:'blob:'+ORIGIN+'/9b1deb4d-3b7d-4bad-9bdd-2b0d7b3dcb6d',filename:'chart.png',size:1024,origin:ORIGIN,picturesDir:DIR});
 assert.ok(options);
 assert.equal(options.defaultPath,path.join(DIR,'chart.png'));
 assert.deepEqual(options.filters,[{name:'PNG image',extensions:['png']}]);
});

test('accepts jpg, jpeg and svg case-insensitively',()=>{
 for(const name of ['a.JPG','b.Jpeg','c.SVG']){
  const options=pictureDownload.decide({url:'blob:'+ORIGIN+'/x',filename:name,size:1,origin:ORIGIN,picturesDir:DIR});
  assert.ok(options,name);
  assert.equal(options.defaultPath,path.join(DIR,name));
 }
});

test('blocks other origins and non-blob URLs',()=>{
 assert.equal(pictureDownload.decide({url:'blob:https://evil.example/x',filename:'a.png',size:1,origin:ORIGIN,picturesDir:DIR}),null);
 assert.equal(pictureDownload.decide({url:'https://example.com/a.png',filename:'a.png',size:1,origin:ORIGIN,picturesDir:DIR}),null);
 assert.equal(pictureDownload.decide({url:'blob:'+ORIGIN,filename:'a.png',size:1,origin:ORIGIN,picturesDir:DIR}),null);
});

test('blocks non-image extensions and missing names',()=>{
 for(const name of ['a.exe','b.docx','','noext']){
  assert.equal(pictureDownload.decide({url:'blob:'+ORIGIN+'/x',filename:name,size:1,origin:ORIGIN,picturesDir:DIR}),null);
 }
});

test('blocks items over 50 MB and allows unknown size',()=>{
 assert.equal(pictureDownload.decide({url:'blob:'+ORIGIN+'/x',filename:'a.png',size:pictureDownload.MAX_BYTES+1,origin:ORIGIN,picturesDir:DIR}),null);
 assert.ok(pictureDownload.decide({url:'blob:'+ORIGIN+'/x',filename:'a.png',size:pictureDownload.MAX_BYTES,origin:ORIGIN,picturesDir:DIR}));
 assert.ok(pictureDownload.decide({url:'blob:'+ORIGIN+'/x',filename:'a.png',origin:ORIGIN,picturesDir:DIR}));
});

test('safeBasename strips path separators and falls back to Image.png',()=>{
 assert.equal(pictureDownload.safeBasename('../../etc/passwd.png'),'passwd.png');
 assert.equal(pictureDownload.safeBasename(''),'Image.png');
});
