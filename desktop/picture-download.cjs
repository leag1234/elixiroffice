// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Decides which session downloads may proceed to a native save dialog. The editor hands rendered
// picture bytes to the desktop layer as a blob: download from the app's own origin; everything
// else (http(s), other origins, non-image names, oversized items) stays canceled.
const path=require('node:path');
const MAX_BYTES=50*1024*1024;
const IMAGE_EXTENSIONS=new Set(['.png','.jpg','.jpeg','.svg']);
const FILTERS={
 '.png':[{name:'PNG image',extensions:['png']}],
 '.jpg':[{name:'JPEG image',extensions:['jpg','jpeg']}],
 '.jpeg':[{name:'JPEG image',extensions:['jpg','jpeg']}],
 '.svg':[{name:'SVG image',extensions:['svg']}]
};
function safeBasename(name){
 const base=path.basename(String(name||'')).replace(/[\x00-\x1f<>:"/\\|?*]/g,'_').trim();
 return base||'Image.png';
}
function decide({url,filename,size,origin,picturesDir}){
 if(typeof url!=='string'||!url.startsWith('blob:'+origin+'/'))return null;
 const name=String(filename||'');
 const extension=path.extname(name).toLowerCase();
 if(!IMAGE_EXTENSIONS.has(extension))return null;
 if(typeof size==='number'&&size>MAX_BYTES)return null;
 const base=safeBasename(name);
 return {
  title:'Save picture',
  defaultPath:path.join(picturesDir||'',base),
  filters:FILTERS[extension]
 };
}
module.exports={decide,safeBasename,MAX_BYTES,IMAGE_EXTENSIONS};
