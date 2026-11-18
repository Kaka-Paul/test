import fs from 'node:fs/promises';
const text=await fs.readFile('D:/Codex-project/英语第6小组汇报/.ppt_build/source-inspect.ndjson','utf8');
const rows=text.trim().split('\n').map(x=>JSON.parse(x));
for(const n of [5,8,11,12,15]){console.log('\nSOURCE SLIDE',n+1);for(const r of rows.filter(r=>r.slideIndex===n && ['textbox','slide'].includes(r.kind)))console.log(JSON.stringify(r));}
