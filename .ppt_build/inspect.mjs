import fs from 'node:fs/promises';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';
const p = await PresentationFile.importPptx(await FileBlob.load('D:/Codex-project/英语第6小组汇报/6.pptx'));
await fs.writeFile('D:/Codex-project/英语第6小组汇报/.ppt_build/source-inspect.ndjson',(await p.inspect({kind:'slide,textbox,image,shape,layout',maxChars:500000})).ndjson);
console.log('SLIDES',p.slides.items.length,Object.getOwnPropertyNames(Object.getPrototypeOf(p.slides)));
for(let i=0;i<p.slides.items.length;i++){
 const s=p.slides.items[i];
 try {const b=await p.export({slide:s,format:'png',scale:0.75}); await fs.writeFile(`D:/Codex-project/英语第6小组汇报/.ppt_build/source-${i+1}.png`,new Uint8Array(await b.arrayBuffer())); console.log('rendered',i+1);}catch(e){console.log('render error',i+1,e.message);}
}
