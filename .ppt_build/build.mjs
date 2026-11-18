import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { pathToFileURL } from 'node:url';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';
const root='D:/Codex-project/英语第6小组汇报';
const skill='C:/Users/JZQ/.codex/plugins/cache/openai-primary-runtime/presentations/26.1029.11837/skills/presentations';
const p=await PresentationFile.importPptx(await FileBlob.load(root+'/6.pptx'));
const records=JSON.parse('[]');
const inspect=(await p.inspect({kind:'slide,textbox',maxChars:500000})).ndjson.trim().split('\n').map(JSON.parse);
const slideIds={};for(const r of inspect)if(r.kind==='slide')slideIds[r.slideIndex]=r.id;
const chosen=[11,12,5,8,15];
p.slides.keep(chosen);
p.slides.reorder(chosen.map(n=>slideIds[n]));
const slides=p.slides.items;
function get(id){return p.resolve(id);}
function text(id,value,pos,size=22,align='left',color='#FFFFFF',bold=false){
 const s=get(id);s.text.autoFit='none';s.text=value;
 s.text.style={typeface:'Calibri Light',fontSize:size,color,alignment:align,verticalAlignment:'middle',insets:0,wrap:'square',autoFit:'none',bold};
 if(pos)s.position=pos;return s;
}
function add(s,name,value,pos,size=20,align='left',color='#FFFFFF',bold=false){
 const sh=s.shapes.add({name,geometry:'textbox',position:pos,fill:'none',line:{fill:'none',width:0}});sh.text=value;sh.text.style={typeface:'Calibri Light',fontSize:size,color,alignment:align,verticalAlignment:'middle',insets:0,wrap:'square',autoFit:'none',bold};return sh;
}
const box=(left,top,width,height)=>({left,top,width,height});
const WHO='https://www.who.int/news-room/fact-sheets/detail/depression';
const NIMH='https://www.nimh.nih.gov/health/publications/depression';
function notes(slide,script,sources){slide.speakerNotes.text=script+'\n\nSources (accessed 30 October 2026):\n'+sources.join('\n');}
function pageHeader(titleId,subtitleId,title,subtitle){text(titleId,title,box(110,17,700,48),36,'center','#F2DD1F');text(subtitleId,subtitle,box(140,63,640,30),20,'center','#F2DD1F');}
// 1: chapter opening, source slide 12.
text('sh/2xw7upwj','Treatment\n& Support',box(280,126,400,112),47,'center','#103B4F');
text('sh/kjq5oru1','PART 4  |  GROUP 6',box(320,245,320,32),21,'center','#103B4F');
text('sh/yh8nmhcv','Professional care. Everyday support.',box(280,289,400,33),20,'center','#103B4F');
get('sh/lkzmhwvm').position=box(468.59,280,22.82,0);
notes(slides[0],`Thank you. We have learned how depression affects people's lives. Now, let's turn to our final question: what help is available, and how can we support someone with depression? The first thing to remember is that depression is treatable. People do not have to face it alone.`,[WHO]);
// 2: two native template columns around original illustration.
pageHeader('sh/1g7apwvy','sh/ba9wfqx0','Professional Treatment','A plan based on individual needs');
text('sh/y9cfa50v','Talk therapy',box(32,263,250,38),29,'center','#F2DD1F');
text('sh/8zexgzix','Work with a trained\nprofessional to develop\nways to cope.',box(27,309,260,100),23,'center');
text('sh/90nypkzi','Medication',box(670,263,250,38),29,'center','#F2DD1F');
text('sh/ne5gnuhc','May help some people.\nDiscuss benefits and risks\nwith a clinician.',box(666,309,260,100),23,'center');
add(slides[1],'treatment-caveat','Some people receive both. Treatment takes time.',box(180,469,600,35),22,'center');
notes(slides[1],`Two main treatment options are talk therapy and medication. Talk therapy means working with a trained professional. It can help people develop new ways of thinking and coping with difficulties. Medication may also help. Some people receive both treatments. The right choice depends on the person's needs and should be discussed with a healthcare professional. Treatment takes time, so we should not expect someone to recover overnight. Seeking help is not a sign of weakness. It is a step toward getting better.`,[WHO,NIMH]);
// 3: original three-branch tree layout.
pageHeader('sh/lsrqp0be','sh/vmtcfudw','Everyday Support','Helpful habits — not a substitute for care');
text('sh/y94v2p0b','Keep a routine',box(28,202,243,37),26,'right','#F2DD1F');
text('sh/x8ve9kzq','Regular sleep\nand meals',box(28,246,243,66),23,'right');
text('sh/kb6d4zih','Stay connected',box(648,320,276,37),26,'left','#F2DD1F');
text('sh/zaxwbuhw','Talk with someone\nyou trust.',box(648,362,276,65),23,'left');
text('sh/c7md0zi5','Move gently',box(28,385,243,37),26,'right','#F2DD1F');
text('sh/o3a5cjud','Try a short walk\nor an activity you enjoy.',box(15,428,256,70),22,'right');
notes(slides[2],`Besides professional treatment, everyday habits can support recovery. These include keeping regular sleeping and eating habits, taking a short walk, and staying connected with friends and family. However, these habits are not a replacement for professional care. We should not suggest that exercise or positive thinking alone can cure depression.`,[WHO,NIMH]);
// 4: original jar illustration and two conversation blocks.
pageHeader('sh/apkra5gb','sh/8vm90fe9','Supporting a Friend','Listen without judgment. Offer practical help.');
text('sh/765cbmtg','“I’m here to listen.”',box(440,176,449,45),31,'left','#F2DD1F');
text('sh/s7edkru1','Instead of “Just cheer up,”\ntry “That sounds difficult.”',box(440,229,458,65),25);
text('sh/t8nudwbm','“Shall we seek help together?”',box(440,340,465,44),28,'left','#E85A5A');
text('sh/mpsb6xcr','Offer to go with them\nto see a counselor.',box(440,393,453,66),25);
add(slides[3],'urgent-help','Immediate danger? Seek local emergency help.',box(440,488,480,34),18,'left');
notes(slides[3],`So, what can we do as classmates? Imagine a friend tells you, “I've been having a really hard time.” Instead of saying, “Just cheer up,” we could say, “That sounds difficult. I'm here to listen.” Listen without judging, and do not rush to give advice. We can also offer practical help. For example, we could ask, “Would you like me to go with you to see a counselor?” Our role is to be a supportive friend, not to become their therapist. If someone is in immediate danger of harming themselves, we should seek urgent help from local emergency services.`,[NIMH,WHO]);
// 5: original closing slide.
text('sh/3qdg7qpo','Understanding,\nnot judgment',box(175,153,610,112),47,'center','#F4E13C');
text('sh/e1sf65o3','Seeking help is a sign of strength.',box(190,285,580,45),29,'center');
text('sh/1ojy10ne','Depression is treatable.\nThank you for listening.',box(225,350,510,70),25,'center');
notes(slides[4],`To conclude, our group has discussed what depression is, how to recognize it, why it matters, and what help is available. If you remember one message from our presentation, let it be this: depression deserves understanding, not judgment. Seeking help is a sign of strength, not weakness. Thank you for listening.`,[WHO,NIMH]);
// 6: references using original closing template, kept after spoken ending.
const ref=slides[4].duplicate();
const sourceText=ref.shapes.items.filter(s=>String(s.text||'').length>0);
for(const s of sourceText)s.text='';
add(ref,'reference-title','References',box(185,103,590,60),42,'center','#F4E13C');
add(ref,'who-reference','World Health Organization\nDepressive disorder (depression)',box(290,183,595,70),27);
const link1=add(ref,'who-link','who.int · Depression fact sheet',box(290,254,595,32),21,'left','#F2DD1F');link1.text.get('who.int · Depression fact sheet').link={uri:WHO,isExternal:true};
add(ref,'nimh-reference','National Institute of Mental Health\nDepression',box(290,315,595,70),25);
const link2=add(ref,'nimh-link','nimh.nih.gov · Depression',box(290,386,595,32),21,'left','#F2DD1F');link2.text.get('nimh.nih.gov · Depression').link={uri:NIMH,isExternal:true};
add(ref,'accessed','Accessed 30 October 2026',box(290,446,595,30),18);
notes(ref,'Reference slide. The source titles are clickable. The spoken presentation ends on the previous slide.',[WHO,NIMH]);
console.log('slideCount',p.slides.items.length,'masters',p.masters.items.length);
await fs.writeFile(root+'/.ppt_build/final-inspect.ndjson',(await p.inspect({kind:'slide,textbox,image,shape,notes',maxChars:500000})).ndjson);
for(let i=0;i<p.slides.items.length;i++){
 const b=await p.export({slide:p.slides.items[i],format:'png',scale:1.5});
 await fs.writeFile(root+`/.ppt_build/final-${i+1}.png`,new Uint8Array(await b.arrayBuffer()));console.log('Rendered',i+1);
}
await (await PresentationFile.exportPptx(p)).save(root+'/.ppt_build/candidate.pptx');
const {finalizePresentation}=await import(pathToFileURL(skill+'/container_tools/artifact_tool_utils.mjs').href);
console.log(await finalizePresentation({workspaceDir:root,candidatePath:root+'/.ppt_build/candidate.pptx',finalPath:root+'/output/Group6_Part4_Treatment_and_Support.pptx',pythonExecutable:'C:/Users/JZQ/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',integrityValidatorPath:skill+'/container_tools/inspect_presentation_package_integrity.py',layoutValidatorPath:skill+'/container_tools/inspect_presentation_layout_geometry.py',layoutArgs:['--expected-slide-size-emu','9144000,5143500','--validate-bullet-geometry','--validate-heading-fit'],explicitTotalSlideCount:6,fontPolicy:{basis:'reference',families:['Calibri Light'],referencePath:root+'/6.pptx',referenceSha256:createHash('sha256').update(await fs.readFile(root+'/6.pptx')).digest('hex')},verifyArtifactToolImport:true,receiptPath:root+'/.ppt_build/validation.json'}));


