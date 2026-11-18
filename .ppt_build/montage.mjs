import sharp from 'sharp';
let tiles=[];for(let i=1;i<=18;i++){tiles.push({input:await sharp(`D:/Codex-project/英语第6小组汇报/.ppt_build/source-${i}.png`).resize(384,216).toBuffer(),left:((i-1)%3)*384,top:Math.floor((i-1)/3)*216});}await sharp({create:{width:1152,height:1296,channels:3,background:'#ffffff'}}).composite(tiles).png().toFile('D:/Codex-project/英语第6小组汇报/.ppt_build/source-montage.png');
