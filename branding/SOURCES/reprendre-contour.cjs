// Reprend le contour blanc du premier logo approuvé, sans changer ses proportions.
const fs = require('fs');
const path = require('path');
const sharp = require('C:/Users/laxim/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/sharp');
(async () => {
  const {data, info} = await sharp(path.join(__dirname, 'concept-declinaisons.png'))
    .extract({left:60, top:40, width:640, height:640}).ensureAlpha().raw().toBuffer({resolveWithObject:true});
  const w=info.width, h=info.height;
  const filled=(x,y)=>{if(x<0||y<0||x>=w||y>=h)return false;const i=(y*w+x)*4;return data[i]>220&&data[i+1]>220&&data[i+2]>220&&data[i+3]>180;};
  const edges=new Map();
  const key=(x,y)=>y*(w+1)+x;
  const add=(x,y,xx,yy)=>{const k=key(x,y); if(!edges.has(k))edges.set(k,[]);edges.get(k).push(key(xx,yy));};
  for(let y=0;y<h;y++)for(let x=0;x<w;x++)if(filled(x,y)){
    if(!filled(x,y-1))add(x,y,x+1,y); if(!filled(x+1,y))add(x+1,y,x+1,y+1);
    if(!filled(x,y+1))add(x+1,y+1,x,y+1); if(!filled(x-1,y))add(x,y+1,x,y);
  }
  const point=k=>[k%(w+1),Math.floor(k/(w+1))];
  const loops=[];
  while(edges.size){const start=edges.keys().next().value;let p=start;const loop=[];
    do{loop.push(point(p));const list=edges.get(p);if(!list)throw Error('Contour ouvert');const next=list.pop();if(!list.length)edges.delete(p);p=next;}while(p!==start);
    loops.push(loop);
  }
  const area=p=>Math.abs(p.reduce((a,v,i)=>{const n=p[(i+1)%p.length];return a+v[0]*n[1]-n[0]*v[1];},0)/2);
  loops.sort((a,b)=>area(b)-area(a));const outline=loops[0];
  function simplify(points,epsilon){
    if(points.length<3)return points;
    const a=points[0],b=points[points.length-1],dx=b[0]-a[0],dy=b[1]-a[1],length=dx*dx+dy*dy;
    let max=0,index=0;
    for(let i=1;i<points.length-1;i++){const p=points[i],t=length?Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/length)):0;
      const d=Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);if(d>max){max=d;index=i;}}
    return max>epsilon?[...simplify(points.slice(0,index+1),epsilon).slice(0,-1),...simplify(points.slice(index),epsilon)]:[a,b];
  }
  let opposite=0,distance=0;
  for(let i=1;i<outline.length;i++){const d=Math.hypot(outline[i][0]-outline[0][0],outline[i][1]-outline[0][1]);if(d>distance){distance=d;opposite=i;}}
  const points=[...simplify(outline.slice(0,opposite+1),0.7).slice(0,-1),...simplify([...outline.slice(opposite),outline[0]],0.7).slice(0,-1)];
  // The original coloured tile occupies x=70..682, y=48..660 in the source.
  const coords=points.map(([x,y])=>[((x+60-70)*120/612+4).toFixed(3),((y+40-48)*120/612+4).toFixed(3)]);
  const d=coords.map((p,i)=>(i?'L':'M')+p.join(' ')).join(' ')+' Z';
  fs.writeFileSync(path.join(__dirname,'korecto-contour.txt'),d+'\n');
  const rendered=await sharp(Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="640" height="640" viewBox="0 0 640 640"><path fill="white" d="${points.map((p,i)=>(i?'L':'M')+p.join(' ')).join(' ')} Z"/></svg>`)).ensureAlpha().raw().toBuffer();
  let intersection=0,union=0;
  for(let y=0;y<h;y++)for(let x=0;x<w;x++){const a=filled(x,y),b=rendered[(y*w+x)*4+3]>127;if(a&&b)intersection++;if(a||b)union++;}
  const overlap=intersection/union;
  if(overlap<0.995)throw Error('Écart excessif avec le contour original : '+overlap);
  fs.writeFileSync(path.join(__dirname,'VERIFICATION_CONTOUR.txt'),`Référence : concept-declinaisons.png, premier logo couleur.\nContour blanc repris sans trait épais ni changement de pente.\nRecouvrement du masque original : ${(100*overlap).toFixed(3)} %.\nPoints : ${points.length}.\n`);
  console.log('Contour original repris : '+(100*overlap).toFixed(3)+' % de recouvrement.');
})().catch(error=>{console.error(error);process.exit(1);});
