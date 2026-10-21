const sharp=require('sharp');
(async()=>{
// chevron line motif, like the FIEPA cover pattern
function lines(x0,y0,n,step,color,op){let s='';for(let i=0;i<n;i++){const o=i*step;s+=`<polyline points="${x0+o},${y0} ${x0+260+o},${y0+300} ${x0+260+o},${y0+900}" fill="none" stroke="${color}" stroke-opacity="${op}" stroke-width="2"/>`;}return s;}
const light=`<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080">
<defs><radialGradient id="g1" cx="1850" cy="-80" r="1000" gradientUnits="userSpaceOnUse"><stop offset="0" stop-color="#D6E8FB" stop-opacity="1"/><stop offset="1" stop-color="#FFFFFF" stop-opacity="0"/></radialGradient>
<radialGradient id="g2" cx="-100" cy="1180" r="800" gradientUnits="userSpaceOnUse"><stop offset="0" stop-color="#E3F0FD" stop-opacity="1"/><stop offset="1" stop-color="#FFFFFF" stop-opacity="0"/></radialGradient>
<linearGradient id="fade" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fff" stop-opacity="0"/><stop offset="1" stop-color="#fff" stop-opacity="1"/></linearGradient>
<mask id="m"><rect x="1300" y="560" width="620" height="520" fill="url(#fade)"/></mask></defs>
<rect width="1920" height="1080" fill="#FFFFFF"/><rect width="1920" height="1080" fill="url(#g1)"/><rect width="1920" height="1080" fill="url(#g2)"/>
<g mask="url(#m)">${lines(1380,560,26,22,'#0B5CAD',0.10)}</g></svg>`;
await sharp(Buffer.from(light)).png().toFile('assets/bg_light.png');
// dark: original FIEPA pattern, slightly deepened at the bottom for contrast
const meta=await sharp('assets/bg_dark_src.png').resize(1920,1080).toBuffer();
const shade=`<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080"><defs><linearGradient id="d" x1="0" y1="0" x2="0.35" y2="1"><stop offset="0" stop-color="#061A3D" stop-opacity="0.15"/><stop offset="1" stop-color="#061A3D" stop-opacity="0.55"/></linearGradient></defs><rect width="1920" height="1080" fill="url(#d)"/></svg>`;
await sharp(meta).composite([{input:Buffer.from(shade)}]).png().toFile('assets/bg_dark.png');
await sharp('assets/bg_dark.png').extract({left:0,top:0,width:640,height:1080}).png().toFile('assets/panel_dark.png');
// navy logo for light slides
const {data,info}=await sharp('assets/logo_white.png').ensureAlpha().raw().toBuffer({resolveWithObject:true});
for(let i=0;i<data.length;i+=4){data[i]=0x0A;data[i+1]=0x2A;data[i+2]=0x5E;}
await sharp(data,{raw:info}).png().toFile('assets/logo_navy.png');
console.log('ok',info.width,info.height);
})();
