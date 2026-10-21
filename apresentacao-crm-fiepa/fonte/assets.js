const sharp=require('sharp');
(async()=>{
// claro: degradê diagonal muito suave (branco -> azul gelo), sem padrões
const light=`<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
<stop offset="0" stop-color="#FFFFFF"/><stop offset="0.55" stop-color="#FAFCFF"/><stop offset="1" stop-color="#E4EFFB"/></linearGradient></defs>
<rect width="1920" height="1080" fill="url(#g)"/></svg>`;
await sharp(Buffer.from(light)).png().toFile('assets/bg_light2.png');
// escuro: degradê marinho profundo + padrão FIEPA bem discreto
const base=`<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
<stop offset="0" stop-color="#071A3A"/><stop offset="0.6" stop-color="#0A2A5E"/><stop offset="1" stop-color="#0D3E86"/></linearGradient></defs>
<rect width="1920" height="1080" fill="url(#g)"/></svg>`;
const pat=await sharp('assets/bg_dark_src.png').resize(1920,1080).ensureAlpha(0.10).png().toBuffer();
await sharp(Buffer.from(base)).composite([{input:pat,blend:'soft-light'}]).png().toFile('assets/bg_dark2.png');
})();
