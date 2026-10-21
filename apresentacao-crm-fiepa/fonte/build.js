// Reimplantação CRM Dynamics — Sistema FIEPA
// Gera deck .pptx editável (formas, textos, tabela e gráfico nativos; ícones SVG recoloríveis)
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const pi = require("react-icons/pi");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const sharp = require("sharp");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "Reimplantacao_CRM_Dynamics_FIEPA.pptx";

const THEME = {
  name: "FIEPA CRM",
  headFontFace: "Segoe UI Semibold",
  bodyFontFace: "Segoe UI",
  colors: {
    dk1: "14233D", lt1: "FFFFFF", dk2: "0A2A5E", lt2: "EAF2FC",
    accent1: "0B5CAD", accent2: "1E8FE1", accent3: "22C3E6",
    accent4: "E5484D", accent5: "1FA67A", accent6: "5B6B82",
    hlink: "0B5CAD", folHlink: "5B6B82",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Sistema FIEPA";
pres.company = "FIEPA";
pres.title = "Reimplantação CRM Dynamics";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const S = pres.ShapeType;

const W = 13.333, M = 0.6, CW = W - 2 * M; // content width 12.133
const LIGHT_FONT = "Segoe UI Light";

// ---------- helpers ----------
function icon(name, hex) {
  const svg = RDS.renderToStaticMarkup(React.createElement(pi["Pi" + name + "Duotone"], { size: 256 }))
    .replace(/currentColor/g, "#" + hex)
    .replace(/ style="[^"]*"/, "");
  return "image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}
const shadow = (o = 0.12, b = 14, off = 4) => ({ type: "outer", color: HEX.dk2, opacity: o, blur: b, offset: off, angle: 90 });

function addIcon(s, name, hex, x, y, d, objectName) {
  s.addImage({ data: icon(name, hex), x, y, w: d, h: d, objectName: objectName || "Ícone " + name });
}
// icon inside a circle / rounded square badge
function badge(s, name, { x, y, d, fill, fillT = 0, iconHex, shape = "circle", scale = 0.58, sh, objectName }) {
  const o = { x, y, w: d, h: d, fill: { color: fill, transparency: fillT }, line: { type: "none" }, objectName: (objectName || name) + " (fundo)" };
  if (sh) o.shadow = sh;
  if (shape === "circle") s.addShape(S.ellipse, o);
  else s.addShape(S.roundRect, { ...o, rectRadius: d * 0.26 });
  const id = d * scale;
  addIcon(s, name, iconHex, x + (d - id) / 2, y + (d - id) / 2, id, (objectName || name) + " (ícone)");
}
function card(s, x, y, w, h, opts = {}) {
  s.addShape(S.roundRect, {
    x, y, w, h, rectRadius: opts.r ?? 0.14,
    fill: { color: opts.fill || C.background1, transparency: opts.fillT || 0 },
    line: opts.line || { color: "DCE7F5", width: 0.75 },
    shadow: opts.noShadow ? undefined : shadow(opts.so ?? 0.10, 16, 4),
    objectName: opts.name || "Cartão",
  });
}
function txt(s, text, o) {
  s.addText(text, { isTextBox: true, margin: 0, valign: "top", fontSize: 14, color: C.text1, ...o });
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "FIEPA Conteudo",
  background: { path: "assets/bg_light.png" },
  objects: [
    { image: { path: "assets/logo_navy.png", x: 11.0, y: 0.52, w: 1.733, h: 0.3 } },
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.46, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2.5, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.76, w: 10.2, h: 0.72, fontSize: 34, color: C.text2, margin: 0, valign: "top", align: "left", fontFace: THEME.headFontFace }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: M, y: 1.5, w: 10.2, h: 0.6, fontSize: 15, color: C.accent6, margin: 0, valign: "top" }, text: "" } },
  ],
  slideNumber: { x: 12.33, y: 7.0, w: 0.4, h: 0.25, fontSize: 10, color: "8A99AF", align: "right", fontFace: "Segoe UI" },
});
pres.defineSlideMaster({ title: "FIEPA Escuro", background: { path: "assets/bg_dark.png" }, objects: [] });
pres.defineSlideMaster({
  title: "FIEPA Limpo",
  background: { path: "assets/bg_light.png" },
  objects: [],
  slideNumber: { x: 12.33, y: 7.0, w: 0.4, h: 0.25, fontSize: 10, color: "8A99AF", align: "right", fontFace: "Segoe UI" },
});

function contentSlide(section, kicker, title, sub) {
  const s = pres.addSlide({ masterName: "FIEPA Conteudo", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "sub" });
  return s;
}

// orbit illustration (cover + closing): CRM in the center, business entities around it
function orbit(s, cx, cy) {
  const ring = (d, t, dash) => s.addShape(S.ellipse, {
    x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { type: "none" },
    line: { color: "FFFFFF", width: 1, transparency: t, dashType: dash || "solid" }, objectName: "Órbita",
  });
  s.addShape(S.ellipse, { x: cx - 2.35, y: cy - 2.35, w: 4.7, h: 4.7, fill: { color: C.accent3, transparency: 93 }, line: { type: "none" }, objectName: "Brilho" });
  ring(4.2, 72, "dash");
  ring(2.75, 60);
  s.addShape(S.ellipse, { x: cx - 0.8, y: cy - 0.8, w: 1.6, h: 1.6, fill: { color: C.background1 }, line: { color: C.accent3, width: 3 }, shadow: { type: "outer", color: "000000", opacity: 0.3, blur: 20, offset: 4, angle: 90 }, objectName: "Núcleo CRM" });
  addIcon(s, "UsersThree", HEX.accent1, cx - 0.5, cy - 0.5, 1.0, "Núcleo CRM (ícone)");
  const nodes = [["Handshake", -90, 2.1], ["ChartPieSlice", -18, 2.1], ["Target", 54, 2.1], ["Files", 126, 2.1], ["CurrencyCircleDollar", 198, 2.1],
                 ["Funnel", 20, 1.375], ["ShieldCheck", 200, 1.375]];
  nodes.forEach(([n, a, r], i) => {
    const d = r > 2 ? 0.72 : 0.5;
    const x = cx + r * Math.cos(a * Math.PI / 180) - d / 2, y = cy + r * Math.sin(a * Math.PI / 180) - d / 2;
    badge(s, n, { x, y, d, fill: r > 2 ? C.background1 : C.accent3, iconHex: r > 2 ? HEX.accent1 : HEX.dk2, scale: 0.6, objectName: "Órbita nó " + (i + 1), sh: { type: "outer", color: "000000", opacity: 0.25, blur: 10, offset: 3, angle: 90 } });
  });
}

// =====================================================================
// 1. CAPA
// =====================================================================
{
  pres.addSection({ title: "Capa" });
  const s = pres.addSlide({ masterName: "FIEPA Escuro", sectionTitle: "Capa" });
  s.addImage({ path: "assets/logo_white.png", x: 0.7, y: 0.7, w: 2.6, h: 0.45, objectName: "Logo FIEPA" });
  txt(s, "SISTEMA FIEPA  ·  PROPOSTA", { x: 0.7, y: 1.75, w: 6, h: 0.3, fontSize: 12, bold: true, color: C.accent3, charSpacing: 3 });
  txt(s, [
    { text: "Reimplantação", options: { breakLine: true } },
    { text: "CRM Dynamics", options: { color: C.accent3 } },
  ], { x: 0.7, y: 2.1, w: 7.2, h: 1.9, fontSize: 54, color: C.background1, fontFace: THEME.headFontFace, lineSpacingMultiple: 0.95 });
  txt(s, "Gestão comercial integrada para o Sistema FIEPA", { x: 0.7, y: 4.08, w: 7.2, h: 0.45, fontSize: 20, italic: true, color: "D7E6F7" });

  orbit(s, 10.35, 2.85);

  const pillars = [
    ["FlowArrow", "PROCESSO ÚNICO", "Etapas e regras comerciais comuns"],
    ["Intersect", "VISÃO INTEGRADA", "de clientes e oportunidades"],
    ["Gauge", "CONTROLE E GESTÃO", "Indicadores e rastreabilidade"],
  ];
  const pw = (CW + 0.1 - 2 * 0.25) / 3;
  pillars.forEach(([ic, t, d], i) => {
    const x = 0.7 + i * (pw + 0.25), y = 5.45;
    s.addShape(S.roundRect, { x, y, w: pw, h: 1.3, rectRadius: 0.14, fill: { color: C.background1, transparency: 88 }, line: { color: C.background1, width: 0.75, transparency: 70 }, objectName: "Pilar " + (i + 1) });
    badge(s, ic, { x: x + 0.28, y: y + 0.29, d: 0.72, fill: C.accent3, iconHex: HEX.dk2, objectName: "Pilar " + (i + 1) });
    txt(s, t, { x: x + 1.2, y: y + 0.3, w: pw - 1.4, h: 0.32, fontSize: 15, bold: true, color: C.background1, charSpacing: 1 });
    txt(s, d, { x: x + 1.2, y: y + 0.66, w: pw - 1.4, h: 0.4, fontSize: 14, color: "D7E6F7" });
  });
  s.addNotes("Abertura: a proposta de reimplantação do CRM Dynamics para o Sistema FIEPA, apoiada em três pilares — processo único, visão integrada e controle e gestão.");
}

// =====================================================================
// 2. AGENDA
// =====================================================================
{
  pres.addSection({ title: "Agenda" });
  const s = pres.addSlide({ masterName: "FIEPA Limpo", sectionTitle: "Agenda" });
  s.addImage({ path: "assets/panel_dark.png", x: 0, y: 0, w: 4.45, h: 7.5, objectName: "Painel lateral" });
  s.addImage({ path: "assets/logo_white.png", x: 0.7, y: 0.7, w: 2.3, h: 0.4, objectName: "Logo FIEPA" });
  txt(s, "ROTEIRO", { x: 0.7, y: 2.75, w: 3.2, h: 0.3, fontSize: 12, bold: true, color: C.accent3, charSpacing: 3 });
  txt(s, "Agenda", { x: 0.7, y: 3.1, w: 3.4, h: 0.95, fontSize: 48, color: C.background1, fontFace: THEME.headFontFace });
  txt(s, "Reimplantação CRM Dynamics", { x: 0.7, y: 4.1, w: 3.4, h: 0.4, fontSize: 15, color: "D7E6F7" });

  const items = [
    ["WarningCircle", "JUSTIFICATIVA", "O CRM atual limita a gestão comercial"],
    ["RocketLaunch", "VALOR PARA A GESTÃO", "O que muda na prática"],
    ["PlugsConnected", "INTEGRAÇÕES E CONTROLE", "Sistemas conectados, acessos controlados"],
    ["FlagBanner", "ENTREGAS", "Etapas do projeto de reimplantação"],
    ["Coins", "INVESTIMENTO", "Duas alternativas de contratação"],
    ["Prohibit", "ESCOPO", "O que não está incluído"],
  ];
  const x0 = 5.15, rowH = 1.0, y0 = 0.82;
  items.forEach(([ic, sec, d], i) => {
    const y = y0 + i * rowH;
    txt(s, String(i + 1).padStart(2, "0"), { x: x0, y: y + 0.12, w: 0.85, h: 0.6, fontSize: 32, color: "B9CFEA", fontFace: LIGHT_FONT });
    badge(s, ic, { x: x0 + 1.0, y: y + 0.13, d: 0.62, fill: C.background2, iconHex: HEX.accent1, objectName: "Agenda " + (i + 1) });
    txt(s, sec, { x: x0 + 1.9, y: y + 0.12, w: 5.5, h: 0.28, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2 });
    txt(s, d, { x: x0 + 1.9, y: y + 0.42, w: 5.6, h: 0.4, fontSize: 19, color: C.text2 });
    if (i < items.length - 1) s.addShape(S.line, { x: x0 + 1.0, y: y + rowH - 0.04, w: 6.55, h: 0, line: { color: "DCE7F5", width: 0.75 }, objectName: "Divisor" });
  });
  s.addNotes("Roteiro da apresentação em seis blocos: justificativa, valor para a gestão, integrações e controle, entregas, investimento e escopo.");
}

// =====================================================================
// 3. 01 – JUSTIFICATIVA
// =====================================================================
{
  pres.addSection({ title: "01 – Justificativa" });
  const s = contentSlide("01 – Justificativa", "01 – JUSTIFICATIVA", "O CRM atual limita a gestão comercial",
    "A reimplantação busca alinhar o processo à estrutura atual e recuperar visibilidade, integração e capacidade de evolução.");
  const pains = [
    ["ArrowsSplit", "PROCESSO DESALINHADO", "O fluxo atual não reflete plenamente a estrutura das entidades."],
    ["EyeSlash", "BAIXA VISÃO GERENCIAL", "Faltam relatórios consistentes para acompanhar funil, conversão e previsão."],
    ["ArrowsClockwise", "RETRABALHO E LEGADO", "Controles paralelos, customizações acumuladas e integrações descontinuadas dificultam a evolução."],
  ];
  const y0 = 2.45, rh = 1.45;
  pains.forEach(([ic, t, d], i) => {
    const y = y0 + i * rh;
    badge(s, ic, { x: M, y, d: 0.95, fill: C.accent4, fillT: 88, iconHex: HEX.accent4, objectName: "Problema " + (i + 1) });
    txt(s, t, { x: 1.85, y: y + 0.02, w: 5.75, h: 0.32, fontSize: 15, bold: true, color: C.text2, charSpacing: 1 });
    txt(s, d, { x: 1.85, y: y + 0.4, w: 5.75, h: 0.7, fontSize: 14, color: C.accent6 });
    if (i < pains.length - 1) s.addShape(S.line, { x: 1.85, y: y + rh - 0.2, w: 5.75, h: 0, line: { color: "D5E2F2", width: 0.75 }, objectName: "Divisor" });
  });
  // painel de consequência
  const px = 8.15, py = 2.35, pw = 4.583, ph = 4.45;
  s.addShape(S.roundRect, { x: px, y: py, w: pw, h: ph, rectRadius: 0.16, fill: { color: C.text2 }, line: { type: "none" }, shadow: shadow(0.25, 20, 6), objectName: "Painel consequências" });
  badge(s, "TrendDown", { x: px + 0.4, y: py + 0.42, d: 0.85, fill: C.accent4, iconHex: "FFFFFF", objectName: "Consequência" });
  txt(s, "CONSEQUÊNCIAS PARA A GESTÃO", { x: px + 0.4, y: py + 1.55, w: pw - 0.8, h: 0.3, fontSize: 12, bold: true, color: C.accent3, charSpacing: 2 });
  txt(s, "Menor capacidade de acompanhar oportunidades, padronizar decisões e planejar resultados com base em informações confiáveis.",
    { x: px + 0.4, y: py + 1.95, w: pw - 0.8, h: 2.2, fontSize: 19, color: C.background1 });
  s.addNotes("Três problemas do CRM atual: processo desalinhado, baixa visão gerencial e retrabalho/legado. A consequência é menor capacidade de acompanhar oportunidades, padronizar decisões e planejar resultados.");
}

// =====================================================================
// 4. 02 – VALOR PARA A GESTÃO
// =====================================================================
{
  pres.addSection({ title: "02 – Valor para a gestão" });
  const s = contentSlide("02 – Valor para a gestão", "02 – VALOR PARA A GESTÃO", "O que muda na prática",
    "O objetivo não é apenas atualizar o sistema: é melhorar a forma de gerir o relacionamento e as oportunidades.");
  const vals = [
    ["UserFocus", "CLIENTE 360º", "Contas, contratos, oportunidades e proposta em um só lugar."],
    ["ShieldCheck", "GOVERNANÇA", "Etapas, responsabilidades e alçadas claras."],
    ["Target", "DECISÃO", "Metas e previsão de vendas"],
    ["Lightning", "EFICIÊNCIA", "Menos controles paralelos e tarefas manuais."],
  ];
  const gap = 0.3, cw = (CW - 3 * gap) / 4, y = 2.35, ch = 3.0;
  vals.forEach(([ic, t, d], i) => {
    const x = M + i * (cw + gap);
    card(s, x, y, cw, ch, { name: "Benefício " + (i + 1) });
    badge(s, ic, { x: x + 0.3, y: y + 0.32, d: 0.9, fill: C.accent1, iconHex: "FFFFFF", shape: "square", objectName: "Benefício " + (i + 1) });
    txt(s, String(i + 1).padStart(2, "0"), { x: x + cw - 1.05, y: y + 0.3, w: 0.75, h: 0.5, fontSize: 26, color: "C9DCF2", align: "right", fontFace: LIGHT_FONT });
    txt(s, t, { x: x + 0.3, y: y + 1.5, w: cw - 0.6, h: 0.32, fontSize: 16, bold: true, color: C.text2, charSpacing: 1 });
    txt(s, d, { x: x + 0.3, y: y + 1.92, w: cw - 0.6, h: 0.9, fontSize: 14, color: C.accent6 });
  });
  // resultado esperado
  const by = 5.66;
  s.addShape(S.roundRect, { x: M, y: by, w: CW, h: 1.12, rectRadius: 0.14, fill: { color: C.text2 }, line: { type: "none" }, shadow: shadow(0.2, 16, 4), objectName: "Resultado esperado", h: 1.2 });
  badge(s, "SealCheck", { x: M + 0.35, y: by + 0.24, d: 0.72, fill: C.accent3, iconHex: HEX.dk2, objectName: "Resultado" });
  txt(s, "RESULTADO ESPERADO", { x: M + 1.35, y: by + 0.2, w: 6, h: 0.28, fontSize: 12, bold: true, color: C.accent3, charSpacing: 2 });
  txt(s, "Um processo comercial comum, com dados rastreáveis e melhor suporte à decisão.", { x: M + 1.35, y: by + 0.5, w: CW - 1.6, h: 0.62, fontSize: 18, color: C.background1 });
  s.addNotes("O ganho vai além da atualização técnica: cliente 360º, governança, decisão apoiada em metas e previsão, e eficiência com menos controles paralelos.");
}

// =====================================================================
// 5. 03 – INTEGRAÇÕES E CONTROLE
// =====================================================================
{
  pres.addSection({ title: "03 – Integrações e controle" });
  const s = contentSlide("03 – Integrações e controle", "03 – INTEGRAÇÕES E CONTROLE", "Sistemas conectados, acessos controlados",
    "A proposta prevê rotas nativas Microsoft e camada Azure para integrações específicas.");
  const cw = 3.75, ch = 1.0, yT = 2.2, yB = 3.62, xL = M, xR = W - M - cw;
  const hubW = 3.5, hubH = 0.92, hubX = (W - hubW) / 2, hubY = 2.86;
  const o2Y = 4.1, o2H = 0.92;
  const lineC = { color: "8FB4DE", width: 1.5 };
  // conectores (desenhados antes, ficam por trás)
  const conn = (x1, y1, x2, y2, extra = {}) => {
    const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), line: { ...lineC, ...extra }, objectName: "Conector" };
    if ((x2 - x1) * (y2 - y1) < 0) o.flipV = true;
    s.addShape(S.line, o);
  };
  conn(xL + cw, yT + ch / 2, hubX + 0.3, hubY + 0.25, { beginArrowType: "triangle", endArrowType: "triangle" });
  conn(xL + cw, yB + ch / 2, hubX + 0.3, hubY + hubH - 0.25, { beginArrowType: "triangle", endArrowType: "triangle" });
  conn(hubX + hubW - 0.3, hubY + 0.25, xR, yT + ch / 2, { endArrowType: "triangle" });
  conn(hubX + hubW - 0.3, hubY + hubH - 0.25, xR, yB + ch / 2, { beginArrowType: "triangle", endArrowType: "triangle" });
  conn(W / 2, hubY + hubH, W / 2, o2Y, { dashType: "dash", color: HEX.accent2 });

  s.addShape(S.roundRect, { x: hubX - 0.16, y: hubY - 0.16, w: hubW + 0.32, h: hubH + 0.32, rectRadius: (hubH + 0.32) / 2, fill: { color: C.accent2, transparency: 82 }, line: { type: "none" }, objectName: "Halo CRM" });
  s.addShape(S.roundRect, { x: hubX, y: hubY, w: hubW, h: hubH, rectRadius: hubH / 2, fill: { color: C.text2 }, line: { type: "none" }, shadow: shadow(0.3, 18, 5), objectName: "Dynamics 365 Sales" });
  badge(s, "UsersThree", { x: hubX + 0.16, y: hubY + 0.15, d: 0.62, fill: C.background1, iconHex: HEX.accent1, objectName: "CRM" });
  txt(s, "DYNAMICS 365 SALES (CRM)", { x: hubX + 0.92, y: hubY, w: hubW - 1.05, h: hubH, fontSize: 14, bold: true, color: C.background1, valign: "middle" });

  const sys = [
    [xL, yT, "Factory", "Dynamics 365 F&SCM", "Contas nos dois sentidos; produtos e preços → CRM"],
    [xL, yB, "ChatsCircle", "Outlook + Teams", "E-mails, reuniões, aprovações e avisos"],
    [xR, yT, "Buildings", "CRM Nacional (CNI)", "Envio a cada alteração, rastreio e reprocessamento"],
    [xR, yB, "Folders", "SharePoint", "Documentos por conta e oportunidade"],
  ];
  sys.forEach(([x, y, ic, t, d]) => {
    card(s, x, y, cw, ch, { name: t });
    badge(s, ic, { x: x + 0.2, y: y + 0.2, d: 0.6, fill: C.background2, iconHex: HEX.accent1, shape: "square", objectName: t });
    txt(s, t, { x: x + 0.98, y: y + 0.14, w: cw - 1.12, h: 0.3, fontSize: 14, bold: true, color: C.text2 });
    txt(s, d, { x: x + 0.98, y: y + 0.45, w: cw - 1.12, h: 0.5, fontSize: 11.5, color: C.accent6 });
  });
  // Onda 2 (futuro)
  const o2W = 3.95, o2X = (W - o2W) / 2;
  s.addShape(S.roundRect, { x: o2X, y: o2Y, w: o2W, h: o2H, rectRadius: 0.14, fill: { color: C.background1 }, line: { color: C.accent2, width: 1.25, dashType: "dash" }, objectName: "Onda 2 · CNPJ" });
  badge(s, "IdentificationCard", { x: o2X + 0.18, y: o2Y + 0.15, d: 0.56, fill: C.accent2, fillT: 85, iconHex: HEX.accent2, shape: "square", objectName: "CNPJ" });
  txt(s, "Onda 2 · Consulta por CNPJ", { x: o2X + 0.9, y: o2Y + 0.12, w: o2W - 1.05, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
  txt(s, "Complemento de cadastro de empresas", { x: o2X + 0.9, y: o2Y + 0.44, w: o2W - 1.0, h: 0.38, fontSize: 11, color: C.accent6 });

  // segurança e privacidade
  const py = 5.18, ph = 1.7;
  s.addShape(S.roundRect, { x: M, y: py, w: CW, h: ph, rectRadius: 0.14, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Painel segurança" });
  addIcon(s, "ShieldCheck", HEX.accent1, M + 0.3, py + 0.16, 0.34, "Escudo");
  txt(s, "SEGURANÇA E PRIVACIDADE", { x: M + 0.72, y: py + 0.2, w: 5, h: 0.28, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2 });
  const sec = [
    ["LockLaminated", "Segregação", "Cada entidade vê e altera seus registros conforme o perfil."],
    ["TreeStructure", "Perfis e hierarquia", "Acesso por função; gestores acompanham suas equipes."],
    ["Password", "Campos sensíveis", "Valores e descontos protegidos por perfil."],
    ["ClockCounterClockwise", "Auditoria", "Registro de alterações nas tabelas principais."],
  ];
  const iw = (CW - 0.6) / 4;
  sec.forEach(([ic, t, d], i) => {
    const x = M + 0.3 + i * iw, y = py + 0.64;
    badge(s, ic, { x, y, d: 0.5, fill: C.text2, iconHex: "FFFFFF", objectName: t });
    txt(s, t, { x: x + 0.64, y: y - 0.02, w: iw - 0.8, h: 0.28, fontSize: 13, bold: true, color: C.text2 });
    txt(s, d, { x: x + 0.64, y: y + 0.29, w: iw - 0.72, h: 0.65, fontSize: 11.5, color: C.accent6 });
  });
  s.addNotes("O Dynamics 365 Sales é o centro: integra-se ao F&SCM (contas nos dois sentidos), ao CRM Nacional da CNI, ao Outlook/Teams e ao SharePoint. A consulta por CNPJ entra na Onda 2. Segurança: segregação por entidade, perfis e hierarquia, campos sensíveis protegidos e auditoria.");
}

// =====================================================================
// 6. 04 – ENTREGAS
// =====================================================================
{
  pres.addSection({ title: "04 – Entregas" });
  const s = contentSlide("04 – Entregas", "04 – ENTREGAS", "Etapas do projeto de reimplantação", "Cada etapa tem um propósito, de forma macro");
  // linha do tempo proporcional (18 + 4 semanas)
  const gap = 0.12, unit = (CW - gap) / 22, w1 = unit * 18, w2 = unit * 4, by = 2.2, bh = 0.78;
  s.addShape(S.roundRect, { x: M, y: by, w: w1, h: bh, rectRadius: 0.14, fill: { color: C.text2 }, line: { type: "none" }, shadow: shadow(0.2, 14, 4), objectName: "Barra Onda 1" });
  badge(s, "RocketLaunch", { x: M + 0.18, y: by + 0.15, d: 0.48, fill: C.accent3, iconHex: HEX.dk2, objectName: "Onda 1" });
  txt(s, "Onda 1 – Núcleo Comercial", { x: M + 0.82, y: by, w: 4.5, h: bh, fontSize: 17, bold: true, color: C.background1, valign: "middle" });
  addIcon(s, "Clock", HEX.accent3, M + w1 - 4.55, by + 0.24, 0.3, "Relógio");
  txt(s, [{ text: "Previsão de execução  ", options: { color: "B9CFEA" } }, { text: "18 semanas ~ 4 meses", options: { bold: true } }],
    { x: M + w1 - 4.18, y: by, w: 4.0, h: bh, fontSize: 13, color: C.background1, valign: "middle", align: "right" });
  const x2 = M + w1 + gap;
  s.addShape(S.roundRect, { x: x2, y: by, w: w2, h: bh, rectRadius: 0.14, fill: { color: C.accent2 }, line: { type: "none" }, shadow: shadow(0.2, 14, 4), objectName: "Barra Onda 2" });
  txt(s, [{ text: "Onda 2", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "4 semanas ~ 1 mês", options: { fontSize: 11.5 } }],
    { x: x2 + 0.15, y: by, w: w2 - 0.25, h: bh, color: C.background1, valign: "middle" });
  // escala
  [[0, "Início"], [18, "Semana 18"], [22, "Semana 22"]].forEach(([wk, l]) => {
    const x = M + wk * unit + (wk === 22 ? gap : 0);
    s.addShape(S.line, { x, y: by + bh + 0.08, w: 0, h: 0.12, line: { color: "9DB4D2", width: 1 }, objectName: "Marca" });
    txt(s, l, { x: wk === 0 ? x : x - 1.2, y: by + bh + 0.22, w: 1.2, h: 0.22, fontSize: 10, color: "8A99AF", align: wk === 0 ? "left" : "right" });
  });

  const o1 = [
    ["ListChecks", "Definição de regras de solução, dados e integrações."],
    ["Buildings", "Ambientes, segurança e estrutura das 4 entidades, com linhas de solução, casas e centros de custo."],
    ["Funnel", "Fluxo guiado de vendas, oportunidade com vários produtos e proposta com modelo."],
    ["ChatsCircle", "Alçadas de aprovação no Teams e Outlook."],
    ["Percent", "Alçada de desconto."],
    ["ChartPieSlice", "Visão 360º, metas, previsão e painel de funil no Power BI."],
    ["PlugsConnected", "Integração com ERP Dynamics e CRM Nacional."],
    ["Database", "Migração de dados em aberto, testes, treinamento e operação assistida."],
  ];
  const gy = 3.55, rh = 0.83, colW = 4.15;
  o1.forEach(([ic, t], i) => {
    const col = Math.floor(i / 4), row = i % 4, x = M + col * (colW + 0.25), y = gy + row * rh;
    badge(s, ic, { x, y: y + 0.02, d: 0.5, fill: C.background2, iconHex: HEX.accent1, shape: "square", objectName: "Entrega " + (i + 1) });
    txt(s, t, { x: x + 0.66, y, w: colW - 0.7, h: 0.72, fontSize: 13, color: C.text1 });
  });
  // painel Onda 2
  const px = 9.35, pw = W - M - px, py = 3.42, ph = 3.43;
  s.addShape(S.roundRect, { x: px, y: py, w: pw, h: ph, rectRadius: 0.14, fill: { color: C.accent2, transparency: 90 }, line: { color: C.accent2, width: 1, transparency: 55 }, objectName: "Painel Onda 2" });
  txt(s, "Onda 2 – Painéis Avançados e CNPJ", { x: px + 0.3, y: py + 0.25, w: pw - 0.6, h: 0.62, fontSize: 15, bold: true, color: C.text2 });
  [["ChartBar", "Painéis Avançados de Produtos e Similaridade no Power BI."], ["IdentificationCard", "Integração para cadastro de Empresas por CNPJ."]].forEach(([ic, t], i) => {
    const y = py + 1.12 + i * 1.08;
    badge(s, ic, { x: px + 0.3, y, d: 0.5, fill: C.accent2, iconHex: "FFFFFF", shape: "square", objectName: "Onda 2 item " + (i + 1) });
    txt(s, t, { x: px + 0.96, y: y - 0.02, w: pw - 1.2, h: 0.9, fontSize: 13, color: C.text1 });
  });
  s.addNotes("Duas ondas. Onda 1 (18 semanas, ~4 meses): o núcleo comercial completo — regras, ambientes, fluxo de vendas, alçadas, Power BI, integrações e migração. Onda 2 (4 semanas, ~1 mês): painéis avançados e cadastro por CNPJ.");
}

// =====================================================================
// 7. 05 – INVESTIMENTO (opções)
// =====================================================================
{
  pres.addSection({ title: "05 – Investimento" });
  const s = contentSlide("05 – Investimento", "05 – INVESTIMENTO", "Duas alternativas de contratação",
    "Modelo Time & Materials: faturamento conforme horas medidas e aprovadas.");
  const cw = (CW - 0.33) / 2, ch = 3.5, y = 2.25;
  const opts = [
    { x: M, dark: false, ic: "Cube", t: "OPÇÃO A · NÚCLEO COMERCIAL", chip: "Onda 1", price: "R$ 378.770,00", detail: "", hours: "2.450 horas" },
    { x: M + cw + 0.33, dark: true, ic: "StackPlus", t: "OPÇÃO B · PROJETO COMPLETO", chip: "Ondas 1 e 2", price: "R$ 415.564,80", detail: "(R$ 378.770,00 + R$ 36.795,00)", hours: "2.688 horas" },
  ];
  opts.forEach((o) => {
    const fg = o.dark ? C.background1 : C.text2, mute = o.dark ? "B9CFEA" : C.accent6;
    if (o.dark) s.addShape(S.roundRect, { x: o.x, y, w: cw, h: ch, rectRadius: 0.16, fill: { color: C.text2 }, line: { type: "none" }, shadow: shadow(0.28, 22, 6), objectName: o.t });
    else card(s, o.x, y, cw, ch, { name: o.t });
    const p = 0.42;
    txt(s, o.t, { x: o.x + p, y: y + 0.42, w: cw - 1.7, h: 0.32, fontSize: 15, bold: true, color: o.dark ? C.accent3 : C.accent1, charSpacing: 1 });
    badge(s, o.ic, { x: o.x + cw - p - 0.85, y: y + 0.32, d: 0.85, fill: o.dark ? C.accent3 : C.background2, iconHex: o.dark ? HEX.dk2 : HEX.accent1, shape: "square", objectName: o.t });
    s.addShape(S.roundRect, { x: o.x + p, y: y + 0.9, w: o.chip.length * 0.105 + 0.45, h: 0.36, rectRadius: 0.18, fill: { color: o.dark ? C.accent2 : C.background2 }, line: { type: "none" }, objectName: "Escopo" });
    txt(s, o.chip, { x: o.x + p, y: y + 0.9, w: o.chip.length * 0.105 + 0.45, h: 0.36, fontSize: 12, bold: true, color: o.dark ? C.background1 : C.accent1, align: "center", valign: "middle" });
    txt(s, o.price, { x: o.x + p, y: y + 1.42, w: cw - 2 * p, h: 0.8, fontSize: 42, color: fg, fontFace: THEME.headFontFace });
    if (o.detail) txt(s, o.detail, { x: o.x + p, y: y + 2.22, w: cw - 2 * p, h: 0.3, fontSize: 13, color: mute });
    s.addShape(S.line, { x: o.x + p, y: y + 2.68, w: cw - 2 * p, h: 0, line: { color: o.dark ? "2D4E85" : "DCE7F5", width: 0.75 }, objectName: "Divisor" });
    addIcon(s, "Clock", o.dark ? HEX.accent3 : HEX.accent1, o.x + p, y + 2.88, 0.38, "Horas");
    txt(s, o.hours, { x: o.x + p + 0.52, y: y + 2.85, w: 3, h: 0.42, fontSize: 20, bold: true, color: fg, valign: "middle" });
  });
  // como funciona o T&M
  const fy = 6.08;
  txt(s, "COMO FUNCIONA", { x: M, y: fy + 0.2, w: 1.9, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 2 });
  const steps = [["Timer", "Horas medidas"], ["CheckCircle", "Horas aprovadas"], ["Receipt", "Faturamento"]];
  const sx0 = 2.65, sw = 3.0, sgap = 0.55;
  steps.forEach(([ic, t], i) => {
    const x = sx0 + i * (sw + sgap);
    s.addShape(S.roundRect, { x, y: fy, w: sw, h: 0.7, rectRadius: 0.35, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Etapa T&M " + (i + 1) });
    badge(s, ic, { x: x + 0.12, y: fy + 0.11, d: 0.48, fill: C.accent1, iconHex: "FFFFFF", objectName: t });
    txt(s, t, { x: x + 0.75, y: fy, w: sw - 0.85, h: 0.7, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    if (i < steps.length - 1) addIcon(s, "CaretRight", HEX.accent2, x + sw + 0.1, fy + 0.18, 0.34, "Seta");
  });
  s.addNotes("Opção A cobre a Onda 1 (núcleo comercial); a Opção B inclui as duas ondas. Modelo Time & Materials: o faturamento acompanha as horas medidas e aprovadas.");
}

// =====================================================================
// 8. 05 – INVESTIMENTO (rateio)
// =====================================================================
{
  const s = contentSlide("05 – Investimento", "05 – INVESTIMENTO", "Rateio por casa", "Visão geral para apoiar a decisão.");
  const colors = [HEX.dk2, HEX.accent1, HEX.accent3];
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Rateio por casa", labels: ["FIEPA", "SESI", "SENAI"], values: [33.33, 33.33, 33.33] }], {
    x: M, y: 2.3, w: 4.2, h: 4.2, holeSize: 64, chartColors: colors, showLegend: false, showValue: false, showPercent: false,
    showTitle: false, dataBorder: { pt: 2, color: "FFFFFF" }, objectName: "Gráfico rateio",
  });
  txt(s, "100%", { x: M + 1.1, y: 3.88, w: 2.0, h: 0.6, fontSize: 32, bold: true, color: C.text2, align: "center" });
  txt(s, "33,33% por casa", { x: M + 0.85, y: 4.5, w: 2.5, h: 0.3, fontSize: 12, color: C.accent6, align: "center" });

  const tx = 5.45, tw = W - M - tx, colW = [2.5, 1.5, 1.64, 1.64];
  const hdr = (t, fill) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: fill }, fontSize: 11.5, align: t === "Casa" ? "left" : "right" } });
  const rows = [
    ["FIEPA (IEL, Observatório, REDES e CIN)", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
    ["SESI", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
    ["SENAI", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
  ];
  const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 13, color: HEX.dk1, align: j ? "right" : "left", fill: { color: i % 2 ? "F5F9FE" : "FFFFFF" }, margin: j ? [0, 0.15, 0, 0.1] : [0, 0.1, 0, 0.45] } })));
  const total = ["TOTAL", "100%", "R$ 378.770,00", "R$ 415.564,80"].map((c, j) => ({ text: c, options: { fontSize: 13, bold: true, color: HEX.dk2, align: j ? "right" : "left", fill: { color: HEX.lt2 }, margin: j ? [0, 0.15, 0, 0.1] : [0, 0.1, 0, 0.45] } }));
  const head = [hdr("Casa", HEX.dk2), hdr("Porcentagem", HEX.dk2), hdr("OPÇÃO A", HEX.accent1), hdr("OPÇÃO B", HEX.accent2)];
  head.forEach((h, j) => (h.options.margin = j ? [0, 0.15, 0, 0.1] : [0, 0.1, 0, 0.45]));
  const ty = 2.78, rH = [0.52, 0.82, 0.6, 0.6, 0.64];
  s.addTable([head, ...body, total], { x: tx, y: ty, w: tw, colW, rowH: rH, valign: "middle", border: { type: "solid", pt: 0.75, color: "DCE7F5" }, fontFace: "Segoe UI", objectName: "Tabela rateio" });
  // marcadores de cor alinhados às linhas
  let yy = ty + rH[0];
  rows.forEach((r, i) => {
    s.addShape(S.ellipse, { x: tx + 0.17, y: yy + rH[i + 1] / 2 - 0.085, w: 0.17, h: 0.17, fill: { color: colors[i] }, line: { type: "none" }, objectName: "Legenda " + r[0].split(" ")[0] });
    yy += rH[i + 1];
  });
  s.addNotes("O investimento é dividido igualmente entre FIEPA, SESI e SENAI (33,33% cada) nas duas opções.");
}

// =====================================================================
// 9. 06 – ESCOPO
// =====================================================================
{
  pres.addSection({ title: "06 – Escopo" });
  const s = contentSlide("06 – Escopo", "06 – ESCOPO", "O que não está incluído", "Limites da proposta.");
  const items = [
    { kind: "x", ic: "CloudSlash", tag: "FORA DO ESCOPO", t: "Licenças Dynamics 365 e consumo Azure não estão incluídos." },
    { kind: "n", n: "40", u: "dias úteis", tag: "MOBILIZAÇÃO", t: "A equipe pode ser mobilizada em até 40 dias úteis após aceite e assinatura." },
    { kind: "n", n: "5", u: "dias úteis", tag: "APROVAÇÃO DE HORAS", t: "Horas medidas mensalmente; a FIEPA tem até 5 dias úteis para aprovar o relatório." },
    { kind: "x", ic: "Headset", tag: "FORA DO ESCOPO", t: "Sustentação após a Onda 2 e integração com WhatsApp não estão incluídas." },
  ];
  const gap = 0.33, cw = (CW - gap) / 2, ch = 2.08;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + gap), y = 2.35 + Math.floor(i / 2) * (ch + 0.3);
    card(s, x, y, cw, ch, { name: "Limite " + (i + 1) });
    const ex = it.kind === "x";
    s.addShape(S.roundRect, { x: x + 0.14, y: y + 0.14, w: 1.75, h: ch - 0.28, rectRadius: 0.1, fill: { color: ex ? C.accent4 : C.accent1, transparency: ex ? 91 : 90 }, line: { type: "none" }, objectName: "Destaque " + (i + 1) });
    if (ex) {
      badge(s, it.ic, { x: x + 0.14 + (1.75 - 0.95) / 2, y: y + 0.38, d: 0.95, fill: C.background1, iconHex: HEX.accent4, sh: shadow(0.12, 10, 2), objectName: "Limite " + (i + 1) });
      addIcon(s, "Prohibit", HEX.accent4, x + 1.3, y + 1.08, 0.36, "Não incluído");
      txt(s, "Não incluído", { x: x + 0.14, y: y + 1.5, w: 1.75, h: 0.3, fontSize: 11, bold: true, color: C.accent4, align: "center" });
    } else {
      txt(s, it.n, { x: x + 0.14, y: y + 0.3, w: 1.75, h: 1.0, fontSize: 54, bold: true, color: C.accent1, align: "center", valign: "middle" });
      txt(s, it.u, { x: x + 0.14, y: y + 1.32, w: 1.75, h: 0.3, fontSize: 13, color: C.accent1, align: "center" });
    }
    txt(s, it.tag, { x: x + 2.2, y: y + 0.38, w: cw - 2.5, h: 0.28, fontSize: 11.5, bold: true, color: ex ? C.accent4 : C.accent1, charSpacing: 2 });
    txt(s, it.t, { x: x + 2.2, y: y + 0.74, w: cw - 2.5, h: 1.1, fontSize: 17, color: C.text2 });
  });
  s.addNotes("Limites da proposta: licenças e Azure fora do escopo; mobilização em até 40 dias úteis após aceite e assinatura; horas medidas mensalmente, com até 5 dias úteis para aprovação; sustentação pós-Onda 2 e WhatsApp fora do escopo.");
}

// =====================================================================
// 10. OBRIGADO
// =====================================================================
{
  pres.addSection({ title: "Encerramento" });
  const s = pres.addSlide({ masterName: "FIEPA Escuro", sectionTitle: "Encerramento" });
  orbit(s, 10.35, 3.75);
  txt(s, "Obrigado!", { x: 0.7, y: 2.35, w: 7, h: 1.25, fontSize: 72, color: C.background1, fontFace: THEME.headFontFace });
  txt(s, "Reimplantação CRM Dynamics", { x: 0.7, y: 3.75, w: 7, h: 0.5, fontSize: 24, bold: true, color: C.accent3 });
  txt(s, "Gestão comercial integrada para o Sistema FIEPA", { x: 0.7, y: 4.3, w: 7, h: 0.45, fontSize: 18, italic: true, color: "D7E6F7" });
  s.addImage({ path: "assets/logo_white.png", x: 0.7, y: 6.25, w: 2.4, h: 0.415, objectName: "Logo FIEPA" });
  s.addNotes("Encerramento e abertura para dúvidas.");
}

// ---------- write + theme + real PNG fallbacks for SVG icons ----------
(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  const svgs = Object.keys(zip.files).filter((f) => /^ppt\/media\/.*\.svg$/.test(f));
  for (const f of svgs) {
    const png = f.replace(/\.svg$/, ".png");
    if (!zip.files[png]) continue;
    const buf = await zip.file(f).async("nodebuffer");
    zip.file(png, await sharp(buf, { density: 300 }).resize(256, 256).png().toBuffer());
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("ok", OUT, svgs.length, "svg icons");
})();
