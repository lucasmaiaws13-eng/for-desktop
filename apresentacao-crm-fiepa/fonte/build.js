// Reimplantação CRM Dynamics — Sistema FIEPA (v3)
// Organização da v1 com o acabamento da v2: tipografia fina, mais respiro,
// ícones de traço fino em círculos discretos. Deck .pptx 100% editável.
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const pi = require("react-icons/pi");
const fs = require("fs");
const JSZip = require("jszip");
const sharp = require("sharp");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "Reimplantacao_CRM_Dynamics_FIEPA.pptx";

const THEME = {
  name: "FIEPA CRM",
  headFontFace: "Segoe UI Light",
  bodyFontFace: "Segoe UI",
  colors: {
    dk1: "1B2A41", lt1: "FFFFFF", dk2: "0A2A5E", lt2: "F1F6FC",
    accent1: "1660C4", accent2: "6B7A90", accent3: "8EC5F2",
    accent4: "DCE6F2", accent5: "3FC1E8", accent6: "C2453E",
    hlink: "1660C4", folHlink: "6B7A90",
  },
};
const HEX = THEME.colors;
const C = { ink: "tx1", navy: "tx2", white: "bg1", pale: "bg2", blue: "accent1", gray: "accent2", sky: "accent3", hair: "accent4", cyan: "accent5", coral: "accent6" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Sistema FIEPA";
pres.company = "FIEPA";
pres.title = "Reimplantação CRM Dynamics";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const S = pres.ShapeType;

const W = 13.333, M = 0.75, CW = W - 2 * M; // 11.833
const LIGHT = "Segoe UI Light", SEMI = "Segoe UI Semibold", BODY = "Segoe UI";
const MUTED_ON_DARK = "C3D3EA";

// ---------- helpers ----------
function icon(name, hex) {
  const svg = RDS.renderToStaticMarkup(React.createElement(pi["Pi" + name + "Light"], { size: 256 }))
    .replace(/currentColor/g, "#" + hex).replace(/ style="[^"]*"/, "");
  return "image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}
function addIcon(s, name, hex, x, y, d) {
  s.addImage({ data: icon(name, hex), x, y, w: d, h: d, objectName: "Ícone " + name });
}
// ícone fino dentro de um círculo discreto (fundo claro + contorno fino)
function ring(s, name, x, y, d, o = {}) {
  s.addShape(S.ellipse, {
    x, y, w: d, h: d,
    fill: o.fill === "none" ? { type: "none" } : { color: o.fill || C.pale, transparency: o.fillT || 0 },
    line: o.line === "none" ? { type: "none" } : { color: o.line || HEX.accent4, width: 0.75, transparency: o.lineT || 0 },
    objectName: (o.name || name) + " (círculo)",
  });
  const id = d * (o.scale || 0.5);
  addIcon(s, name, o.icon || HEX.accent1, x + (d - id) / 2, y + (d - id) / 2, id);
}
const soft = () => ({ type: "outer", color: HEX.dk2, opacity: 0.07, blur: 18, offset: 3, angle: 90 });
function card(s, x, y, w, h, o = {}) {
  s.addShape(S.roundRect, {
    x, y, w, h, rectRadius: o.r ?? 0.08,
    fill: { color: o.fill || C.white, transparency: o.fillT || 0 },
    line: o.line === "none" ? { type: "none" } : { color: o.line || HEX.accent4, width: 0.75, transparency: o.lineT || 0 },
    shadow: o.shadow === false ? undefined : soft(),
    objectName: o.name || "Cartão",
  });
}
function txt(s, text, o) {
  s.addText(text, { isTextBox: true, margin: 0, valign: "top", fontFace: BODY, fontSize: 14, color: C.ink, ...o });
}
const kicker = (s, t, x, y, w, color = C.blue, size = 10.5) =>
  txt(s, t, { x, y, w, h: 0.25, fontSize: size, fontFace: SEMI, color, charSpacing: 1.5 });
const hline = (s, x, y, w, color = HEX.accent4) => s.addShape(S.line, { x, y, w, h: 0, line: { color, width: 0.75 }, objectName: "Linha" });

// ---------- layouts ----------
const FOOT = { text: { text: "Reimplantação CRM Dynamics  ·  Sistema FIEPA", options: { x: M, y: 7.0, w: 6, h: 0.22, fontSize: 9, color: "8C99AD", fontFace: BODY, margin: 0 } } };
const NUM = { x: 12.18, y: 7.0, w: 0.4, h: 0.22, fontSize: 9, color: "8C99AD", align: "right", fontFace: BODY };
pres.defineSlideMaster({
  title: "FIEPA Conteudo",
  background: { path: "assets/bg_light2.png" },
  objects: [
    { image: { path: "assets/logo_navy.png", x: 11.33, y: 0.62, w: 1.25, h: 0.216 } },
    FOOT,
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.62, w: 8, h: 0.25, fontSize: 10.5, fontFace: SEMI, color: C.blue, charSpacing: 1.5, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.92, w: 10.2, h: 0.7, fontSize: 34, fontFace: LIGHT, color: C.navy, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: M, y: 1.66, w: 10, h: 0.5, fontSize: 14, fontFace: BODY, color: C.gray, margin: 0, valign: "top" }, text: "" } },
  ],
  slideNumber: NUM,
});
pres.defineSlideMaster({ title: "FIEPA Limpo", background: { path: "assets/bg_light2.png" }, objects: [FOOT], slideNumber: NUM });
pres.defineSlideMaster({ title: "FIEPA Escuro", background: { path: "assets/bg_dark2.png" }, objects: [] });

function contentSlide(section, k, title, sub) {
  const s = pres.addSlide({ masterName: "FIEPA Conteudo", sectionTitle: section });
  s.addText(k, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "sub" });
  return s;
}

// órbita (capa e encerramento): CRM no centro, entidades de negócio ao redor
function orbit(s, cx, cy) {
  const circ = (d, line, t, dash, fill) => s.addShape(S.ellipse, {
    x: cx - d / 2, y: cy - d / 2, w: d, h: d,
    fill: fill ? { color: fill.c, transparency: fill.t } : { type: "none" },
    line: line ? { color: line, width: 0.75, transparency: t, dashType: dash || "solid" } : { type: "none" }, objectName: "Órbita",
  });
  circ(4.9, null, 0, null, { c: C.white, t: 96 });
  circ(4.2, "FFFFFF", 70, "dash");
  circ(2.75, "FFFFFF", 78);
  s.addShape(S.ellipse, { x: cx - 0.78, y: cy - 0.78, w: 1.56, h: 1.56, fill: { color: C.white }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 24, offset: 4, angle: 90 }, objectName: "Núcleo CRM" });
  addIcon(s, "UsersThree", HEX.accent1, cx - 0.42, cy - 0.42, 0.84);
  const nodes = [["Handshake", -90, 2.1], ["ChartPieSlice", -18, 2.1], ["Target", 54, 2.1], ["Files", 126, 2.1], ["CurrencyCircleDollar", 198, 2.1],
                 ["Funnel", 20, 1.375], ["ShieldCheck", 200, 1.375]];
  nodes.forEach(([n, a, r]) => {
    const big = r > 2, d = big ? 0.66 : 0.46;
    const x = cx + r * Math.cos(a * Math.PI / 180) - d / 2, y = cy + r * Math.sin(a * Math.PI / 180) - d / 2;
    ring(s, n, x, y, d, { fill: big ? HEX.dk2 : HEX.dk2, line: big ? "FFFFFF" : HEX.accent5, lineT: big ? 55 : 30, icon: big ? "FFFFFF" : HEX.accent5, scale: 0.5, name: "Órbita " + n });
  });
}

// =====================================================================
// 1. CAPA
// =====================================================================
{
  pres.addSection({ title: "Capa" });
  const s = pres.addSlide({ masterName: "FIEPA Escuro", sectionTitle: "Capa" });
  s.addImage({ path: "assets/logo_white.png", x: M, y: 0.7, w: 2.2, h: 0.38, objectName: "Logo FIEPA" });
  kicker(s, "PROPOSTA  ·  SISTEMA FIEPA", M, 1.95, 6, C.cyan);
  txt(s, [
    { text: "Reimplantação", options: { fontFace: LIGHT, breakLine: true } },
    { text: "CRM Dynamics", options: { fontFace: SEMI } },
  ], { x: M, y: 2.25, w: 7.4, h: 1.85, fontSize: 54, color: C.white, lineSpacingMultiple: 0.95 });
  txt(s, "Gestão comercial integrada para o Sistema FIEPA", { x: M, y: 4.18, w: 7.4, h: 0.4, fontSize: 18, fontFace: LIGHT, color: MUTED_ON_DARK });

  orbit(s, 10.3, 2.9);

  const pillars = [
    ["FlowArrow", "PROCESSO ÚNICO", "Etapas e regras comerciais comuns"],
    ["Intersect", "VISÃO INTEGRADA", "de clientes e oportunidades"],
    ["Gauge", "CONTROLE E GESTÃO", "Indicadores e rastreabilidade"],
  ];
  const gap = 0.3, pw = (CW - 2 * gap) / 3, y = 5.5;
  pillars.forEach(([ic, t, d], i) => {
    const x = M + i * (pw + gap);
    card(s, x, y, pw, 1.22, { fill: C.white, fillT: 93, line: "FFFFFF", lineT: 78, shadow: false, name: "Pilar " + (i + 1) });
    ring(s, ic, x + 0.3, y + 0.31, 0.6, { fill: "none", line: HEX.accent5, lineT: 20, icon: HEX.accent5, name: "Pilar " + (i + 1) });
    txt(s, t, { x: x + 1.1, y: y + 0.3, w: pw - 1.25, h: 0.25, fontSize: 11, fontFace: SEMI, color: C.white, charSpacing: 1.5 });
    txt(s, d, { x: x + 1.1, y: y + 0.6, w: pw - 1.25, h: 0.35, fontSize: 14, fontFace: LIGHT, color: MUTED_ON_DARK });
  });
  s.addNotes("Abertura: a proposta de reimplantação do CRM Dynamics para o Sistema FIEPA, apoiada em três pilares — processo único, visão integrada e controle e gestão.");
}

// =====================================================================
// 2. AGENDA
// =====================================================================
{
  pres.addSection({ title: "Agenda" });
  const s = pres.addSlide({ masterName: "FIEPA Limpo", sectionTitle: "Agenda" });
  s.addImage({ path: "assets/panel_dark2.png", x: 0, y: 0, w: 4.45, h: 7.5, objectName: "Painel lateral" });
  s.addImage({ path: "assets/logo_white.png", x: M, y: 0.7, w: 2.0, h: 0.346, objectName: "Logo FIEPA" });
  kicker(s, "ROTEIRO", M, 2.75, 3, C.cyan);
  txt(s, "Agenda", { x: M, y: 3.05, w: 3.4, h: 0.95, fontSize: 48, fontFace: LIGHT, color: C.white });
  txt(s, "Reimplantação CRM Dynamics", { x: M, y: 4.05, w: 3.4, h: 0.35, fontSize: 14, fontFace: LIGHT, color: MUTED_ON_DARK });

  const items = [
    ["WarningCircle", "JUSTIFICATIVA", "O CRM atual limita a gestão comercial"],
    ["RocketLaunch", "VALOR PARA A GESTÃO", "O que muda na prática"],
    ["PlugsConnected", "INTEGRAÇÕES E CONTROLE", "Sistemas conectados, acessos controlados"],
    ["FlagBanner", "ENTREGAS", "Etapas do projeto de reimplantação"],
    ["Coins", "INVESTIMENTO", "Duas alternativas de contratação"],
    ["Prohibit", "ESCOPO", "O que não está incluído"],
  ];
  const x0 = 5.2, xw = W - M - x0, y0 = 0.82, rh = 0.98;
  items.forEach(([ic, sec, d], i) => {
    const y = y0 + i * rh;
    txt(s, String(i + 1).padStart(2, "0"), { x: x0, y: y + 0.2, w: 0.75, h: 0.5, fontSize: 26, fontFace: LIGHT, color: "A9C1E2" });
    ring(s, ic, x0 + 0.95, y + 0.17, 0.62, { name: "Agenda " + (i + 1) });
    kicker(s, sec, x0 + 1.85, y + 0.17, xw - 1.85, C.blue, 9.5);
    txt(s, d, { x: x0 + 1.85, y: y + 0.43, w: xw - 1.85, h: 0.4, fontSize: 19, fontFace: LIGHT, color: C.navy });
    if (i < items.length - 1) hline(s, x0 + 0.95, y + rh - 0.02, xw - 0.95);
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
  const y0 = 2.55, rh = 1.42, tx = M + 1.15, tw = 6.1;
  pains.forEach(([ic, t, d], i) => {
    const y = y0 + i * rh;
    ring(s, ic, M, y, 0.8, { name: "Problema " + (i + 1) });
    kicker(s, t, tx, y + 0.04, tw, C.navy, 11);
    txt(s, d, { x: tx, y: y + 0.38, w: tw, h: 0.7, fontSize: 14, color: C.gray, lineSpacingMultiple: 1.1 });
    if (i < pains.length - 1) hline(s, tx, y + rh - 0.2, tw);
  });
  const px = 8.3, pw = W - M - px, py = 2.5, ph = 4.2;
  s.addShape(S.roundRect, { x: px, y: py, w: pw, h: ph, rectRadius: 0.08, fill: { color: C.navy }, line: { type: "none" }, shadow: { type: "outer", color: HEX.dk2, opacity: 0.18, blur: 24, offset: 6, angle: 90 }, objectName: "Painel consequências" });
  ring(s, "TrendDown", px + 0.45, py + 0.42, 0.8, { fill: "none", line: HEX.accent5, lineT: 25, icon: HEX.accent5, name: "Consequência" });
  kicker(s, "CONSEQUÊNCIAS PARA A GESTÃO", px + 0.45, py + 1.48, pw - 0.9, C.cyan);
  txt(s, "Menor capacidade de acompanhar oportunidades, padronizar decisões e planejar resultados com base em informações confiáveis.",
    { x: px + 0.45, y: py + 1.85, w: pw - 0.9, h: 2.1, fontSize: 17, fontFace: LIGHT, color: C.white, lineSpacingMultiple: 1.1 });
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
  const gap = 0.3, cw = (CW - 3 * gap) / 4, y = 2.42, ch = 3.08, p = 0.34;
  vals.forEach(([ic, t, d], i) => {
    const x = M + i * (cw + gap);
    card(s, x, y, cw, ch, { name: "Benefício " + (i + 1) });
    ring(s, ic, x + p, y + p, 0.8, { name: "Benefício " + (i + 1) });
    txt(s, String(i + 1).padStart(2, "0"), { x: x + cw - p - 0.8, y: y + p + 0.05, w: 0.8, h: 0.45, fontSize: 22, fontFace: LIGHT, color: "BCD0EA", align: "right" });
    kicker(s, t, x + p, y + 1.48, cw - 2 * p, C.navy, 11);
    txt(s, d, { x: x + p, y: y + 1.84, w: cw - 2 * p, h: 0.9, fontSize: 14, color: C.gray, lineSpacingMultiple: 1.1 });
  });
  const by = 5.7, bh = 1.1;
  s.addShape(S.roundRect, { x: M, y: by, w: CW, h: bh, rectRadius: 0.08, fill: { color: C.navy }, line: { type: "none" }, shadow: { type: "outer", color: HEX.dk2, opacity: 0.16, blur: 20, offset: 5, angle: 90 }, objectName: "Resultado esperado" });
  ring(s, "SealCheck", M + 0.4, by + (bh - 0.66) / 2, 0.66, { fill: "none", line: HEX.accent5, lineT: 25, icon: HEX.accent5, name: "Resultado" });
  kicker(s, "RESULTADO ESPERADO", M + 1.35, by + 0.24, 5, C.cyan);
  txt(s, "Um processo comercial comum, com dados rastreáveis e melhor suporte à decisão.", { x: M + 1.35, y: by + 0.52, w: CW - 1.6, h: 0.42, fontSize: 17, fontFace: LIGHT, color: C.white });
  s.addNotes("O ganho vai além da atualização técnica: cliente 360º, governança, decisão apoiada em metas e previsão, e eficiência com menos controles paralelos.");
}

// =====================================================================
// 5. 03 – INTEGRAÇÕES E CONTROLE
// =====================================================================
{
  pres.addSection({ title: "03 – Integrações e controle" });
  const s = contentSlide("03 – Integrações e controle", "03 – INTEGRAÇÕES E CONTROLE", "Sistemas conectados, acessos controlados",
    "A proposta prevê rotas nativas Microsoft e camada Azure para integrações específicas.");
  const cw = 3.75, ch = 1.02, yT = 2.38, yB = 3.72, xL = M, xR = W - M - cw;
  const hubW = 3.3, hubH = 0.86, hubX = (W - hubW) / 2, hubY = 2.9;
  const o2W = 3.95, o2X = (W - o2W) / 2, o2Y = 4.12, o2H = 0.86;
  const lineC = { color: "A9C1E2", width: 0.75 };
  const conn = (x1, y1, x2, y2, extra = {}) => {
    const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), line: { ...lineC, ...extra }, objectName: "Conector" };
    if ((x2 - x1) * (y2 - y1) < 0) o.flipV = true;
    s.addShape(S.line, o);
  };
  const arr = { beginArrowType: "triangle", endArrowType: "triangle" };
  conn(xL + cw, yT + ch / 2, hubX + 0.35, hubY + 0.2, arr);
  conn(xL + cw, yB + ch / 2, hubX + 0.35, hubY + hubH - 0.2, arr);
  conn(hubX + hubW - 0.35, hubY + 0.2, xR, yT + ch / 2, { endArrowType: "triangle" });
  conn(hubX + hubW - 0.35, hubY + hubH - 0.2, xR, yB + ch / 2, arr);
  conn(W / 2, hubY + hubH, W / 2, o2Y, { dashType: "dash", color: HEX.accent1 });

  s.addShape(S.roundRect, { x: hubX - 0.14, y: hubY - 0.14, w: hubW + 0.28, h: hubH + 0.28, rectRadius: (hubH + 0.28) / 2, fill: { color: C.pale }, line: { color: HEX.accent4, width: 0.75 }, objectName: "Halo CRM" });
  s.addShape(S.roundRect, { x: hubX, y: hubY, w: hubW, h: hubH, rectRadius: hubH / 2, fill: { color: C.navy }, line: { type: "none" }, shadow: { type: "outer", color: HEX.dk2, opacity: 0.2, blur: 18, offset: 4, angle: 90 }, objectName: "Dynamics 365 Sales" });
  ring(s, "UsersThree", hubX + 0.15, hubY + 0.15, 0.56, { fill: "none", line: "FFFFFF", lineT: 50, icon: "FFFFFF", name: "CRM" });
  txt(s, "DYNAMICS 365 SALES (CRM)", { x: hubX + 0.85, y: hubY, w: hubW - 0.95, h: hubH, fontSize: 12, fontFace: SEMI, color: C.white, valign: "middle", charSpacing: 1 });

  const sys = [
    [xL, yT, "Factory", "Dynamics 365 F&SCM", "Contas nos dois sentidos; produtos e preços → CRM"],
    [xL, yB, "ChatsCircle", "Outlook + Teams", "E-mails, reuniões, aprovações e avisos"],
    [xR, yT, "Buildings", "CRM Nacional (CNI)", "Envio a cada alteração, rastreio e reprocessamento"],
    [xR, yB, "Folders", "SharePoint", "Documentos por conta e oportunidade"],
  ];
  sys.forEach(([x, y, ic, t, d]) => {
    card(s, x, y, cw, ch, { name: t });
    ring(s, ic, x + 0.22, y + 0.22, 0.58, { name: t });
    txt(s, t, { x: x + 0.98, y: y + 0.17, w: cw - 1.12, h: 0.28, fontSize: 14, fontFace: SEMI, color: C.navy });
    txt(s, d, { x: x + 0.98, y: y + 0.47, w: cw - 1.12, h: 0.48, fontSize: 11.5, color: C.gray });
  });
  s.addShape(S.roundRect, { x: o2X, y: o2Y, w: o2W, h: o2H, rectRadius: 0.08, fill: { color: C.white }, line: { color: HEX.accent1, width: 0.75, dashType: "dash" }, objectName: "Onda 2 · CNPJ" });
  ring(s, "IdentificationCard", o2X + 0.2, o2Y + 0.15, 0.56, { name: "CNPJ" });
  txt(s, "Onda 2 · Consulta por CNPJ", { x: o2X + 0.92, y: o2Y + 0.14, w: o2W - 1.02, h: 0.28, fontSize: 13, fontFace: SEMI, color: C.blue });
  txt(s, "Complemento de cadastro de empresas", { x: o2X + 0.92, y: o2Y + 0.45, w: o2W - 1.02, h: 0.3, fontSize: 11, color: C.gray });

  const py = 5.28, ph = 1.55;
  s.addShape(S.roundRect, { x: M, y: py, w: CW, h: ph, rectRadius: 0.08, fill: { color: C.pale }, line: { type: "none" }, objectName: "Painel segurança" });
  kicker(s, "SEGURANÇA E PRIVACIDADE", M + 0.35, py + 0.24, 5);
  const sec = [
    ["LockLaminated", "Segregação", "Cada entidade vê e altera seus registros conforme o perfil."],
    ["TreeStructure", "Perfis e hierarquia", "Acesso por função; gestores acompanham suas equipes."],
    ["Password", "Campos sensíveis", "Valores e descontos protegidos por perfil."],
    ["ClockCounterClockwise", "Auditoria", "Registro de alterações nas tabelas principais."],
  ];
  const iw = (CW - 0.7) / 4;
  sec.forEach(([ic, t, d], i) => {
    const x = M + 0.35 + i * iw, y = py + 0.62;
    ring(s, ic, x, y, 0.5, { fill: C.white, name: t });
    txt(s, t, { x: x + 0.65, y: y - 0.01, w: iw - 0.8, h: 0.26, fontSize: 13, fontFace: SEMI, color: C.navy });
    txt(s, d, { x: x + 0.65, y: y + 0.28, w: iw - 0.8, h: 0.55, fontSize: 11, color: C.gray });
  });
  s.addNotes("O Dynamics 365 Sales é o centro: integra-se ao F&SCM (contas nos dois sentidos), ao CRM Nacional da CNI, ao Outlook/Teams e ao SharePoint. A consulta por CNPJ entra na Onda 2. Segurança: segregação por entidade, perfis e hierarquia, campos sensíveis protegidos e auditoria.");
}

// =====================================================================
// 6. 04 – ENTREGAS
// =====================================================================
{
  pres.addSection({ title: "04 – Entregas" });
  const s = contentSlide("04 – Entregas", "04 – ENTREGAS", "Etapas do projeto de reimplantação", "Cada etapa tem um propósito, de forma macro");
  // barras proporcionais: 18 + 4 semanas
  const gap = 0.1, unit = (CW - gap) / 22, w1 = unit * 18, w2 = unit * 4, by = 2.42, bh = 0.7, x2 = M + w1 + gap;
  s.addShape(S.roundRect, { x: M, y: by, w: w1, h: bh, rectRadius: 0.08, fill: { color: C.navy }, line: { type: "none" }, shadow: { type: "outer", color: HEX.dk2, opacity: 0.16, blur: 16, offset: 4, angle: 90 }, objectName: "Barra Onda 1" });
  ring(s, "RocketLaunch", M + 0.17, by + 0.13, 0.44, { fill: "none", line: HEX.accent5, lineT: 25, icon: HEX.accent5, name: "Onda 1" });
  txt(s, "Onda 1 – Núcleo Comercial", { x: M + 0.8, y: by, w: 4.5, h: bh, fontSize: 16, fontFace: SEMI, color: C.white, valign: "middle" });
  txt(s, [{ text: "Previsão de execução   ", options: { color: MUTED_ON_DARK, fontFace: LIGHT } }, { text: "18 semanas ~ 4 meses", options: { fontFace: SEMI } }],
    { x: M + w1 - 4.6, y: by, w: 4.35, h: bh, fontSize: 12.5, color: C.white, valign: "middle", align: "right" });
  s.addShape(S.roundRect, { x: x2, y: by, w: w2, h: bh, rectRadius: 0.08, fill: { color: C.blue }, line: { type: "none" }, shadow: { type: "outer", color: HEX.dk2, opacity: 0.16, blur: 16, offset: 4, angle: 90 }, objectName: "Barra Onda 2" });
  txt(s, [{ text: "Onda 2", options: { fontFace: SEMI, fontSize: 13.5, breakLine: true } }, { text: "4 semanas ~ 1 mês", options: { fontFace: LIGHT, fontSize: 11.5 } }],
    { x: x2 + 0.2, y: by, w: w2 - 0.3, h: bh, color: C.white, valign: "middle" });
  [[M, "Início", "left"], [x2 - gap / 2, "Semana 18", "center"], [M + CW, "Semana 22", "right"]].forEach(([x, l, al]) => {
    txt(s, l, { x: al === "left" ? x : al === "right" ? x - 1.5 : x - 0.75, y: by + bh + 0.12, w: 1.5, h: 0.2, fontSize: 9.5, color: "8C99AD", align: al });
  });

  const o1 = [
    "Definição de regras de solução, dados e integrações.",
    "Ambientes, segurança e estrutura das 4 entidades, com linhas de solução, casas e centros de custo.",
    "Fluxo guiado de vendas, oportunidade com vários produtos e proposta com modelo.",
    "Alçadas de aprovação no Teams e Outlook.",
    "Alçada de desconto.",
    "Visão 360º, metas, previsão e painel de funil no Power BI.",
    "Integração com ERP Dynamics e CRM Nacional.",
    "Migração de dados em aberto, testes, treinamento e operação assistida.",
  ];
  const iy = 3.72, rh = 0.76, colW = 3.75;
  const item = (x, y, w, n, t) => {
    txt(s, n, { x, y: y + 0.01, w: 0.42, h: 0.22, fontSize: 10.5, fontFace: SEMI, color: C.blue });
    txt(s, t, { x: x + 0.45, y: y - 0.02, w: w - 0.45, h: 0.66, fontSize: 12.5, color: C.ink });
  };
  o1.forEach((t, i) => {
    const col = Math.floor(i / 4), row = i % 4;
    item(M + col * (colW + 0.35), iy + row * rh, colW, String(i + 1).padStart(2, "0"), t);
  });
  const px = 8.85, pw = W - M - px, py = 3.5, ph = 3.3;
  s.addShape(S.roundRect, { x: px, y: py, w: pw, h: ph, rectRadius: 0.08, fill: { color: C.pale }, line: { type: "none" }, objectName: "Painel Onda 2" });
  kicker(s, "ONDA 2", px + 0.35, py + 0.32, pw - 0.7);
  txt(s, "Painéis Avançados e CNPJ", { x: px + 0.35, y: py + 0.6, w: pw - 0.6, h: 0.4, fontSize: 16, fontFace: LIGHT, color: C.navy });
  ["Painéis Avançados de Produtos e Similaridade no Power BI.", "Integração para cadastro de Empresas por CNPJ."]
    .forEach((t, i) => item(px + 0.35, py + 1.35 + i * 0.9, pw - 0.65, String(i + 1).padStart(2, "0"), t));
  s.addNotes("Duas ondas. Onda 1 (18 semanas, ~4 meses): o núcleo comercial completo — regras, ambientes, fluxo de vendas, alçadas, Power BI, integrações e migração. Onda 2 (4 semanas, ~1 mês): painéis avançados e cadastro por CNPJ.");
}

// =====================================================================
// 7. 05 – INVESTIMENTO (opções)
// =====================================================================
{
  pres.addSection({ title: "05 – Investimento" });
  const s = contentSlide("05 – Investimento", "05 – INVESTIMENTO", "Duas alternativas de contratação",
    "Modelo Time & Materials: faturamento conforme horas medidas e aprovadas.");
  const gap = 0.35, cw = (CW - gap) / 2, y = 2.42, ch = 3.42, p = 0.45;
  const opts = [
    { x: M, dark: false, ic: "Cube", k: "OPÇÃO A · NÚCLEO COMERCIAL", chip: "Onda 1", price: "R$ 378.770,00", detail: "", hours: "2.450 horas" },
    { x: M + cw + gap, dark: true, ic: "StackPlus", k: "OPÇÃO B · PROJETO COMPLETO", chip: "Ondas 1 e 2", price: "R$ 415.564,80", detail: "(R$ 378.770,00 + R$ 36.795,00)", hours: "2.688 horas" },
  ];
  opts.forEach((o) => {
    const fg = o.dark ? C.white : C.navy, mute = o.dark ? MUTED_ON_DARK : C.gray;
    if (o.dark) s.addShape(S.roundRect, { x: o.x, y, w: cw, h: ch, rectRadius: 0.08, fill: { color: C.navy }, line: { type: "none" }, shadow: { type: "outer", color: HEX.dk2, opacity: 0.2, blur: 24, offset: 6, angle: 90 }, objectName: o.k });
    else card(s, o.x, y, cw, ch, { name: o.k });
    kicker(s, o.k, o.x + p, y + 0.45, cw - 1.8, o.dark ? C.cyan : C.blue, 11);
    ring(s, o.ic, o.x + cw - p - 0.72, y + 0.32, 0.72, o.dark ? { fill: "none", line: HEX.accent5, lineT: 25, icon: HEX.accent5, name: o.k } : { name: o.k });
    const chipW = o.chip.length * 0.095 + 0.5;
    s.addShape(S.roundRect, { x: o.x + p, y: y + 0.85, w: chipW, h: 0.32, rectRadius: 0.16, fill: o.dark ? { type: "none" } : { color: C.pale }, line: o.dark ? { color: HEX.accent5, width: 0.75, transparency: 30 } : { type: "none" }, objectName: "Escopo" });
    txt(s, o.chip, { x: o.x + p, y: y + 0.85, w: chipW, h: 0.32, fontSize: 11, fontFace: SEMI, color: o.dark ? C.cyan : C.blue, align: "center", valign: "middle" });
    txt(s, o.price, { x: o.x + p, y: y + 1.38, w: cw - 2 * p, h: 0.8, fontSize: 42, fontFace: LIGHT, color: fg });
    if (o.detail) txt(s, o.detail, { x: o.x + p, y: y + 2.18, w: cw - 2 * p, h: 0.26, fontSize: 12, color: mute });
    hline(s, o.x + p, y + 2.62, cw - 2 * p, o.dark ? "2C4D86" : HEX.accent4);
    addIcon(s, "Clock", o.dark ? HEX.accent5 : HEX.accent1, o.x + p, y + 2.83, 0.34);
    txt(s, o.hours, { x: o.x + p + 0.5, y: y + 2.78, w: 3, h: 0.44, fontSize: 19, fontFace: LIGHT, color: fg, valign: "middle" });
  });
  // como funciona o T&M
  const fy = 6.1, sw = 2.9, sgap = 0.5, sx0 = M + 1.95;
  kicker(s, "COMO FUNCIONA", M, fy + 0.2, 2, C.gray, 10);
  [["Timer", "Horas medidas"], ["CheckCircle", "Horas aprovadas"], ["Receipt", "Faturamento"]].forEach(([ic, t], i) => {
    const x = sx0 + i * (sw + sgap);
    s.addShape(S.roundRect, { x, y: fy, w: sw, h: 0.64, rectRadius: 0.32, fill: { color: C.white }, line: { color: HEX.accent4, width: 0.75 }, objectName: "Etapa T&M " + (i + 1) });
    ring(s, ic, x + 0.1, fy + 0.1, 0.44, { name: t });
    txt(s, t, { x: x + 0.7, y: fy, w: sw - 0.8, h: 0.64, fontSize: 14, color: C.navy, valign: "middle" });
    if (i < 2) addIcon(s, "CaretRight", HEX.accent3, x + sw + 0.11, fy + 0.17, 0.3);
  });
  s.addNotes("Opção A cobre a Onda 1 (núcleo comercial); a Opção B inclui as duas ondas. Modelo Time & Materials: o faturamento acompanha as horas medidas e aprovadas.");
}

// =====================================================================
// 8. 05 – INVESTIMENTO (rateio)
// =====================================================================
{
  const s = contentSlide("05 – Investimento", "05 – INVESTIMENTO", "Rateio por casa", "Visão geral para apoiar a decisão.");
  const colors = [HEX.dk2, HEX.accent1, HEX.accent3];
  const d = 3.7, cy0 = 2.55;
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Rateio por casa", labels: ["FIEPA", "SESI", "SENAI"], values: [33.33, 33.33, 33.33] }], {
    x: M, y: cy0, w: d, h: d, holeSize: 78, chartColors: colors, showLegend: false, showValue: false, showPercent: false,
    showTitle: false, dataBorder: { pt: 2.5, color: "FFFFFF" }, objectName: "Gráfico rateio",
  });
  txt(s, "33,33%", { x: M, y: cy0 + d / 2 - 0.4, w: d, h: 0.5, fontSize: 28, fontFace: LIGHT, color: C.navy, align: "center" });
  txt(s, "por casa", { x: M, y: cy0 + d / 2 + 0.12, w: d, h: 0.25, fontSize: 11, color: C.gray, align: "center" });

  const tx = 4.85, tw = W - M - tx;
  card(s, tx, cy0 - 0.05, tw, d + 0.1, { name: "Cartão tabela" });
  const ix = tx + 0.3, iw = tw - 0.6, colW = [2.6, 1.45, 1.54, 1.543];
  const hair = { type: "solid", pt: 0.75, color: HEX.accent4 }, none = { type: "none" }, strong = { type: "solid", pt: 1, color: HEX.dk2 };
  const cell = (t, j, o = {}) => ({ text: t, options: { align: j ? "right" : "left", margin: j ? [0, 0.05, 0, 0.1] : [0, 0.1, 0, 0.38], ...o } });
  const head = ["CASA", "PORCENTAGEM", "OPÇÃO A", "OPÇÃO B"].map((t, j) => cell(t, j, { fontSize: 9.5, fontFace: SEMI, color: j === 3 ? HEX.accent1 : HEX.accent2, charSpacing: 1, border: [none, none, strong, none] }));
  const rows = [
    ["FIEPA (IEL, Observatório, REDES e CIN)", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
    ["SESI", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
    ["SENAI", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
  ].map((r) => r.map((c, j) => cell(c, j, { fontSize: 13, color: HEX.dk1, border: [none, none, hair, none] })));
  const total = ["TOTAL", "100%", "R$ 378.770,00", "R$ 415.564,80"].map((c, j) => cell(c, j, { fontSize: 13.5, fontFace: SEMI, color: HEX.dk2, border: [strong, none, none, none] }));
  const rH = [0.45, 0.78, 0.58, 0.58, 0.62], tH = rH.reduce((a, b) => a + b), ty = cy0 + (d - tH) / 2;
  s.addTable([head, ...rows, total], { x: ix, y: ty, w: iw, colW, rowH: rH, valign: "middle", fontFace: BODY, objectName: "Tabela rateio" });
  let yy = ty + rH[0];
  ["FIEPA", "SESI", "SENAI"].forEach((n, i) => {
    s.addShape(S.ellipse, { x: ix + 0.12, y: yy + rH[i + 1] / 2 - 0.07, w: 0.14, h: 0.14, fill: { color: colors[i] }, line: { type: "none" }, objectName: "Legenda " + n });
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
    { ex: true, ic: "CloudSlash", tag: "FORA DO ESCOPO", t: "Licenças Dynamics 365 e consumo Azure não estão incluídos." },
    { ex: false, n: "40", tag: "MOBILIZAÇÃO", t: "A equipe pode ser mobilizada em até 40 dias úteis após aceite e assinatura." },
    { ex: false, n: "5", tag: "APROVAÇÃO DE HORAS", t: "Horas medidas mensalmente; a FIEPA tem até 5 dias úteis para aprovar o relatório." },
    { ex: true, ic: "Headset", tag: "FORA DO ESCOPO", t: "Sustentação após a Onda 2 e integração com WhatsApp não estão incluídas." },
  ];
  const gap = 0.35, cw = (CW - gap) / 2, ch = 1.95, zw = 1.7;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + gap), y = 2.5 + Math.floor(i / 2) * (ch + 0.3);
    card(s, x, y, cw, ch, { name: "Limite " + (i + 1) });
    s.addShape(S.roundRect, { x: x + 0.12, y: y + 0.12, w: zw, h: ch - 0.24, rectRadius: 0.06, fill: { color: it.ex ? "FBF1F0" : HEX.lt2 }, line: { type: "none" }, objectName: "Destaque " + (i + 1) });
    if (it.ex) {
      ring(s, it.ic, x + 0.12 + (zw - 0.8) / 2, y + 0.42, 0.8, { fill: C.white, line: "F0D5D2", icon: HEX.accent6, name: "Limite " + (i + 1) });
      txt(s, "Não incluído", { x: x + 0.12, y: y + 1.35, w: zw, h: 0.24, fontSize: 10.5, fontFace: SEMI, color: C.coral, align: "center" });
    } else {
      txt(s, it.n, { x: x + 0.12, y: y + 0.3, w: zw, h: 0.95, fontSize: 52, fontFace: LIGHT, color: C.blue, align: "center", valign: "middle" });
      txt(s, "dias úteis", { x: x + 0.12, y: y + 1.32, w: zw, h: 0.26, fontSize: 12, color: C.blue, align: "center" });
    }
    kicker(s, it.tag, x + zw + 0.5, y + 0.42, cw - zw - 0.8, it.ex ? C.coral : C.blue, 10);
    txt(s, it.t, { x: x + zw + 0.5, y: y + 0.76, w: cw - zw - 0.8, h: 1.0, fontSize: 16, fontFace: LIGHT, color: C.navy, lineSpacingMultiple: 1.05 });
  });
  s.addNotes("Limites da proposta: licenças e Azure fora do escopo; mobilização em até 40 dias úteis após aceite e assinatura; horas medidas mensalmente, com até 5 dias úteis para aprovação; sustentação pós-Onda 2 e WhatsApp fora do escopo.");
}

// =====================================================================
// 10. OBRIGADO
// =====================================================================
{
  pres.addSection({ title: "Encerramento" });
  const s = pres.addSlide({ masterName: "FIEPA Escuro", sectionTitle: "Encerramento" });
  orbit(s, 10.3, 3.75);
  txt(s, "Obrigado!", { x: M, y: 2.5, w: 7, h: 1.2, fontSize: 70, fontFace: LIGHT, color: C.white });
  txt(s, [{ text: "Reimplantação CRM Dynamics", options: { fontFace: SEMI, color: C.cyan, breakLine: true } },
          { text: "Gestão comercial integrada para o Sistema FIEPA", options: { fontFace: LIGHT, color: MUTED_ON_DARK } }],
    { x: M, y: 3.85, w: 7, h: 0.85, fontSize: 18, lineSpacingMultiple: 1.25 });
  s.addImage({ path: "assets/logo_white.png", x: M, y: 6.3, w: 2.0, h: 0.346, objectName: "Logo FIEPA" });
  s.addNotes("Encerramento e abertura para dúvidas.");
}

// ---------- write + theme + PNG fallbacks for SVG icons ----------
(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  const svgs = Object.keys(zip.files).filter((f) => /^ppt\/media\/.*\.svg$/.test(f));
  for (const f of svgs) {
    const png = f.replace(/\.svg$/, ".png");
    if (!zip.files[png]) continue;
    zip.file(png, await sharp(await zip.file(f).async("nodebuffer"), { density: 300 }).resize(256, 256).png().toBuffer());
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("ok", OUT, svgs.length, "svg icons");
})();
