// Reimplantação CRM Dynamics — Sistema FIEPA (v2: minimalista)
// Deck .pptx editável: textos, formas, tabela e gráfico nativos; ícones SVG de traço fino
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
    accent4: "D9E2EE", accent5: "3FC1E8", accent6: "0D3E86",
    hlink: "1660C4", folHlink: "6B7A90",
  },
};
const HEX = THEME.colors;
const C = { ink: "tx1", navy: "tx2", white: "bg1", pale: "bg2", blue: "accent1", gray: "accent2", sky: "accent3", hair: "accent4", cyan: "accent5", deep: "accent6" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Sistema FIEPA";
pres.company = "FIEPA";
pres.title = "Reimplantação CRM Dynamics";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const S = pres.ShapeType;

const W = 13.333, M = 0.75, CW = W - 2 * M; // 11.833
const LIGHT = "Segoe UI Light", SEMI = "Segoe UI Semibold", BODY = "Segoe UI";

// ---------- helpers ----------
function icon(name, hex) {
  const svg = RDS.renderToStaticMarkup(React.createElement(pi["Pi" + name + "Light"], { size: 256 }))
    .replace(/currentColor/g, "#" + hex).replace(/ style="[^"]*"/, "");
  return "image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}
function addIcon(s, name, hex, x, y, d) {
  s.addImage({ data: icon(name, hex), x, y, w: d, h: d, objectName: "Ícone " + name });
}
function txt(s, text, o) {
  s.addText(text, { isTextBox: true, margin: 0, valign: "top", fontFace: BODY, fontSize: 14, color: C.ink, ...o });
}
const kicker = (s, t, x, y, w, color = C.blue) =>
  txt(s, t, { x, y, w, h: 0.25, fontSize: 10.5, fontFace: SEMI, color, charSpacing: 1.5 });
function hline(s, x, y, w, color = HEX.accent4, width = 0.75) {
  s.addShape(S.line, { x, y, w, h: 0, line: { color, width }, objectName: "Linha" });
}
function vline(s, x, y, h, color = HEX.accent4, width = 0.75) {
  s.addShape(S.line, { x, y, w: 0, h, line: { color, width }, objectName: "Linha" });
}

// ---------- layouts ----------
const footer = [
  { image: { path: "assets/logo_navy.png", x: 11.33, y: 0.62, w: 1.25, h: 0.216 } },
  { text: { text: "Reimplantação CRM Dynamics  ·  Sistema FIEPA", options: { x: M, y: 7.0, w: 6, h: 0.22, fontSize: 9, color: "8C99AD", fontFace: BODY, margin: 0 } } },
];
pres.defineSlideMaster({
  title: "FIEPA Conteudo",
  background: { path: "assets/bg_light2.png" },
  objects: [
    ...footer,
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.62, w: 8, h: 0.25, fontSize: 10.5, fontFace: SEMI, color: C.blue, charSpacing: 1.5, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.92, w: 10.2, h: 0.7, fontSize: 34, fontFace: LIGHT, color: C.navy, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: M, y: 1.66, w: 10, h: 0.5, fontSize: 14, fontFace: BODY, color: C.gray, margin: 0, valign: "top" }, text: "" } },
  ],
  slideNumber: { x: 12.18, y: 7.0, w: 0.4, h: 0.22, fontSize: 9, color: "8C99AD", align: "right", fontFace: BODY },
});
pres.defineSlideMaster({
  title: "FIEPA Limpo",
  background: { path: "assets/bg_light2.png" },
  objects: [footer[1]],
  slideNumber: { x: 12.18, y: 7.0, w: 0.4, h: 0.22, fontSize: 9, color: "8C99AD", align: "right", fontFace: BODY },
});
pres.defineSlideMaster({ title: "FIEPA Escuro", background: { path: "assets/bg_dark2.png" }, objects: [] });

function contentSlide(section, k, title, sub) {
  const s = pres.addSlide({ masterName: "FIEPA Conteudo", sectionTitle: section });
  s.addText(k, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "sub" });
  return s;
}

// =====================================================================
// 1. CAPA
// =====================================================================
{
  pres.addSection({ title: "Capa" });
  const s = pres.addSlide({ masterName: "FIEPA Escuro", sectionTitle: "Capa" });
  s.addImage({ path: "assets/logo_white.png", x: M, y: 0.7, w: 2.0, h: 0.346, objectName: "Logo FIEPA" });
  kicker(s, "PROPOSTA  ·  SISTEMA FIEPA", M, 2.3, 6, C.cyan);
  txt(s, [
    { text: "Reimplantação", options: { fontFace: LIGHT, breakLine: true } },
    { text: "CRM Dynamics", options: { fontFace: SEMI } },
  ], { x: M, y: 2.62, w: 9, h: 1.95, fontSize: 58, color: C.white, lineSpacingMultiple: 0.95 });
  txt(s, "Gestão comercial integrada para o Sistema FIEPA", { x: M, y: 4.62, w: 9, h: 0.4, fontSize: 18, fontFace: LIGHT, color: "C9D8EE" });

  const cols = [
    ["PROCESSO ÚNICO", "Etapas e regras comerciais comuns"],
    ["VISÃO INTEGRADA", "de clientes e oportunidades"],
    ["CONTROLE E GESTÃO", "Indicadores e rastreabilidade"],
  ];
  const gap = 0.45, cw = (CW - 2 * gap) / 3, y = 5.75;
  cols.forEach(([t, d], i) => {
    const x = M + i * (cw + gap);
    hline(s, x, y, cw, "5B7BB0", 0.75);
    txt(s, String(i + 1).padStart(2, "0"), { x, y: y + 0.22, w: 0.5, h: 0.25, fontSize: 11, fontFace: SEMI, color: C.cyan });
    txt(s, t, { x: x + 0.5, y: y + 0.22, w: cw - 0.5, h: 0.25, fontSize: 11, fontFace: SEMI, color: C.white, charSpacing: 1.5 });
    txt(s, d, { x: x + 0.5, y: y + 0.52, w: cw - 0.5, h: 0.32, fontSize: 14, fontFace: LIGHT, color: "C9D8EE" });
  });
  s.addNotes("Abertura: a proposta de reimplantação do CRM Dynamics para o Sistema FIEPA, apoiada em três pilares — processo único, visão integrada e controle e gestão.");
}

// =====================================================================
// 2. AGENDA
// =====================================================================
{
  pres.addSection({ title: "Agenda" });
  const s = pres.addSlide({ masterName: "FIEPA Limpo", sectionTitle: "Agenda" });
  s.addImage({ path: "assets/logo_navy.png", x: M, y: 0.62, w: 1.25, h: 0.216, objectName: "Logo FIEPA" });
  kicker(s, "ROTEIRO", M, 2.6, 3);
  txt(s, "Agenda", { x: M, y: 2.9, w: 3.6, h: 0.9, fontSize: 48, fontFace: LIGHT, color: C.navy });
  txt(s, "Reimplantação CRM Dynamics", { x: M, y: 3.85, w: 3.6, h: 0.35, fontSize: 14, color: C.gray });

  const items = [
    ["JUSTIFICATIVA", "O CRM atual limita a gestão comercial"],
    ["VALOR PARA A GESTÃO", "O que muda na prática"],
    ["INTEGRAÇÕES E CONTROLE", "Sistemas conectados, acessos controlados"],
    ["ENTREGAS", "Etapas do projeto de reimplantação"],
    ["INVESTIMENTO", "Duas alternativas de contratação"],
    ["ESCOPO", "O que não está incluído"],
  ];
  const x0 = 5.3, xw = W - M - x0, y0 = 0.85, rh = 0.95;
  items.forEach(([sec, d], i) => {
    const y = y0 + i * rh;
    txt(s, String(i + 1).padStart(2, "0"), { x: x0, y: y + 0.2, w: 0.8, h: 0.5, fontSize: 24, fontFace: LIGHT, color: C.blue });
    txt(s, sec, { x: x0 + 0.95, y: y + 0.17, w: xw - 0.95, h: 0.22, fontSize: 9.5, fontFace: SEMI, color: C.gray, charSpacing: 1.5 });
    txt(s, d, { x: x0 + 0.95, y: y + 0.42, w: xw - 0.95, h: 0.4, fontSize: 19, fontFace: LIGHT, color: C.navy });
    hline(s, x0, y + rh, xw);
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
  const gap = 0.6, cw = (CW - 2 * gap) / 3, y = 2.6;
  pains.forEach(([ic, t, d], i) => {
    const x = M + i * (cw + gap);
    if (i) vline(s, x - gap / 2, y, 1.85);
    addIcon(s, ic, HEX.accent1, x, y, 0.46);
    txt(s, t, { x, y: y + 0.72, w: cw, h: 0.25, fontSize: 11, fontFace: SEMI, color: C.navy, charSpacing: 1.5 });
    txt(s, d, { x, y: y + 1.07, w: cw, h: 0.8, fontSize: 14, color: C.gray, lineSpacingMultiple: 1.1 });
  });
  // consequência — faixa marinho de largura total
  const by = 4.95;
  s.addShape(S.rect, { x: 0, y: by, w: W, h: 1.85, fill: { color: C.navy }, line: { type: "none" }, objectName: "Faixa consequências" });
  kicker(s, "CONSEQUÊNCIAS PARA A GESTÃO", M, by + 0.42, 6, C.cyan);
  txt(s, "Menor capacidade de acompanhar oportunidades, padronizar decisões e planejar resultados com base em informações confiáveis.",
    { x: M, y: by + 0.78, w: 10.8, h: 0.85, fontSize: 21, fontFace: LIGHT, color: C.white, lineSpacingMultiple: 1.05 });
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
  const gx = M, gw = 7.9, gy = 2.55, cellW = gw / 2, cellH = 1.95;
  vline(s, gx + cellW, gy, cellH * 2);
  hline(s, gx, gy + cellH, gw);
  vals.forEach(([ic, t, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = gx + col * cellW + (col ? 0.45 : 0), y = gy + row * cellH + (row ? 0.35 : 0);
    const w = cellW - 0.45;
    addIcon(s, ic, HEX.accent1, x, y, 0.42);
    txt(s, t, { x: x + 0.62, y: y + 0.09, w: w - 0.62, h: 0.25, fontSize: 11, fontFace: SEMI, color: C.navy, charSpacing: 1.5 });
    txt(s, d, { x: x + 0.62, y: y + 0.45, w: w - 0.62, h: 0.8, fontSize: 14, color: C.gray, lineSpacingMultiple: 1.1 });
  });
  // resultado esperado
  const px = 9.15, pw = W - M - px, py = gy, ph = cellH * 2;
  s.addShape(S.rect, { x: px, y: py, w: pw, h: ph, fill: { color: C.navy }, line: { type: "none" }, objectName: "Resultado esperado" });
  kicker(s, "RESULTADO ESPERADO", px + 0.4, py + 0.45, pw - 0.8, C.cyan);
  txt(s, "Um processo comercial comum, com dados rastreáveis e melhor suporte à decisão.",
    { x: px + 0.4, y: py + 0.85, w: pw - 0.8, h: 2.4, fontSize: 21, fontFace: LIGHT, color: C.white, lineSpacingMultiple: 1.1 });
  s.addNotes("O ganho vai além da atualização técnica: cliente 360º, governança, decisão apoiada em metas e previsão, e eficiência com menos controles paralelos.");
}

// =====================================================================
// 5. 03 – INTEGRAÇÕES E CONTROLE
// =====================================================================
{
  pres.addSection({ title: "03 – Integrações e controle" });
  const s = contentSlide("03 – Integrações e controle", "03 – INTEGRAÇÕES E CONTROLE", "Sistemas conectados, acessos controlados",
    "A proposta prevê rotas nativas Microsoft e camada Azure para integrações específicas.");
  const cx = W / 2, cy = 3.33, r = 0.85;
  const lw = 4.0, yT = 2.5, yB = 3.7;
  const nodes = [
    { x: M, y: yT, al: "right", t: "Dynamics 365 F&SCM", d: "Contas nos dois sentidos; produtos e preços → CRM" },
    { x: M, y: yB, al: "right", t: "Outlook + Teams", d: "E-mails, reuniões, aprovações e avisos" },
    { x: W - M - lw, y: yT, al: "left", t: "CRM Nacional (CNI)", d: "Envio a cada alteração, rastreio e reprocessamento" },
    { x: W - M - lw, y: yB, al: "left", t: "SharePoint", d: "Documentos por conta e oportunidade" },
  ];
  nodes.forEach((n) => {
    const ax = n.al === "right" ? n.x + lw + 0.2 : n.x - 0.2, ay = n.y + 0.14;
    const dx = cx - ax, dy = cy - ay, L = Math.hypot(dx, dy);
    const ex = cx - (dx / L) * (r + 0.12), ey = cy - (dy / L) * (r + 0.12);
    const o = { x: Math.min(ax, ex), y: Math.min(ay, ey), w: Math.abs(ex - ax), h: Math.abs(ey - ay), line: { color: "9DB8DA", width: 0.75 }, objectName: "Conexão " + n.t };
    if ((ex - ax) * (ey - ay) < 0) o.flipV = true;
    s.addShape(S.line, o);
    s.addShape(S.ellipse, { x: ax - 0.045, y: ay - 0.045, w: 0.09, h: 0.09, fill: { color: C.blue }, line: { type: "none" }, objectName: "Ponto" });
    txt(s, n.t, { x: n.x, y: n.y, w: lw, h: 0.3, fontSize: 15, fontFace: SEMI, color: C.navy, align: n.al });
    txt(s, n.d, { x: n.x, y: n.y + 0.36, w: lw, h: 0.5, fontSize: 12.5, color: C.gray, align: n.al });
  });
  s.addShape(S.ellipse, { x: cx - r - 0.12, y: cy - r - 0.12, w: 2 * r + 0.24, h: 2 * r + 0.24, fill: { type: "none" }, line: { color: "B9CFEA", width: 0.75 }, objectName: "Anel" });
  s.addShape(S.ellipse, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: C.navy }, line: { type: "none" }, objectName: "Dynamics 365 Sales" });
  txt(s, [{ text: "DYNAMICS 365", options: { breakLine: true } }, { text: "SALES (CRM)" }],
    { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fontSize: 11, fontFace: SEMI, color: C.white, align: "center", valign: "middle", charSpacing: 2 });
  // Onda 2
  s.addShape(S.line, { x: cx, y: cy + r + 0.12, w: 0, h: 0.3, line: { color: HEX.accent1, width: 0.75, dashType: "dash" }, objectName: "Conexão Onda 2" });
  txt(s, "ONDA 2", { x: cx - 2, y: cy + r + 0.5, w: 4, h: 0.22, fontSize: 9.5, fontFace: SEMI, color: C.blue, charSpacing: 1.5, align: "center" });
  txt(s, "Consulta por CNPJ", { x: cx - 2, y: cy + r + 0.74, w: 4, h: 0.3, fontSize: 14, fontFace: SEMI, color: C.navy, align: "center" });
  txt(s, "Complemento de cadastro de empresas", { x: cx - 2, y: cy + r + 1.06, w: 4, h: 0.28, fontSize: 12.5, color: C.gray, align: "center" });

  // segurança e privacidade
  const sy = 5.68;
  hline(s, M, sy, CW);
  kicker(s, "SEGURANÇA E PRIVACIDADE", M, sy + 0.2, 4);
  const sec = [
    ["Segregação", "Cada entidade vê e altera seus registros conforme o perfil."],
    ["Perfis e hierarquia", "Acesso por função; gestores acompanham suas equipes."],
    ["Campos sensíveis", "Valores e descontos protegidos por perfil."],
    ["Auditoria", "Registro de alterações nas tabelas principais."],
  ];
  const sx0 = M + 3.0, sw = (W - M - sx0) / 4;
  sec.forEach(([t, d], i) => {
    const x = sx0 + i * sw;
    txt(s, t, { x, y: sy + 0.17, w: sw - 0.3, h: 0.28, fontSize: 13, fontFace: SEMI, color: C.navy });
    txt(s, d, { x, y: sy + 0.48, w: sw - 0.3, h: 0.6, fontSize: 11.5, color: C.gray });
  });
  s.addNotes("O Dynamics 365 Sales é o centro: integra-se ao F&SCM (contas nos dois sentidos), ao CRM Nacional da CNI, ao Outlook/Teams e ao SharePoint. A consulta por CNPJ entra na Onda 2. Segurança: segregação por entidade, perfis e hierarquia, campos sensíveis protegidos e auditoria.");
}

// =====================================================================
// 6. 04 – ENTREGAS
// =====================================================================
{
  pres.addSection({ title: "04 – Entregas" });
  const s = contentSlide("04 – Entregas", "04 – ENTREGAS", "Etapas do projeto de reimplantação", "Cada etapa tem um propósito, de forma macro");

  // linha do tempo proporcional: 18 + 4 semanas
  const ty = 2.5, unit = CW / 22, x18 = M + unit * 18, xEnd = M + CW;
  txt(s, "ONDA 1", { x: M, y: ty - 0.32, w: 2, h: 0.22, fontSize: 9.5, fontFace: SEMI, color: C.navy, charSpacing: 1.5 });
  txt(s, "ONDA 2", { x: x18 + 0.12, y: ty - 0.32, w: 2, h: 0.22, fontSize: 9.5, fontFace: SEMI, color: C.blue, charSpacing: 1.5 });
  s.addShape(S.rect, { x: M, y: ty, w: x18 - M - 0.04, h: 0.06, fill: { color: C.navy }, line: { type: "none" }, objectName: "Linha do tempo Onda 1" });
  s.addShape(S.rect, { x: x18 + 0.04, y: ty, w: xEnd - x18 - 0.04, h: 0.06, fill: { color: C.sky }, line: { type: "none" }, objectName: "Linha do tempo Onda 2" });
  [[M, "Início", "left"], [x18, "Semana 18", "center"], [xEnd, "Semana 22", "right"]].forEach(([x, l, al]) => {
    s.addShape(S.ellipse, { x: x - 0.07, y: ty - 0.04, w: 0.14, h: 0.14, fill: { color: C.white }, line: { color: HEX.dk2, width: 1 }, objectName: "Marco " + l });
    txt(s, l, { x: al === "left" ? x : al === "right" ? x - 1.5 : x - 0.75, y: ty + 0.18, w: 1.5, h: 0.2, fontSize: 9.5, color: C.gray, align: al });
  });

  // colunas: Onda 1 (8 entregas) | Onda 2 (2 entregas)
  const y0 = 3.08, c2x = 8.9, c2w = W - M - c2x;
  vline(s, c2x - 0.42, y0, 3.72);
  const head = (x, w, name, dur) => {
    txt(s, name, { x, y: y0, w, h: 0.4, fontSize: 19, fontFace: LIGHT, color: C.navy });
    txt(s, [{ text: "Previsão de execução: ", options: { color: C.gray } }, { text: dur, options: { fontFace: SEMI, color: C.navy } }],
      { x, y: y0 + 0.46, w, h: 0.25, fontSize: 11.5 });
  };
  head(M, 7.6, "Núcleo Comercial", "18 semanas ~ 4 meses");
  head(c2x, c2w, "Painéis Avançados e CNPJ", "4 semanas ~ 1 mês");

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
  const iy = y0 + 1.0, rh = 0.72, colW = 3.6;
  const item = (x, y, w, n, t) => {
    txt(s, n, { x, y: y + 0.01, w: 0.4, h: 0.2, fontSize: 10, fontFace: SEMI, color: C.blue });
    txt(s, t, { x: x + 0.4, y: y - 0.02, w: w - 0.4, h: 0.64, fontSize: 12, color: C.ink });
  };
  o1.forEach((t, i) => {
    const col = Math.floor(i / 4), row = i % 4;
    item(M + col * (colW + 0.3), iy + row * rh, colW, String(i + 1).padStart(2, "0"), t);
  });
  ["Painéis Avançados de Produtos e Similaridade no Power BI.", "Integração para cadastro de Empresas por CNPJ."]
    .forEach((t, i) => item(c2x, iy + i * rh, c2w, String(i + 1).padStart(2, "0"), t));
  s.addNotes("Duas ondas. Onda 1 (18 semanas, ~4 meses): o núcleo comercial completo — regras, ambientes, fluxo de vendas, alçadas, Power BI, integrações e migração. Onda 2 (4 semanas, ~1 mês): painéis avançados e cadastro por CNPJ.");
}

// =====================================================================
// 7. 05 – INVESTIMENTO (opções)
// =====================================================================
{
  pres.addSection({ title: "05 – Investimento" });
  const s = contentSlide("05 – Investimento", "05 – INVESTIMENTO", "Duas alternativas de contratação",
    "Modelo Time & Materials: faturamento conforme horas medidas e aprovadas.");
  const gap = 0.4, cw = (CW - gap) / 2, y = 2.5, ch = 4.15;
  const opts = [
    { x: M, dark: false, k: "OPÇÃO A", name: "Núcleo Comercial", scope: "Onda 1", price: "R$ 378.770,00", detail: "", hours: "2.450" },
    { x: M + cw + gap, dark: true, k: "OPÇÃO B", name: "Projeto Completo", scope: "Ondas 1 e 2", price: "R$ 415.564,80", detail: "(R$ 378.770,00 + R$ 36.795,00)", hours: "2.688" },
  ];
  opts.forEach((o) => {
    const fg = o.dark ? C.white : C.navy, mute = o.dark ? "B7C9E4" : C.gray, p = 0.5;
    s.addShape(S.rect, { x: o.x, y, w: cw, h: ch, fill: { color: o.dark ? C.navy : C.white }, line: o.dark ? { type: "none" } : { color: HEX.accent4, width: 0.75 }, objectName: o.k });
    kicker(s, o.k, o.x + p, y + 0.5, 3, o.dark ? C.cyan : C.blue);
    txt(s, o.name, { x: o.x + p, y: y + 0.8, w: cw - 2 * p, h: 0.45, fontSize: 24, fontFace: LIGHT, color: fg });
    txt(s, o.scope, { x: o.x + p, y: y + 1.3, w: cw - 2 * p, h: 0.28, fontSize: 13, color: mute });
    txt(s, o.price, { x: o.x + p, y: y + 1.9, w: cw - 2 * p, h: 0.8, fontSize: 44, fontFace: LIGHT, color: fg });
    if (o.detail) txt(s, o.detail, { x: o.x + p, y: y + 2.72, w: cw - 2 * p, h: 0.26, fontSize: 12, color: mute });
    hline(s, o.x + p, y + 3.18, cw - 2 * p, o.dark ? "2C4D86" : HEX.accent4);
    txt(s, [{ text: o.hours, options: { fontSize: 24, fontFace: LIGHT, color: fg } }, { text: "  horas", options: { fontSize: 13, color: mute } }],
      { x: o.x + p, y: y + 3.32, w: cw - 2 * p, h: 0.5, valign: "bottom" });
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
    x: M, y: cy0, w: d, h: d, holeSize: 80, chartColors: colors, showLegend: false, showValue: false, showPercent: false,
    showTitle: false, dataBorder: { pt: 2.5, color: "FFFFFF" }, objectName: "Gráfico rateio",
  });
  txt(s, "33,33%", { x: M, y: cy0 + d / 2 - 0.4, w: d, h: 0.5, fontSize: 28, fontFace: LIGHT, color: C.navy, align: "center" });
  txt(s, "por casa", { x: M, y: cy0 + d / 2 + 0.12, w: d, h: 0.25, fontSize: 11, color: C.gray, align: "center" });

  const tx = 5.35, tw = W - M - tx, colW = [2.7, 1.45, 1.539, 1.539];
  const hair = { type: "solid", pt: 0.75, color: HEX.accent4 }, none = { type: "none" }, strong = { type: "solid", pt: 1, color: HEX.dk2 };
  const cell = (t, j, o = {}) => ({ text: t, options: { align: j ? "right" : "left", margin: j ? [0, 0.05, 0, 0.1] : [0, 0.1, 0, 0.42], ...o } });
  const head = ["CASA", "PORCENTAGEM", "OPÇÃO A", "OPÇÃO B"].map((t, j) => cell(t, j, { fontSize: 9.5, fontFace: SEMI, color: HEX.accent2, charSpacing: 1, border: [none, none, strong, none] }));
  const rows = [
    ["FIEPA (IEL, Observatório, REDES e CIN)", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
    ["SESI", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
    ["SENAI", "33,33%", "R$ 126.256,67", "R$ 138.521,60"],
  ].map((r) => r.map((c, j) => cell(c, j, { fontSize: 13, color: HEX.dk1, border: [none, none, hair, none] })));
  const total = ["TOTAL", "100%", "R$ 378.770,00", "R$ 415.564,80"].map((c, j) => cell(c, j, { fontSize: 13.5, fontFace: SEMI, color: HEX.dk2, border: [strong, none, none, none] }));
  const rH = [0.45, 0.78, 0.58, 0.58, 0.62], tH = rH.reduce((a, b) => a + b), ty = cy0 + (d - tH) / 2;
  s.addTable([head, ...rows, total], { x: tx, y: ty, w: tw, colW, rowH: rH, valign: "middle", fontFace: BODY, objectName: "Tabela rateio" });
  let yy = ty + rH[0];
  ["FIEPA", "SESI", "SENAI"].forEach((n, i) => {
    s.addShape(S.ellipse, { x: tx + 0.14, y: yy + rH[i + 1] / 2 - 0.07, w: 0.14, h: 0.14, fill: { color: colors[i] }, line: { type: "none" }, objectName: "Legenda " + n });
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
  const y0 = 2.55, colW = (CW - 0.9) / 2, x2 = M + colW + 0.9;
  vline(s, M + colW + 0.45, y0, 4.1);
  kicker(s, "NÃO INCLUÍDO", M, y0, 4);
  kicker(s, "CONDIÇÕES", x2, y0, 4);
  const rowY = [y0 + 0.55, y0 + 2.4];
  [
    "Licenças Dynamics 365 e consumo Azure não estão incluídos.",
    "Sustentação após a Onda 2 e integração com WhatsApp não estão incluídas.",
  ].forEach((t, i) => {
    const y = rowY[i];
    s.addShape(S.ellipse, { x: M, y: y + 0.05, w: 0.5, h: 0.5, fill: { type: "none" }, line: { color: HEX.accent2, width: 0.75 }, objectName: "Marcador" });
    addIcon(s, "X", HEX.accent2, M + 0.13, y + 0.18, 0.24);
    txt(s, t, { x: M + 0.8, y: y + 0.02, w: colW - 0.8, h: 1.1, fontSize: 19, fontFace: LIGHT, color: C.navy, lineSpacingMultiple: 1.05 });
    if (i === 0) hline(s, M, y0 + 2.0, colW);
  });
  [
    ["40", "A equipe pode ser mobilizada em até 40 dias úteis após aceite e assinatura."],
    ["5", "Horas medidas mensalmente; a FIEPA tem até 5 dias úteis para aprovar o relatório."],
  ].forEach(([n, t], i) => {
    const y = rowY[i];
    txt(s, n, { x: x2, y: y - 0.12, w: 1.1, h: 0.8, fontSize: 48, fontFace: LIGHT, color: C.blue });
    txt(s, "dias úteis", { x: x2, y: y + 0.7, w: 1.2, h: 0.22, fontSize: 10.5, color: C.gray });
    txt(s, t, { x: x2 + 1.35, y: y + 0.02, w: colW - 1.35, h: 1.1, fontSize: 17, color: C.ink, lineSpacingMultiple: 1.05 });
    if (i === 0) hline(s, x2, y0 + 2.0, colW);
  });
  s.addNotes("Limites da proposta: licenças e Azure fora do escopo; mobilização em até 40 dias úteis após aceite e assinatura; horas medidas mensalmente, com até 5 dias úteis para aprovação; sustentação pós-Onda 2 e WhatsApp fora do escopo.");
}

// =====================================================================
// 10. OBRIGADO
// =====================================================================
{
  pres.addSection({ title: "Encerramento" });
  const s = pres.addSlide({ masterName: "FIEPA Escuro", sectionTitle: "Encerramento" });
  txt(s, "Obrigado", { x: M, y: 2.55, w: 9, h: 1.2, fontSize: 72, fontFace: LIGHT, color: C.white });
  txt(s, [{ text: "Reimplantação CRM Dynamics", options: { fontFace: SEMI, color: C.white, breakLine: true } },
          { text: "Gestão comercial integrada para o Sistema FIEPA", options: { fontFace: LIGHT, color: "C9D8EE" } }],
    { x: M, y: 3.95, w: 9, h: 0.85, fontSize: 18, lineSpacingMultiple: 1.25 });
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
