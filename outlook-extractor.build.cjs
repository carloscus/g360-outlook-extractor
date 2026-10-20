// G360 Outlook Extractor — deck build script
// Ground: sandwich (ink cover/close, paper body)
// Motif: G360 dot-chevron isotype as recurring marker
// Dominance: accent #00d084 on paper, paper on accent (ink slides)
const tokensPath = ".slides/tokens.cjs";
const T = require(require("path").resolve(tokensPath));
const PptxGenJS = T.loadPptxgen();

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.author = "G360";
pptx.title = "G360 Outlook Extractor";

// ── tokens ─────────────────────────────────────────────────────────
const A     = "#" + T.colours.accent;
const INK   = "#" + T.colours.ink;
const PAPER = "#" + T.colours.paper;
const MUTED = "#" + T.colours.muted;
const SURF  = "#" + T.colours.surface;
const DANGER= "#" + T.colours.danger;
const WARN  = "#" + T.colours.warning;
const SUCC  = "#" + T.colours.success;
const INFO  = "#" + T.colours.info;
const SCALE = T.scale;
const MX    = T.margins.x;
const MY    = T.margins.top;

function titleSlide(bg, title, sub, meta) {
  const s = pptx.addSlide();
  s.background = { color: bg };
  [{ c: MUTED, o: 0.5 }, { c: A, o: 1 }, { c: MUTED, o: 0.5 }].forEach((d, i) =>
    s.addShape("oval", { x: MX+0.05, y: MY+0.8+i*0.2, w: 0.1, h: 0.1,
      fill: { color: d.c, transparency: (1-d.o)*100 } })
  );
  s.addText(">", { x: MX+0.2, y: MY+0.72, w: 0.25, h: 0.22,
    fontSize: SCALE.caption, color: A, fontFace: T.fonts.body, bold: true });
  s.addText(title, { x: MX+0.6, y: MY+0.6, w: 8.2, h: 0.7,
    fontSize: SCALE.title, color: bg===INK ? PAPER : INK,
    fontFace: T.fonts.heading, bold: true, valign: "top", slidesLead: true });
  if (sub) s.addText(sub, { x: MX+0.3, y: MY+1.3, w: 8.2, h: 0.4,
    fontSize: SCALE.body, color: MUTED, fontFace: T.fonts.body, valign: "top" });
  if (meta) s.addText(meta, { x: MX+0.3, y: 4.8, w: 6, h: 0.3,
    fontSize: SCALE.caption, color: MUTED, fontFace: T.fonts.body, align: "left" });
  return s;
}

function bulletSlide(title, bullets) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  s.addShape("rect", { x: MX+0.02, y: MY+0.02, w: 0.07, h: 0.07,
    fill: { color: A }, line: { color: A, width: 0.5 } });
  s.addText(title, { x: MX+0.3, y: MY+0.1, w: 8.2, h: 0.45,
    fontSize: SCALE.h1, color: INK, fontFace: T.fonts.heading, bold: true,
    slidesLead: true });
  bullets.forEach((t, i) => {
    const y = MY + 0.65 + i * 0.32;
    s.addShape("oval", { x: MX+0.35, y, w: 0.05, h: 0.05,
      fill: { color: A }, line: { color: A, width: 0.5 } });
    s.addText(t, { x: MX+0.5, y, w: 8.2, h: 0.27,
      fontSize: SCALE.body, color: INK, fontFace: T.fonts.body, valign: "top" });
  });
  return s;
}

function colorSlide() {
  const s = pptx.addSlide();
  s.background = { color: INK };
  s.addText("Paleta G360", { x: MX+0.3, y: MY+0.1, w: 8.2, h: 0.45,
    fontSize: SCALE.h1, color: PAPER, fontFace: T.fonts.heading, bold: true,
    slidesLead: true });
  const swatches = [
    { label: "Accent",    hex: "#00D084", bg: A },
    { label: "Ink",       hex: "#0B1220", bg: INK },
    { label: "Surface",   hex: "#1A2332", bg: SURF },
    { label: "Muted",     hex: "#94A3B8", bg: MUTED },
    { label: "Danger",    hex: "#B91C1C", bg: DANGER },
    { label: "Warning",   hex: "#F59E0B", bg: WARN },
    { label: "Success",   hex: "#22C55E", bg: SUCC },
  ];
  const cols=4, swW=1.9, swH=0.85, gap=0.15;
  const sx=MX+0.3, sy=MY+0.65;
  swatches.forEach((sw, i) => {
    const col=i%cols, row=Math.floor(i/cols);
    const x=sx+col*(swW+gap), y=sy+row*(swH+gap);
    s.addShape("rect", { x, y, w:swW, h:swH*0.6, fill:{color:sw.bg} });
    s.addText(sw.label, { x, y:y+swH*0.62, w:swW, h:0.2,
      fontSize:SCALE.caption, color:PAPER, fontFace:T.fonts.body, align:"center", valign:"top" });
    s.addText(sw.hex, { x, y:y+swH*0.62+0.2, w:swW, h:0.2,
      fontSize:SCALE.caption, color:MUTED, fontFace:T.fonts.body, align:"center", valign:"top" });
  });
  return s;
}

// ── Slides ─────────────────────────────────────────────────────────
titleSlide(INK, "G360 Outlook Extractor", "Auditoría y extracción inteligente de correos MAPI", "v1.1.0 · G360 Desktop");
bulletSlide("¿Qué es?", [
  "Extrae correos con filtros: fecha, remitente, palabra clave",
  "Descarga adjuntos organizados con renombrado inteligente",
  "Genera archivo de contexto (.txt) con metadatos del correo",
  "Soporta subcarpetas y múltiples cuentas MAPI",
]);
bulletSlide("El problema", [
  "Adjuntos importantes dispersos en múltiples carpetas",
  "Búsqueda manual que toma horas",
  "Dificultad para auditar qué se recibió y cuándo",
  "Falta de trazabilidad en documentos empresariales",
]);
bulletSlide("La solución", [
  "Filtrado MAPI en tiempo real — sin descargar todo",
  "Extracción selectiva por extensión (Excel, PDF, Docs)",
  "Código OC extraído automáticamente del nombre del archivo",
  "Reporte de auditoría en tiempo real con tabla interactiva",
]);
bulletSlide("Características principales", [
  "Filtros: fechas, keyword, remitente, nombre de adjunto",
  "Subcarpetas: procesamiento recursivo con un solo clic",
  "Tema claro/oscuro: paleta G360 oficial, contraste WCAG AA",
  "Deduplicación: adjuntos con código OC se agrupan",
  "Errores traducidos: mensajes COM legibles",
]);
colorSlide();

// Cierre
{
  const s = pptx.addSlide();
  s.background = { color: INK };
  s.addText("G360 Desktop", { x:MX, y:MY+1.5, w:8.2, h:0.8,
    fontSize:SCALE.display, color:PAPER, fontFace:T.fonts.heading, bold:true, align:"center",
    slidesLead:true });
  const cx=6.5, cy=0.8;
  [{c:MUTED,o:0.5},{c:A,o:1},{c:MUTED,o:0.5}].forEach((d,i) =>
    s.addShape("oval",{x:cx-0.05,y:cy+i*0.25,w:0.12,h:0.12,fill:{color:d.c,tran:(1-d.o)*100}})
  );
  s.addText(">", { x:cx+0.08, y:cy-0.02, w:0.35, h:0.3,
    fontSize:22, color:A, fontFace:T.fonts.body, bold:true });
  s.addText("Powered by G360", { x:MX, y:4.8, w:8.2, h:0.4,
    fontSize:SCALE.body, color:MUTED, fontFace:T.fonts.body, align:"center" });
}

pptx.writeFile({ fileName: "outlook-extractor.pptx" })
  .then(() => console.log("PPTX written: outlook-extractor.pptx"))
  .catch(err => console.error("Error:", err));
