import JSZip from "jszip";
import {
  DOMParser,
  XMLSerializer,
  type Element as XmlElement,
  type Node as XmlNode,
} from "@xmldom/xmldom";
import fs from "fs/promises";
import path from "path";
import type { Newsletter } from "./parseNewsletter";

/**
 * Builds the report by re-using the AKM Global template (.docx) so the banner,
 * logo, header/footer, colours, tables and fonts are identical to the original.
 * Only the text / rows / links are replaced.
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HYPERLINK_TYPE = `${R}/hyperlink`;
const TEMPLATE_PATH = path.join(process.cwd(), "templates", "compliance-template.docx");

type El = XmlElement;

// ---------- small DOM helpers ----------
const kids = (el: El, name: string): El[] =>
  Array.from(el.childNodes).filter(
    (n): n is El => n.nodeType === 1 && (n as El).localName === name && (n as El).namespaceURI === W
  );
const desc = (el: El, name: string): El[] =>
  Array.from(el.getElementsByTagNameNS(W, name)) as El[];
const removeNode = (n: XmlNode) => n.parentNode?.removeChild(n);

function setRunText(run: El, text: string) {
  const doc = run.ownerDocument;
  if (!doc) return;

  let t = kids(run, "t")[0];
  if (!t) {
    t = doc.createElementNS(W, "w:t");
    run.appendChild(t);
  }
  while (t.firstChild) t.removeChild(t.firstChild);
  t.appendChild(doc.createTextNode(text));
  t.setAttribute("xml:space", "preserve");
}

/** Put `text` in the first run of paragraph `p`, keeping its formatting; drop the other runs. */
function setPText(p: El, text: string) {
  const runs = kids(p, "r");
  if (!runs.length) return;
  setRunText(runs[0], text);
  runs.slice(1).forEach(removeNode);
}

const firstP = (tc: El) => kids(tc, "p")[0];
const pad2 = (n: number) => String(n).padStart(2, "0");

function addPPr(p: El, childName: string) {
  const doc = p.ownerDocument;
  if (!doc) return;

  let pPr = kids(p, "pPr")[0];
  if (!pPr) {
    pPr = doc.createElementNS(W, "w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  if (kids(pPr, childName).length) return;
  const el = doc.createElementNS(W, `w:${childName}`);
  const style = kids(pPr, "pStyle")[0];
  if (style && style.nextSibling) pPr.insertBefore(el, style.nextSibling);
  else if (style) pPr.appendChild(el);
  else pPr.insertBefore(el, pPr.firstChild);
}

function addCantSplit(tr: El) {
  const doc = tr.ownerDocument;
  if (!doc) return;

  let trPr = kids(tr, "trPr")[0];
  if (!trPr) {
    trPr = doc.createElementNS(W, "w:trPr");
    tr.insertBefore(trPr, kids(tr, "tc")[0]);
  }
  if (!kids(trPr, "cantSplit").length) trPr.appendChild(doc.createElementNS(W, "w:cantSplit"));
}

// ---------- text helpers ----------
const clean = (s: string) => s.replace(/\s*(\.{3}|…)\s*$/, "").replace(/\s+/g, " ").trim();
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];

function parseDate(input: string) {
  const m = input.match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})/);
  let day: number, month: number, year: number;
  if (m && MONTHS.findIndex((x) => x.toLowerCase() === m[2].toLowerCase()) >= 0) {
    day = +m[1];
    month = MONTHS.findIndex((x) => x.toLowerCase() === m[2].toLowerCase());
    year = +m[3];
  } else {
    const d = new Date();
    day = d.getDate(); month = d.getMonth(); year = d.getFullYear();
  }
  const suffix = day % 10 === 1 && day !== 11 ? "ST" : day % 10 === 2 && day !== 12 ? "ND" : day % 10 === 3 && day !== 13 ? "RD" : "TH";
  return {
    long: `${day} ${MONTHS[month]} ${year}`,                                   // 17 September 2026
    heading: `${day}${suffix} ${MONTHS[month].toUpperCase()} ${year}`,          // 17TH SEPTEMBER 2026
    short: `${pad2(day)}-${MONTHS[month].slice(0, 3)}-${String(year).slice(-2)}`, // 17-Sep-26
    fileName: `${day}_${MONTHS[month]}_${year}`,
  };
}

const GLANCE_ORDER = ["labour", "finance & taxation", "secretarial", "ehs", "commercial", "industry specific"];
const norm = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();

export function reportFileName(nl: Newsletter) {
  return `Daily_Regulatory_Compliance_Update_${parseDate(nl.date).fileName}.docx`;
}

// ---------- main ----------
export async function buildDocx(nl: Newsletter): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await fs.readFile(TEMPLATE_PATH));
  const parser = new DOMParser();
  const dom = parser.parseFromString(await zip.file("word/document.xml")!.async("string"), "text/xml");
  const rels = parser.parseFromString(await zip.file("word/_rels/document.xml.rels")!.async("string"), "text/xml");

  const body = dom.getElementsByTagNameNS(W, "body")[0] as El;
  const c = Array.from(body.childNodes).filter((n): n is El => n.nodeType === 1);
  if (c.length !== 36) throw new Error("Unexpected template layout");

  const date = parseDate(nl.date);
  const total = nl.categories.reduce((n, k) => n + k.items.length, 0);

  // 1. Reporting date (top right)
  {
    const runs = kids(c[1], "r");
    setRunText(runs[1], date.long);
    runs.slice(2).forEach(removeNode);
  }

  // 2. "Today at a glance" counters
  {
    const tcs = kids(kids(c[3], "tr")[0], "tc");
    setPText(firstP(tcs[0]), pad2(total));
    const count = (name: string) =>
      nl.categories.filter((k) => norm(k.name) === name).reduce((n, k) => n + k.items.length, 0);
    GLANCE_ORDER.forEach((name, i) => setPText(firstP(tcs[i + 1]), pad2(count(name))));
  }

  // 3. "Total vendor updates: N | Suggested reading order…"
  {
    const runs = kids(c[4], "r");
    setRunText(runs[1], String(total));
    removeNode(runs[2]);
  }

  // 4. Reporting approach paragraph (contains the date)
  {
    const p = kids(kids(kids(kids(c[5], "tr")[0], "tc")[0], "p")[1] as El, "r")[0];
    const t = kids(p, "t")[0].textContent || "";
    setRunText(p, t.replace(/\d{1,2} [A-Za-z]+ \d{4}/, date.long));
  }

  // 5. "17TH SEPTEMBER 2026 | 32 UPDATES"
  setPText(c[7], `${date.heading} | ${total} UPDATES`);

  // 6. Category sections: clone the template's header table + data table
  const catProto = c[9].cloneNode(true) as El;
  const tblProto = c[10].cloneNode(true) as El;
  const spacerA = c[11].cloneNode(true) as El;
  const spacerB = c[15].cloneNode(true) as El;
  const anchor = c[33];
  for (let i = 9; i <= 32; i++) removeNode(c[i]);

  // wipe the sample hyperlinks from the relationships file
  const relRoot = rels.documentElement;
  if (!relRoot) throw new Error("Missing document relationships root");

  Array.from(relRoot.getElementsByTagName("Relationship"))
    .filter((r) => (r as El).getAttribute("Type") === HYPERLINK_TYPE)
    .forEach((r) => removeNode(r));

  let linkNo = 0;
  const addLink = (target: string) => {
    const id = `rIdLnk${++linkNo}`;
    const r = rels.createElementNS(relRoot.namespaceURI, "Relationship");
    r.setAttribute("Id", id);
    r.setAttribute("Type", HYPERLINK_TYPE);
    r.setAttribute("Target", target);
    r.setAttribute("TargetMode", "External");
    relRoot.appendChild(r);
    return id;
  };

  nl.categories.forEach((cat, ci) => {
    if (!cat.items.length) return;

    // section header bar
    const hdr = catProto.cloneNode(true) as El;
    const hcells = desc(hdr, "tc");
    setPText(firstP(hcells[0]), cat.name.toUpperCase());
    setPText(firstP(hcells[1]), `${pad2(cat.items.length)} UPDATES`);
    desc(hdr, "p").forEach((p) => addPPr(p, "keepNext"));
    body.insertBefore(hdr, anchor);

    // data table
    const tbl = tblProto.cloneNode(true) as El;
    const rows = kids(tbl, "tr");
    const rowProto = rows[1].cloneNode(true) as El;
    rows.slice(1).forEach(removeNode);
    desc(rows[0], "p").forEach((p) => addPPr(p, "keepNext"));
    addCantSplit(rows[0]);

    cat.items.forEach((item, i) => {
      const tr = rowProto.cloneNode(true) as El;
      const tcs = kids(tr, "tc");
      setPText(firstP(tcs[0]), String(i + 1));
      setPText(firstP(tcs[1]), clean(item.title));
      setPText(firstP(tcs[2]), item.region);
      setPText(firstP(tcs[3]), date.short);
      setPText(firstP(tcs[4]), clean(item.summary));
      const link = kids(firstP(tcs[5]), "hyperlink")[0];
      link.setAttributeNS(R, "r:id", addLink(item.link));
      addCantSplit(tr);
      tbl.appendChild(tr);
    });
    body.insertBefore(tbl, anchor);

    if (ci < nl.categories.length - 1) {
      body.insertBefore(spacerA.cloneNode(true), anchor);
      body.insertBefore(spacerB.cloneNode(true), anchor);
    }
  });

  // remove duplicate paragraph ids created by cloning
  for (const el of Array.from(dom.getElementsByTagName("*"))) {
    el.removeAttribute("w14:paraId");
    el.removeAttribute("w14:textId");
  }

  const ser = new XMLSerializer();
  zip.file("word/document.xml", ser.serializeToString(dom));
  zip.file("word/_rels/document.xml.rels", ser.serializeToString(rels));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}