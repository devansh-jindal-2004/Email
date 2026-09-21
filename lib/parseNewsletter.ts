import * as cheerio from "cheerio";

export interface UpdateItem {
  title: string;
  region: string;
  summary: string;
  link: string;
}
export interface Category {
  name: string;
  count?: number;
  items: UpdateItem[];
}
export interface Newsletter {
  date: string; // e.g. "17th September 2026"
  categories: Category[];
}

const clean = (s: string) =>
  s.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();

const DATE_RE = /^(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})/;
const CAT_RE = /^(.*?)\s*\((\d+)\s+updates?\)\s*$/i;

export function parseNewsletter(html: string): Newsletter {
  const $ = cheerio.load(html);
  let date = "";
  const categories: Category[] = [];
  let current: Category | null = null;

  // h4 = date line / category heading, h3 = individual update (in document order)
  $("h4, h3").each((_, el) => {
    const tag = el.tagName.toLowerCase();
    const text = clean($(el).text());

    if (tag === "h4") {
      const d = text.match(DATE_RE);
      if (d) {
        date = d[1];
        return;
      }
      const c = text.match(CAT_RE);
      if (c) {
        current = { name: c[1], count: Number(c[2]), items: [] };
        categories.push(current);
      }
      return;
    }

    // h3 -> update item
    const a = $(el).find("a[href]").first();
    const link = a.attr("href");
    if (!link || !current) return;

    // the item's own <table> holds title, region (<b>) and summary
    const row = $(el).closest("table");
    const region = clean(row.find("b").first().text());
    const summaryLink = row.find("p a[href]").last();
    const summary = clean(summaryLink.text());

    current.items.push({ title: text, region, summary, link });
  });

  return { date, categories };
}