import { NextRequest, NextResponse } from "next/server";
import { parseNewsletter } from "@/lib/parseNewsletter";
import { buildDocx, reportFileName } from "@/lib/buildDocx";

export const runtime = "nodejs";

// POST /api/compliance-docx
//   - Content-Type: application/json  -> { "html": "<email html>" }
//   - Content-Type: text/html | text/plain -> raw email HTML as body
export async function POST(req: NextRequest) {
  try {
    const ct = req.headers.get("content-type") || "";
    const html = ct.includes("application/json")
      ? (await req.json()).html
      : await req.text();

    if (!html || typeof html !== "string") {
      return NextResponse.json({ error: "Email HTML is required" }, { status: 400 });
    }

    const newsletter = parseNewsletter(html);
    if (!newsletter.categories.length) {
      return NextResponse.json({ error: "No updates found in the email" }, { status: 422 });
    }

    const buffer = await buildDocx(newsletter);
    const filename = reportFileName(newsletter);

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message ?? "Failed to generate document" }, { status: 500 });
  }
}