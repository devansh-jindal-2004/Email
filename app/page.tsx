"use client";

import { useRef, useState, DragEvent, ClipboardEvent } from "react";

type Status = "idle" | "loading" | "success" | "error";

export default function Home() {
  const [html, setHtml] = useState("");
  const [status, setStatus] = useState<Status>("idle");
  const [message, setMessage] = useState("");
  const [dragging, setDragging] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  async function generate(source: string) {
    if (!source.trim()) {
      setStatus("error");
      setMessage("Please paste or upload the email HTML first.");
      return;
    }
    setStatus("loading");
    setMessage("Generating Word file…");

    try {
      const res = await fetch("/api/mail", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html: source }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Request failed (${res.status})`);
      }

      // pick the filename from the response header
      const cd = res.headers.get("Content-Disposition") || "";
      const name =
        cd.match(/filename="?([^";]+)"?/)?.[1] ||
        "Daily_Regulatory_Compliance_Update.docx";

      // trigger download
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus("success");
      setMessage(`Downloaded ${name}`);
    } catch (e: any) {
      setStatus("error");
      setMessage(e?.message || "Something went wrong.");
    }
  }

  async function handleFile(file: File) {
    const text = await file.text();
    setHtml(text);
    generate(text); // auto-send on upload / drop
  }

  function onDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  }

  function onPaste(e: ClipboardEvent<HTMLTextAreaElement>) {
    const pasted = e.clipboardData.getData("text");
    if (pasted.includes("<") && pasted.includes(">")) {
      setHtml(pasted);
      e.preventDefault();
      generate(pasted); // auto-send on paste
    }
  }

  const color =
    status === "error" ? "#b42318" : status === "success" ? "#067647" : "#475467";

  return (
    <main style={styles.page}>
      <div style={styles.card}>
        <h1 style={styles.h1}>
          <span style={{ color: "#4891D1" }}>Daily</span> Regulatory Compliance Update
        </h1>
        <p style={styles.sub}>
          Upload or paste the newsletter email HTML. The Word file downloads automatically.
        </p>

        {/* Drop zone */}
        <div
          onClick={() => fileRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          style={{
            ...styles.drop,
            borderColor: dragging ? "#4891D1" : "#d0d5dd",
            background: dragging ? "#eef6fd" : "#fafafa",
          }}
        >
          <strong>Drop .html file here</strong>
          <span style={{ color: "#667085", fontSize: 13 }}>or click to browse</span>
          <input
            ref={fileRef}
            type="file"
            accept=".html,.htm,.txt,text/html,text/plain"
            hidden
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) handleFile(f);
              e.target.value = "";
            }}
          />
        </div>

        <div style={styles.or}>or paste the HTML source</div>

        <textarea
          value={html}
          onChange={(e) => setHtml(e.target.value)}
          onPaste={onPaste}
          placeholder="<html>…email HTML source…</html>"
          style={styles.textarea}
        />

        <button
          onClick={() => generate(html)}
          disabled={status === "loading"}
          style={{ ...styles.button, opacity: status === "loading" ? 0.6 : 1 }}
        >
          {status === "loading" ? "Generating…" : "Generate & Download .docx"}
        </button>

        {message && <p style={{ ...styles.msg, color }}>{message}</p>}
      </div>
    </main>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    background: "#f2f2f2",
    display: "flex",
    justifyContent: "center",
    alignItems: "flex-start",
    padding: "40px 16px",
    fontFamily: "Arial, Helvetica, sans-serif",
  },
  card: {
    width: "100%",
    maxWidth: 680,
    background: "#fff",
    borderRadius: 12,
    padding: 28,
    boxShadow: "0 1px 3px rgba(0,0,0,.1)",
  },
  h1: { fontSize: 24, margin: "0 0 6px", textAlign: "center", color: "#191919" },
  sub: { textAlign: "center", color: "#667085", margin: "0 0 22px", fontSize: 14 },
  drop: {
    border: "2px dashed #d0d5dd",
    borderRadius: 10,
    padding: "28px 16px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: 4,
    cursor: "pointer",
  },
  or: { textAlign: "center", color: "#98a2b3", fontSize: 13, margin: "16px 0 8px" },
  textarea: {
    width: "100%",
    height: 180,
    boxSizing: "border-box",
    border: "1px solid #d0d5dd",
    borderRadius: 8,
    padding: 10,
    fontFamily: "monospace",
    fontSize: 12,
    resize: "vertical",
  },
  button: {
    width: "100%",
    marginTop: 14,
    padding: "12px 0",
    background: "#2E5CAA",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    fontSize: 15,
    cursor: "pointer",
  },
  msg: { textAlign: "center", marginTop: 14, fontSize: 14 },
};