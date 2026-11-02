"""Export memo.md to memo.pdf.

Draws the football-field chart (memo_football_field.svg) from the Summary sheet of the recalculated
nike_model.xlsx, converts the Markdown subset used in memo.md to HTML, and prints it to PDF with the
headless Microsoft Edge or Google Chrome already on the machine (no extra Python packages).

Usage:
    python export_memo.py

Recalculate nike_model.xlsx first (open and save it in Excel, or run it through LibreOffice):
openpyxl writes formulas without values, and the chart is drawn from the calculated values.
"""
import html
import os
import re
import subprocess
import tempfile
import time

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
MODEL = os.path.join(HERE, "nike_model.xlsx")
MEMO = os.path.join(HERE, "memo.md")
SVG = os.path.join(HERE, "memo_football_field.svg")
PDF = os.path.join(HERE, "memo.pdf")
BROWSERS = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe"]

CSS = """
@page { size: Letter; margin: 0.5in 0.55in; }
body { font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 1.32; color: #111; }
h1 { font-size: 15pt; color: #1F3864; margin: 0 0 2pt; }
h2 { font-size: 10.5pt; color: #1F3864; border-bottom: 1px solid #1F3864; margin: 9pt 0 3pt; padding-bottom: 1pt; }
p { margin: 3pt 0; }
ul, ol { margin: 2pt 0 2pt 15pt; padding: 0; }
li { margin: 1.5pt 0; }
table { border-collapse: collapse; width: 100%; margin: 4pt 0; font-size: 8.3pt; }
th { background: #1F3864; color: #fff; padding: 2pt 4pt; text-align: center; }
td { border-bottom: 1px solid #ccc; padding: 2pt 4pt; text-align: center; }
td:first-child, th:first-child { text-align: left; }
.figure { text-align: center; margin: 4pt 0; }
em { color: #444; }
"""


def football_field_svg():
    """Horizontal range bars (low to high) with a base-case tick and the current price as a dashed line."""
    ws = load_workbook(MODEL, data_only=True)["Summary"]
    rows = {str(ws[f"A{r}"].value): r for r in range(1, ws.max_row + 1) if ws[f"A{r}"].value}
    price = ws[f"C{rows['Current share price (close)']}"].value
    if price is None:
        raise SystemExit("nike_model.xlsx has no calculated values: recalculate it first.")
    first, last = rows["Method"] + 1, rows["Current share price (reference line)"] - 1
    bars = [(ws[f"A{r}"].value, ws[f"C{r}"].value, ws[f"D{r}"].value, ws[f"E{r}"].value) for r in range(first, last + 1)]

    left, width, row_h, top = 250, 400, 26, 28
    x_max = 10 * (int(max(b[2] for b in bars) * 1.08) // 10 + 1)
    x = lambda v: left + v / x_max * width
    height = top + row_h * len(bars) + 44
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {left + width + 20} {height}" width="{left + width + 20}" '
           f'height="{height}" font-family="Arial" font-size="11">']
    for v in range(0, x_max + 1, 10):  # gridlines and axis labels
        out.append(f'<line x1="{x(v):.1f}" y1="{top - 6}" x2="{x(v):.1f}" y2="{top + row_h * len(bars)}" stroke="#ddd"/>')
        out.append(f'<text x="{x(v):.1f}" y="{top + row_h * len(bars) + 14}" text-anchor="middle" fill="#555">${v}</text>')
    for i, (label, lo, hi, base) in enumerate(bars):
        y = top + i * row_h
        out.append(f'<text x="{left - 8}" y="{y + 15}" text-anchor="end">{html.escape(label)}</text>')
        out.append(f'<rect x="{x(lo):.1f}" y="{y + 4}" width="{x(hi) - x(lo):.1f}" height="{row_h - 9}" fill="#1F3864"/>')
        out.append(f'<text x="{x(lo) - 4:.1f}" y="{y + 15}" text-anchor="end" font-size="9" fill="#333">${lo:.0f}</text>')
        out.append(f'<text x="{x(hi) + 4:.1f}" y="{y + 15}" font-size="9" fill="#333">${hi:.0f}</text>')
        if base is not None:
            out.append(f'<line x1="{x(base):.1f}" y1="{y + 2}" x2="{x(base):.1f}" y2="{y + row_h - 3}" stroke="#F2B705" stroke-width="3"/>')
    out.append(f'<line x1="{x(price):.1f}" y1="{top - 10}" x2="{x(price):.1f}" y2="{top + row_h * len(bars)}" '
               f'stroke="#C00000" stroke-width="1.5" stroke-dasharray="4 3"/>')
    out.append(f'<text x="{x(price):.1f}" y="{top - 14}" text-anchor="middle" fill="#C00000">Current ${price:.2f}</text>')
    out.append(f'<text x="{left + width}" y="{height - 4}" text-anchor="end" fill="#555" font-size="9">'
               'Bars: low-high range; yellow tick: base / median; dashed line: current price. US$ per share</text>')
    out.append("</svg>")
    with open(SVG, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print(f"Wrote {SVG} ({len(bars)} methods, current price ${price:.2f})")


def inline(text):
    """Bold, italic, inline code and links inside one line of Markdown."""
    text = html.escape(text, quote=False)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"\*(.+?)\*", r"<em>\1</em>", text)
    text = re.sub(r"`(.+?)`", r"<code>\1</code>", text)
    return re.sub(r"\[(.+?)\]\((.+?)\)", r'<a href="\2">\1</a>', text)


def markdown_to_html(md):
    """The subset memo.md uses: headings, paragraphs, bullet / numbered lists, pipe tables, SVG images."""
    out, lines, i = [], md.splitlines(), 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line:
            i += 1
        elif line.startswith("#"):
            level = len(line) - len(line.lstrip("#"))
            out.append(f"<h{level}>{inline(line[level:].strip())}</h{level}>")
            i += 1
        elif m := re.match(r"!\[(.*?)\]\((.+?)\)", line):
            with open(os.path.join(HERE, m.group(2)), encoding="utf-8") as f:
                out.append(f'<div class="figure">{f.read()}</div>')
            i += 1
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-+:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            head = "".join(f"<th>{inline(c)}</th>" for c in rows[0])
            body = "".join("<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r) + "</tr>" for r in rows[1:])
            out.append(f"<table><tr>{head}</tr>{body}</table>")
        elif re.match(r"(- |\d+\. )", line):
            tag = "ul" if line.startswith("- ") else "ol"
            items = []  # [text, [sub-bullets]]; one level of nesting (indented "  - ")
            while i < len(lines) and re.match(r"\s*(- |\d+\. )", lines[i]):
                text = re.sub(r"^\s*(- |\d+\. )", "", lines[i])
                if lines[i].startswith(" ") and items:
                    items[-1][1].append(text)
                else:
                    items.append([text, []])
                i += 1
            lis = "".join(f"<li>{inline(t)}" + ("<ul>" + "".join(f"<li>{inline(s)}</li>" for s in subs) + "</ul>" if subs else "")
                          + "</li>" for t, subs in items)
            out.append(f"<{tag}>{lis}</{tag}>")
        else:
            para = []
            while i < len(lines) and lines[i].strip() and not re.match(r"(#|\||- |\d+\. |!\[)", lines[i]):
                para.append(lines[i].strip())
                i += 1
            out.append(f"<p>{inline(' '.join(para))}</p>")
    return "\n".join(out)


def main():
    football_field_svg()
    with open(MEMO, encoding="utf-8") as f:
        body = markdown_to_html(f.read())
    page = f'<!doctype html><html><head><meta charset="utf-8"><title>NKE memo</title><style>{CSS}</style></head><body>{body}</body></html>'
    browser = next((b for b in BROWSERS if os.path.exists(b)), None)
    if browser is None:
        raise SystemExit("No Edge or Chrome found to print the PDF.")
    with tempfile.TemporaryDirectory() as tmp:
        page_path = os.path.join(tmp, "memo.html")
        with open(page_path, "w", encoding="utf-8") as f:
            f.write(page)
        if os.path.exists(PDF):
            os.remove(PDF)
        subprocess.run([browser, "--headless", "--disable-gpu", "--no-pdf-header-footer", "--no-first-run",
                        f"--user-data-dir={tmp}", f"--print-to-pdf={PDF}", "file:///" + page_path.replace("\\", "/")],
                       check=True, timeout=120)
        # On Windows the Edge launcher can return before the PDF is written: wait for the file to appear and stop growing.
        size = -1
        for _ in range(60):
            time.sleep(1)
            if os.path.exists(PDF) and os.path.getsize(PDF) == size > 0:
                break
            size = os.path.getsize(PDF) if os.path.exists(PDF) else -1
        else:
            raise SystemExit("The browser did not write the PDF within 60 seconds.")
    print(f"Wrote {PDF} ({os.path.getsize(PDF):,} bytes)")


if __name__ == "__main__":
    main()
