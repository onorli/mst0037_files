#!/usr/bin/env python3
"""Build one landing page per spreadsheet in files/ (standard library only).

Each page shows the sheet as a read-only table that works on a phone, a toggle for
formulas, a button that opens the real file in Microsoft's free online viewer, and a
download button. Run after changing a spreadsheet:  python3 build_pages.py
"""
import html, pathlib, re, urllib.parse, zipfile
import xml.etree.ElementTree as ET

ROOT = pathlib.Path(__file__).resolve().parent
RAW = "https://raw.githubusercontent.com/onorli/mst0037_files/main/files/"
SITE = "https://onorli.github.io/mst0037_files/"
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# filename -> (slug, title, booklet section, one-paragraph description)
PAGES = {
    "simple_stock_model.xlsx": ("simple-stock-model", "Simple stock model", "1.1",
        "One-step binomial model of the stock price: today's price, the up and down prices, the probability of the up state, the expected price at time T and the required return."),
    "option_2per_replicate.xlsx": ("option-2per-replicate", "Call option price: replicating portfolio", "3.1",
        "Prices the call by building a portfolio of shares and borrowing that pays exactly what the call pays in both states. Shows the hedge ratio Delta, the borrowed amount B and the resulting option price."),
    "option_2per_risk_neutral.xlsx": ("option-2per-risk-neutral", "Call option price: risk-neutral approach", "3.2",
        "Prices the same call with the risk-neutral probability: compute the probability that makes the stock earn the risk-free rate, then discount the expected payoff at the risk-free rate."),
    "stock_more_steps.xlsx": ("stock-more-steps", "Stock price: more steps", "4.1",
        "Three-step binomial tree for the stock price built from the up factor u and the down factor d = 1/u. Each node is the previous node times u or d."),
    "option_more_steps.xlsx": ("option-more-steps", "Call option price: more steps", "4.2",
        "Values the call on the three-step tree. The Price sheet holds the stock tree, the Option sheet the payoffs at the final nodes, the risk-neutral probabilities of reaching them and the discounted option value."),
    "binomial_vs_bs.xlsx": ("binomial-vs-bs", "Binomial model versus Black-Scholes", "5.3 and 7.3",
        "Four sheets: the Black-Scholes calculator of section 5.3 as described in the booklet; a one-step and a multi-step binomial tree with u and d derived from the volatility and the risk-neutral probability per step; and the Black-Scholes price and put value computed from the same inputs, so the two models can be compared as the number of steps grows."),
}

def col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0); n = 0
    for ch in letters: n = n * 26 + ord(ch) - 64
    return n

def col_letters(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

def read_workbook(path):
    z = zipfile.ZipFile(path); shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    sheets = []
    for s in wb.find("m:sheets", NS):
        target = rels[s.get(RID)]; target = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
        cells = {}
        for c in ET.fromstring(z.read(target)).iter("{%s}c" % NS["m"]):
            t = c.get("t"); v = c.find("m:v", NS); f = c.find("m:f", NS); val = None
            if v is not None: val = shared[int(v.text)] if t == "s" else v.text
            elif t == "inlineStr": val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
            formula = ("=" + f.text) if f is not None and f.text else None
            if val is not None or formula: cells[c.get("r")] = (val, formula)
        sheets.append((s.get("name"), cells))
    return sheets

def fmt(val):
    if val is None: return ""
    try:
        x = float(val)
    except ValueError:
        return val
    if x == int(x) and abs(x) < 1e12: return str(int(x))
    return f"{x:.4f}".rstrip("0").rstrip(".")

def sheet_table(name, cells, show_name):
    rows = [int(re.search(r"\d+", r).group(0)) for r in cells]; cols = [col_index(r) for r in cells]
    r0, r1, c0, c1 = min(rows), max(rows), min(cols), max(cols)
    out = []
    if show_name: out.append(f"<h2>Sheet: {html.escape(name)}</h2>")
    out.append('<div class="wrap"><table class="sheet"><thead><tr><th></th>')
    out += [f"<th>{col_letters(c)}</th>" for c in range(c0, c1 + 1)]
    out.append("</tr></thead><tbody>")
    for r in range(r0, r1 + 1):
        out.append(f"<tr><th>{r}</th>")
        for c in range(c0, c1 + 1):
            ref = f"{col_letters(c)}{r}"; val, formula = cells.get(ref, (None, None))
            cls = ' class="f"' if formula else ""
            data = f' data-formula="{html.escape(formula)}"' if formula else ""
            out.append(f"<td{cls}{data}><span>{html.escape(fmt(val))}</span></td>")
        out.append("</tr>")
    out.append("</tbody></table></div>")
    return "\n".join(out)

CSS = """
:root{--bg:#fff;--fg:#1b1b1b;--muted:#5f6b73;--line:#d9dee2;--acc:#1f78b4;--fcell:#eef6fb}
@media (prefers-color-scheme:dark){:root:not([data-theme=light]){--bg:#15181b;--fg:#e8eaec;--muted:#9aa5ad;--line:#333a40;--acc:#6ab0de;--fcell:#1d2a34}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
main{max-width:860px;margin:0 auto;padding:24px 16px 48px}h1{font-size:1.5rem;margin:0 0 4px}h2{font-size:1.1rem;margin:28px 0 8px}
.sec{color:var(--muted);margin:0 0 16px}p{margin:0 0 16px}
.btns{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 20px}
.btn{display:inline-block;padding:10px 16px;border-radius:8px;border:1px solid var(--acc);color:var(--acc);text-decoration:none;font-weight:600}
.btn.primary{background:var(--acc);color:#fff}
.wrap{overflow-x:auto;border:1px solid var(--line);border-radius:8px}
table.sheet{border-collapse:collapse;font-size:14px;min-width:100%}
.sheet th,.sheet td{border:1px solid var(--line);padding:6px 8px;white-space:nowrap;text-align:left}
.sheet th{background:var(--line);color:var(--muted);font-weight:500;text-align:center}
.sheet td.f{background:var(--fcell)}
.hint{color:var(--muted);font-size:14px;margin:8px 0 0}
.toggle{margin:12px 0 0;padding:8px 14px;border-radius:8px;border:1px solid var(--line);background:var(--bg);color:var(--fg);font:inherit;cursor:pointer}
footer{margin-top:40px;color:var(--muted);font-size:14px}a{color:var(--acc)}
"""

JS = """
document.querySelector('.toggle').addEventListener('click',function(){
  var on=this.dataset.on==='1';document.querySelectorAll('td.f').forEach(function(td){
    td.querySelector('span').textContent=on?td.dataset.value:td.dataset.formula;});
  this.dataset.on=on?'0':'1';this.textContent=on?'Show formulas':'Show values';});
document.querySelectorAll('td.f').forEach(function(td){td.dataset.value=td.querySelector('span').textContent;});
"""

def page(fname, slug, title, section, desc, sheets):
    raw = RAW + fname
    viewer = "https://view.officeapps.live.com/op/view.aspx?src=" + urllib.parse.quote(raw, safe="")
    tables = "\n".join(sheet_table(n, c, len(sheets) > 1) for n, c in sheets)
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title><style>{CSS}</style></head>
<body><main>
<h1>{html.escape(title)}</h1>
<p class="sec">Options booklet, section {section} &middot; MST0037</p>
<p>{html.escape(desc)}</p>
<div class="btns">
<a class="btn primary" href="{viewer}">Open in browser</a>
<a class="btn" href="{raw}">Download .xlsx</a>
</div>
{tables}
<p class="hint">Shaded cells are formulas. On a PC, download the file to change the inputs and see the results update.</p>
<button class="toggle" data-on="0">Show formulas</button>
<footer><a href="{SITE}">All course files</a></footer>
</main><script>{JS}</script></body></html>
"""

def index(entries):
    items = "\n".join(f'<li><a href="{SITE}{slug}/">{html.escape(title)}</a> <span class="sec">section {sec}</span></li>' for slug, title, sec in entries)
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>MST0037 course files</title><style>{CSS} li{{margin:8px 0}} .sec{{display:inline}}</style></head>
<body><main><h1>MST0037 course files</h1>
<p>Spreadsheets referred to by the QR codes in the Options booklet and lecture notes. Each page shows the sheet, opens it in the browser, or downloads the file.</p>
<ul>{items}</ul></main></body></html>
"""

if __name__ == "__main__":
    entries = []
    for fname, (slug, title, sec, desc) in PAGES.items():
        sheets = read_workbook(ROOT / "files" / fname)
        d = ROOT / slug; d.mkdir(exist_ok=True)
        (d / "index.html").write_text(page(fname, slug, title, sec, desc, sheets))
        entries.append((slug, title, sec)); print("wrote", slug + "/index.html")
    (ROOT / "index.html").write_text(index(entries)); print("wrote index.html")
