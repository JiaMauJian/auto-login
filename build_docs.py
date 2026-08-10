"""把 .md 說明文件轉成單一檔案的 HTML（樣式內嵌、可離線雙擊開啟）。

用法：
    python build_docs.py                 # 轉換預設的兩份文件到 dist\
    python build_docs.py a.md b.md       # 指定要轉換的檔案

給沒有 VS Code／Markdown 編輯器的電腦看用的：.md 在 Windows 預設會被記事本開啟，
表格和換行會糊成一團；轉成 HTML 後用任何瀏覽器雙擊就能正常閱讀。
"""

import sys
from pathlib import Path

import markdown

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
DEFAULT_DOCS = ["簡易說明.md", "詳細說明.md"]

# 字型只用 Windows 內建的，不連外部 CDN（那台電腦可能沒網路，也不該有外連）。
CSS = """
:root {
  --bg: #ffffff;
  --fg: #1f2328;
  --muted: #656d76;
  --border: #d8dee4;
  --code-bg: #f2f4f6;
  --block-bg: #f6f8fa;
  --accent: #0969da;
  --warn-bg: #fff8e6;
  --warn-border: #e6b800;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #14171a;
    --fg: #e4e6e8;
    --muted: #9198a1;
    --border: #30363d;
    --code-bg: #22272c;
    --block-bg: #1b1f23;
    --accent: #5aa9f8;
    --warn-bg: #2b2412;
    --warn-border: #9a7b00;
  }
}
* { box-sizing: border-box; }
html { -webkit-text-size-adjust: 100%; }
body {
  margin: 0;
  background: var(--bg);
  color: var(--fg);
  font-family: "Microsoft JhengHei UI", "Microsoft JhengHei", "Segoe UI",
               "PingFang TC", "Noto Sans TC", sans-serif;
  font-size: 17px;
  line-height: 1.85;
}
main {
  max-width: 780px;
  margin: 0 auto;
  padding: 48px 24px 96px;
}
h1, h2, h3 { line-height: 1.4; font-weight: 700; }
h1 {
  font-size: 1.9em;
  margin: 0 0 32px;
  padding-bottom: 14px;
  border-bottom: 2px solid var(--border);
}
h2 {
  font-size: 1.4em;
  margin: 48px 0 16px;
  padding-bottom: 8px;
  border-bottom: 1px solid var(--border);
}
h3 { font-size: 1.15em; margin: 32px 0 12px; }
p, ul, ol { margin: 0 0 16px; }
ul, ol { padding-left: 1.6em; }
li { margin: 6px 0; }
li > ul, li > ol { margin: 6px 0; }
hr { border: 0; border-top: 1px solid var(--border); margin: 40px 0; }
a { color: var(--accent); }
strong { font-weight: 700; }

code {
  font-family: Consolas, "Cascadia Mono", "Courier New", monospace;
  font-size: 0.88em;
  background: var(--code-bg);
  border-radius: 4px;
  padding: 0.15em 0.4em;
  word-break: break-all;
}
pre {
  background: var(--block-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 14px 16px;
  overflow-x: auto;
  margin: 0 0 16px;
}
pre code {
  background: none;
  padding: 0;
  font-size: 0.92em;
  line-height: 1.7;
  white-space: pre;
  word-break: normal;
}

/* 表格可能比手機螢幕寬，讓它自己捲，不要撐破整頁 */
.table-wrap { overflow-x: auto; margin: 0 0 20px; }
table { border-collapse: collapse; width: 100%; }
th, td {
  border: 1px solid var(--border);
  padding: 9px 14px;
  text-align: left;
  vertical-align: top;
}
th { background: var(--block-bg); font-weight: 700; white-space: nowrap; }

blockquote {
  margin: 0 0 16px;
  padding: 2px 16px;
  border-left: 4px solid var(--border);
  color: var(--muted);
}

@media print {
  body { font-size: 12pt; }
  main { max-width: none; padding: 0; }
  pre, .table-wrap { break-inside: avoid; }
}
"""

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>{css}</style>
</head>
<body>
<main>
{body}
</main>
</body>
</html>
"""


def convert(md_path: Path, out_dir: Path) -> Path:
    text = md_path.read_text(encoding="utf-8")
    html = markdown.markdown(
        text,
        extensions=["tables", "fenced_code", "sane_lists", "nl2br"],
    )
    # 讓寬表格自己水平捲動，而不是把整頁撐寬
    html = html.replace("<table>", '<div class="table-wrap"><table>')
    html = html.replace("</table>", "</table></div>")

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / (md_path.stem + ".html")
    out_path.write_text(
        TEMPLATE.format(title=md_path.stem, css=CSS, body=html),
        encoding="utf-8",
    )
    return out_path


def main() -> int:
    names = sys.argv[1:] or DEFAULT_DOCS
    for name in names:
        src = Path(name)
        if not src.is_absolute():
            src = ROOT / name
        if not src.exists():
            print(f"找不到 {src}", file=sys.stderr)
            return 1
        out = convert(src, DIST)
        print(f"{src.name} -> {out.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
