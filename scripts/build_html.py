#!/usr/bin/env python3
"""Render the winners as a web page for review.

    python3 scripts/build_html.py [--lang ar|en] [-o output/path.html]

A reading view of the same data the Word document is built from, so the
write-ups can be reviewed without opening Word.
"""

import argparse
import html
import pathlib

import yaml

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
READY = {"confirmed", "drafted", "approved"}

STYLE = """
:root{
  --teal:#3D959D; --teal-ink:#2A6B72; --gold:#B99664;
  --paper:#FAF8F4; --surface:#FFFFFF; --ink:#182224; --ink-soft:#5B6D70;
  --line:#E4E1D8; --wait:#8A7855; --wait-bg:#F6F1E6;
  color-scheme:light;
}
@media (prefers-color-scheme:dark){ :root:not([data-theme="light"]){
  --teal:#5EB4BC; --teal-ink:#8ED2D8; --gold:#CCAC79;
  --paper:#121A1B; --surface:#192325; --ink:#E7EDED; --ink-soft:#9CAFB1;
  --line:#2A3738; --wait:#C7B289; --wait-bg:#241F15;
  color-scheme:dark;
}}
:root[data-theme="dark"]{
  --teal:#5EB4BC; --teal-ink:#8ED2D8; --gold:#CCAC79;
  --paper:#121A1B; --surface:#192325; --ink:#E7EDED; --ink-soft:#9CAFB1;
  --line:#2A3738; --wait:#C7B289; --wait-bg:#241F15;
  color-scheme:dark;
}
*{box-sizing:border-box}
body{
  margin:0; background:var(--paper); color:var(--ink);
  font-family:"Noto Naskh Arabic","Times New Roman",serif;
  font-size:17px; line-height:1.95;
}
.wrap{max-width:860px; margin:0 auto; padding-inline:16px; padding-block:40px 72px}
.kufi{font-family:"Noto Kufi Arabic","Segoe UI",system-ui,sans-serif}
.eyebrow{
  font-family:"Noto Kufi Arabic",system-ui,sans-serif; font-size:12px;
  letter-spacing:.14em; color:var(--ink-soft); margin:0 0 10px;
}
h1{
  font-family:"Noto Kufi Arabic",system-ui,sans-serif; font-size:clamp(25px,5vw,34px);
  line-height:1.5; margin:0 0 6px; text-wrap:balance;
}
.sub{color:var(--ink-soft); font-size:15px; margin:0}
.tally{
  display:flex; flex-wrap:wrap; gap:8px; margin:22px 0 0; padding:0; list-style:none;
  font-family:"Noto Kufi Arabic",system-ui,sans-serif; font-size:13px;
}
.tally li{
  background:var(--surface); border:1px solid var(--line); border-radius:999px;
  padding:5px 14px; color:var(--ink-soft);
}
.tally b{color:var(--ink); font-variant-numeric:tabular-nums}
.section{margin-top:52px}
.section__band{
  background:var(--gold); color:#fff; border-radius:7px;
  font-family:"Noto Kufi Arabic",system-ui,sans-serif; font-size:17px;
  padding:11px 18px; margin:0 0 26px; text-align:center;
}
.award{margin-bottom:34px; border:1px solid var(--line); border-radius:9px; overflow:hidden; background:var(--surface)}
.award__head{background:var(--teal); color:#fff; padding:12px 18px}
.award__head h3{font-family:"Noto Kufi Arabic",system-ui,sans-serif; font-size:16px; margin:0; font-weight:700}
.award__head p{margin:3px 0 0; font-size:12.5px; opacity:.92; font-family:system-ui,sans-serif; direction:ltr; text-align:right}
.winner{padding:22px 18px; border-top:1px solid var(--line); display:flex; gap:16px}
.winner:first-of-type{border-top:0}
.num{
  flex:0 0 28px; height:28px; border-radius:50%; background:var(--teal); color:#fff;
  display:grid; place-items:center; font-family:system-ui,sans-serif; font-size:13px;
  font-variant-numeric:tabular-nums; margin-top:5px;
}
.winner__body{min-width:0; flex:1}
.name{font-family:"Noto Kufi Arabic",system-ui,sans-serif; font-size:16.5px; line-height:1.75; margin:0 0 6px}
.affil{margin:0 0 14px; color:var(--ink-soft); font-size:15px}
.affil--missing{color:var(--wait)}
.brief p{margin:0 0 13px}
.brief p:last-child{margin-bottom:0}
.wait{
  background:var(--wait-bg); border:1px dashed var(--wait); color:var(--wait);
  border-radius:7px; padding:11px 15px; font-size:14.5px;
  font-family:"Noto Kufi Arabic",system-ui,sans-serif;
}
footer{margin-top:56px; padding-top:22px; border-top:1px solid var(--line); color:var(--ink-soft); font-size:13.5px}
@media (max-width:480px){ .winner{gap:12px; padding-inline:14px} body{font-size:16px} }
"""


def esc(s):
    return html.escape(str(s or ""))


def paras(text):
    blocks = [" ".join(b.split()) for b in (text or "").split("\n\n") if b.strip()]
    return "\n".join(f"<p>{esc(b)}</p>" for b in blocks)


def build(lang):
    cfg = yaml.safe_load((DATA / "awards.yml").read_text(encoding="utf-8"))
    winners = {}
    for f in sorted((DATA / "winners").glob("*.yml")):
        d = yaml.safe_load(f.read_text(encoding="utf-8"))
        winners.setdefault(d["award"], []).append(d)

    named = written = total = 0
    body = []
    for sec in cfg["sections"]:
        body.append('<section class="section">')
        body.append(f'<h2 class="section__band">{esc(sec["heading_ar"])}</h2>')
        for award in sec["awards"]:
            slots = sorted(winners.get(award["id"], []), key=lambda w: w["id"])
            body.append('<article class="award">')
            body.append('<div class="award__head">'
                        f'<h3>{esc(award["name_ar"])}</h3>'
                        f'<p>{esc(award["name_en"])}</p></div>')
            for i, w in enumerate(slots, 1):
                total += 1
                ready = w.get("status") in READY and w.get("name_ar")
                bio = w.get(f"bio_{lang}") or ""
                if ready:
                    named += 1
                if bio.strip():
                    written += 1
                body.append('<div class="winner">')
                body.append(f'<div class="num">{i}</div>')
                body.append('<div class="winner__body">')
                if ready:
                    body.append(f'<h4 class="name">{esc(w["name_ar"])}</h4>')
                    affil = w.get("notes_ar")
                    cls = "affil" if affil else "affil affil--missing"
                    body.append(f'<p class="{cls}">{esc(affil) if affil else "الجهة غير محدّدة بعد"}</p>')
                    body.append(f'<div class="brief">{paras(bio)}</div>' if bio.strip()
                                else '<p class="wait">النبذة قيد الإعداد — بانتظار البحث وأسماء الباحثين.</p>')
                else:
                    body.append(f'<p class="wait">{esc(cfg["placeholder_ar"])}</p>')
                body.append("</div></div>")
            body.append("</article>")
        body.append("</section>")

    head = [
        '<p class="eyebrow kufi">مؤسسة حمدان بن راشد للعلوم الطبية والتربوية</p>',
        f'<h1>{esc(cfg["title_ar"])}</h1>',
        '<p class="sub">نسخة للمراجعة الداخلية — غير مخصّصة للنشر.</p>',
        '<ul class="tally">'
        f'<li>الفائزون: <b>{named}</b> من <b>{total}</b></li>'
        f'<li>النبذ المكتملة: <b>{written}</b> من <b>{total}</b></li>'
        '<li>الجوائز: <b>5</b></li></ul>',
    ]

    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{esc(cfg['title_ar'])}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Kufi+Arabic:wght@400;700&family=Noto+Naskh+Arabic:wght@400;700&display=swap">
<style>{STYLE}</style>
</head>
<body>
<main class="wrap">
{''.join(head)}
{''.join(body)}
<footer>
يُبنى هذا العرض من ملفات الفائزين في المستودع، وهو مطابق لمحتوى وثيقة الإعلان.
كل معلومة فيه مستندة إلى وثائق الترشيح وتقارير التحكيم.
</footer>
</main>
</body>
</html>"""


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-l", "--lang", choices=("ar", "en"), default="ar")
    ap.add_argument("-o", "--output", type=pathlib.Path)
    args = ap.parse_args()
    out = args.output or ROOT / "output" / f"winners-2026-{args.lang}.html"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(build(args.lang), encoding="utf-8")
    print(f"written to {out.relative_to(ROOT)}  ({out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
