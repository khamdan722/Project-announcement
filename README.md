# Project-announcement

Content and tooling for the **2026 medical awards winners announcement** —
مؤسسة حمدان بن راشد للعلوم الطبية والتربوية / Hamdan Bin Rashid Foundation for
Medical and Educational Sciences.

The deliverable is the winners document handed to the media team, in the same
layout as the 2024 release. Arbitration results arrive over time, so winner
content lives in small per-winner files and the document is regenerated from
whatever is confirmed at that moment.

## Layout

```
data/awards.yml       the five awards, their order, sections and winner counts
data/winners/*.yml    one file per winner slot (10 in total)
scripts/              build_press_release.py — renders the Word document
reference/            the 2024 release, kept as the styling and tone reference
output/               the generated document handed to the media team
```

## The five awards

| # | Award | Winners | Type |
|---|---|---|---|
| 1 | الجائزة العربية للأبحاث في القطاع الصحي — Arab Award for Research in Healthcare | 2 | research papers |
| 2 | الجائزة العربية في العلوم الوراثية — Arab Award in Genetics | 1 | personality |
| 3 | جائزة أفضل بحث في القطاع الصحي — Best Research in Healthcare | 3 | research papers |
| 4 | جائزة حمدان للمتميزين في القطاع الصحي — Hamdan Award for Distinguished Personalities | 2 | personalities |
| 5 | جائزة الابتكار في القطاع الصحي — Innovation in Healthcare Award | 2 | healthcare projects |

Awards 1–2 sit under **أولاً: جوائز العالم العربي**, awards 3–5 under
**ثانياً: جوائز دولة الإمارات العربية المتحدة**. There is no ranking — all
winners are equal.

## Adding a winner

Open the winner's file in `data/winners/`, fill in the three content fields and
move its status forward:

| field | column in the document | language |
|---|---|---|
| `name_ar` | الاسم — paper/project title, or the person's name | Arabic |
| `notes_ar` | الملاحظات — research team and institution, or the person's position | Arabic |
| `bio_en` | نبذة عن الفائز — 250–400 word write-up | English |

`status` moves `pending → confirmed → drafted → approved`. Anything still
`pending` is printed in the document as **بانتظار نتائج التحكيم**, so a current
version can be produced at any time.

List every document a write-up was built from under `sources`. Facts that no
source supports do not go in.

## Building the document

```sh
pip install -r requirements.txt
python3 scripts/build_press_release.py
```

This writes `output/winners-press-release-2026.docx` and prints how many names
and write-ups are in place per award. The document matches the 2024 release:
Dubai font, right-to-left tables, teal `#3D959D` award headers, gold `#B99664`
section bands, Letter page with 0.5" margins.

Use `-o path.docx` to write somewhere else.
