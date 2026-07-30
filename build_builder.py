#!/usr/bin/env python3
"""
Assemble the self-contained "Cot State Builder.html" from:
  - builder_src.html   (the drag-&-drop UI + in-browser extractor, with placeholders)
  - template.html      (the dashboard shell, embedded as base64)
  - vendor/xlsx.full.min.js  (SheetJS, inlined)

Run this only when you change builder_src.html or template.html.
Day-to-day, users just open "Cot State Builder.html" and drop in a spreadsheet.
"""
import base64, json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))

def read(p):
    return open(os.path.join(HERE, p), encoding="utf-8").read()

def main():
    sheetjs = read(os.path.join("vendor", "xlsx.full.min.js"))
    if "</script" in sheetjs.lower():
        sys.exit("SheetJS contains a </script> — cannot inline safely.")
    tpl_b64 = base64.b64encode(read("template.html").encode("utf-8")).decode("ascii")

    # geo.json is the single source of truth for unit locations — bake it in.
    geo = {k: v for k, v in json.loads(read("geo.json")).items() if not k.startswith("_")}
    geo_json = json.dumps(geo, ensure_ascii=False)

    src = read("builder_src.html")
    for marker in ("/*SHEETJS*/", "/*TEMPLATE_B64*/", "/*GEO_JSON*/"):
        if marker not in src:
            sys.exit(f"Placeholder {marker} missing from builder_src.html")

    out = (src.replace("/*GEO_JSON*/", geo_json, 1)
              .replace("/*SHEETJS*/", sheetjs, 1)
              .replace("/*TEMPLATE_B64*/", tpl_b64, 1))
    # Friendly name for sending, plus index.html so GitHub Pages serves it at the site root.
    for name in ("Cot State Builder.html", "index.html"):
        open(os.path.join(HERE, name), "w", encoding="utf-8").write(out)
        print(f"Wrote {name}  ({round(len(out)/1024)} KB)")

if __name__ == "__main__":
    main()
