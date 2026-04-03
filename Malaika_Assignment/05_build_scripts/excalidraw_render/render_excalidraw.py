"""Render Excalidraw JSON to PNG using Playwright + headless Chromium.

Usage:
    python render_excalidraw.py <path.excalidraw> [--output path.png] [--scale 2]
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path


def validate(data):
    errs = []
    if data.get("type") != "excalidraw":
        errs.append(f"Expected type 'excalidraw', got '{data.get('type')}'")
    if "elements" not in data:
        errs.append("Missing 'elements' array")
    elif not isinstance(data["elements"], list) or len(data["elements"]) == 0:
        errs.append("'elements' array is empty or not a list")
    return errs


def bounding_box(elements):
    xs, ys = [], []
    for el in elements:
        if el.get("isDeleted"): continue
        x, y, w, h = el.get("x",0), el.get("y",0), el.get("width",0), el.get("height",0)
        if el.get("type") in ("arrow","line") and "points" in el:
            for px, py in el["points"]:
                xs += [x+px]; ys += [y+py]
        else:
            xs += [x, x+abs(w)]; ys += [y, y+abs(h)]
    if not xs: return 0,0,800,600
    return min(xs), min(ys), max(xs), max(ys)


def render(src: Path, out: Path | None = None, scale=2, max_w=1920) -> Path:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("ERROR: playwright not installed. Run: pip install playwright && playwright install chromium", file=sys.stderr)
        sys.exit(1)

    data = json.loads(src.read_text(encoding="utf-8"))
    errs = validate(data)
    if errs:
        for e in errs: print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    els = [e for e in data["elements"] if not e.get("isDeleted")]
    mn_x, mn_y, mx_x, mx_y = bounding_box(els)
    pad = 80
    vp_w = min(int(mx_x - mn_x + pad*2), max_w)
    vp_h = max(int(mx_y - mn_y + pad*2), 600)

    if out is None:
        out = src.with_suffix(".png")

    tmpl = Path(__file__).parent / "render_template.html"
    if not tmpl.exists():
        print(f"ERROR: Template not found: {tmpl}", file=sys.stderr)
        sys.exit(1)

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=True)
        except Exception as e:
            if "Executable doesn't exist" in str(e):
                print("ERROR: Run: playwright install chromium", file=sys.stderr)
                sys.exit(1)
            raise

        page = browser.new_page(viewport={"width": vp_w, "height": vp_h}, device_scale_factor=scale)
        page.goto(tmpl.as_uri())
        page.wait_for_function("window.__moduleReady === true", timeout=30000)

        result = page.evaluate(f"window.renderDiagram({json.dumps(data)})")
        if not result or not result.get("success"):
            msg = result.get("error","unknown") if result else "null"
            print(f"ERROR: Render failed: {msg}", file=sys.stderr)
            browser.close(); sys.exit(1)

        page.wait_for_function("window.__renderComplete === true", timeout=15000)
        svg = page.query_selector("#root svg")
        if svg is None:
            print("ERROR: No SVG element found.", file=sys.stderr)
            browser.close(); sys.exit(1)

        svg.screenshot(path=str(out))
        browser.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=Path)
    ap.add_argument("--output", "-o", type=Path, default=None)
    ap.add_argument("--scale", "-s", type=int, default=2)
    ap.add_argument("--width", "-w", type=int, default=1920)
    args = ap.parse_args()
    if not args.input.exists():
        print(f"ERROR: Not found: {args.input}", file=sys.stderr); sys.exit(1)
    png = render(args.input, args.output, args.scale, args.width)
    print(str(png))

if __name__ == "__main__":
    main()
