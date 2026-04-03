"""
Generates hr_dialogue.excalidraw -- Manager vs HR Manager dialogue illustration
showing the rational vs behavioural debate in HRM.

Uses the two-element labeling format required by raw .excalidraw files:
  - shape element with boundElements referencing text id
  - text element with containerId referencing shape id
"""
import json
from pathlib import Path

OUT = Path(__file__).parent / "hr_dialogue.excalidraw"

BLUE_BG     = "#dbe4ff"
BLUE_STR    = "#4a9eed"
BLUE_DARK   = "#1971c2"
GREEN_BG    = "#d3f9d8"
GREEN_STR   = "#22c55e"
GREEN_DARK  = "#15803d"
YELLOW_BG   = "#fff3bf"
YELLOW_STR  = "#f59e0b"
GREY        = "#868e96"


def el_base(type_, id_, x, y, w, h, **kw):
    base = dict(
        type=type_, id=id_, x=x, y=y, width=w, height=h,
        angle=0, strokeColor="#1e1e1e", backgroundColor="transparent",
        fillStyle="solid", strokeWidth=2, strokeStyle="solid",
        roughness=1, opacity=100, isDeleted=False,
        groupIds=[], frameId=None, boundElements=[]
    )
    base.update(kw)
    return base


def rect(id_, x, y, w, h, bg="transparent", stroke="#1e1e1e", sw=2,
         rounded=True, bound_text_ids=None, **kw):
    el = el_base("rectangle", id_, x, y, w, h,
                 backgroundColor=bg, strokeColor=stroke, strokeWidth=sw, **kw)
    if rounded:
        el["roundness"] = {"type": 3}
    if bound_text_ids:
        el["boundElements"] = [{"type": "text", "id": t} for t in bound_text_ids]
    return el


def ellipse(id_, x, y, w, h, bg="transparent", stroke="#1e1e1e", sw=2, **kw):
    return el_base("ellipse", id_, x, y, w, h,
                   backgroundColor=bg, strokeColor=stroke, strokeWidth=sw,
                   roughness=0, **kw)


def text_in(id_, container_id, text, x, y, w, h,
            size=16, color="#1e1e1e", align="center"):
    return dict(
        type="text", id=id_, x=x, y=y, width=w, height=h,
        angle=0, strokeColor=color, backgroundColor="transparent",
        fillStyle="solid", strokeWidth=1, strokeStyle="solid",
        roughness=0, opacity=100, isDeleted=False,
        groupIds=[], frameId=None, boundElements=[],
        containerId=container_id,
        text=text, fontSize=size, fontFamily=1,
        textAlign=align, verticalAlignment="middle",
        baseline=int(size * 0.8),
        lineHeight=1.25
    )


def standalone_text(id_, text, x, y, size=18, color="#1e1e1e", align="left"):
    est_w = len(text.split('\n')[0]) * size * 0.55
    return dict(
        type="text", id=id_, x=x, y=y, width=max(int(est_w), 200), height=int(size*1.5),
        angle=0, strokeColor=color, backgroundColor="transparent",
        fillStyle="solid", strokeWidth=1, strokeStyle="solid",
        roughness=0, opacity=100, isDeleted=False,
        groupIds=[], frameId=None, boundElements=[],
        containerId=None,
        text=text, fontSize=size, fontFamily=1,
        textAlign=align, verticalAlignment="top",
        baseline=int(size*0.8), lineHeight=1.25
    )


def arrow(id_, x1, y1, x2, y2, color="#555555", sw=2, dashed=False):
    dx, dy = x2-x1, y2-y1
    el = el_base("arrow", id_, x1, y1, abs(dx) or 1, abs(dy) or 1,
                 strokeColor=color, strokeWidth=sw, roughness=0,
                 endArrowhead="arrow", startArrowhead=None)
    el["points"] = [[0, 0], [dx, dy]]
    if dashed:
        el["strokeStyle"] = "dashed"
    return el


# ─── Build elements ────────────────────────────────────────────────────────────

elements = []

# Title
elements.append(rect("title_box", 200, 15, 1000, 55,
                      bg="#2C3E50", stroke="#2C3E50", rounded=True,
                      bound_text_ids=["title_txt"]))
elements.append(text_in("title_txt", "title_box",
                         "The Case for the Behavioural Approach\n"
                         "A Dialogue: Manager vs HR Manager  |  MGMT268 Topic 4",
                         200, 15, 1000, 55, size=16, color="#ffffff"))

# Divider
elements.append(arrow("divider", 50, 82, 1350, 82, color="#cccccc", sw=1))

# ── Manager figure (LEFT, x-centre ~150) ───────────────────────────────────
# Head
elements.append(ellipse("m_head", 120, 95, 60, 60,
                         bg=BLUE_BG, stroke=BLUE_STR, sw=2))
# Body
elements.append(rect("m_body", 107, 157, 86, 65,
                      bg=BLUE_BG, stroke=BLUE_STR, sw=2, rounded=True))
# Name tag
elements.append(rect("m_name", 75, 230, 150, 35,
                      bg=BLUE_STR, stroke=BLUE_DARK, sw=2, rounded=True,
                      bound_text_ids=["m_name_txt"]))
elements.append(text_in("m_name_txt", "m_name",
                         "Manager", 75, 230, 150, 35,
                         size=16, color="#ffffff"))

# ── HR Manager figure (RIGHT, x-centre ~1250) ───────────────────────────────
elements.append(ellipse("hr_head", 1220, 95, 60, 60,
                         bg=GREEN_BG, stroke=GREEN_STR, sw=2))
elements.append(rect("hr_body", 1207, 157, 86, 65,
                      bg=GREEN_BG, stroke=GREEN_STR, sw=2, rounded=True))
elements.append(rect("hr_name", 1175, 230, 150, 35,
                      bg=GREEN_STR, stroke=GREEN_DARK, sw=2, rounded=True,
                      bound_text_ids=["hr_name_txt"]))
elements.append(text_in("hr_name_txt", "hr_name",
                         "HR Manager", 1175, 230, 150, 35,
                         size=16, color="#ffffff"))

# ── Exchange 1: Manager speaks ───────────────────────────────────────────────
L1_X, L1_Y, L1_W, L1_H = 250, 100, 520, 75
elements.append(rect("l1_box", L1_X, L1_Y, L1_W, L1_H,
                      bg=BLUE_BG, stroke=BLUE_STR, sw=2, rounded=True,
                      bound_text_ids=["l1_txt"]))
elements.append(text_in("l1_txt", "l1_box",
                         "Taylor's 'one best way' gives us maximum efficiency.\n"
                         "Isn't the rational approach to job design sufficient?",
                         L1_X, L1_Y, L1_W, L1_H, size=15))
# Tail arrow → Manager
elements.append(arrow("al1", L1_X, L1_Y + L1_H//2, L1_X - 8, L1_Y + L1_H//2,
                       color=BLUE_STR, sw=2))

# ── Exchange 2: HR Manager responds ─────────────────────────────────────────
R1_X, R1_Y, R1_W, R1_H = 420, 195, 730, 90
elements.append(rect("r1_box", R1_X, R1_Y, R1_W, R1_H,
                      bg=GREEN_BG, stroke=GREEN_STR, sw=2, rounded=True,
                      bound_text_ids=["r1_txt"]))
elements.append(text_in("r1_txt", "r1_box",
                         "Workers were literally called 'hands' — not people.\n"
                         "Scientific Management was obsolete by the 1920s due to low morale.\n"
                         "The Hawthorne Studies (Mayo, 1927) proved the human side matters.",
                         R1_X, R1_Y, R1_W, R1_H, size=15))
# Tail arrow → HR Manager
elements.append(arrow("ar1", R1_X + R1_W, R1_Y + R1_H//2,
                       R1_X + R1_W + 8, R1_Y + R1_H//2,
                       color=GREEN_STR, sw=2))

# ── Exchange 3: Manager ──────────────────────────────────────────────────────
L2_X, L2_Y, L2_W, L2_H = 250, 305, 480, 70
elements.append(rect("l2_box", L2_X, L2_Y, L2_W, L2_H,
                      bg=BLUE_BG, stroke=BLUE_STR, sw=2, rounded=True,
                      bound_text_ids=["l2_txt"]))
elements.append(text_in("l2_txt", "l2_box",
                         "But General Motors uses assembly lines and it works...\n"
                         "Efficiency is still the goal, isn't it?",
                         L2_X, L2_Y, L2_W, L2_H, size=15))
elements.append(arrow("al2", L2_X, L2_Y + L2_H//2, L2_X - 8, L2_Y + L2_H//2,
                       color=BLUE_STR, sw=2))

# ── Exchange 4: HR Manager ───────────────────────────────────────────────────
R2_X, R2_Y, R2_W, R2_H = 350, 395, 800, 90
elements.append(rect("r2_box", R2_X, R2_Y, R2_W, R2_H,
                      bg=GREEN_BG, stroke=GREEN_STR, sw=2, rounded=True,
                      bound_text_ids=["r2_txt"]))
elements.append(text_in("r2_txt", "r2_box",
                         "Volvo competes on quality, not cost — same product, different approach.\n"
                         "Getting technical conditions right is not enough.\n"
                         "You also have to get the human conditions right.",
                         R2_X, R2_Y, R2_W, R2_H, size=15))
elements.append(arrow("ar2", R2_X + R2_W, R2_Y + R2_H//2,
                       R2_X + R2_W + 8, R2_Y + R2_H//2,
                       color=GREEN_STR, sw=2))

# ── Exchange 5: Manager asks the key question ────────────────────────────────
L3_X, L3_Y, L3_W, L3_H = 250, 505, 390, 55
elements.append(rect("l3_box", L3_X, L3_Y, L3_W, L3_H,
                      bg=BLUE_BG, stroke=BLUE_STR, sw=2, rounded=True,
                      bound_text_ids=["l3_txt"]))
elements.append(text_in("l3_txt", "l3_box",
                         "So what is the actual solution?",
                         L3_X, L3_Y, L3_W, L3_H, size=15))
elements.append(arrow("al3", L3_X, L3_Y + L3_H//2, L3_X - 8, L3_Y + L3_H//2,
                       color=BLUE_STR, sw=2))

# ── Exchange 6: HR Manager — key conclusion (highlighted) ───────────────────
R3_X, R3_Y, R3_W, R3_H = 300, 580, 850, 115
elements.append(rect("r3_box", R3_X, R3_Y, R3_W, R3_H,
                      bg=GREEN_BG, stroke=GREEN_DARK, sw=3, rounded=True,
                      bound_text_ids=["r3_txt"]))
elements.append(text_in("r3_txt", "r3_box",
                         "Hackman & Oldham's Job Characteristics Model (1976):\n"
                         "Skill Variety  |  Task Identity  |  Task Significance\n"
                         "Autonomy  |  Feedback\n"
                         "These five dimensions drive motivation, quality, and retention.",
                         R3_X, R3_Y, R3_W, R3_H, size=15))
elements.append(arrow("ar3", R3_X + R3_W, R3_Y + R3_H//2,
                       R3_X + R3_W + 8, R3_Y + R3_H//2,
                       color=GREEN_DARK, sw=2))

# ── Footer takeaway ──────────────────────────────────────────────────────────
F_X, F_Y, F_W, F_H = 150, 720, 1100, 60
elements.append(rect("footer_box", F_X, F_Y, F_W, F_H,
                      bg=YELLOW_BG, stroke=YELLOW_STR, sw=2, rounded=True,
                      bound_text_ids=["footer_txt"]))
elements.append(text_in("footer_txt", "footer_box",
                         "'You cannot manage Human Resources without managing "
                         "the human beings they belong to.'  \u2014 MGMT268 Topic 4",
                         F_X, F_Y, F_W, F_H, size=15, color="#5c4000"))

# Figure caption
elements.append(standalone_text("fig_cap",
    "Figure 4: Illustration — The Rational vs Behavioural Debate in HRM",
    350, 795, size=13, color=GREY))

# ─── Write file ────────────────────────────────────────────────────────────────
doc = {
    "type": "excalidraw",
    "version": 2,
    "source": "https://excalidraw.com",
    "elements": elements,
    "appState": {
        "gridSize": None,
        "viewBackgroundColor": "#ffffff"
    },
    "files": {}
}

OUT.write_text(json.dumps(doc, indent=2), encoding="utf-8")
print(f"Saved: {OUT}  ({len(elements)} elements)")
