"""
Creates the Hackman-Oldham Job Characteristics Model diagram
for Malaika's MGMT268 essay.
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.patheffects as pe
from pathlib import Path

OUTPUT = Path(__file__).parent / "hackman_oldham_model.png"

fig, ax = plt.subplots(1, 1, figsize=(16, 9))
ax.set_xlim(0, 16)
ax.set_ylim(0, 9)
ax.axis('off')
fig.patch.set_facecolor('#FAFAFA')

# ── Colours ──────────────────────────────────────────────────
C1 = '#1B4F72'   # dark blue  – core characteristics
C2 = '#117A65'   # dark green – psychological states
C3 = '#784212'   # dark brown – outcomes
ARROW = '#566573'
HEADER_FG = 'white'

def box(ax, x, y, w, h, color, text, fontsize=9.5, text_color='white', radius=0.25):
    b = FancyBboxPatch((x, y), w, h,
                       boxstyle=f"round,pad=0.05,rounding_size={radius}",
                       linewidth=0, facecolor=color, zorder=3)
    ax.add_patch(b)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=fontsize, color=text_color, fontweight='bold',
            wrap=True, zorder=4,
            multialignment='center')

def header(ax, x, y, w, h, color, text):
    b = FancyBboxPatch((x, y), w, h,
                       boxstyle="round,pad=0.05,rounding_size=0.2",
                       linewidth=0, facecolor=color, zorder=3)
    ax.add_patch(b)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=11, color=HEADER_FG, fontweight='bold', zorder=4,
            multialignment='center')

def arrow(ax, x1, y1, x2, y2):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", color=ARROW,
                                lw=2, mutation_scale=18),
                zorder=2)

# ── Column headers ────────────────────────────────────────────
header(ax, 0.3,  7.5, 4.8, 0.7, C1, "CORE JOB\nCHARACTERISTICS")
header(ax, 5.7,  7.5, 4.8, 0.7, C2, "CRITICAL\nPSYCHOLOGICAL STATES")
header(ax, 11.1, 7.5, 4.5, 0.7, C3, "PERSONAL &\nWORK OUTCOMES")

# ── Column 1: Core Job Characteristics ───────────────────────
chars = [
    "Skill\nVariety",
    "Task\nIdentity",
    "Task\nSignificance",
    "Autonomy",
    "Feedback",
]
col1_x, col1_w, col1_h = 0.3, 4.8, 1.1
for i, c in enumerate(chars):
    y = 6.1 - i * 1.25
    box(ax, col1_x, y, col1_w, col1_h, C1, c, fontsize=10)

# ── Column 2: Critical Psychological States ───────────────────
states = [
    "Experienced\nMeaningfulness\nof the Work",
    "Experienced\nResponsibility\nfor Outcomes",
    "Knowledge\nof Results",
]
col2_x, col2_w = 5.7, 4.8
state_heights = [1.6, 1.6, 1.2]
state_ys = [5.6, 3.8, 2.4]
for s, h, y in zip(states, state_heights, state_ys):
    box(ax, col2_x, y, col2_w, h, C2, s, fontsize=9.5)

# ── Column 3: Outcomes ────────────────────────────────────────
outcomes = [
    ("High Internal\nMotivation",    5.9),
    ("High Job\nSatisfaction",       4.7),
    ("High Work\nQuality",           3.5),
    ("Low Absenteeism\n& Turnover",  2.3),
]
outcomes = [(t, float(y)) for t, y in outcomes]
col3_x, col3_w, col3_h = 11.1, 4.5, 0.95
for text, y in outcomes:
    box(ax, col3_x, y, col3_w, col3_h, C3, text, fontsize=9.5)

# ── Bracket arrows col1 → col2 ────────────────────────────────
# Skill Variety, Task Identity, Task Significance → Meaningfulness
for row_y in [5.65, 4.4, 3.15]:
    arrow(ax, col1_x + col1_w, row_y + col1_h/2, col2_x, state_ys[0] + state_heights[0]/2)

# Autonomy → Responsibility
arrow(ax, col1_x + col1_w, 1.9 + col1_h/2, col2_x, state_ys[1] + state_heights[1]/2)

# Feedback → Knowledge of Results
arrow(ax, col1_x + col1_w, 0.65 + col1_h/2, col2_x, state_ys[2] + state_heights[2]/2)

# ── Arrows col2 → col3 ───────────────────────────────────────
for sy, sh in zip(state_ys, state_heights):
    mid_y = sy + sh / 2
    for _otext, oy in outcomes:
        arrow(ax, col2_x + col2_w, mid_y, col3_x, float(oy) + col3_h/2)

# ── Growth Need Strength note ─────────────────────────────────
ax.text(8.1, 1.1,
        "Moderated by: Employee Growth Need Strength",
        ha='center', va='center', fontsize=9,
        style='italic', color='#5D6D7E',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='#EBF5FB', edgecolor='#AED6F1', lw=1.2))

# ── Title ─────────────────────────────────────────────────────
ax.text(8, 8.6,
        "Figure 1: Hackman & Oldham's Job Characteristics Model (1976)",
        ha='center', va='center', fontsize=13, fontweight='bold', color='#1C2833')

plt.tight_layout(pad=0.5)
plt.savefig(str(OUTPUT), dpi=180, bbox_inches='tight', facecolor='#FAFAFA')
plt.close()
print(f"Diagram saved: {OUTPUT}")
