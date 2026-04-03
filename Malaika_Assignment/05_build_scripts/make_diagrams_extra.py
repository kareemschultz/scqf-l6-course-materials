"""
Generates three supplementary visuals:
  1. comparison_table.png   -- Scientific Management vs Behavioural Approach
  2. gns_diagram.png        -- Growth Need Strength as JCM moderator
  3. animation_frame.png    -- Extracted frame from HRMDialogueScene.mp4
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.patheffects as pe
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
from pathlib import Path
import numpy as np

BASE  = Path(__file__).parent
VIDEO = BASE / "hr_animation/media/videos/main/1080p60/HRMDialogueScene.mp4"


# ============================================================
# 1. Comparison table: Scientific Management vs Behavioural
# ============================================================
def make_comparison():
    rows = [
        ("View of Worker",    "A replaceable unit / 'hand'",       "A whole person with psychological needs"),
        ("Motivation",        "Pay & financial incentives only",    "Meaningful work, autonomy, feedback"),
        ("Control",           "Top-down; one prescribed 'best way'","Worker has autonomy over how job is done"),
        ("Primary Goal",      "Maximum efficiency & output",        "Motivation + quality + low turnover"),
        ("Key Weakness",      "Ignores human psychology",           "Needs careful implementation per role"),
        ("Historical Origin", "Taylor (1909)",                      "Hackman & Oldham (1976); Hawthorne (1927)"),
    ]

    fig, ax = plt.subplots(figsize=(13, 5.5))
    ax.set_xlim(0, 13)
    ax.set_ylim(0, len(rows) + 1.4)
    ax.axis('off')
    fig.patch.set_facecolor('#FAFAFA')

    # Column x positions
    cx = [0.15, 3.5, 8.3]
    col_w = [3.2, 4.6, 4.6]

    # Header row
    header_labels = ["Dimension", "Scientific Management\n(Rational Approach)", "Behavioural Approach"]
    header_colors = ['#2C3E50', '#C0392B', '#1A5276']
    for i, (lbl, col, cw) in enumerate(zip(header_labels, cx, col_w)):
        rect = FancyBboxPatch((cx[i] - 0.1, len(rows) + 0.3), cw, 0.95,
                              boxstyle="round,pad=0.05", linewidth=0,
                              facecolor=header_colors[i])
        ax.add_patch(rect)
        ax.text(cx[i] + cw / 2 - 0.1, len(rows) + 0.78, lbl,
                ha='center', va='center', fontsize=10, fontweight='bold',
                color='white', wrap=True,
                multialignment='center')

    # Data rows
    row_bg = ['#FFFFFF', '#F2F3F4']
    for r, (dim, sm, beh) in enumerate(reversed(rows)):
        y = r
        bg = row_bg[r % 2]
        for i, (txt, cw) in enumerate(zip([dim, sm, beh], col_w)):
            rect = FancyBboxPatch((cx[i] - 0.1, y + 0.05), cw, 0.85,
                                  boxstyle="round,pad=0.03", linewidth=0.5,
                                  edgecolor='#CCCCCC', facecolor=bg)
            ax.add_patch(rect)
            color = '#2C3E50' if i == 0 else ('#922B21' if i == 1 else '#1A5276')
            fw = 'bold' if i == 0 else 'normal'
            ax.text(cx[i] + cw / 2 - 0.1, y + 0.48, txt,
                    ha='center', va='center', fontsize=9,
                    color=color, fontweight=fw,
                    multialignment='center', wrap=True)

    ax.set_title("Figure 2: Scientific Management vs the Behavioural Approach -- A Comparative Overview",
                 fontsize=11, fontstyle='italic', pad=10, color='#2C3E50')

    out = BASE / "comparison_table.png"
    plt.tight_layout()
    plt.savefig(str(out), dpi=150, bbox_inches='tight', facecolor='#FAFAFA')
    plt.close()
    print(f"Saved: {out}")


# ============================================================
# 2. Growth Need Strength -- JCM moderator diagram
# ============================================================
def make_gns():
    fig, ax = plt.subplots(figsize=(13, 5))
    ax.set_xlim(0, 13)
    ax.set_ylim(0, 5)
    ax.axis('off')
    fig.patch.set_facecolor('#FAFAFA')

    def box(x, y, w, h, label, sublabel, fc, ec, fs=9.5):
        rect = FancyBboxPatch((x, y), w, h,
                              boxstyle="round,pad=0.15", linewidth=1.5,
                              edgecolor=ec, facecolor=fc, zorder=3)
        ax.add_patch(rect)
        ax.text(x + w / 2, y + h / 2 + (0.18 if sublabel else 0),
                label, ha='center', va='center',
                fontsize=fs, fontweight='bold', color='white',
                multialignment='center', zorder=4)
        if sublabel:
            ax.text(x + w / 2, y + h / 2 - 0.28,
                    sublabel, ha='center', va='center',
                    fontsize=7.5, color='#DDDDDD',
                    multialignment='center', zorder=4)

    def arrow(x1, x2, y, color='#555555'):
        ax.annotate('', xy=(x2, y), xytext=(x1, y),
                    arrowprops=dict(arrowstyle='->', color=color,
                                   lw=2.0, mutation_scale=18))

    # Column 1: Core Job Characteristics
    c1_items = ["Skill Variety", "Task Identity", "Task Significance", "Autonomy", "Feedback"]
    box(0.2, 0.4, 2.8, 4.2, "Core Job\nCharacteristics", None, '#1A5276', '#1A5276', fs=10)
    for i, item in enumerate(c1_items):
        rect = FancyBboxPatch((0.35, 0.55 + i * 0.72), 2.5, 0.6,
                              boxstyle="round,pad=0.05", linewidth=1,
                              edgecolor='#AED6F1', facecolor='#2E86C1', zorder=4)
        ax.add_patch(rect)
        ax.text(1.6, 0.85 + i * 0.72, item,
                ha='center', va='center', fontsize=8.5,
                color='white', zorder=5)

    arrow(3.0, 3.8, 2.5)

    # Column 2: Critical Psychological States
    c2_items = ["Experienced\nMeaningfulness", "Experienced\nResponsibility", "Knowledge\nof Results"]
    box(3.8, 0.4, 3.0, 4.2, "Critical Psychological\nStates", None, '#1E8449', '#1E8449', fs=10)
    for i, item in enumerate(c2_items):
        rect = FancyBboxPatch((3.95, 0.6 + i * 1.1), 2.7, 0.95,
                              boxstyle="round,pad=0.05", linewidth=1,
                              edgecolor='#A9DFBF', facecolor='#27AE60', zorder=4)
        ax.add_patch(rect)
        ax.text(5.3, 1.08 + i * 1.1, item,
                ha='center', va='center', fontsize=8.5,
                color='white', multialignment='center', zorder=5)

    arrow(6.8, 7.6, 2.5)

    # Column 3: Outcomes
    c3_items = ["High internal\nmotivation", "High satisfaction", "High quality work", "Low absenteeism\n& turnover"]
    box(7.6, 0.4, 3.0, 4.2, "Personal & Work\nOutcomes", None, '#7D6608', '#7D6608', fs=10)
    for i, item in enumerate(c3_items):
        rect = FancyBboxPatch((7.75, 0.55 + i * 0.85), 2.7, 0.72,
                              boxstyle="round,pad=0.05", linewidth=1,
                              edgecolor='#F9E79F', facecolor='#D4AC0D', zorder=4)
        ax.add_patch(rect)
        ax.text(9.1, 0.91 + i * 0.85, item,
                ha='center', va='center', fontsize=8.5,
                color='white', multialignment='center', zorder=5)

    # GNS moderator arrow (curved, from bottom)
    ax.annotate('', xy=(5.3, 0.55), xytext=(1.6, 0.4),
                arrowprops=dict(
                    arrowstyle='->', color='#884EA0', lw=2.2,
                    connectionstyle='arc3,rad=-0.35', mutation_scale=16))
    ax.text(3.45, 0.12, "Moderated by: Growth Need Strength (GNS)",
            ha='center', va='center', fontsize=9, color='#884EA0',
            fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.3', fc='#F5EEF8', ec='#884EA0', lw=1.5))

    ax.set_title(
        "Figure 3: Hackman & Oldham's Job Characteristics Model -- "
        "including Growth Need Strength as Moderating Variable (1976)",
        fontsize=10, fontstyle='italic', pad=10, color='#2C3E50')

    out = BASE / "gns_diagram.png"
    plt.tight_layout()
    plt.savefig(str(out), dpi=150, bbox_inches='tight', facecolor='#FAFAFA')
    plt.close()
    print(f"Saved: {out}")


# ============================================================
# 3. Extract a frame from the animation MP4
# ============================================================
def extract_frame():
    out = BASE / "animation_frame.png"
    if not VIDEO.exists():
        print(f"  Video not found: {VIDEO}")
        return

    try:
        import imageio.v3 as iio
        # Read one frame from roughly the middle of the video
        # Use index=-1 to get metadata first
        meta = iio.immeta(str(VIDEO), plugin="pyav")
        fps      = float(meta.get("fps", 24))
        duration = float(meta.get("duration", 60))
        target_s = duration * 0.45          # ~45% in -- T-chart scene
        frame_idx = int(fps * target_s)

        frame = iio.imread(str(VIDEO), index=frame_idx, plugin="pyav")
        iio.imwrite(str(out), frame)
        print(f"Saved: {out}  (frame {frame_idx} @ ~{target_s:.1f}s)")
    except Exception as e:
        print(f"  imageio extraction failed ({e}), generating placeholder frame...")
        _make_placeholder_frame(out)


def _make_placeholder_frame(out):
    """Fallback: draw a labelled placeholder instead."""
    fig, ax = plt.subplots(figsize=(10, 5.6), facecolor='#1a1a2e')
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5.6)
    ax.axis('off')

    ax.text(5, 4.8, "HR Management: Rational vs Behavioural Approach",
            ha='center', va='center', fontsize=14, fontweight='bold',
            color='white')
    ax.text(5, 4.2, "Animated Dialogue -- Manager & HR Manager",
            ha='center', va='center', fontsize=10, color='#AAAAFF')

    # T-chart
    ax.plot([5, 5], [0.3, 3.8], color='white', lw=2)
    ax.plot([0.5, 9.5], [3.8, 3.8], color='white', lw=2)
    ax.text(2.5, 4.1, "Scientific Management", ha='center', fontsize=10,
            color='#FF6B6B', fontweight='bold')
    ax.text(7.5, 4.1, "Behavioural Approach", ha='center', fontsize=10,
            color='#6BFF6B', fontweight='bold')

    left_items  = ["One best way (Taylor, 1909)", "Workers as 'hands'", "Top-down control"]
    right_items = ["Skill variety & autonomy", "Hawthorne Studies (Mayo, 1927)", "Job enrichment"]
    for i, (l, r) in enumerate(zip(left_items, right_items)):
        y = 3.3 - i * 0.85
        ax.text(2.5, y, l, ha='center', fontsize=9, color='#FFCCCC')
        ax.text(7.5, y, r, ha='center', fontsize=9, color='#CCFFCC')

    ax.text(5, 0.1, "[Storyboard frame -- companion animation: HRMDialogueScene.mp4]",
            ha='center', va='center', fontsize=8, color='#888888', fontstyle='italic')

    plt.tight_layout(pad=0.2)
    plt.savefig(str(out), dpi=130, bbox_inches='tight', facecolor='#1a1a2e')
    plt.close()
    print(f"Saved placeholder frame: {out}")


# ============================================================
if __name__ == '__main__':
    make_comparison()
    make_gns()
    extract_frame()
    print("\nAll extra diagrams done.")
