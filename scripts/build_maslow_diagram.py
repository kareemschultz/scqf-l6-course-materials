"""Create blank Maslow pyramid diagram using matplotlib."""
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os

fig, ax = plt.subplots(1, 1, figsize=(8, 6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 7)
ax.set_aspect('equal')
ax.axis('off')
fig.patch.set_facecolor('white')

# Pyramid layers (bottom to top)
layers = [
    (0.5, 0.2, 9.0, 1.1, '#D5E8D4', '1. Physiological Needs', '[Workplace example...]'),
    (1.4, 1.5, 7.2, 1.1, '#DAE8FC', '2. Safety Needs', '[Workplace example...]'),
    (2.3, 2.8, 5.4, 1.1, '#FFF2CC', '3. Social / Belonging', '[Workplace example...]'),
    (3.2, 4.1, 3.6, 1.1, '#F8CECC', '4. Esteem Needs', '[Workplace example...]'),
    (4.1, 5.4, 1.8, 0.9, '#E1D5E7', '5. Self-Actualisation', ''),
]

for x, y, w, h, color, label, placeholder in layers:
    rect = patches.FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.05",
        facecolor=color, edgecolor='#333333', linewidth=1.5
    )
    ax.add_patch(rect)
    ax.text(x + w/2, y + h/2 + 0.1, label,
            ha='center', va='center', fontsize=10, fontweight='bold', color='#333333')
    if placeholder:
        ax.text(x + w/2, y + h/2 - 0.25, placeholder,
                ha='center', va='center', fontsize=8, fontstyle='italic', color='#888888')

ax.text(5.0, 6.7, "Maslow's Hierarchy of Needs - BLANK TEMPLATE",
        ha='center', va='center', fontsize=13, fontweight='bold', color='#2F5496')
ax.text(5.0, -0.2, '[Replace placeholder text with your own workplace examples]',
        ha='center', va='center', fontsize=9, fontstyle='italic', color='#888888')

out = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    'SCQF_L6_SUPPORT_PACK', 'J22A76_Management_People_Finance', 'Maslow_Blank_Diagram.png'
)
fig.savefig(out, dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print(f'Created {out}')
