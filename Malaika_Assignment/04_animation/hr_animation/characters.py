"""
characters.py — CorporateCharacter and speech bubble helpers for HRM animation.
"""
import textwrap
from manim import (
    VGroup, Circle, RoundedRectangle, Text, Polygon,
    DOWN, UP, LEFT, RIGHT,
    DARK_BLUE, WHITE, BLACK,
    ManimColor,
)


def wrap_text(text: str, width: int = 38) -> str:
    """Word-wrap text to a max character width per line."""
    return "\n".join(textwrap.wrap(text, width=width))


class CorporateCharacter(VGroup):
    """
    A simple corporate silhouette character built from VMobjects.

    Parameters
    ----------
    color : ManimColor
        Fill color for head and body.
    name : str
        Label displayed below the character.
    """

    HEAD_RADIUS = 0.45
    BODY_WIDTH = 0.9
    BODY_HEIGHT = 1.4
    BODY_CORNER_RADIUS = 0.15

    def __init__(self, color: ManimColor, name: str, **kwargs):
        super().__init__(**kwargs)
        self.char_color = color
        self.char_name = name

        # Head
        self.head = Circle(radius=self.HEAD_RADIUS, color=color, fill_color=color, fill_opacity=1)

        # Body (torso)
        self.body = RoundedRectangle(
            width=self.BODY_WIDTH,
            height=self.BODY_HEIGHT,
            corner_radius=self.BODY_CORNER_RADIUS,
            color=color,
            fill_color=color,
            fill_opacity=1,
        )

        # Position head centered above body
        self.body.next_to(self.head, DOWN, buff=0.1)

        # Name label
        self.label = Text(name, font_size=22, color=BLACK)
        self.label.next_to(self.body, DOWN, buff=0.2)

        self.add(self.head, self.body, self.label)

    def highlight(self):
        """Return an animation that scales the character up slightly."""
        from manim import ScaleInPlace, Indicate
        return ScaleInPlace(self, 1.12)

    def reset(self):
        """Return an animation that returns the character to normal scale."""
        from manim import ScaleInPlace
        return ScaleInPlace(self, 1 / 1.12)


def create_speech_bubble(
    text: str,
    side: str = "right",
    font_size: int = 20,
    bubble_color: ManimColor = WHITE,
    text_color: ManimColor = BLACK,
) -> VGroup:
    """
    Build a speech bubble VGroup with a pointer tail.

    Parameters
    ----------
    text : str
        The dialogue text (already word-wrapped is fine; will be re-wrapped
        if single line exceeds 38 chars).
    side : str
        "left" — bubble sits on the left, tail points right (toward right char).
        "right" — bubble sits on the right, tail points left (toward left char).
    font_size : int
        Font size of the bubble text.
    bubble_color : ManimColor
        Background color of the bubble rectangle.
    text_color : ManimColor
        Color of the dialogue text.

    Returns
    -------
    VGroup
        A VGroup containing the bubble rectangle, tail polygon, and text.
    """
    wrapped = wrap_text(text, width=38)

    bubble_text = Text(wrapped, font_size=font_size, color=text_color)
    bubble_text.set_line_spacing(0.4)

    # Pad the rectangle around the text
    pad_h = 0.35
    pad_v = 0.28
    rect_w = bubble_text.width + 2 * pad_h
    rect_h = bubble_text.height + 2 * pad_v

    rect = RoundedRectangle(
        width=rect_w,
        height=rect_h,
        corner_radius=0.18,
        color=BLACK,
        fill_color=bubble_color,
        fill_opacity=1,
        stroke_width=1.5,
    )

    bubble_text.move_to(rect.get_center())

    # Pointer triangle — points outward toward the character
    tail_offset_x = rect_w / 2
    if side == "right":
        # Tail on the left edge, pointing left
        tip = rect.get_left() + LEFT * 0.4
        top = rect.get_left() + UP * 0.18
        bot = rect.get_left() + DOWN * 0.18
    else:
        # Tail on the right edge, pointing right
        tip = rect.get_right() + RIGHT * 0.4
        top = rect.get_right() + UP * 0.18
        bot = rect.get_right() + DOWN * 0.18

    tail = Polygon(
        tip, top, bot,
        color=BLACK,
        fill_color=bubble_color,
        fill_opacity=1,
        stroke_width=1.5,
    )

    bubble = VGroup(rect, tail, bubble_text)
    return bubble
