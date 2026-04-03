"""
main.py — HRM Dialogue Animation
================================
Animated educational video illustrating the difference between the
Rational Approach and the Behavioral Approach to Human Resource Management.

Usage:
    manim -pqh main.py HRMDialogueScene   # High quality, 1080p60
    manim -pql main.py HRMDialogueScene   # Low quality, quick preview
"""

import textwrap

from manim import (
    # Scene
    Scene,
    # Mobjects
    VGroup, Text, Rectangle, RoundedRectangle, Circle,
    Line, Polygon, Arrow, DashedLine,
    # Colors
    WHITE, BLACK, DARK_BLUE, TEAL, GREY, GREY_A, GREY_B, GREY_C,
    YELLOW, GREEN, RED, BLUE, ORANGE,
    ManimColor,
    # Animations
    Write, FadeIn, FadeOut, Create, GrowFromCenter,
    AnimationGroup, LaggedStart, LaggedStartMap,
    ScaleInPlace,
    # Directions & constants
    UP, DOWN, LEFT, RIGHT, ORIGIN,
    # Utilities
    config,
    # Value tracker
    ValueTracker,
)
from manim_voiceover import VoiceoverScene
from manim_voiceover.services.gtts import GTTSService

from characters import CorporateCharacter, create_speech_bubble

# ─── Layout constants ────────────────────────────────────────────────────────
BG_COLOR = "#F0F0F0"
TITLE_BAR_COLOR = "#1A2B4A"  # Dark navy
TITLE_TEXT_COLOR = WHITE

MANAGER_COLOR = DARK_BLUE
HR_COLOR = TEAL

MANAGER_POS = LEFT * 4.2
HR_POS = RIGHT * 4.2

TITLE_BAR_Y = -3.5   # Bottom of screen

# Speech bubble vertical centre position (above characters)
BUBBLE_Y = 1.5

# ─── TTS service factory ──────────────────────────────────────────────────────

def _uk_service():
    """British English accent — used for the Manager."""
    return GTTSService(lang="en", tld="co.uk")


def _us_service():
    """US English accent — used for the HR Manager."""
    return GTTSService(lang="en", tld="com")


# ─── Scene ────────────────────────────────────────────────────────────────────

class HRMDialogueScene(VoiceoverScene):
    """
    Full HRM dialogue animation.

    Scene structure:
      1. Intro  — title + subtitle + characters walk in
      2. Lines 1-6 — opening dialogue (rational approach discussed)
      3. Midpoint T-chart — Rational vs Behavioral comparison card
      4. Lines 7-11 — conclusion dialogue (behavioral need established)
      5. Outro  — conclusion statement + "End of Lesson"
    """

    # ── Lifecycle ──────────────────────────────────────────────────────────────

    def construct(self):
        self.set_speech_service(_us_service())   # Default; switched per line

        # Set background color
        self.camera.background_color = BG_COLOR

        # Build persistent scene elements
        self.manager, self.hr_manager = self._build_characters()
        self.title_bar, self.title_bar_text = self._build_title_bar()

        # Run sections
        self._intro()
        self._dialogue_section_one()
        self._midpoint_chart()
        self._dialogue_section_two()
        self._outro()

    # ── Scene builders ─────────────────────────────────────────────────────────

    def _build_characters(self):
        manager = CorporateCharacter(
            color=MANAGER_COLOR,
            name="Manager",
        )
        manager.move_to(MANAGER_POS + DOWN * 0.5)

        hr_manager = CorporateCharacter(
            color=HR_COLOR,
            name="HR Manager",
        )
        hr_manager.move_to(HR_POS + DOWN * 0.5)

        return manager, hr_manager

    def _build_title_bar(self):
        bar = Rectangle(
            width=config.frame_width,
            height=0.72,
            fill_color=TITLE_BAR_COLOR,
            fill_opacity=1,
            stroke_width=0,
        )
        bar.move_to([0, TITLE_BAR_Y, 0])

        label = Text(
            "Human Resource Management: Rational vs Behavioral Approach",
            font_size=18,
            color=TITLE_TEXT_COLOR,
        )
        label.move_to(bar.get_center())

        return bar, label

    # ── Intro ──────────────────────────────────────────────────────────────────

    def _intro(self):
        # Large centred title
        title = Text(
            "Rational vs. Behavioral HRM",
            font_size=48,
            color=DARK_BLUE,
            weight="BOLD",
        )
        title.move_to(UP * 1.5)

        subtitle = Text(
            "Why the Behavioral Approach Matters",
            font_size=28,
            color="#444444",
        )
        subtitle.next_to(title, DOWN, buff=0.4)

        # Animate title in
        self.play(Write(title), run_time=1.8)
        self.play(FadeIn(subtitle, shift=UP * 0.2), run_time=1.0)
        self.wait(1.0)

        # Fade title/subtitle out
        self.play(FadeOut(title), FadeOut(subtitle), run_time=0.8)

        # Slide characters in from off-screen
        self.manager.move_to(LEFT * 10 + DOWN * 0.5)   # Start off left edge
        self.hr_manager.move_to(RIGHT * 10 + DOWN * 0.5)  # Start off right edge

        # Title bar slides up from bottom
        self.title_bar.move_to([0, -5, 0])
        self.title_bar_text.move_to([0, -5, 0])

        self.play(
            self.manager.animate.move_to(MANAGER_POS + DOWN * 0.5),
            self.hr_manager.animate.move_to(HR_POS + DOWN * 0.5),
            self.title_bar.animate.move_to([0, TITLE_BAR_Y, 0]),
            self.title_bar_text.animate.move_to([0, TITLE_BAR_Y, 0]),
            run_time=1.5,
        )

        self.wait(2.0)

    # ── speak_line helper ──────────────────────────────────────────────────────

    def speak_line(
        self,
        character: CorporateCharacter,
        voiceover_text: str,
        bubble_side: str,
        bubble_anchor_x: float,
    ):
        """
        Animate a single dialogue line:
          1. Highlight the speaking character.
          2. Show a speech bubble with word-wrapped text.
          3. Play TTS audio, running a gentle idle animation during playback.
          4. Fade bubble out, reset character scale.

        Parameters
        ----------
        character : CorporateCharacter
            The character who is speaking.
        voiceover_text : str
            The text spoken aloud (and shown in the bubble).
        bubble_side : str
            "right" — bubble on the right side of character (tail points left).
            "left"  — bubble on the left side of character (tail points right).
        bubble_anchor_x : float
            X-coordinate to centre the bubble at.
        """
        # 1. Highlight character
        self.play(character.highlight(), run_time=0.3)

        # 2. Build and position speech bubble
        bubble = create_speech_bubble(
            text=voiceover_text,
            side=bubble_side,
            font_size=19,
            bubble_color=WHITE,
            text_color=BLACK,
        )
        # Position: vertically above the character, horizontally offset
        bubble.move_to([bubble_anchor_x, BUBBLE_Y + 0.8, 0])

        # Keep bubble inside frame (clamp x)
        half_w = bubble.width / 2
        max_x = config.frame_width / 2 - half_w - 0.15
        new_x = max(-max_x, min(max_x, bubble.get_center()[0]))
        bubble.move_to([new_x, bubble.get_center()[1], 0])

        self.play(FadeIn(bubble, scale=0.92), run_time=0.4)

        # 3. TTS audio — gentle bob animation during playback
        with self.voiceover(text=voiceover_text) as tracker:
            # Bob character up/down during speech
            self.play(
                character.animate.shift(UP * 0.06),
                run_time=min(0.4, tracker.duration / 2),
            )
            remaining = tracker.duration - 0.4
            if remaining > 0.1:
                self.play(
                    character.animate.shift(DOWN * 0.06),
                    run_time=min(0.4, remaining),
                )
                still = remaining - 0.4
                if still > 0.0:
                    self.wait(still)

        # 4. Fade out bubble, reset character
        self.play(FadeOut(bubble), character.reset(), run_time=0.45)
        self.wait(0.5)

    # ── Dialogue section 1 (lines 1–6) ────────────────────────────────────────

    def _dialogue_section_one(self):
        # Add persistent elements to scene if not yet added
        self.add(self.manager, self.hr_manager, self.title_bar, self.title_bar_text)

        # ── Line 1: Manager
        self.set_speech_service(_uk_service())
        self.speak_line(
            character=self.manager,
            voiceover_text=(
                "I don't understand what's going wrong. Productivity has dropped, "
                "employees are disengaged, and some of our best people have started leaving."
            ),
            bubble_side="right",
            bubble_anchor_x=0.0,
        )

        # ── Line 2: HR Manager
        self.set_speech_service(_us_service())
        self.speak_line(
            character=self.hr_manager,
            voiceover_text=(
                "I've been looking into this. The issue may be rooted in how we have "
                "designed jobs here. We may be relying too heavily on the rational approach."
            ),
            bubble_side="left",
            bubble_anchor_x=0.0,
        )

        # ── Line 3: Manager
        self.set_speech_service(_uk_service())
        self.speak_line(
            character=self.manager,
            voiceover_text=(
                "But our tasks are clearly defined and optimised for efficiency. "
                "That is exactly what Scientific Management teaches — one best way to do every job."
            ),
            bubble_side="right",
            bubble_anchor_x=0.0,
        )

        # ── Line 4: HR Manager
        self.set_speech_service(_us_service())
        self.speak_line(
            character=self.hr_manager,
            voiceover_text=(
                "Taylor's model was groundbreaking — but it treated workers like machine parts. "
                "Workers were literally called hands — farm hands, factory hands — not people."
            ),
            bubble_side="left",
            bubble_anchor_x=0.0,
        )

        # ── Line 5: Manager
        self.set_speech_service(_uk_service())
        self.speak_line(
            character=self.manager,
            voiceover_text=(
                "So are you saying the rational approach is not sufficient for managing "
                "human resources in this organization?"
            ),
            bubble_side="right",
            bubble_anchor_x=0.0,
        )

        # ── Line 6: HR Manager
        self.set_speech_service(_us_service())
        self.speak_line(
            character=self.hr_manager,
            voiceover_text=(
                "Exactly. Getting the technical conditions right is not enough — "
                "you also have to get the human conditions right. "
                "That is what the behavioural approach addresses."
            ),
            bubble_side="left",
            bubble_anchor_x=0.0,
        )

    # ── Midpoint T-chart ───────────────────────────────────────────────────────

    def _midpoint_chart(self):
        """Animated comparison card between the two HRM approaches."""

        # Temporarily move characters off-screen to free space for the chart
        self.play(
            self.manager.animate.shift(LEFT * 1.5),
            self.hr_manager.animate.shift(RIGHT * 1.5),
            run_time=0.6,
        )

        # ── Card background
        card = RoundedRectangle(
            width=9.5,
            height=5.2,
            corner_radius=0.25,
            fill_color=WHITE,
            fill_opacity=1,
            stroke_color="#CCCCCC",
            stroke_width=2,
        )
        card.move_to(UP * 0.4)

        # ── Card title
        card_title = Text(
            "Comparing HRM Approaches",
            font_size=26,
            color=DARK_BLUE,
            weight="BOLD",
        )
        card_title.move_to(card.get_top() + DOWN * 0.45)

        # ── Divider line
        divider = Line(
            start=[0, card.get_top()[1] - 0.85, 0],
            end=[0, card.get_bottom()[1] + 0.2, 0],
            color="#BBBBBB",
            stroke_width=1.5,
        )

        # ── Column headers
        rational_header = Text("Rational Approach", font_size=22, color=DARK_BLUE, weight="BOLD")
        rational_header.move_to(LEFT * 2.35 + UP * 0.55)

        behavioral_header = Text("Behavioral Approach", font_size=22, color=TEAL, weight="BOLD")
        behavioral_header.move_to(RIGHT * 2.35 + UP * 0.55)

        # ── Bullet items
        rational_bullets = [
            "• One best way (Taylor, 1909)",
            "• Task simplification",
            "• Manager decides; worker executes",
            "• Workers as 'hands', not people",
        ]
        behavioral_bullets = [
            "• Skill variety & autonomy",
            "• Task identity & significance",
            "• Feedback & responsibility",
            "• Hawthorne Studies (Mayo, 1927)",
        ]

        def make_bullet_group(items, x_offset, color):
            group = VGroup()
            for i, item in enumerate(items):
                t = Text(item, font_size=19, color=color)
                t.move_to([x_offset, 0.1 - i * 0.6, 0])
                t.align_to(Line(LEFT * 0.1, RIGHT * 0.1).move_to([x_offset, 0, 0]), LEFT)
                group.add(t)
            return group

        r_bullets = make_bullet_group(rational_bullets, x_offset=-3.8, color="#1A2B4A")
        b_bullets = make_bullet_group(behavioral_bullets, x_offset=0.35, color="#007070")

        # Align bullet groups inside card
        r_bullets.move_to(LEFT * 2.35 + DOWN * 0.45)
        b_bullets.move_to(RIGHT * 2.35 + DOWN * 0.45)

        # ── Animate card in
        self.play(FadeIn(card, scale=0.95), run_time=0.5)
        self.play(Write(card_title), run_time=0.8)
        self.play(Create(divider), run_time=0.4)
        self.play(
            FadeIn(rational_header, shift=RIGHT * 0.1),
            FadeIn(behavioral_header, shift=LEFT * 0.1),
            run_time=0.5,
        )

        # Stagger bullet points in
        for r_item, b_item in zip(r_bullets, b_bullets):
            self.play(
                FadeIn(r_item, shift=RIGHT * 0.15),
                FadeIn(b_item, shift=LEFT * 0.15),
                run_time=0.4,
            )
            self.wait(0.1)

        self.wait(3.0)

        # ── Fade chart out
        chart_group = VGroup(card, card_title, divider, rational_header, behavioral_header, r_bullets, b_bullets)
        self.play(FadeOut(chart_group, scale=0.95), run_time=0.7)

        # Return characters to position
        self.play(
            self.manager.animate.shift(RIGHT * 1.5),
            self.hr_manager.animate.shift(LEFT * 1.5),
            run_time=0.6,
        )
        self.wait(0.4)

    # ── Dialogue section 2 (lines 7–11) ───────────────────────────────────────

    def _dialogue_section_two(self):

        # ── Line 7: Manager
        self.set_speech_service(_uk_service())
        self.speak_line(
            character=self.manager,
            voiceover_text=(
                "I assumed that once employees knew what to do and were paid fairly, "
                "that was enough. Are you saying there is more to it?"
            ),
            bubble_side="right",
            bubble_anchor_x=0.0,
        )

        # ── Line 8: HR Manager
        self.set_speech_service(_us_service())
        self.speak_line(
            character=self.hr_manager,
            voiceover_text=(
                "Research shows employees need skill variety, task identity, autonomy, "
                "and feedback built into their roles — not just pay. "
                "That is the Hackman and Oldham Job Characteristics Model."
            ),
            bubble_side="left",
            bubble_anchor_x=0.0,
        )

        # ── Line 9: Manager
        self.set_speech_service(_uk_service())
        self.speak_line(
            character=self.manager,
            voiceover_text=(
                "So by only optimising tasks for efficiency, we may have removed "
                "the very things that make the work meaningful to our employees?"
            ),
            bubble_side="right",
            bubble_anchor_x=0.0,
        )

        # ── Line 10: HR Manager
        self.set_speech_service(_us_service())
        self.speak_line(
            character=self.hr_manager,
            voiceover_text=(
                "Precisely. The Hawthorne Studies in 1927 proved that simply paying "
                "attention to workers — treating them as human beings — increased "
                "their productivity more than any technical change did."
            ),
            bubble_side="left",
            bubble_anchor_x=0.0,
        )

        # ── Line 11: Manager
        self.set_speech_service(_uk_service())
        self.speak_line(
            character=self.manager,
            voiceover_text=(
                "I see it now. We need both — rational design for efficiency, "
                "and the behavioural approach to keep people motivated, "
                "satisfied, and committed to doing their best work."
            ),
            bubble_side="right",
            bubble_anchor_x=0.0,
        )

    # ── Outro ──────────────────────────────────────────────────────────────────

    def _outro(self):
        # Characters face each other: nudge them slightly inward
        self.play(
            self.manager.animate.shift(RIGHT * 0.6),
            self.hr_manager.animate.shift(LEFT * 0.6),
            run_time=0.7,
        )

        # ── Handshake / agreement icon between characters (stylised check in circle)
        icon_circle = Circle(
            radius=0.52,
            color=GREEN,
            fill_color=GREEN,
            fill_opacity=0.15,
            stroke_width=3,
        )
        icon_circle.move_to(ORIGIN + UP * 0.5)

        check_line1 = Line(
            start=icon_circle.get_center() + LEFT * 0.25 + DOWN * 0.05,
            end=icon_circle.get_center() + LEFT * 0.02 + DOWN * 0.28,
            color=GREEN,
            stroke_width=4,
        )
        check_line2 = Line(
            start=check_line1.get_end(),
            end=icon_circle.get_center() + RIGHT * 0.32 + UP * 0.22,
            color=GREEN,
            stroke_width=4,
        )
        icon = VGroup(icon_circle, check_line1, check_line2)

        self.play(GrowFromCenter(icon_circle), run_time=0.6)
        self.play(Create(check_line1), Create(check_line2), run_time=0.5)

        # ── Conclusion text
        conclusion = Text(
            "Effective HRM combines BOTH the\nRational AND Behavioral approaches.",
            font_size=27,
            color=DARK_BLUE,
            weight="BOLD",
            line_spacing=0.5,
        )
        conclusion.move_to(UP * 2.5)

        self.play(FadeIn(conclusion, shift=UP * 0.2), run_time=1.0)
        self.wait(2.5)

        # ── Fade everything except end card
        all_current = VGroup(
            self.manager,
            self.hr_manager,
            self.title_bar,
            self.title_bar_text,
            icon,
            conclusion,
        )
        self.play(FadeOut(all_current), run_time=1.2)

        # ── End of Lesson card
        end_card_bg = Rectangle(
            width=config.frame_width,
            height=config.frame_height,
            fill_color=TITLE_BAR_COLOR,
            fill_opacity=1,
            stroke_width=0,
        )
        end_title = Text(
            "End of Lesson",
            font_size=54,
            color=WHITE,
            weight="BOLD",
        )
        end_title.move_to(UP * 0.4)

        watermark = Text(
            "HRM 101",
            font_size=26,
            color="#8899BB",
        )
        watermark.move_to(DOWN * 1.2)

        self.play(FadeIn(end_card_bg), run_time=0.6)
        self.play(Write(end_title), run_time=1.2)
        self.play(FadeIn(watermark, shift=UP * 0.1), run_time=0.8)
        self.wait(3.0)


# ── Direct execution support ──────────────────────────────────────────────────

if __name__ == "__main__":
    import subprocess, sys
    quality = sys.argv[1] if len(sys.argv) > 1 else "l"  # default: low quality preview
    subprocess.run(
        ["manim", f"-pq{quality}", __file__, "HRMDialogueScene"],
        check=True,
    )
