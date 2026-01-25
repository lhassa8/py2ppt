"""Animation and transition XML handling.

Provides dataclasses and XML generation for slide transitions
and shape animations in PresentationML format.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from lxml import etree

from .ns import qn

# Transition type mappings to PresentationML
TRANSITION_MAP = {
    "fade": "fade",
    "push": "push",
    "wipe": "wipe",
    "split": "split",
    "blinds": "blinds",
    "checker": "checker",
    "circle": "circle",
    "dissolve": "dissolve",
    "comb": "comb",
    "cover": "cover",
    "cut": "cut",
    "diamond": "diamond",
    "newsflash": "newsflash",
    "plus": "plus",
    "pull": "pull",
    "random": "random",
    "randombar": "randomBar",
    "strips": "strips",
    "wedge": "wedge",
    "wheel": "wheel",
    "zoom": "zoom",
    "none": None,
}

# Animation effect mappings
ANIMATION_MAP = {
    # Entrance effects
    "appear": {"preset_class": "entr", "preset_id": 1},
    "fade": {"preset_class": "entr", "preset_id": 10},
    "fly_in": {"preset_class": "entr", "preset_id": 2},
    "float_in": {"preset_class": "entr", "preset_id": 42},
    "split": {"preset_class": "entr", "preset_id": 16},
    "wipe": {"preset_class": "entr", "preset_id": 22},
    "shape": {"preset_class": "entr", "preset_id": 17},
    "wheel": {"preset_class": "entr", "preset_id": 21},
    "random_bars": {"preset_class": "entr", "preset_id": 14},
    "grow_turn": {"preset_class": "entr", "preset_id": 37},
    "zoom": {"preset_class": "entr", "preset_id": 23},
    "swivel": {"preset_class": "entr", "preset_id": 19},
    "bounce": {"preset_class": "entr", "preset_id": 26},
    # Emphasis effects
    "pulse": {"preset_class": "emph", "preset_id": 5},
    "color_pulse": {"preset_class": "emph", "preset_id": 18},
    "teeter": {"preset_class": "emph", "preset_id": 7},
    "spin": {"preset_class": "emph", "preset_id": 8},
    "grow_shrink": {"preset_class": "emph", "preset_id": 6},
    "desaturate": {"preset_class": "emph", "preset_id": 9},
    "darken": {"preset_class": "emph", "preset_id": 10},
    "lighten": {"preset_class": "emph", "preset_id": 11},
    "transparency": {"preset_class": "emph", "preset_id": 12},
    "object_color": {"preset_class": "emph", "preset_id": 13},
    # Exit effects
    "disappear": {"preset_class": "exit", "preset_id": 1},
    "fade_out": {"preset_class": "exit", "preset_id": 10},
    "fly_out": {"preset_class": "exit", "preset_id": 2},
    "float_out": {"preset_class": "exit", "preset_id": 42},
    "split_out": {"preset_class": "exit", "preset_id": 16},
    "wipe_out": {"preset_class": "exit", "preset_id": 22},
    "shape_out": {"preset_class": "exit", "preset_id": 17},
    "wheel_out": {"preset_class": "exit", "preset_id": 21},
    "random_bars_out": {"preset_class": "exit", "preset_id": 14},
    "shrink_turn": {"preset_class": "exit", "preset_id": 37},
    "zoom_out": {"preset_class": "exit", "preset_id": 23},
    "swivel_out": {"preset_class": "exit", "preset_id": 19},
    "bounce_out": {"preset_class": "exit", "preset_id": 26},
    # Motion paths
    "path_right": {"preset_class": "path", "preset_id": 1},
    "path_left": {"preset_class": "path", "preset_id": 2},
    "path_up": {"preset_class": "path", "preset_id": 3},
    "path_down": {"preset_class": "path", "preset_id": 4},
}


@dataclass
class SlideTransition:
    """Slide transition configuration.

    Attributes:
        type: Transition type name (e.g., "fade", "push", "wipe")
        duration: Duration in milliseconds (default 1000)
        advance_on_click: Whether to advance on mouse click
        advance_time: Auto-advance time in milliseconds (None = no auto-advance)
        direction: Direction for applicable transitions (e.g., "l", "r", "u", "d")
    """

    type: str = "fade"
    duration: int = 1000
    advance_on_click: bool = True
    advance_time: int | None = None
    direction: str | None = None

    def to_xml(self) -> etree._Element:
        """Generate transition XML element."""
        # Create p:transition element
        transition = etree.Element(qn("p:transition"))

        # Set duration in 1/1000 of a second format (spd attribute uses predefined values)
        if self.duration <= 500:
            transition.set("spd", "fast")
        elif self.duration >= 1500:
            transition.set("spd", "slow")
        else:
            transition.set("spd", "med")

        # Set advance options
        transition.set("advClick", "1" if self.advance_on_click else "0")

        if self.advance_time is not None:
            transition.set("advTm", str(self.advance_time))

        # Add transition type element
        trans_type = TRANSITION_MAP.get(self.type.lower())
        if trans_type:
            type_elem = etree.SubElement(transition, qn(f"p:{trans_type}"))

            # Add direction if specified
            if self.direction:
                type_elem.set("dir", self.direction)

        return transition


@dataclass
class ShapeAnimation:
    """Shape animation configuration.

    Attributes:
        shape_id: The shape ID to animate
        effect: Animation effect name (e.g., "appear", "fade", "fly_in")
        trigger: When to trigger ("onClick", "withPrevious", "afterPrevious")
        duration: Duration in milliseconds
        delay: Delay before starting in milliseconds
        direction: Direction for applicable effects
        sequence: Sequence number for ordering
    """

    shape_id: int
    effect: str = "appear"
    trigger: Literal["onClick", "withPrevious", "afterPrevious"] = "onClick"
    duration: int = 500
    delay: int = 0
    direction: str | None = None
    sequence: int = 1

    def get_preset_info(self) -> dict:
        """Get preset class and ID for the effect."""
        return ANIMATION_MAP.get(self.effect, {"preset_class": "entr", "preset_id": 1})


@dataclass
class AnimationSequence:
    """A sequence of animations for a slide.

    Attributes:
        animations: List of shape animations in order
    """

    animations: list[ShapeAnimation] = field(default_factory=list)

    def to_xml(self) -> etree._Element | None:
        """Generate timing XML for the animations."""
        if not self.animations:
            return None

        # Create p:timing element
        timing = etree.Element(qn("p:timing"))

        # Create time node list
        tn_lst = etree.SubElement(timing, qn("p:tnLst"))

        # Create parallel time node (container for all animations)
        par = etree.SubElement(tn_lst, qn("p:par"))
        ctn = etree.SubElement(par, qn("p:cTn"))
        ctn.set("id", "1")
        ctn.set("dur", "indefinite")
        ctn.set("restart", "never")
        ctn.set("nodeType", "tmRoot")

        # Child container for sequences
        child_tn_lst = etree.SubElement(ctn, qn("p:childTnLst"))

        # Create a sequence container
        seq = etree.SubElement(child_tn_lst, qn("p:seq"))
        seq.set("concurrent", "1")
        seq.set("nextAc", "seek")

        seq_ctn = etree.SubElement(seq, qn("p:cTn"))
        seq_ctn.set("id", "2")
        seq_ctn.set("dur", "indefinite")
        seq_ctn.set("nodeType", "mainSeq")

        seq_child_lst = etree.SubElement(seq_ctn, qn("p:childTnLst"))

        # Previous condition container
        prev_cond = etree.SubElement(seq, qn("p:prevCondLst"))
        cond = etree.SubElement(prev_cond, qn("p:cond"))
        cond.set("evt", "onPrev")
        cond.set("delay", "0")
        tgt_el = etree.SubElement(cond, qn("p:tgtEl"))
        sld_tgt = etree.SubElement(tgt_el, qn("p:sldTgt"))  # noqa: F841

        # Next condition container
        next_cond = etree.SubElement(seq, qn("p:nextCondLst"))
        cond = etree.SubElement(next_cond, qn("p:cond"))
        cond.set("evt", "onNext")
        cond.set("delay", "0")
        tgt_el = etree.SubElement(cond, qn("p:tgtEl"))
        sld_tgt = etree.SubElement(tgt_el, qn("p:sldTgt"))  # noqa: F841

        # Add each animation as a parallel container
        node_id = 3
        for i, anim in enumerate(self.animations):
            node_id = self._add_animation_node(
                seq_child_lst, anim, node_id, i == 0
            )

        # Build target list
        bld_lst = etree.SubElement(timing, qn("p:bldLst"))
        for anim in self.animations:
            bld_p = etree.SubElement(bld_lst, qn("p:bldP"))
            bld_p.set("spid", str(anim.shape_id))
            bld_p.set("grpId", "0")

        return timing

    def _add_animation_node(
        self,
        parent: etree._Element,
        anim: ShapeAnimation,
        start_id: int,
        is_first: bool,
    ) -> int:
        """Add animation node to parent container.

        Returns the next available node ID.
        """
        preset = anim.get_preset_info()
        node_id = start_id

        # Create parallel container for this animation
        par = etree.SubElement(parent, qn("p:par"))
        ctn = etree.SubElement(par, qn("p:cTn"))
        ctn.set("id", str(node_id))
        node_id += 1
        ctn.set("fill", "hold")

        # Add start condition
        st_cond_lst = etree.SubElement(ctn, qn("p:stCondLst"))
        cond = etree.SubElement(st_cond_lst, qn("p:cond"))

        if anim.trigger == "onClick" or is_first:
            cond.set("delay", "0" if anim.delay == 0 else str(anim.delay))
        elif anim.trigger == "withPrevious":
            cond.set("delay", str(anim.delay))
        else:  # afterPrevious
            cond.set("delay", str(anim.delay))
            cond.set("evt", "onEnd")
            tn = etree.SubElement(cond, qn("p:tn"))
            tn.set("val", str(node_id - 2))

        # Child container for effect
        child_lst = etree.SubElement(ctn, qn("p:childTnLst"))

        # Effect container
        par2 = etree.SubElement(child_lst, qn("p:par"))
        ctn2 = etree.SubElement(par2, qn("p:cTn"))
        ctn2.set("id", str(node_id))
        node_id += 1
        ctn2.set("presetID", str(preset["preset_id"]))
        ctn2.set("presetClass", preset["preset_class"])
        ctn2.set("presetSubtype", "0")
        ctn2.set("fill", "hold")
        ctn2.set("nodeType", "clickEffect" if anim.trigger == "onClick" else "afterEffect")

        # Add start condition for effect
        st_cond_lst2 = etree.SubElement(ctn2, qn("p:stCondLst"))
        cond2 = etree.SubElement(st_cond_lst2, qn("p:cond"))
        cond2.set("delay", "0")

        # Child for actual animation behaviors
        child_lst2 = etree.SubElement(ctn2, qn("p:childTnLst"))

        # Add set behavior (visibility)
        self._add_set_behavior(child_lst2, anim, node_id)
        node_id += 1

        # Add animation effect based on type
        if preset["preset_class"] == "entr" and anim.effect == "fade":
            self._add_fade_effect(child_lst2, anim, node_id)
            node_id += 1

        return node_id

    def _add_set_behavior(
        self,
        parent: etree._Element,
        anim: ShapeAnimation,
        node_id: int,
    ) -> None:
        """Add set behavior for visibility."""
        set_elem = etree.SubElement(parent, qn("p:set"))
        cbn = etree.SubElement(set_elem, qn("p:cBhvr"))
        ctn = etree.SubElement(cbn, qn("p:cTn"))
        ctn.set("id", str(node_id))
        ctn.set("dur", "1")
        ctn.set("fill", "hold")

        st_cond_lst = etree.SubElement(ctn, qn("p:stCondLst"))
        cond = etree.SubElement(st_cond_lst, qn("p:cond"))
        cond.set("delay", "0")

        tgt_el = etree.SubElement(cbn, qn("p:tgtEl"))
        sp_tgt = etree.SubElement(tgt_el, qn("p:spTgt"))
        sp_tgt.set("spid", str(anim.shape_id))

        attr_name_lst = etree.SubElement(cbn, qn("p:attrNameLst"))
        attr_name = etree.SubElement(attr_name_lst, qn("p:attrName"))
        attr_name.text = "style.visibility"

        to_elem = etree.SubElement(set_elem, qn("p:to"))
        str_val = etree.SubElement(to_elem, qn("p:strVal"))
        str_val.set("val", "visible")

    def _add_fade_effect(
        self,
        parent: etree._Element,
        anim: ShapeAnimation,
        node_id: int,
    ) -> None:
        """Add fade animation effect."""
        anim_effect = etree.SubElement(parent, qn("p:animEffect"))
        anim_effect.set("transition", "in")
        anim_effect.set("filter", "fade")

        cbn = etree.SubElement(anim_effect, qn("p:cBhvr"))
        ctn = etree.SubElement(cbn, qn("p:cTn"))
        ctn.set("id", str(node_id))
        ctn.set("dur", str(anim.duration))

        tgt_el = etree.SubElement(cbn, qn("p:tgtEl"))
        sp_tgt = etree.SubElement(tgt_el, qn("p:spTgt"))
        sp_tgt.set("spid", str(anim.shape_id))


def parse_transition(element: etree._Element) -> SlideTransition | None:
    """Parse a p:transition element into SlideTransition."""
    if element is None:
        return None

    # Get speed
    spd = element.get("spd", "med")
    if spd == "fast":
        duration = 500
    elif spd == "slow":
        duration = 2000
    else:
        duration = 1000

    # Get advance options
    adv_click = element.get("advClick", "1") == "1"
    adv_tm = element.get("advTm")
    advance_time = int(adv_tm) if adv_tm else None

    # Find transition type
    trans_type = "none"
    direction = None
    for child in element:
        tag = etree.QName(child.tag).localname
        for name, pml_name in TRANSITION_MAP.items():
            if pml_name == tag:
                trans_type = name
                direction = child.get("dir")
                break
        if trans_type != "none":
            break

    return SlideTransition(
        type=trans_type,
        duration=duration,
        advance_on_click=adv_click,
        advance_time=advance_time,
        direction=direction,
    )
