"""Import-time speed-ups that keep extract_review output identical.

``parse._walk_part`` still builds the heading list in document order; this
module only replaces the O(n) nearest-heading scan with a bisect after the
walk finishes. Tests that construct ``_PartData`` directly call ``finalize``.
"""

from __future__ import annotations

from bisect import bisect_right

from . import parse as _parse_mod
from .parse import _PartData
from .parse import _walk_part as _walk_part_orig


def finalize(self) -> None:
    self.headings.sort(key=lambda h: h[0])
    self._heading_idx = [h[0] for h in self.headings]


def nearest_heading(self, idx: int | None) -> str:
    if idx is None or not self.headings:
        return ""
    index = getattr(self, "_heading_idx", None)
    if not index:
        finalize(self)
        index = self._heading_idx
    i = bisect_right(index, idx) - 1
    return self.headings[i][1] if i >= 0 else ""


def _walk_part(container, label, state):
    _walk_part_orig(container, label, state)
    state.parts[label].finalize()


_PartData.finalize = finalize
_PartData.nearest_heading = nearest_heading
# extract_review calls parse._walk_part; rebind the name it already imported.
_parse_mod._walk_part = _walk_part
