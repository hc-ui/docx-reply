"""Heading lookup must match the old linear scan after the bisect speed-up."""

from docx_reply.parse import _PartData


def test_nearest_heading_matches_linear_scan():
    pd = _PartData("")
    for i in range(0, 400, 7):
        pd.headings.append((i, f"标题{i}"))
    pd.finalize()
    for idx in [*range(0, 400, 3), None, -1, 399]:
        best = ""
        if idx is not None:
            for h_idx, h_text in pd.headings:
                if h_idx <= idx:
                    best = h_text
                else:
                    break
        assert pd.nearest_heading(idx) == best
