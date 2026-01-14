import re

def normalize_input(text: str) -> str:
    s = text.strip()
    s = s.replace("$$", "")
    s = s.replace("\\begin{equation}", "")
    s = s.replace("\\end{equation}", "")
    s = s.replace("Before:", "")
    s = s.replace("After:", "")
    s = s.replace("Style 1:", "")
    s = s.replace("Style 2:", "")
    s = s.replace("Style 3:", "")
    s = re.sub(r"\\\\\[[^\]]*\]", r"\\\\", s)
    s = re.sub(r"\s+", " ", s)
    return s
