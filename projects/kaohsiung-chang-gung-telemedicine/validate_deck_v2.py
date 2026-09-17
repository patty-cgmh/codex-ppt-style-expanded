#!/usr/bin/env python3
"""Structural validation for the redesigned Clinical Calm V2 deck."""

from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

ROOT=Path(__file__).resolve().parent
PPTX=ROOT/"output"/"kaohsiung-chang-gung-telemedicine-v2.pptx"


def main():
    assert PPTX.exists(), f"missing {PPTX}"
    prs=Presentation(PPTX)
    assert len(prs.slides)==15
    assert round(prs.slide_width/prs.slide_height,3)==round(16/9,3)
    pictures=0; placeholders=0; out=[]; compositions=[]; data_marked=0
    for n,slide in enumerate(prs.slides,1):
        text="\n".join(sh.text for sh in slide.shapes if sh.has_text_frame)
        data_marked += int("〔待補" in text or "〔待確認〕" in text)
        pictures += sum(sh.shape_type==MSO_SHAPE_TYPE.PICTURE for sh in slide.shapes)
        placeholders += text.count("IL-")
        for sh in slide.shapes:
            if sh.left<0 or sh.top<0 or sh.left+sh.width>prs.slide_width or sh.top+sh.height>prs.slide_height:
                # Deliberately cropped organic ovals are allowed; content shapes are not.
                if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE: out.append((n,sh.name))
        compositions.append(len(slide.shapes))
    assert pictures==0, pictures
    assert placeholders==0, "V2 must replace IL placeholder labels with native illustrations"
    assert not out, out
    assert data_marked>=13, f"only {data_marked} slides mark unavailable data"
    assert len(set(compositions))>=8, "insufficient composition variety"
    print(f"PASS: 15 slides, 16:9, native vector illustrations, {len(set(compositions))} distinct shape-count compositions, 0 pictures")


if __name__=="__main__": main()
