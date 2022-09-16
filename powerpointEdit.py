from pptx import Presentation
from pptx.util import Inches


ppt = Presentation("resource/report.pptx")
slide = ppt.slides[2]

for shape in slide.shapes:
    print(shape.shape_type)