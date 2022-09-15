from pptx import Presentation
from pptx.util import Inches


ppt = Presentation("resource/report.potx")
slide = ppt.slides[0]

for shape in slide.shapes:
    print(shape.shape_type)