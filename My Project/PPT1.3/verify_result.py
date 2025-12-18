import win32com.client
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
result_path = os.path.join(base_dir, "result_friday.pptx")

try:
    app = win32com.client.Dispatch("PowerPoint.Application")
    # app.Visible = True
    pres = app.Presentations.Open(result_path)
    print(f"Total Slide count: {pres.Slides.Count}")
    
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        print(f"Slide {i}: Layout {slide.Layout}")
        
    pres.Close()
except Exception as e:
    print(f"Error: {e}")
