import win32com.client
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.join(base_dir, "004.pptx")

try:
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(template_path)
    print(f"Slide count: {pres.Slides.Count}")
    pres.Close()
except Exception as e:
    print(f"Error: {e}")
