import win32com.client
import os

def analyze_ppt(ppt_path):
    print(f"Analyzing: {ppt_path}")
    if not os.path.exists(ppt_path):
        print("File not found.")
        return

    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        pres = app.Presentations.Open(ppt_path)
        
        print(f"Total Slides: {pres.Slides.Count}")
        
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            print(f"\n--- Slide {i} ---")
            text_shapes = []
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text = shape.TextFrame.TextRange.Text
                    print(f"  Shape: {shape.Name}, Top: {shape.Top}, Left: {shape.Left}, Text: '{text}'")
                    text_shapes.append(shape)
            
            if text_shapes:
                text_shapes.sort(key=lambda s: s.Top, reverse=True)
                print(f"  -> Bottom-most text shape: {text_shapes[0].Name} (Text: '{text_shapes[0].TextFrame.TextRange.Text}')")

        pres.Close()
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    analyze_ppt(os.path.join(base_dir, "result_friday.pptx"))
