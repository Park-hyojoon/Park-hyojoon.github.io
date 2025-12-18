import os
from main import generate_ppt

def test_error_handling():
    print("--- Starting Error Handling Verification ---")
    
    # 1. Test with non-existent file (Should return Warning)
    print("\nTest 1: Non-existent file")
    songs_before = ["non_existent_song.ppt"]
    songs_after = []
    template_path = "004.pptx" # Assumed to exist
    output_path = "test_output.pptx"
    
    # Create dummy template if needed
    if not os.path.exists(template_path):
        print("Creating dummy template for test...")
        import win32com.client
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Add()
        pres.Slides.Add(1, 12) # Blank
        pres.Slides.Add(2, 12)
        pres.Slides.Add(3, 12)
        pres.SaveAs(os.path.abspath(template_path))
        pres.Close()

    errors, warnings = generate_ppt(songs_before, songs_after, os.path.abspath(template_path), os.path.abspath(output_path), "Title", "Range", "Body")
    
    print(f"Errors: {errors}")
    print(f"Warnings: {warnings}")
    
    if any("File not found" in w for w in warnings):
        print("PASS: Correctly identified non-existent file.")
    else:
        print("FAIL: Did not warn about non-existent file.")

    # 2. Test with invalid PPT file (Should return Error)
    print("\nTest 2: Invalid PPT file")
    invalid_ppt = "invalid_song.ppt"
    with open(invalid_ppt, "w") as f:
        f.write("This is not a PPT file.")
        
    songs_before = [os.path.abspath(invalid_ppt)]
    
    errors, warnings = generate_ppt(songs_before, songs_after, os.path.abspath(template_path), os.path.abspath(output_path), "Title", "Range", "Body")
    
    print(f"Errors: {errors}")
    print(f"Warnings: {warnings}")
    
    if any("Failed to convert" in e for e in errors):
        print("PASS: Correctly identified conversion failure.")
    else:
        print("FAIL: Did not report conversion failure.")

    # Cleanup
    if os.path.exists(invalid_ppt):
        os.remove(invalid_ppt)
    if os.path.exists(output_path):
        os.remove(output_path)
        
    print("\n--- Verification Complete ---")

if __name__ == "__main__":
    test_error_handling()
