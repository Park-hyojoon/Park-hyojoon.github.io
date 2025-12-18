import win32com.client
import os
import time
import pythoncom
import traceback

class PowerPointManager:
    """
    Context manager to ensure PowerPoint application is properly closed.
    Prevents 'File in use' and 'Server execution failed' errors by handling cleanup.
    """
    def __init__(self):
        self.app = None
        self.presentations = []

    def __enter__(self):
        try:
            self.app = win32com.client.Dispatch("PowerPoint.Application")
            self.app.Visible = True
            return self
        except Exception as e:
            print(f"Failed to initialize PowerPoint: {e}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Close all opened presentations
        for pres in self.presentations:
            try:
                pres.Close()
            except:
                pass
        
        # Quit Application
        if self.app:
            try:
                self.app.Quit()
            except:
                pass
        
        # Release COM object
        self.app = None

    def open_presentation(self, path):
        if not self.app:
            raise Exception("PowerPoint app is not initialized.")
        try:
            pres = self.app.Presentations.Open(path)
            self.presentations.append(pres)
            return pres
        except Exception as e:
            print(f"Error opening {path}: {e}")
            raise

    def close_presentation(self, pres):
        if pres in self.presentations:
            try:
                pres.Close()
            except:
                pass
            self.presentations.remove(pres)

def convert_ppt_to_pptx(ppt_mgr, ppt_path):
    """Converts a .ppt file to .pptx format using the existing PowerPoint manager."""
    pptx_path = ppt_path + "x"
    
    # Check if file exists and is valid
    if os.path.exists(pptx_path):
        if os.path.getsize(pptx_path) > 0:
            print(f"File already exists and is valid: {pptx_path}")
            return pptx_path
        else:
            print(f"File exists but is empty, deleting: {pptx_path}")
            try:
                os.remove(pptx_path)
            except Exception as e:
                print(f"Warning: Could not delete empty file {pptx_path}: {e}")
    
    print(f"Converting {ppt_path} to {pptx_path}...")
    
    try:
        presentation = ppt_mgr.open_presentation(ppt_path)
        presentation.SaveAs(pptx_path, 24) # 24 is ppSaveAsOpenXMLPresentation
        ppt_mgr.close_presentation(presentation)
        print(f"Conversion successful: {pptx_path}")
        return pptx_path
    except Exception as e:
        print(f"Error converting {ppt_path}: {e}")
        raise Exception(f"Failed to convert {os.path.basename(ppt_path)}. Error: {str(e)}")

def setup_worship_title(slide, new_title):
    """
    Finds a text box on the slide containing '기도회' and replaces it with new_title.
    Preserves existing formatting as much as possible by setting TextRange.Text.
    """
    try:
        found = False
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text
                # Check for key keywords that identify the title box
                if "기도회" in text or "예배" in text:
                    shape.TextFrame.TextRange.Text = new_title
                    found = True
                    # Optional: We could break here, but if there are multiple parts (unlikely), 
                    # we might want to check them. But usually title is one box.
                    print(f"Updated worship title to: {new_title}")
                    break
        
        if not found:
            print(f"Warning: Could not find a text box containing '기도회' or '예배' on Slide {slide.SlideIndex}.")
            
    except Exception as e:
        print(f"Error updating worship title on Slide {slide.SlideIndex}: {e}")

def setup_bible_slide(slide, text):
    """Updates the bottom-most text box on the given slide with text and centers all text boxes."""
    try:
        slide_width = slide.Parent.PageSetup.SlideWidth
        text_shapes = []
        
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_shapes.append(shape)
        
        if not text_shapes:
            print(f"No text shapes found on Slide {slide.SlideIndex}.")
            return

        # Sort by Top position (descending) to find the bottom-most shape
        text_shapes.sort(key=lambda s: s.Top, reverse=True)
        
        target_shape = text_shapes[0]
        # Clear existing text first to avoid formatting issues
        target_shape.TextFrame.TextRange.Text = text
        
        # Center align ALL text boxes on the slide
        for shape in text_shapes:
            try:
                # Align text to center (ppAlignCenter = 2)
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                # Align shape to center of slide
                shape.Left = (slide_width - shape.Width) / 2
            except Exception as align_err:
                print(f"Could not align shape {shape.Name}: {align_err}")
                
    except Exception as e:
        print(f"Error updating Slide {slide.SlideIndex if 'slide' in locals() else 'Unknown'}: {e}")

def setup_bible_body_slide(slide, chapter_verse, body_text):
    """Updates Slide 5 with Chapter/Verse (top) and Body (bottom) text, and centers them."""
    try:
        slide_width = slide.Parent.PageSetup.SlideWidth
        text_shapes = []
        
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_shapes.append(shape)
        
        if len(text_shapes) < 2:
            print(f"Warning: Slide {slide.SlideIndex} needs at least 2 text boxes, found {len(text_shapes)}.")
            if not text_shapes:
                return

        # Sort by Top position (ascending)
        text_shapes.sort(key=lambda s: s.Top)
        
        # Top-most is Chapter/Verse
        chapter_shape = text_shapes[0]
        chapter_shape.TextFrame.TextRange.Text = chapter_verse
        
        # Bottom-most is Body
        if len(text_shapes) >= 2:
            body_shape = text_shapes[-1]
            body_shape.TextFrame.TextRange.Text = body_text
        
        # Center align ALL text boxes
        for shape in text_shapes:
            try:
                # Align text to center
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                # Align shape to center of slide
                shape.Left = (slide_width - shape.Width) / 2
            except Exception as align_err:
                print(f"Could not align shape {shape.Name}: {align_err}")
                
    except Exception as e:
        print(f"Error updating Slide {slide.SlideIndex if 'slide' in locals() else 'Unknown'}: {e}")

def setup_sermon_title_slide(slide, title):
    """
    Finds a text box on the sermon slide (Slide 6) and replaces it with the title.
    Looks for placeholders like 'Sermon Title', '설교 제목', etc.
    """
    try:
        found = False
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text
                # Check for keywords
                if "Sermon" in text or "Title" in text or "설교" in text or "제목" in text:
                    shape.TextFrame.TextRange.Text = title
                    found = True
                    print(f"Updated Sermon Title slide.")
                    break
        
        if not found:
             print(f"Warning: Could not identify 'Sermon Title' box on Slide {slide.SlideIndex}.")

    except Exception as e:
        print(f"Error updating Sermon Title slide: {e}")

def generate_ppt(songs_before, songs_after, template_path, output_path, worship_title, bible_title, bible_range, bible_body, sermon_title=""):
    print(f"Template Path: {template_path}")
    print(f"Output File: {output_path}")

    errors = []
    warnings = []
    
    # Verify template exists before starting PowerPoint
    if not os.path.exists(template_path):
         msg = f"Template file not found: {template_path}"
         print(f"Error: {msg}")
         errors.append(msg)
         return errors, warnings

    # Use Context Manager for safety
    try:
        with PowerPointManager() as ppt_mgr:
            
            # Helper to process files (convert .ppt to .pptx)
            def process_file_list(file_list):
                processed = []
                if not file_list:
                    return processed
                    
                for file_path in file_list:
                    if not os.path.exists(file_path):
                        msg = f"File not found: {file_path}"
                        print(f"Warning: {msg}")
                        warnings.append(msg)
                        continue
                        
                    if file_path.lower().endswith(".ppt"):
                        # Convert to .pptx using the SAME ppt_mgr instance
                        try:
                            converted_path = convert_ppt_to_pptx(ppt_mgr, file_path)
                            if converted_path:
                                processed.append(converted_path)
                        except Exception as e:
                            msg = f"Failed to convert {os.path.basename(file_path)}: {str(e)}"
                            print(msg)
                            errors.append(msg)
                    elif file_path.lower().endswith(".pptx"):
                        processed.append(file_path)
                    else:
                        msg = f"Skipping unsupported file type: {os.path.basename(file_path)}"
                        print(msg)
                        warnings.append(msg)
                return processed

            print("Processing 'Before Sermon' songs...")
            songs_before_bible = process_file_list(songs_before)
            
            print("Processing 'After Sermon' songs...")
            songs_after_bible = process_file_list(songs_after)

            # Open Template
            print(f"Opening template: {template_path}")
            # We open it as a copy to avoid locking the template, but SaveAs handles this too.
            # Using Open() is fine as long as we SaveAs immediately.
            main_pres = ppt_mgr.open_presentation(template_path)
            
            # Ensure output directory exists
            output_path = os.path.abspath(output_path)
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            try:
                main_pres.SaveAs(output_path)
                print(f"Saved initial copy to: {output_path}")
            except Exception as e:
                # If we can't save, it's critical.
                raise Exception(f"Error saving to {output_path}: {e}")

            # Basic Validation
            if main_pres.Slides.Count < 3:
                raise Exception("Template must have at least 3 slides.")

            # Update Slide 1: Worship Title
            setup_worship_title(main_pres.Slides(1), worship_title)
            
            # Update Slide 1 & 4 with Bible Reference
            setup_bible_slide(main_pres.Slides(1), bible_title)
            
            if main_pres.Slides.Count >= 4:
                 setup_bible_slide(main_pres.Slides(4), bible_title)
            
            # Update Slide 5 with Bible Body (Splitting logic)
            if main_pres.Slides.Count >= 5:
                bible_parts = [part.strip() for part in bible_body.split('/')]
                
                # Start at Slide 5
                current_bible_slide_index = 5
                
                for i, part in enumerate(bible_parts):
                    # Always verify the slide exists at the expected index
                    if current_bible_slide_index > main_pres.Slides.Count:
                        raise Exception(f"Logic Error: Expected slide at {current_bible_slide_index} but Count is {main_pres.Slides.Count}")
                        
                    current_slide = main_pres.Slides(current_bible_slide_index)
                    
                    if i == 0:
                        # First part: modify existing Slide 5
                        setup_bible_body_slide(current_slide, bible_range, part)
                    else:
                        # Subsequent parts: Copy previous slide
                        main_pres.Slides(current_bible_slide_index).Copy()
                        
                        # Paste after current slide
                        # Note: Paste usually pastes AFTER the current selection or at the end? 
                        # To be safe, we select the current slide, then Paste.
                        main_pres.Slides(current_bible_slide_index).Select()
                        ppt_mgr.app.CommandBars.ExecuteMso("PasteSourceFormatting")
                        time.sleep(0.5) # Wait for paste
                        
                        # The new slide should be at index + 1
                        current_bible_slide_index += 1
                        
                        # Verify we have the new slide
                        if current_bible_slide_index > main_pres.Slides.Count:
                             # Wait a bit longer if needed
                             time.sleep(1)
                             if current_bible_slide_index > main_pres.Slides.Count:
                                 raise Exception("Paste failed: New slide not found.")

                        setup_bible_body_slide(main_pres.Slides(current_bible_slide_index), bible_range, part)
            else:
                warnings.append("Warning: Slide 5 not found in template.")

            # Sermon Title Logic (Wednesday Mode)
            # If sermon_title is provided, we expect a Sermon Slide (Slide 6 original).
            # Due to Bible splitting (if any), the Sermon Slide is at 'current_bible_slide_index + 1'.
            if sermon_title:
                sermon_slide_index = current_bible_slide_index + 1
                if main_pres.Slides.Count >= sermon_slide_index:
                    print(f"Updating Sermon Title on Slide {sermon_slide_index}...")
                    setup_sermon_title_slide(main_pres.Slides(sermon_slide_index), sermon_title)
                else:
                    msg = "Wednesday Mode selected but Slide 6 (Sermon Title) not found in template."
                    print(msg)
                    warnings.append(msg)

            # --- Songs Insertion Logic ---
            
            # Break Slide is originally Slide 3. 
            # We want to COPY Slide 3 to insert as a break.
            break_slide_index = 3
            
            # 1. Insert BEFORE Bible (After Slide 3)
            # The insertion point starts after Slide 3
            current_insert_index = 3
            
            def insert_songs_at(songs_list, target_index):
                nonlocal current_insert_index
                
                for song_path in songs_list:
                    print(f"Inserting song: {os.path.basename(song_path)}")
                    try:
                        # Open song using the manager (so it gets closed properly)
                        song_pres = ppt_mgr.open_presentation(song_path)
                        song_slide_count = song_pres.Slides.Count
                        song_pres.Slides.Range().Copy()
                        ppt_mgr.close_presentation(song_pres) # Close immediately after copy
                        
                        # Paste into Main
                        # We want to paste AFTER 'target_index'
                        # To paste after slide N, we select slide N.
                        main_pres.Slides(target_index).Select()
                        ppt_mgr.app.CommandBars.ExecuteMso("PasteSourceFormatting")
                        time.sleep(1)
                        
                        # Update index: we added N slides
                        target_index += song_slide_count
                        
                        # Insert Break Slide AFTER the song
                        main_pres.Slides(break_slide_index).Copy()
                        main_pres.Slides(target_index).Select()
                        ppt_mgr.app.CommandBars.ExecuteMso("PasteSourceFormatting")
                        time.sleep(0.5)
                        
                        # Update index: we added 1 break slide
                        target_index += 1
                        
                    except Exception as e:
                        msg = f"Error inserting song {os.path.basename(song_path)}: {e}"
                        print(msg)
                        errors.append(msg)
                
                return target_index

            # Process Before Bible Songs
            current_insert_index = insert_songs_at(songs_before_bible, current_insert_index)
            
            # 2. Insert Break Slide AFTER Bible section
            # The Bible section ends at the last Bible body slide.
            # But wait, our 'current_insert_index' tracking for "Songs Before" stopped right before the Bible section started?
            # No, the logic in the original code was: 
            # - Insert songs after Slide 3 (Break Slide).
            # - Then later, Slide 4 (Bible Title) and Slide 5+ (Body) come AFTER that.
            # 
            # CRITICAL CORRECTION: 
            # When we insert slides at index 3, the existing slides (4, 5, etc.) shift DOWN.
            # So Slide 4 (originally) becomes Slide 4 + N_inserted.
            # We need to be careful. The original code did insertions *before* touching Bible slides?
            # NO: The original code updated Bible slides FIRST (lines 224-257), THEN inserted songs (lines 264+).
            # If we updated Bible slides first, Slide 4 and 5 are fixed content.
            # 
            # BUT, the original code inserted "Batch 1" at `current_insert_index = 3`. 
            # If we insert at 3, the newly updated Bible slides (originally at 4, 5...) will be pushed down.
            # This is CORRECT behavior if we want Songs -> Break -> Bible -> Break -> Songs.
            #
            # However, we must ensure `break_slide_index` (3) is still valid? Yes, Slide 3 stays at 3 unless we insert *before* 3.
            # We are inserting *after* 3. So Slide 3 is safe.
            # 
            # What about the Bible slides we just updated? 
            # We updated them *before* inserting songs.
            # When we insert songs after Slide 3, the Bible slides (which were at 4, 5...) shift to (4+N, 5+N...).
            # This is fine, we don't need to reference them by index anymore.
            
            # --- Insert Break Slide AFTER the Bible Section ---
            # Where is the end of the Bible section?
            # It WAS at the end of the presentation before we added "Songs After".
            # Actually, "Songs After" are appended to the very end.
            # So we can just append a Break Slide at the current end (which is the end of Bible body), 
            # THEN append "Songs After".
            
            print("Inserting Break Slide after Bible slides...")
            main_pres.Slides(break_slide_index).Copy()
            # Paste at the end
            main_pres.Slides(main_pres.Slides.Count).Select()
            ppt_mgr.app.CommandBars.ExecuteMso("PasteSourceFormatting")
            time.sleep(0.5)
            
            # Now insert "Songs After" at the very end
            current_end_index = main_pres.Slides.Count
            insert_songs_at(songs_after_bible, current_end_index)

            print("Inserted songs and Break Slides.")
            
            main_pres.Save()
            print(f"Final save to: {output_path}")

    except Exception as e:
        msg = f"An unexpected error occurred: {e}"
        print(msg)
        traceback.print_exc()
        errors.append(msg)
    
    return errors, warnings

def main():
    # Only for testing, not used by GUI directly
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ppt_dir = os.path.join(base_dir, "ppt")
    template_path = os.path.join(base_dir, "004.pptx")
    output_path = os.path.join(base_dir, "result_friday.pptx")
    
    bible_title = "베드로전서 1:1"
    bible_range = "베드로전서 1:1-2"
    bible_body = "1   ... / 2   ..."
    worship_title = "금요 기도회" # Test value

    songs_before = []
    songs_after = []
    
    if os.path.exists(ppt_dir):
        all_songs = [os.path.join(ppt_dir, f) for f in os.listdir(ppt_dir) if f.lower().endswith(('.ppt', '.pptx')) and not f.startswith("~$")]
        all_songs.sort()
        songs_before = all_songs[:1]
        songs_after = all_songs[1:]

    generate_ppt(songs_before, songs_after, template_path, output_path, worship_title, bible_title, bible_range, bible_body)

if __name__ == "__main__":
    main()
