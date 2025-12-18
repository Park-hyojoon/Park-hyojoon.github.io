import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import threading
import pythoncom
from main import generate_ppt

import datetime

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT Automation Tool")
        self.root.geometry("1200x720") # Increased width to 1050 for path visibility

        # Variables
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Worship Title (Default: 금요 기도회)
        self.worship_title_var = tk.StringVar(value="금요 기도회")
        
        # 1) PPT Folder (Songs) -> D:\05. Download
        self.ppt_dir_var = tk.StringVar(value=r"D:\05. Download")
        
        # 2) Template File -> D:\02. 열띰!\02. 교회\03. 금요기도회 PPT\004.pptx
        # Assuming 004.pptx is the filename inside that folder
        # 2) Template File -> D:\02. 열띰!\02. 교회\03. 금요기도회 PPT\friday.pptx
        # We start with Friday default
        self.template_path_var = tk.StringVar(value=r"D:\02. 열띰!\02. 교회\03. 금요기도회 PPT\friday.pptx")
        
        self.is_wednesday_var = tk.BooleanVar(value=False)
        self.sermon_title_var = tk.StringVar(value="")
        
        self.bible_title_var = tk.StringVar(value="")
        # self.bible_range_var removed as requested
        
        # Calculate next Friday for default filename
        today = datetime.date.today()
        friday = today + datetime.timedelta((4 - today.weekday()) % 7)
        default_filename = f"{friday.strftime('%Y년 %m월 %d일')} 금요기도회.pptx"
        
        # 3) Output File -> D:\02. 열띰!\02. 교회\03. 금요기도회 PPT
        self.output_path_var = tk.StringVar(value=os.path.join(r"D:\02. 열띰!\02. 교회\03. 금요기도회 PPT", default_filename))
        
        # UI Elements
        self.create_widgets()
        
        # Initial population - Removed as requested
        # self.populate_song_lists()

        # Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        about_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="About", menu=about_menu)
        about_menu.add_command(label="Info", command=self.show_about)

    def show_about(self):
        messagebox.showinfo("About", "2025년 12월 5일 FridayWorshipPPT v1.35 완성")

    def create_widgets(self):
        # Main Container (PanedWindow or just Frames)
        # Using Grid to allocate more weight to Left Frame (approx 60/40 split)
        main_container = tk.Frame(self.root)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        main_container.grid_columnconfigure(0, weight=5, uniform="group1") # Left Frame (70%)
        main_container.grid_columnconfigure(1, weight=5, uniform="group1") # Right Frame (30%)
        main_container.grid_rowconfigure(0, weight=1)

        # Left Frame (Settings & Lists)
        left_frame = tk.Frame(main_container)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # Right Frame (Inputs & Action)
        right_frame = tk.Frame(main_container)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # === LEFT FRAME CONTENT ===

        # 1. PPT Directory
        tk.Label(left_frame, text="PPT Folder (Songs):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 2))
        frame_ppt = tk.Frame(left_frame)
        frame_ppt.pack(fill="x", pady=(0, 10))
        tk.Entry(frame_ppt, textvariable=self.ppt_dir_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame_ppt, text="Browse", command=self.browse_ppt_dir).pack(side="right", padx=2)
        
        # Tools Row
        frame_tools = tk.Frame(left_frame)
        frame_tools.pack(fill="x", pady=(0, 10))
        tk.Button(frame_tools, text="Refresh", command=self.populate_song_lists).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(frame_tools, text="Delete All", command=self.clear_all_lists).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(frame_tools, text="FIX PPT", command=self.reset_powerpoint, bg="#ffcccc").pack(side="left", fill="x", expand=True, padx=2)

        # 2. Template & Mode
        tk.Label(left_frame, text="Template & Mode:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 2))
        
        # Checkbox
        chk_wed = tk.Checkbutton(left_frame, text="Wednesday Mode", 
                                 variable=self.is_wednesday_var, command=self.toggle_mode)
        chk_wed.pack(anchor="w", pady=(0, 2))

        frame_tpl = tk.Frame(left_frame)
        frame_tpl.pack(fill="x", pady=(0, 10))
        tk.Entry(frame_tpl, textvariable=self.template_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame_tpl, text="Browse", command=self.browse_template).pack(side="right", padx=2)

        # 3. Output File
        tk.Label(left_frame, text="Output File:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 2))
        frame_out = tk.Frame(left_frame)
        frame_out.pack(fill="x", pady=(0, 10))
        tk.Entry(frame_out, textvariable=self.output_path_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame_out, text="Browse", command=self.browse_output).pack(side="right", padx=2)

        # 4. Songs Before Sermon
        tk.Label(left_frame, text="Songs Before Sermon:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 2))
        frame_before = tk.Frame(left_frame)
        frame_before.pack(fill="both", expand=True, pady=(0, 5))
        
        sb_before = tk.Scrollbar(frame_before)
        sb_before.pack(side="right", fill="y")
        
        self.list_before = tk.Listbox(frame_before, selectmode=tk.EXTENDED, yscrollcommand=sb_before.set, height=5)
        self.list_before.pack(side="left", fill="both", expand=True)
        sb_before.config(command=self.list_before.yview)
        
        # Controls Before
        frame_btns_before = tk.Frame(left_frame)
        frame_btns_before.pack(fill="x", pady=(0, 10))
        tk.Button(frame_btns_before, text="\u2191", width=3, command=lambda: self.move_up(self.list_before)).pack(side="left", padx=2)
        tk.Button(frame_btns_before, text="\u2193", width=3, command=lambda: self.move_down(self.list_before)).pack(side="left", padx=2)
        tk.Button(frame_btns_before, text="Del", width=4, command=lambda: self.delete_song(self.list_before)).pack(side="left", padx=2)
        tk.Button(frame_btns_before, text="Clear", width=5, command=lambda: self.clear_all(self.list_before)).pack(side="left", padx=2)
        tk.Button(frame_btns_before, text="To After \u2193", command=self.move_to_after).pack(side="right", padx=2)

        # 5. Songs After Sermon
        tk.Label(left_frame, text="Songs After Sermon:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 2))
        frame_after = tk.Frame(left_frame)
        frame_after.pack(fill="both", expand=True, pady=(0, 5))
        
        sb_after = tk.Scrollbar(frame_after)
        sb_after.pack(side="right", fill="y")
        
        self.list_after = tk.Listbox(frame_after, selectmode=tk.EXTENDED, yscrollcommand=sb_after.set, height=5)
        self.list_after.pack(side="left", fill="both", expand=True)
        sb_after.config(command=self.list_after.yview)

        # Controls After
        frame_btns_after = tk.Frame(left_frame)
        frame_btns_after.pack(fill="x", pady=(0, 0))
        tk.Button(frame_btns_after, text="\u2191", width=3, command=lambda: self.move_up(self.list_after)).pack(side="left", padx=2)
        tk.Button(frame_btns_after, text="\u2193", width=3, command=lambda: self.move_down(self.list_after)).pack(side="left", padx=2)
        tk.Button(frame_btns_after, text="Del", width=4, command=lambda: self.delete_song(self.list_after)).pack(side="left", padx=2)
        tk.Button(frame_btns_after, text="Clear", width=5, command=lambda: self.clear_all(self.list_after)).pack(side="left", padx=2)
        tk.Button(frame_btns_after, text="\u2191 To Before", command=self.move_to_before).pack(side="right", padx=2)


        # === RIGHT FRAME CONTENT ===

        # 1. Worship Title
        tk.Label(right_frame, text="Worship Title (Slide 1):").pack(anchor="w", pady=(0, 2))
        entry_worship = tk.Entry(right_frame, textvariable=self.worship_title_var)
        entry_worship.pack(fill="x", pady=(0, 10))

        # 2. Sermon Title
        tk.Label(right_frame, text="Sermon Title (Slide 6 - Wed Only):").pack(anchor="w", pady=(0, 2))
        entry_sermon = tk.Entry(right_frame, textvariable=self.sermon_title_var)
        entry_sermon.pack(fill="x", pady=(0, 10))

        # 3. Bible Chapter
        tk.Label(right_frame, text="Bible Chapter/Verse (All Slides):").pack(anchor="w", pady=(0, 2))
        entry_title = tk.Entry(right_frame, textvariable=self.bible_title_var)
        entry_title.pack(fill="x", pady=(0, 10))

        # 4. Bible Body
        tk.Label(right_frame, text="Bible Body (Slide 5) - Use '/' to split:", font=("Arial", 9)).pack(anchor="w", pady=(0, 2))
        # Enable Undo here
        self.bible_body_text = scrolledtext.ScrolledText(right_frame, height=20, undo=True)
        self.bible_body_text.pack(fill="both", expand=True, pady=(0, 10))
        self.bible_body_text.insert("1.0", "")
        
        # Tab Binding
        def focus_next_widget(event):
            event.widget.tk_focusNext().focus()
            return "break"
        self.bible_body_text.bind("<Tab>", focus_next_widget)

        # 5. Generate Button
        btn_gen = tk.Button(right_frame, text="Generate PPT", command=self.start_generation, bg="lightblue", font=("Arial", 12, "bold"), height=2)
        btn_gen.pack(fill="x", pady=(0, 0))

    def toggle_mode(self):
        """Switches template filename, output directory, and output filename based on checkbox"""
        today = datetime.date.today()
        
        if self.is_wednesday_var.get():
            # Wednesday Mode
            # 1. Template Path
            # Explicitly set to the requested Wednesday path
            new_tpl_path = r"D:\02. 열띰!\02. 교회\04. 수요기도회 PPT\wednesday.pptx"
            
            # 2. Date Calculation (Next Wednesday)
            target_weekday = 2 # Wednesday
            days_ahead = target_weekday - today.weekday()
            if days_ahead <= 0: # Target day already happened this week
                days_ahead += 7
            next_date = today + datetime.timedelta(days_ahead)
            
            # 3. Output Path
            base_output_dir = r"D:\02. 열띰!\02. 교회\04. 수요기도회 PPT"
            
            filename = f"{next_date.strftime('%Y년 %m월 %d일')} 수요기도회.pptx"
            
        else:
            # Friday Mode (Default)
            # 1. Template Path
            new_tpl_path = r"D:\02. 열띰!\02. 교회\03. 금요기도회 PPT\friday.pptx"
            
            # 2. Date Calculation (Next Friday)
            target_weekday = 4 # Friday
            days_ahead = target_weekday - today.weekday()
            if days_ahead <= 0:
                days_ahead += 7
            next_date = today + datetime.timedelta(days_ahead)
            
            # 3. Output Path
            base_output_dir = r"D:\02. 열띰!\02. 교회\03. 금요기도회 PPT"
            
            filename = f"{next_date.strftime('%Y년 %m월 %d일')} 금요기도회.pptx"

        # Apply changes
        self.template_path_var.set(new_tpl_path)
        
        # Output
        new_output_path = os.path.join(base_output_dir, filename)
        self.output_path_var.set(new_output_path)

    def browse_ppt_dir(self):
        # Users want to see files to verify they are in the right folder.
        # So we use askopenfilename but strictly to get the directory.
        initial = self.ppt_dir_var.get()
        if not os.path.exists(initial):
            initial = os.getcwd()
            
        paths = filedialog.askopenfilenames(
            title="Select song files (Directory will be selected)",
            initialdir=initial,
            filetypes=[("Song Files", "*.pptx;*.ppt"), ("All Files", "*.*")]
        )
        
        if paths:
            # multiple files might be selected, just take the first one to get the directory
            path = paths[0]
            directory = os.path.dirname(path)
            self.ppt_dir_var.set(os.path.normpath(directory))
            self.populate_song_lists()

    def reset_powerpoint(self):
        """Force kills PowerPoint processes to fix lock issues"""
        if messagebox.askyesno("Confirm", "This will close ALL PowerPoint windows. Continue?"):
            try:
                os.system("taskkill /IM POWERPNT.EXE /F")
                messagebox.showinfo("Success", "PowerPoint has been reset.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to reset PowerPoint: {e}")

    def clear_all_lists(self):
        self.list_before.delete(0, tk.END)
        self.list_after.delete(0, tk.END)

    def populate_song_lists(self):
        ppt_dir = self.ppt_dir_var.get()
        self.list_before.delete(0, tk.END)
        self.list_after.delete(0, tk.END)
        
        if os.path.exists(ppt_dir):
            # STRICTLY filter only .pptx (case insensitive)
            files = [f for f in os.listdir(ppt_dir) if f.lower().endswith('.pptx') and not f.startswith("~$")]
            files.sort()
            
            # Default split: First 2 to Before, Rest to After
            for i, f in enumerate(files):
                if i < 2:
                    self.list_before.insert(tk.END, f)
                else:
                    self.list_after.insert(tk.END, f)

    def move_up(self, listbox):
        try:
            selection = listbox.curselection()
            if not selection:
                return
            
            # Convert to list and sort
            selection = sorted(list(selection))
            
            # If any item is already at the top, we can't move the block up if it's contiguous with top
            # But standard behavior is to move all movable items up.
            # Let's iterate from top to bottom of selection
            
            for index in selection:
                if index > 0:
                    text = listbox.get(index)
                    listbox.delete(index)
                    listbox.insert(index - 1, text)
                    listbox.selection_set(index - 1)
        except Exception:
            pass

    def move_down(self, listbox):
        try:
            selection = listbox.curselection()
            if not selection:
                return
            
            # Convert to list and sort descending
            selection = sorted(list(selection), reverse=True)
            
            for index in selection:
                if index < listbox.size() - 1:
                    text = listbox.get(index)
                    listbox.delete(index)
                    listbox.insert(index + 1, text)
                    listbox.selection_set(index + 1)
        except Exception:
            pass

    def delete_song(self, listbox):
        try:
            selection = listbox.curselection()
            if not selection:
                return
            
            # Delete in reverse order to maintain indices
            for index in sorted(list(selection), reverse=True):
                listbox.delete(index)
        except Exception:
            pass

    def clear_all(self, listbox):
        listbox.delete(0, tk.END)

    def move_to_after(self):
        try:
            selection = self.list_before.curselection()
            if not selection:
                return
            
            # Get items
            items = [self.list_before.get(i) for i in selection]
            
            # Delete from source (reverse order)
            for index in sorted(list(selection), reverse=True):
                self.list_before.delete(index)
                
            # Insert into target (at top, in order)
            # To keep their relative order, insert them in reverse order at index 0?
            # No, if we have [A, B] selected, we want [A, B] at top of After.
            # So insert B at 0, then A at 0? No, that gives [A, B].
            # Wait: Insert A at 0 -> [A, ...]. Insert B at 0 -> [B, A, ...]. Reversed.
            # So we should insert in reverse order of appearance in 'items' to preserve order at top.
            
            for item in reversed(items):
                self.list_after.insert(0, item)
                self.list_after.selection_set(0)
                
        except Exception:
            pass

    def move_to_before(self):
        try:
            selection = self.list_after.curselection()
            if not selection:
                return
            
            # Get items
            items = [self.list_after.get(i) for i in selection]
            
            # Delete from source (reverse order)
            for index in sorted(list(selection), reverse=True):
                self.list_after.delete(index)
                
            # Insert into target (at bottom)
            for item in items:
                self.list_before.insert(tk.END, item)
                self.list_before.selection_set(tk.END)
                
        except Exception:
            pass

    def browse_template(self):
        initial = os.path.dirname(self.template_path_var.get())
        if not os.path.exists(initial):
            initial = os.getcwd()
            
        path = filedialog.askopenfilename(initialdir=initial, filetypes=[("PowerPoint Files", "*.pptx;*.ppt")])
        if path:
            self.template_path_var.set(os.path.normpath(path))

    def browse_output(self):
        initial = os.path.dirname(self.output_path_var.get())
        if not os.path.exists(initial):
            initial = os.getcwd()
            
        # Suggest the current filename
        initial_file = os.path.basename(self.output_path_var.get())
        
        path = filedialog.asksaveasfilename(initialdir=initial, initialfile=initial_file, filetypes=[("PowerPoint Files", "*.pptx")])
        if path:
            if not path.lower().endswith(".pptx"):
                path += ".pptx"
            self.output_path_var.set(os.path.normpath(path))

    def start_generation(self):
        # Gather inputs
        ppt_dir = self.ppt_dir_var.get()
        template_path = self.template_path_var.get()
        output_path = self.output_path_var.get()
        bible_title = self.bible_title_var.get()
        worship_title = self.worship_title_var.get()
        sermon_title = self.sermon_title_var.get() if self.is_wednesday_var.get() else ""
        # bible_range = self.bible_range_var.get() # Removed
        bible_body = self.bible_body_text.get("1.0", "end-1c")
        
        # Get songs from listboxes
        files_before = self.list_before.get(0, tk.END)
        files_after = self.list_after.get(0, tk.END)
        
        songs_before = [os.path.join(ppt_dir, f) for f in files_before]
        songs_after = [os.path.join(ppt_dir, f) for f in files_after]
        
        # Run in a separate thread
        # Pass bible_title for both title and range arguments
        threading.Thread(target=self.run_logic, args=(songs_before, songs_after, template_path, output_path, worship_title, bible_title, bible_title, bible_body, sermon_title)).start()

    def run_logic(self, songs_before, songs_after, template_path, output_path, worship_title, bible_title, bible_range, bible_body, sermon_title=""):
        pythoncom.CoInitialize()
        try:
            errors, warnings = generate_ppt(songs_before, songs_after, template_path, output_path, worship_title, bible_title, bible_range, bible_body, sermon_title)
            
            msg = ""
            if errors:
                msg += "Errors occurred:\n" + "\n".join([f"- {e}" for e in errors]) + "\n\n"
            
            if warnings:
                msg += "Warnings:\n" + "\n".join([f"- {w}" for w in warnings]) + "\n\n"
                
            if not errors:
                msg += f"Presentation generated successfully!\nSaved to: {output_path}"
                if warnings:
                    messagebox.showwarning("Completed with Warnings", msg)
                else:
                    messagebox.showinfo("Success", msg)
                
                # Auto Open File
                try:
                    os.startfile(output_path)
                except Exception as e:
                    print(f"Could not auto-open file: {e}")

            else:
                messagebox.showerror("Error", msg)
                
        except Exception as e:
            messagebox.showerror("Error", f"An critical error occurred:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
