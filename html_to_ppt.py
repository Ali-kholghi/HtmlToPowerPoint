import io
import os
import time
import pathlib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Import optional image processing to prevent layout distortion
try:
    from PIL import Image
    has_pil = True
except ImportError:
    has_pil = False

# Import automation libraries
from pptx import Presentation
from pptx.util import Inches
from playwright.sync_api import sync_playwright


def convert_html_to_pptx(html_path, ppt_path, mode, status_callback):
    """
    Handles the browser rendering, screenshot capture, and PowerPoint construction.
    """
    try:
        status_callback("Starting web browser engine...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            
            # Use standard 16:9 widescreen dimensions for the browser viewport
            viewport_width = 1920
            viewport_height = 1080
            
            context = browser.new_context(viewport={"width": viewport_width, "height": viewport_height},
                device_scale_factor=2
                )
            page = context.new_page()
            
            status_callback("Loading HTML file...")
            # Resolve the absolute local path to a file URI
            file_url = pathlib.Path(html_path).absolute().as_uri()
            page.goto(file_url, wait_until="networkidle")
            
            # Allow dynamic content or local fonts a moment to load
            time.sleep(1.0)
            
            screenshots = []
            
            if mode == "scroll":
                status_callback("Analyzing page height...")
                # Scroll to the bottom first to trigger any lazy-loaded content
                total_height = page.evaluate("() => document.body.scrollHeight")
                current = 0
                while current < total_height:
                    page.evaluate(f"window.scrollTo(0, {current})")
                    time.sleep(0.1)
                    current += 500
                    total_height = page.evaluate("() => document.body.scrollHeight")
                
                # Scroll back to top
                page.evaluate("window.scrollTo(0, 0)")
                time.sleep(0.5)
                
                # Capture screenshots viewport-by-viewport
                current_scroll = 0
                page_idx = 1
                while current_scroll < total_height:
                    status_callback(f"Capturing vertical section {page_idx}...")
                    page.evaluate(f"window.scrollTo(0, {current_scroll})")
                    time.sleep(0.5)  # Let any scrolling animations settle
                    
                    img_bytes = page.screenshot(full_page=False)
                    screenshots.append(img_bytes)
                    
                    current_scroll += viewport_height
                    page_idx += 1
                    
            elif mode == "section":
                status_callback("Scanning for <section> tags...")
                sections = page.locator("section")
                count = sections.count()
                
                if count == 0:
                    status_callback("No <section> elements found. Falling back to Vertical Scroll...")
                    # Fallback to scroll logic if no sections are present
                    total_height = page.evaluate("() => document.body.scrollHeight")
                    current_scroll = 0
                    while current_scroll < total_height:
                        page.evaluate(f"window.scrollTo(0, {current_scroll})")
                        time.sleep(0.5)
                        img_bytes = page.screenshot(full_page=False)
                        screenshots.append(img_bytes)
                        current_scroll += viewport_height
                else:
                    for i in range(count):
                        status_callback(f"Capturing section {i+1} of {count}...")
                        section = sections.nth(i)
                        section.scroll_into_view_if_needed()
                        time.sleep(0.5)
                        img_bytes = section.screenshot()
                        screenshots.append(img_bytes)
            
            browser.close()
            
            if not screenshots:
                raise Exception("No visual content was captured from the HTML file.")
            
            status_callback("Assembling PowerPoint slides...")
            prs = Presentation()
            
            # Configure presentation to standard widescreen (16:9)
            slide_width_in = 13.333
            slide_height_in = 7.5
            prs.slide_width = Inches(slide_width_in)
            prs.slide_height = Inches(slide_height_in)
            
            blank_layout = prs.slide_layouts[6]  # Blank layout slide
            
            for i, img_data in enumerate(screenshots):
                status_callback(f"Adding slide {i+1} of {len(screenshots)}...")
                slide = prs.slides.add_slide(blank_layout)
                image_stream = io.BytesIO(img_data)
                
                # If Pillow is installed, calculate aspect ratio to center & fit without stretching
                if has_pil:
                    img = Image.open(image_stream)
                    img_width, img_height = img.size
                    
                    slide_ratio = slide_width_in / slide_height_in
                    img_ratio = img_width / img_height
                    
                    if img_ratio > slide_ratio:
                        # Image is wider than slide ratio: fit to width, center vertically
                        display_width = slide_width_in
                        display_height = slide_width_in / img_ratio
                        left = 0
                        top = (slide_height_in - display_height) / 2
                    else:
                        # Image is taller than slide ratio: fit to height, center horizontally
                        display_height = slide_height_in
                        display_width = slide_height_in * img_ratio
                        left = (slide_width_in - display_width) / 2
                        top = 0
                        
                    image_stream.seek(0)  # Reset buffer position
                    slide.shapes.add_picture(
                        image_stream, 
                        Inches(left), 
                        Inches(top), 
                        width=Inches(display_width), 
                        height=Inches(display_height)
                    )
                else:
                    # Fallback to full stretching if Pillow is missing
                    slide.shapes.add_picture(
                        image_stream, 
                        Inches(0), 
                        Inches(0), 
                        width=prs.slide_width, 
                        height=prs.slide_height
                    )
            
            status_callback("Writing PowerPoint file to disk...")
            prs.save(ppt_path)
            status_callback("Conversion complete.")
            
    except Exception as e:
        status_callback(f"Error: {str(e)}")
        raise e


# --- GUI Implementation ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("HTML to PowerPoint Converter")
        self.root.geometry("550x300")
        self.root.resizable(False, False)
        
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        
        # UI Elements Container
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. HTML File Selection
        ttk.Label(main_frame, text="Select HTML File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.html_entry = ttk.Entry(main_frame, width=40)
        self.html_entry.grid(row=0, column=1, pady=5, padx=5)
        self.browse_html_btn = ttk.Button(main_frame, text="Browse", command=self.browse_html)
        self.browse_html_btn.grid(row=0, column=2, pady=5)
        
        # 2. PowerPoint Output Location
        ttk.Label(main_frame, text="Save PowerPoint As:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.ppt_entry = ttk.Entry(main_frame, width=40)
        self.ppt_entry.grid(row=1, column=1, pady=5, padx=5)
        self.browse_ppt_btn = ttk.Button(main_frame, text="Browse", command=self.browse_ppt)
        self.browse_ppt_btn.grid(row=1, column=2, pady=5)
        
        # 3. Capture Mode Choice
        ttk.Label(main_frame, text="Capture Mode:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.mode_var = tk.StringVar(value="scroll")
        
        mode_frame = ttk.Frame(main_frame)
        mode_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=10)
        
        self.scroll_radio = ttk.Radiobutton(
            mode_frame, text="Vertical Scroll (for standard web pages)", 
            variable=self.mode_var, value="scroll"
        )
        self.scroll_radio.pack(anchor=tk.W)
        
        self.section_radio = ttk.Radiobutton(
            mode_frame, text="Slide/Section tags (for discrete slideshow containers)", 
            variable=self.mode_var, value="section"
        )
        self.section_radio.pack(anchor=tk.W)
        
        # Divider
        ttk.Separator(main_frame, orient='horizontal').grid(row=3, column=0, columnspan=3, sticky='ew', pady=10)
        
        # 4. Status Indicator
        self.status_label = ttk.Label(main_frame, text="Ready", font=("Arial", 9, "italic"), foreground="gray")
        self.status_label.grid(row=4, column=0, columnspan=2, sticky=tk.W)
        
        # 5. Convert Button
        self.convert_btn = ttk.Button(main_frame, text="Convert to PPTX", command=self.start_conversion)
        self.convert_btn.grid(row=4, column=2, sticky=tk.E)

    def browse_html(self):
        file_path = filedialog.askopenfilename(
            title="Select HTML File",
            filetypes=[("HTML Files", "*.html *.htm"), ("All Files", "*.*")]
        )
        if file_path:
            self.html_entry.delete(0, tk.END)
            self.html_entry.insert(0, file_path)
            
            # Automatically suggest an output file name based on the HTML file name
            suggested_ppt = os.path.splitext(file_path)[0] + ".pptx"
            self.ppt_entry.delete(0, tk.END)
            self.ppt_entry.insert(0, suggested_ppt)

    def browse_ppt(self):
        file_path = filedialog.asksaveasfilename(
            title="Save PowerPoint As",
            defaultextension=".pptx",
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")]
        )
        if file_path:
            self.ppt_entry.delete(0, tk.END)
            self.ppt_entry.insert(0, file_path)

    def update_status(self, text):
        self.status_label.config(text=text)
        self.root.update_idletasks()

    def start_conversion(self):
        html_file = self.html_entry.get()
        ppt_file = self.ppt_entry.get()
        mode = self.mode_var.get()
        
        if not html_file or not os.path.exists(html_file):
            messagebox.showerror("Error", "Please select a valid HTML file.")
            return
        if not ppt_file:
            messagebox.showerror("Error", "Please specify an output location.")
            return
            
        # Disable inputs to prevent interaction during conversion
        self.convert_btn.config(state="disabled")
        self.browse_html_btn.config(state="disabled")
        self.browse_ppt_btn.config(state="disabled")
        self.scroll_radio.config(state="disabled")
        self.section_radio.config(state="disabled")
        
        # Run conversion in a separate thread so the Tkinter UI doesn't freeze
        def run():
            try:
                convert_html_to_pptx(html_file, ppt_file, mode, self.update_status)
                messagebox.showinfo("Success", f"Presentation saved successfully:\n{ppt_file}")
            except Exception as e:
                messagebox.showerror("Conversion Failed", f"An error occurred:\n{str(e)}")
            finally:
                # Re-enable inputs
                self.convert_btn.config(state="normal")
                self.browse_html_btn.config(state="normal")
                self.browse_ppt_btn.config(state="normal")
                self.scroll_radio.config(state="normal")
                self.section_radio.config(state="normal")
                self.update_status("Ready")
                
        threading.Thread(target=run, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()