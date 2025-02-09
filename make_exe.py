"""
PDFMute - Version 2

Changes from v1 to v2:
1. Professional Layout and UI:
   - We now use a two-column layout:
     a. Left panel for user controls (Algorithm, Color Choice, 'About' Button).
     b. Right panel for PDF/Docx loading, progress bar, and preview.
   - Center panel for the "Go" button and animated GIF.
   - Overall styling uses custom ttk styles and a consistent color theme.

2. Improved Code Organization:
   - Clear separation of responsibilities:
     - "GUI Setup" code is grouped together.
     - "Logic / Worker" functions remain external.
     - "Conversion + Red Removal" flows are integrated in a simpler manner.

3. Enhanced Feedback:
   - Status label shows "Ready", "Working...", "Done", or "Error" states more clearly.
   - The animated GIF is displayed near the progress bar for better visibility.

4. Additional Minor Tweaks:
   - More robust docx2pdf usage (try/except).
   - On finishing tasks, UI resets properly.
   - Code refactoring and improved naming.
"""

import io
import os
import sys
import time
import threading
import webbrowser
import uuid
from pathlib import Path

import fitz  # PyMuPDF
import numpy as np
from PIL import Image, ImageTk
from tkinter import Tk, Frame, Canvas, Label, Button, Toplevel, filedialog, StringVar
from tkinter import ttk

# Torch is only used in the remove_red_pixels_gpu function
from torch import tensor, device
from torch.cuda import is_available as cuda_is_available

import tempfile
import pythoncom
import win32com.client
import logging

os.chdir(sys._MEIPASS)




# ------------- RED REMOVAL LOGIC (CPU) ------------- #
def remove_red_pixels(input_pdf, output_pdf, progress_callback, color):
    """
    Remove red (or pink) pixels from a PDF by converting them to white/black.
    Uses CPU-based approach with PyMuPDF + PIL.
    """
    doc = fitz.open(input_pdf)
    new_doc = fitz.Document()
    colorrgb = (255, 255, 255) if color == 'white' else (0, 0, 0)

    # Pre-defined target colors (e.g. pink-ish or near red) with delta tolerance
    target_colors = [
        ((224, 202, 202), 5),
        ((218, 203, 204), 5),
        ((229, 220, 220), 5),
        ((230, 212, 220), 5),
        ((215, 190, 197), 5),
        ((254, 251, 249), 5),
        ((197, 193, 194), 5),
        ((197, 195, 196), 5),
        ((198, 194, 195), 5),
        ((200, 192, 195), 5),
        ((200, 195, 195), 5),
        ((200, 196, 195), 5),
        ((201, 193, 194), 5),
        ((202, 185, 187), 5),
        ((203, 199, 198), 5),
        ((205, 203, 204), 5),
    ]

    total_pages = len(doc)
    for page_number, page in enumerate(doc):
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        pixels = img.load()

        # Pass 1: Intense red + target colors
        for y in range(img.height):
            for x in range(img.width):
                r, g, b = pixels[x, y]

                # Condition for intense red
                if r > 150 and r > g * 1.2 and r > b * 1.5 and (r + g + b) > 100:
                    pixels[x, y] = colorrgb
                else:
                    for ccheck, delta in target_colors:
                        if abs(r - ccheck[0]) <= delta and \
                           abs(g - ccheck[1]) <= delta and \
                           abs(b - ccheck[2]) <= delta:
                            pixels[x, y] = colorrgb
                            break

        # Pass 2: Additional pass for leftover reds if color is white
        if color == 'white':
            for y in range(img.height):
                for x in range(img.width):
                    r, g, b = pixels[x, y]
                    if r > g and r > b and r > 180:
                        pixels[x, y] = (255, 255, 255)

        # Save to PDF
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=100)
        img_byte_arr.seek(0)
        new_page = new_doc.new_page(width=pix.width, height=pix.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr.read())

        # Update progress
        progress_callback(((page_number + 1) / total_pages) * 100)

    # Save the processed PDF
    new_doc.save(output_pdf)
    doc.close()
    new_doc.close()


# ------------- RED REMOVAL LOGIC (GPU) ------------- #
def remove_red_pixels_gpu(input_pdf, output_pdf, progress_callback, color):
    """
    Remove red (or pink) pixels from a PDF by converting them to white/black.
    Uses GPU-based approach via PyTorch Tensors.
    """
    doc = fitz.open(input_pdf)
    new_doc = fitz.Document()
    device_gpu = device("cuda" if cuda_is_available() else "cpu")
    colorrgb = (255.0, 255.0, 255.0) if color == 'white' else (0.0, 0.0, 0.0)

    total_pages = len(doc)
    for page_number, page in enumerate(doc):
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Convert to tensor on GPU/CPU
        img_tensor = tensor(np.array(img), device=device_gpu).permute(2, 0, 1).float()

        # Pass 1: Strong red
        red_mask = (img_tensor[0] > 150) & (img_tensor[0] > img_tensor[1] * 1.2) & (img_tensor[0] > img_tensor[2] * 1.5)
        img_tensor[:, red_mask] = tensor(colorrgb, device=device_gpu).view(3, 1)

        # Pass 2: Lighter pinkish
        pink_mask = (img_tensor[0] > 140) & (img_tensor[0] > img_tensor[1] * 1.1) & (img_tensor[0] > img_tensor[2] * 1.2)
        img_tensor[:, pink_mask] = tensor(colorrgb, device=device_gpu).view(3, 1)

        # Convert back to PIL
        final_img = Image.fromarray(img_tensor.byte().permute(1, 2, 0).cpu().numpy())
        img_byte_arr = io.BytesIO()
        final_img.save(img_byte_arr, format='JPEG', quality=100)
        img_byte_arr.seek(0)

        # Insert into new PDF
        new_page = new_doc.new_page(width=pix.width, height=pix.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr.read())

        # Update progress
        progress_callback(((page_number + 1) / total_pages) * 100)

    new_doc.save(output_pdf)
    doc.close()
    new_doc.close()


# ------------- DOCX / DOC -> PDF CONVERSION ------------- #
def convert_docx_to_pdf(input_file, output_file=None):
    """
    Convert DOCX to PDF, either to a temporary file or specified output path
    Returns the path to the converted PDF file
    """
    try:
        converter = DocxConverter()
        temp_pdf = converter.convert(input_file)

        if output_file:
            import shutil
            shutil.copy2(temp_pdf, output_file)
            return output_file

        return temp_pdf

    except Exception as e:
        raise Exception(f"Conversion failed: {str(e)}")

# ------------- PDF PREVIEW ------------- #
def preview_pdf_page(pdf_path):
    """
    Returns a PIL.Image of the first page of the PDF (at a reduced DPI).
    """
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=100)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


# ------------- MAIN APPLICATION CLASS ------------- #
class PDFMuteApp(Tk):
    def __init__(self):
        super().__init__()

        # Add signal handlers for proper cleanup
        import signal
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)
        self.title("PDFMute v2 - Professional Edition")
        self.geometry("1250x850")
        self.minsize(1000, 700)

        # State tracking
        self.input_file = None
        self.output_file = None
        self.running = False
        self.dot_count = 0
        self.threads = []

        # Default selection
        self.algorithm_choice = StringVar(value="CPU")
        self.color_choice = StringVar(value="white")

        # Color palette
        self.bg_color = "#EFEFEF"         # Light gray background
        self.primary_color = "#006666"    # Primary teal
        self.accent_color = "#009999"     # Lighter teal accent
        self.highlight_color = "#00CC99"  # Highlight color (greenish-teal)
        self.font_color = "#333333"       # Dark text
        self.red_color = "#FF4D4D"

        # Configure the style
        self._configure_styles()

        # Build main layout
        self._build_layout()
        self.docx_converter = DocxConverter()

    def _signal_handler(self, signum, frame):
        """Handle termination signals"""
        self.on_closing()
    # ------------------- STYLES ------------------- #
    def _configure_styles(self):
        """
        Configure ttk styles for a professional look.
        """
        style = ttk.Style(self)
        style.theme_use("clam")

        # General label
        style.configure("TLabel", background=self.bg_color, foreground=self.font_color, font=("Helvetica", 12))

        # Frame
        style.configure("TFrame", background=self.bg_color)

        # Button
        style.configure("TButton",
                        background=self.primary_color,
                        foreground="#FFFFFF",
                        font=("Helvetica", 12, "bold"),
                        borderwidth=0,
                        padding=5)

        style.map("TButton",
                  background=[("active", self.highlight_color),
                              ("disabled", "#A0A0A0")])

        # Progress bar
        style.configure("Green.Horizontal.TProgressbar",
                        troughcolor="#FFFFFF",
                        background=self.highlight_color,
                        bordercolor=self.bg_color,
                        lightcolor=self.highlight_color,
                        darkcolor=self.highlight_color)

    # ------------------- LAYOUT ------------------- #
    def _build_layout(self):
        """
        Build the main layout:
        Left column for controls, center for 'Go' & GIF, right for preview & status.
        """
        self.configure(bg=self.bg_color)

        # Main container frames
        container = ttk.Frame(self, style="TFrame")
        container.pack(fill="both", expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(container, style="TFrame")
        left_frame.pack(side="left", fill="y", padx=(0, 10))

        center_frame = ttk.Frame(container, style="TFrame")
        center_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_frame = ttk.Frame(container, style="TFrame")
        right_frame.pack(side="right", fill="y")

        # ---- LEFT FRAME: Algorithm & Color selection ----
        left_label = ttk.Label(left_frame, text="Algorithm & Color", style="TLabel")
        left_label.pack(anchor="nw", pady=(0, 10))

        # Algorithm Options
        alg_frame = ttk.LabelFrame(left_frame, text="Pick an Algorithm:")
        alg_frame.pack(fill="x", pady=(0, 10))

        cpu_button = ttk.Radiobutton(alg_frame, text="CPU - Slower, Higher Quality",
                                     value="CPU", variable=self.algorithm_choice)
        gpu_button = ttk.Radiobutton(alg_frame, text="GPU - Faster, Lower Quality",
                                     value="GPU", variable=self.algorithm_choice)
        cpu_button.pack(anchor="w", pady=2)
        gpu_button.pack(anchor="w", pady=2)

        # Color Options
        color_frame = ttk.LabelFrame(left_frame, text="Turn Red To:")
        color_frame.pack(fill="x", pady=(0, 10))

        white_button = ttk.Radiobutton(color_frame, text="White", value="white",
                                       variable=self.color_choice)
        black_button = ttk.Radiobutton(color_frame, text="Black", value="black",
                                       variable=self.color_choice)
        white_button.pack(anchor="w", pady=2)
        black_button.pack(anchor="w", pady=2)

        # About Button
        about_btn = ttk.Button(left_frame, text="About", command=self._show_about)
        about_btn.pack(anchor="nw", pady=(10, 0))
        # ---- CENTER FRAME: "Load / Save / Go" + Activity Indicator & GIF ----
        center_controls = ttk.Frame(center_frame, style="TFrame")
        center_controls.pack(anchor="n", pady=10, fill="x")

        # Row 1: Load / Save
        load_button = ttk.Button(center_controls, text="Load PDF/DOCX", command=self._on_load_click)
        load_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.save_button = ttk.Button(center_controls, text="Save as PDF",
                                      command=self._on_save_click, state="disabled")
        self.save_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Row 2: Go Button
        self.go_button = ttk.Button(center_controls, text="Go", command=self._on_go_click, state="disabled")
        self.go_button.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")

        center_controls.columnconfigure(0, weight=1)
        center_controls.columnconfigure(1, weight=1)

        # Row 3: Status + Progress + GIF
        status_frame = ttk.Frame(center_frame, style="TFrame")
        status_frame.pack(anchor="n", fill="x")

        self.status_label = ttk.Label(status_frame, text="Ready", style="TLabel")
        self.status_label.pack(side="left", padx=5)

        self.progress_bar = ttk.Progressbar(status_frame, style="Green.Horizontal.TProgressbar",
                                            orient="horizontal",
                                            length=250,
                                            mode="determinate")
        self.progress_bar.pack(side="left", padx=5)

        # Animated GIF container
        self.gif_label = Label(status_frame, bg=self.bg_color)
        self.gif_label.pack(side="left", padx=5)
        self.gif_frames = []
        self.gif_running = False
        self.gif_index = 0

        # ---- RIGHT FRAME: PDF Preview Canvas ----
        preview_lbl = ttk.Label(right_frame, text="Preview", style="TLabel")
        preview_lbl.pack(anchor="nw", pady=(0, 5))

        self.preview_canvas = Canvas(right_frame, bg="#CCCCCC", width=595, height=842)
        self.preview_canvas.pack(pady=5, padx=5)
    # ------------------- EVENT HANDLERS ------------------- #
    def _on_load_click(self):
        """
        Triggered when user clicks "Load PDF/DOCX".
        """
        input_path = filedialog.askopenfilename(
            filetypes=[
                ("PDF / Word files", "*.pdf *.doc *.docx"),
                ("PDF files", "*.pdf"),
                ("Word files", "*.doc *.docx"),
            ]
        )
        if not input_path:
            return

        try:
            # שמירת הנתיב המקורי
            self.original_file_path = input_path

            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.pdf':
                self.input_file = input_path
            else:
                # Convert DOCX/DOC to temporary PDF in a separate thread
                self.status_label.config(text="Converting document...")
                conversion_thread = threading.Thread(target=self._convert_docx_thread, args=(input_path,))
                conversion_thread.start()
                self.threads.append(conversion_thread)

        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")


    def _convert_docx_thread(self, input_path):
        try:
            self.input_file = self.docx_converter.convert(input_path)
            self._preview_pdf(self.input_file)
            self.save_button.config(state="normal")
            self.status_label.config(text="Loaded successfully")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
    def _on_save_click(self):
        """
        Triggered when user clicks "Save as PDF".
        """
        if not self.input_file:
            self.status_label.config(text="No File Loaded!")
            return

        # Get original file name without extension
        original_name = os.path.splitext(os.path.basename(self.original_file_path))[0]
        default_output = f"{original_name}_MuteRed.pdf"

        out_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=default_output
        )
        if out_path:
            self.output_file = out_path
            self.go_button.config(state="normal")
            self.status_label.config(text="Ready to process")
        else:
            self.status_label.config(text="Save was canceled")
    def _on_go_click(self):
        """
        Triggered when the user clicks "Go".
        Starts the red removal process with a background thread.
        """
        if not (self.input_file and self.output_file):
            self.status_label.config(text="No valid input/output")
            return

        self.status_label.config(text="Working...")
        self.running = True
        self.go_button.config(state="disabled")
        self.save_button.config(state="disabled")
        self._start_gif_animation("busy.gif")

        # Start thread for red removal
        thread_target = self._process_thread_cpu if self.algorithm_choice.get() == "CPU" else self._process_thread_gpu
        process_thread = threading.Thread(target=thread_target, args=(self.input_file, self.output_file))
        process_thread.daemon = True
        process_thread.start()
        self.threads.append(process_thread)

        # Start a separate thread to update the button text with "Working..."
        text_updater = threading.Thread(target=self._update_go_button_text, daemon=True)
        text_updater.start()

    # ------------------- RED REMOVAL THREAD WRAPPERS ------------------- #
    def _process_thread_cpu(self, input_pdf, output_pdf):
        """
        Thread wrapper for CPU-based red removal.
        """
        try:
            remove_red_pixels(input_pdf, output_pdf, self._update_progress, self.color_choice.get())
            self.status_label.config(text="Done!")
        except Exception as e:
            self.status_label.config(text=f"Error: {e}")
        finally:
            self._cleanup_after_processing()

    def _process_thread_gpu(self, input_pdf, output_pdf):
        """
        Thread wrapper for GPU-based red removal.
        """
        try:
            remove_red_pixels_gpu(input_pdf, output_pdf, self._update_progress, self.color_choice.get())
            self.status_label.config(text="Done!")
        except Exception as e:
            self.status_label.config(text=f"Error: {e}")
        finally:
            self._cleanup_after_processing()

    # ------------------- GIF ANIMATION / UI UPDATES ------------------- #
    def _start_gif_animation(self, gif_path):
        """
        Loads the GIF frames and starts animating them.
        """
        self.gif_frames = []
        self.gif_index = 0
        self.gif_running = True

        try:
            gif_img = Image.open(gif_path)
            while True:
                self.gif_frames.append(ImageTk.PhotoImage(gif_img.copy()))
                gif_img.seek(gif_img.tell() + 1)
        except EOFError:
            pass
        except Exception as e:
            print(f"Failed to load GIF: {e}")

        self._animate_gif()

    def _animate_gif(self):
        """
        Recursive function that updates the gif_label with the next frame every 50ms.
        """
        if self.gif_running and self.gif_frames:
            frame = self.gif_frames[self.gif_index]
            self.gif_label.config(image=frame)
            self.gif_index = (self.gif_index + 1) % len(self.gif_frames)
            self.after(50, self._animate_gif)
        else:
            self.gif_label.config(image=None)

    def _update_go_button_text(self):
        """
        Updates the "Go" button text to "Working...", cycling dots, while running is True.
        """
        while self.running:
            self.dot_count = (self.dot_count % 3) + 1
            dots = "." * self.dot_count
            self.go_button.config(text=f"Working{dots}")
            time.sleep(0.5)

    def _update_progress(self, value):
        """
        Callback to update the progress bar from 0 to 100.
        """
        self.progress_bar["value"] = value
        self.update_idletasks()

    def _cleanup_after_processing(self):
        """
        Actions to take after finishing or failing the process.
        """
        try:
            self.running = False
            self.gif_running = False
            if hasattr(self, 'go_button'):
                self.go_button.config(text="Go")
                self.go_button.config(state="normal")
            if hasattr(self, 'save_button'):
                self.save_button.config(state="normal")
            if hasattr(self, 'progress_bar'):
                self._update_progress(0)
        except:
            pass
    # ------------------- PREVIEW FUNCTION ------------------- #
    def _preview_pdf(self, pdf_path):
        """
        Loads the first page of a PDF into the preview canvas.
        """
        img = preview_pdf_page(pdf_path)
        ratio = min(595 / img.width, 842 / img.height)  # approximate A4 scaling
        new_size = (int(img.width * ratio), int(img.height * ratio))
        preview_img = ImageTk.PhotoImage(img.resize(new_size, Image.ANTIALIAS))
        self.preview_canvas.delete("all")
        x_center = (595 - new_size[0]) // 2
        y_center = (842 - new_size[1]) // 2
        self.preview_canvas.create_image(x_center, y_center, anchor="nw", image=preview_img)
        self.preview_canvas.image = preview_img  # keep a reference

    # ------------------- ABOUT DIALOG ------------------- #
    def _show_about(self):
        """
        Opens a 'About' window describing the app.
        """
        top = Toplevel(self)
        top.title("About PDFMute v2")
        top.geometry("300x150")

        info = (
            "PDFMute v2 - Professional Edition\n"
            "Version: 2.0\n\n"
            "Created by: Amit Hacoon\n"
            "GitHub: https://github.com/amithacoon/pdfmute"
        )
        lbl = ttk.Label(top, text=info, justify="center")
        lbl.pack(padx=10, pady=10)

        link_btn = ttk.Button(top, text="Open GitHub", command=lambda: webbrowser.open("https://github.com/amithacoon/pdfmute"))
        link_btn.pack()

        ttk.Button(top, text="Close", command=top.destroy).pack(pady=5)

    # ------------------- CLEAN EXIT ------------------- #
    def on_closing(self):
        """
        Overridden method when closing the main app window.
        """
        try:
            # Clean up temp files
            import tempfile
            from pathlib import Path
            temp_dir = Path(tempfile.gettempdir()) / 'pdfmute_temp'
            if temp_dir.exists():
                import shutil
                shutil.rmtree(temp_dir)

            # Stop running processes
            self.running = False
            self.gif_running = False
            for t in self.threads:
                if t.is_alive():
                    t.join(timeout=1.0)

            self.destroy()

        except:
            self.destroy()


import os
import tempfile
import time
import uuid
from pathlib import Path

import pythoncom
import win32com.client
import logging

class DocxConverter:
    def __init__(self):
        self.temp_dir = Path(tempfile.gettempdir()) / 'pdfmute_temp'
        self.temp_dir.mkdir(exist_ok=True)

        # Setup logging
        self.logger = logging.getLogger('DocxConverter')
        self.logger.setLevel(logging.DEBUG)

        if not self.logger.handlers:
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            log_file = self.temp_dir / 'conversion.log'
            file_handler = logging.FileHandler(str(log_file), encoding='utf-8')
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

        pythoncom.CoInitialize()
        self.logger.info("COM Initialized successfully")

    def _convert_single_file(self, word_app, input_path, output_path):
        """Single conversion attempt with detailed error logging"""
        self.logger.debug(f"Starting single file conversion")
        self.logger.debug(f"Input exists: {os.path.exists(input_path)}")

        doc = None
        try:
            self.logger.debug("Attempting to open document")
            try:
                doc = word_app.Documents.Open(
                    FileName=str(input_path),
                    ReadOnly=1,  # Use numeric value instead of constants.wdReadOnly
                    Visible=0,  # Use numeric value instead of constants.wdFalse
                    ConfirmConversions=0  # Use numeric value instead of constants.wdFalse
                )
                self.logger.info("Document opened successfully")
            except Exception as e:
                self.logger.error(f"Error opening document: {str(e)}")
                raise

            time.sleep(1)  # Allow document to load

            self.logger.debug("Attempting to save as PDF")
            try:
                doc.SaveAs(
                    FileName=str(output_path),
                    FileFormat=17,  # Use numeric value instead of constants.wdFormatPDF
                    AddToRecentFiles=0  # Use numeric value instead of constants.wdFalse
                )
                self.logger.info("Document saved as PDF")
            except Exception as e:
                self.logger.error(f"Error saving document as PDF: {str(e)}")
                raise

            time.sleep(1)  # Allow save to complete

            # Verify the output file exists
            if os.path.exists(output_path):
                self.logger.info("Output file created successfully")
                return True
            else:
                self.logger.error("Output file was not created")
                return False

        except Exception as e:
            self.logger.error(f"Error during conversion: {str(e)}")
            raise

        finally:
            if doc:
                try:
                    self.logger.debug("Attempting to close document")
                    doc.Close(SaveChanges=0)  # Use numeric value instead of constants.wdDoNotSaveChanges
                    self.logger.debug("Document closed")
                except Exception as e:
                    self.logger.error(f"Error closing document: {str(e)}")

            # Additional cleanup
            try:
                self.logger.debug("Attempting to quit Word application")
                word_app.Quit()
                self.logger.debug("Word application quit")
            except Exception as e:
                self.logger.error(f"Error quitting Word application: {str(e)}")

            # Wait for Word process to fully terminate
            time.sleep(2)

    def convert(self, input_file):
        """Convert with retry logic and detailed logging"""
        self.logger.info(f"Starting conversion process for {input_file}")

        try:
            input_path = Path(input_file).resolve()
            temp_pdf = self.temp_dir / f"{uuid.uuid4()}.pdf"

            self.logger.debug(f"Using temp file: {temp_pdf}")

            if not input_path.exists():
                self.logger.error(f"Input file not found: {input_path}")
                raise FileNotFoundError(f"Input file not found: {input_path}")

            # Check for crashed Word instances and close them
            self.logger.debug("Checking for crashed Word instances")
            self._close_crashed_word_instances()

            self.logger.debug("Creating Word application")
            word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False

            # Check if Word application is properly initialized
            try:
                word.Documents.Count
            except Exception as e:
                self.logger.error(f"Error initializing Word application: {str(e)}")
                raise

            max_retries = 3
            retry_delay = 1

            for attempt in range(max_retries):
                try:
                    self.logger.info(f"Conversion attempt {attempt + 1} of {max_retries}")

                    if self._convert_single_file(word, input_path, temp_pdf):
                        self.logger.info("Conversion successful")
                        return str(temp_pdf)

                except Exception as e:
                    self.logger.error(f"Attempt {attempt + 1} failed: {str(e)}")
                    if attempt == max_retries - 1:
                        raise

                    self.logger.info(f"Waiting {retry_delay} seconds before retry")
                    time.sleep(retry_delay)
                    retry_delay *= 2

                finally:
                    try:
                        word.Quit()
                        self.logger.debug("Word application closed")
                    except Exception as e:
                        self.logger.error(f"Error closing Word: {str(e)}")

        except Exception as e:
            self.logger.error(f"Conversion failed: {str(e)}")
            raise Exception(f"Conversion error: {str(e)}")

    def _close_crashed_word_instances(self):
        """Close crashed Word instances"""
        import psutil

        for proc in psutil.process_iter():
            try:
                if proc.name().lower() == "winword.exe":
                    if proc.status() == psutil.STATUS_ZOMBIE or proc.status() == psutil.STATUS_DEAD:
                        self.logger.debug(f"Closing crashed Word process with PID: {proc.pid}")
                        proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
# ------------- ENTRY POINT ------------- #
if __name__ == "__main__":
    try:
        app = PDFMuteApp()
        app.protocol("WM_DELETE_WINDOW", app.on_closing)
        app.mainloop()
    except Exception as e:
        print(f"Error: {e}")
        import os
        os._exit(1)