import fitz  # PyMuPDF
import numpy as np
import io
import torch
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image, ImageTk
import threading
import os
import time
def remove_red_pixels(input_pdf, output_pdf, progress_callback):
    doc = fitz.open(input_pdf)
    new_doc = fitz.Document()

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

        # First pass - convert intense red pixels and target colors
        for y in range(img.height):
            for x in range(img.width):
                r, g, b = pixels[x, y]
                if r > 150 and r > g * 1.2 and r > b * 1.5 and (r + g + b) > 100:
                    pixels[x, y] = (255, 255, 255)
                else:
                    for color, delta in target_colors:
                        if abs(r - color[0]) <= delta and abs(g - color[1]) <= delta and abs(b - color[2]) <= delta:
                            pixels[x, y] = (255, 255, 255)
                            break

        # Second pass - convert remaining reddish pixels (adjust thresholds as needed)
        for y in range(img.height):
            for x in range(img.width):
                r, g, b = pixels[x, y]
                if r > g and r > b and r > 180:  # Adjust threshold for reddishness
                    pixels[x, y] = (255, 255, 255)

        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=100)
        img_byte_arr.seek(0)
        new_page = new_doc.new_page(width=pix.width, height=pix.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr.read())
        progress_percent = ((page_number + 1) / total_pages) * 100
        progress_callback(progress_percent)  # Update the progress

    new_doc.save(output_pdf)
    doc.close()
    new_doc.close()

def remove_red_pixels_gpu(input_pdf, output_pdf, progress_callback):
    # Open the PDF document
    doc = fitz.open(input_pdf)
    new_doc = fitz.Document()
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    total_pages = len(doc)

    for page_number, page in enumerate(doc):
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_tensor = torch.tensor(np.array(img), device=device)
        img_tensor = img_tensor.permute(2, 0, 1).float()  # Convert to C, H, W format

        # Define a mask for red removal
        red_mask = (img_tensor[0] > 150) & (img_tensor[0] > img_tensor[1] * 1.2) & (img_tensor[0] > img_tensor[2] * 1.5)
        img_tensor[:, red_mask] = torch.tensor([255.0, 255.0, 255.0], device=device).view(3, 1)

        # Convert back to PIL Image to save in PDF
        img_tensor = img_tensor.byte().permute(1, 2, 0).cpu().numpy()
        img = Image.fromarray(img_tensor)
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=100)
        img_byte_arr.seek(0)

        new_page = new_doc.new_page(width=pix.width, height=pix.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr.read())

        progress_percent = ((page_number + 1) / total_pages) * 100
        progress_callback(progress_percent)  # Update the progress

    new_doc.save(output_pdf)
    doc.close()
    new_doc.close()


def preview_pdf_page(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)  # Load the first page
    pix = page.get_pixmap(dpi=100)  # Render page to an image
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


class PDFRedRemoverApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('PDF Mute')
        self.geometry('1000x800')

        # Color Palette
        self.bg_color = "#f2f2f2"  # Light gray background
        self.button_color = "#00a884"  # Teal for buttons
        self.button_hover_color = "#00876c"  # Darker teal on hover
        self.red_color = "#ff4d4d"  # Red for active state

        # Configure Style
        style = ttk.Style()
        style.theme_use("clam")

        # Customize Button Style
        style.configure("TButton", font=("Helvetica", 14), padding=10,
                        background=self.button_color, foreground="white",
                        borderradius=5)

        # Configure Label Style
        style.configure("TLabel", font=("Helvetica", 12))

        # Main Frames
        control_frame = tk.Frame(self, bg=self.bg_color)
        control_frame.pack(fill='x', padx=20, pady=10)
        preview_frame = tk.Frame(self, bg=self.bg_color)
        preview_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # Algorithm selection frame
        algorithm_frame = tk.Frame(control_frame, bg=self.bg_color)
        algorithm_frame.pack(side='left', padx=10)

        selection_label = tk.Label(algorithm_frame, text="Pick an Algorithm:", background=self.bg_color, font=("Helvetica", 12))
        selection_label.pack(side='top', padx=10, pady=(0, 10))  # Add some padding to separate from the radio buttons

        # Algorithm selection radio buttons with labels
        self.algorithm = tk.StringVar()
        cpu_button = ttk.Radiobutton(algorithm_frame, text='CPU', value='CPU', variable=self.algorithm)
        cpu_label = ttk.Label(algorithm_frame, text='- Slow and High quality result', background=self.bg_color)
        gpu_button = ttk.Radiobutton(algorithm_frame, text='GPU', value='GPU', variable=self.algorithm)
        gpu_label = ttk.Label(algorithm_frame, text='- Fast and Low quality result', background=self.bg_color)
        about_button = tk.Button(algorithm_frame, text='About', command=self.show_about, font=('Helvetica', 12, 'bold'), bg=self.bg_color)


        cpu_button.pack(anchor='w')
        cpu_label.pack(anchor='w')
        gpu_button.pack(anchor='w')
        gpu_label.pack(anchor='w')
        about_button.pack(side='top', pady=(10, 0))  # Positioned at the top of the frame, under the radio buttons

        self.algorithm.set('CPU')  # Default selection

        # Buttons
        self.load_button = tk.Button(control_frame, text='Load PDF', command=self.load_pdf)
        self.load_button.pack(side='left', padx=(10, 20))
        self.save_button = tk.Button(control_frame, text='Save as PDF', command=self.set_output, state='disabled')
        self.save_button.pack(side='left')

        # Go Button with Color Change
        self.go_button = tk.Button(control_frame, text='Go', command=self.process_pdf,
                                   font=("Helvetica", 16, "bold"),
                                   bg=self.button_color, fg="white",
                                   activebackground=self.red_color,
                                   borderwidth=0, relief="flat", state='disabled')
        self.go_button.pack(side='left', padx=(10, 20))

        # Progress bar and activity indicator
        # Configure the progress bar style with green color
        style.configure('Green.Horizontal.TProgressbar', troughcolor=self.bg_color, background=self.button_color)
        self.progress = ttk.Progressbar(control_frame, style='Green.Horizontal.TProgressbar', length=200,
                                        mode='determinate')
        self.progress.pack(side='left', padx=(10, 20))
        self.activity_indicator = tk.Label(control_frame, text=" ", font=('Helvetica', 12), bg=self.bg_color)
        self.activity_indicator.pack(side='left', padx=(10, 0))

        # GIF Frame with fixed size
        gif_frame = tk.Frame(control_frame, width=250, height=250, bg=self.bg_color)
        gif_frame.pack(side='left', padx=(10, 0))
        gif_frame.pack_propagate(False)  # Prevent the frame from resizing
        self.gif_label = tk.Label(gif_frame, bg=self.bg_color)
        self.gif_label.pack(fill='both', expand=True)
        self.gif_frames = []  # Initialize gif_frames to avoid AttributeError

        # Preview Canvas
        self.preview_canvas = tk.Canvas(preview_frame, bg='grey', width=595, height=842)
        self.preview_canvas.pack(pady=20)

    def show_about(self):
        top = tk.Toplevel(self)
        top.title("About PDF Mute")

        message = "Version: 1.1\nCreated by Amit Hacoon"
        msg_label = tk.Label(top, text=message)
        msg_label.pack(pady=(10, 5))

        # Link to GitHub
        link_label = tk.Label(top, text="GitHub Repository", fg="blue", cursor="hand2")
        link_label.pack()
        link_label.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/amithacoon/pdfmute"))

        # Close button for the dialog
        close_button = tk.Button(top, text="Close", command=top.destroy)
        close_button.pack(pady=(5, 10))

    def load_gif(self, gif_path):
        self.gif_frames = []
        self.gif_index = 0
        self.gif = Image.open(gif_path)
        try:
            while True:
                self.gif_frames.append(ImageTk.PhotoImage(self.gif.copy()))
                self.gif.seek(self.gif.tell() + 1)
        except EOFError:
            pass  # End of GIF file

    def animate_gif(self):
        if self.running:
            frame = self.gif_frames[self.gif_index]
            self.gif_label.config(image=frame)
            self.gif_index = (self.gif_index + 1) % len(self.gif_frames)
            self.after(50, self.animate_gif)

    def animate_activity(self):
        chars = ""
        while self.running:
            for char in chars:
                if not self.running:
                    break
                self.activity_indicator.config(text=char)
                self.update()
                time.sleep(0.1)
        self.activity_indicator.config(text=" ")  # Reset to blank when not processing

    def load_pdf(self):
        self.filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.filename:
            image = preview_pdf_page(self.filename)
            self.preview_image = ImageTk.PhotoImage(image.resize((595, 842)))  # Resize for A4 proportion
            self.preview_canvas.create_image(298, 421, image=self.preview_image)  # Center the image
            self.save_button.config(state='normal')  # Enable save button

    def set_output(self):
        default_output_pdf = os.path.splitext(self.filename)[0] + "_MuteRed.pdf"
        self.output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile=default_output_pdf)
        if self.output_pdf:
            self.go_button.config(state='normal')  # Enable go button

    def process_pdf(self):
        self.go_button.config(bg=self.red_color)
        self.running = True
        threading.Thread(target=self.animate_activity).start()
        if not self.gif_frames:  # Ensure GIF has been loaded
            self.load_gif('busy.gif')  # Provide the correct path
        threading.Thread(target=self.animate_gif).start()
        process_func = self.process_thread_cpu if self.algorithm.get() == 'CPU' else self.process_thread_gpu
        threading.Thread(target=process_func, args=(self.filename, self.output_pdf)).start()

    def process_thread_cpu(self, input_pdf, output_pdf):
        try:
            remove_red_pixels(input_pdf, output_pdf, self.update_progress)
            self.activity_indicator.config(text="Done!")  # Update text to Done when complete
        except Exception as e:
            self.activity_indicator.config(text="Error!")  # Show error in the activity indicator
        finally:
            self.running = False
            self.save_button.config(state='normal')
            self.go_button.config(state='normal')
            self.gif_label.config(image='')  # Hide GIF when done
            self.go_button.config(bg=self.button_color)
            self.activity_indicator.config(text="Done!")  # Update text to "Done!" when complete
        pass
    def process_thread_gpu(self, input_pdf, output_pdf):
        try:
            remove_red_pixels_gpu(input_pdf, output_pdf, self.update_progress)
            self.activity_indicator.config(text="Done!")  # Update text to Done when complete
        except Exception as e:
            self.activity_indicator.config(text="Error!")  # Show error in the activity indicator
        finally:
            self.running = False
            self.save_button.config(state='normal')
            self.go_button.config(state='normal')
            self.gif_label.config(image='')  # Hide GIF when done
            self.go_button.config(bg=self.button_color)
            self.activity_indicator.config(text="Done!")  # Update text to "Done!" when complete
        pass
    def update_progress(self, progress):
        self.progress['value'] = progress
        self.update_idletasks()

if __name__ == '__main__':
    app = PDFRedRemoverApp()
    app.mainloop()


