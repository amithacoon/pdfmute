import fitz  # PyMuPDF
from PIL import Image
import io  # For handling byte streams
import os  # For handling file directories


def remove_red_pixels(input_pdf, output_pdf):
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

    for page in doc:
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

    new_doc.save(output_pdf)


def process_directory(source_dir, target_dir):
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)  # Create target directory if it doesn't exist

    for filename in os.listdir(source_dir):
        if filename.endswith('.pdf'):
            input_path = os.path.join(source_dir, filename)
            output_path = os.path.join(target_dir, filename)
            remove_red_pixels(input_path, output_path)
            print(f"Processed {filename} and saved to {target_dir}")


# Example usage
source_directory = "exams"
target_directory = "no solution"
process_directory(source_directory, target_directory)


