# PDF Red Pixel Remover

## Description
This script processes PDF files to remove red pixels, which can be used to clean scanned documents or images within the PDFs that contain unwanted red marks or highlights.

## Setup
To run this script, you need Python installed on your machine along with some packages which can be installed via pip. Ensure you have Python 3.x installed and then run the following command to install dependencies:

pip install -r requirements.txt


## Usage
To use the script, place your PDF files in a source directory. The script processes all PDF files in this directory and saves the processed files in a target directory without red pixels.

1. Update the `source_directory` and `target_directory` variables in the script to point to your source files and desired output location.
2. Run the script with the following command:

python remove_red_pixels.py


The script will process each PDF file and output the cleaned PDFs to the specified target directory, printing a message for each processed file.

## Contributing
Feel free to fork this project and submit pull requests with enhancements or fixes. If you find a bug or have a feature request, please open an issue.

## License
This project is open-source and available under the MIT License.
