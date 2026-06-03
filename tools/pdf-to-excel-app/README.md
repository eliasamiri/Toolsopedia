# PDF to Excel Converter

This project is a web application that allows users to upload PDF files, convert them to Excel format, and download the converted files. It is built using Flask, a lightweight WSGI web application framework in Python.

## Project Structure

```
pdf-to-excel-app
├── src
│   ├── app.py                # Entry point of the application
│   ├── controllers
│   │   └── convert.py        # Handles PDF upload and conversion
│   ├── services
│   │   └── converter.py       # Logic for converting PDF to Excel
│   ├── routes
│   │   └── index.py          # Defines application routes
│   └── utils
│       └── file_handler.py    # Utility functions for file handling
├── static
│   └── css
│       └── style.css         # CSS styles for the application
├── templates
│   ├── index.html            # Main HTML template for file upload
│   └── download.html         # Template for download link
├── uploads                   # Directory for temporary uploaded files
├── converted                 # Directory for converted Excel files
├── requirements.txt          # Project dependencies
└── README.md                 # Project documentation
```

## Setup Instructions

1. **Clone the repository:**
   ```
   git clone <repository-url>
   cd pdf-to-excel-app
   ```

2. **Create a virtual environment:**
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install dependencies:**
   ```
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```
   python src/app.py
   ```

5. **Access the application:**
   Open your web browser and go to `http://127.0.0.1:5000`.

## Usage

- Upload a PDF file using the form on the main page.
- After the upload, the application will convert the PDF to an Excel file.
- A download link for the converted file will be provided on the next page.

## Dependencies

- Flask
- PyPDF2 (or any other library for PDF manipulation)
- openpyxl (or any other library for Excel file creation)

## License

This project is licensed under the MIT License.