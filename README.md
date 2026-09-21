# Survey Processing Portal

A web application for processing educational survey documents. Converts PDF and DOCX survey files into professionally formatted reports with automatic percentage calculations.

## Features

- **Process Survey Documents**: Upload DOCX files with survey data
- **PDF to DOCX Conversion**: Convert PDF survey results to editable DOCX format
- **Automatic Calculations**: Calculates percentages based on total responses
- **Professional Formatting**: Generates styled reports with:
  - Header with campus name and grade
  - Grade section information table
  - Color-coded tables with dark blue headers
  - Teacher, Subject, Responses, and Percentage columns

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python app.py
```

3. Open browser at: `http://localhost:5000`

## Deployment

### Deploy to Render (Free)

1. Push code to GitHub
2. Go to https://render.com
3. Create new Web Service
4. Connect your GitHub repository
5. Set build command: `pip install -r requirements.txt`
6. Set start command: `gunicorn app:app`
7. Deploy!

## Usage

1. Fill in the form fields:
   - Campus Name (e.g., "Primary Campus")
   - Grade-Section (e.g., "Grade-3")
   - Survey Date
   - Total Students
   - Total Responses

2. Upload your survey file (DOCX or PDF)

3. Click "Process Document" or "Convert to DOCX"

4. Download the formatted result

## Technologies

- **Backend**: Flask (Python)
- **Document Processing**: python-docx, pdfplumber
- **Frontend**: HTML, CSS, JavaScript
- **Deployment**: Gunicorn

## License

MIT License
