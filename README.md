# RAPydity

A tool for processing Reasonable Adjustment Plans (RAPs) and applying extra time accommodations to Canvas assignments.

## Features

- Process RAP PDF files to extract student information and extra time requirements
- Manage multiple courses and their associated RAP folders
- Apply extra time accommodations to Canvas quizzes and assignments
- User-friendly GUI interface
- Secure handling of sensitive data

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/rapydity.git
   cd rapydity
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Initial Setup

1. Launch the application:
   ```bash
   python rapydity.py
   ```

2. On first run, you'll need to:
   - Enter your Canvas API token
   - Configure the shared RAP folder location

## Usage

1. Select a course from the dropdown menu
2. Use "Update from RAP PDFs" to process new RAP files
3. View the extracted data using "View Extra Time Data"
4. Apply extra time to assignments using "Apply Extra Time"

## Security

- Sensitive data (Canvas tokens, RAP PDFs, student data) is stored locally and never committed to version control
- Configuration files are automatically created in the application directory
- All data processing is done locally on your machine

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Florian Breuer 