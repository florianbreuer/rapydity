# RAPydity

A tool for processing Reasonable Adjustment Plans (RAPs) and applying extra time accommodations to Canvas assignments.

## Features

- Read RAP data from a single CSV file (exported from ServiceNow)
- Automatically filter to students enrolled in the selected course
- Apply extra time accommodations to Canvas quizzes and assignments
- User-friendly GUI interface
- Legacy support for processing individual RAP PDF files
- Secure handling of sensitive data

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/florianbreuer/rapydity.git
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
   - Select the RAP CSV file (exported from ServiceNow with columns `u_student_id`, `u_exam_time`, `u_requested_for1`)

## Usage

1. Select a course from the dropdown menu
2. Use "Update from RAP CSV" to process the RAP data for the selected course
3. View the extracted data using "View Extra Time Data"
4. Apply extra time to assignments using "Apply Extra Time"
5. Or use the "Just Do It!" button to update from RAP CSV and apply extra time to all quizzes in one step

Use "Change RAP File..." at the bottom of the window to select a different RAP CSV file at any time.

## Security

- Sensitive data (Canvas tokens, RAP data, student data) is stored locally and never committed to version control
- Configuration files are automatically created in the application directory
- All data processing is done locally on your machine

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Florian Breuer
