import os
import csv
import PyPDF2
import requests
from pathlib import Path
from dataclasses import dataclass
import re
from typing import Optional, Dict, List
import configparser
import logging
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import platform
import sys

__version__ = "2.0.0"

@dataclass
class CourseConfig:
    course_id: str
    course_name: str
    end_at: Optional[str] = None
    csv_file: Optional[Path] = None

    def __post_init__(self):
        # Set default CSV filename if not provided
        if self.csv_file is None:
            self.csv_file = Path(f'extra_time_{self.course_id}.csv')
        # Convert Path strings to Path objects if needed
        if isinstance(self.csv_file, str):
            self.csv_file = Path(self.csv_file)

class CourseManager:
    def __init__(self, canvas_api, logger=None):
        self.canvas_api = canvas_api
        self.logger = logger or logging.getLogger(__name__)
        self.config_file = Path('courses.ini')

        # Load RAP CSV file path from config
        config = configparser.ConfigParser()
        if self.config_file.exists():
            config.read(self.config_file)
            if 'General' in config and config['General'].get('rap_csv_file'):
                self.rap_csv_file = Path(config['General']['rap_csv_file'])
            else:
                self.rap_csv_file = None
            # Keep legacy shared_rap_folder for PDF fallback
            if 'General' in config and config['General'].get('shared_rap_folder'):
                self.shared_rap_folder = Path(config['General']['shared_rap_folder'])
            else:
                self.shared_rap_folder = None
        else:
            self.rap_csv_file = None
            self.shared_rap_folder = None

        self.courses: Dict[str, CourseConfig] = {}
        self._load_config()
        
    def _load_config(self):
        """Load course configurations from file"""
        config = configparser.ConfigParser()

        if not self.config_file.exists():
            # Create default config
            general = {}
            if self.rap_csv_file:
                general['rap_csv_file'] = str(self.rap_csv_file)
            if self.shared_rap_folder:
                general['shared_rap_folder'] = str(self.shared_rap_folder)
            if general:
                config['General'] = general
            with open(self.config_file, 'w') as f:
                config.write(f)
        else:
            config.read(self.config_file)

            # Load course configurations
            for section in config.sections():
                if section.startswith('Course.'):
                    course_id = section.split('.')[1]
                    end_at = config[section].get('end_at') or None
                    self.logger.debug(f"Loading course {course_id} from config with end_at: {end_at}")
                    course_config = CourseConfig(
                        course_id=course_id,
                        course_name=config[section]['name'],
                        end_at=end_at,
                        csv_file=config[section].get('csv_file')
                    )
                    self.courses[course_id] = course_config
        
    def save_config(self):
        """Save current configuration to file"""
        config = configparser.ConfigParser()

        # Save general settings
        general = {}
        if self.rap_csv_file:
            general['rap_csv_file'] = str(self.rap_csv_file)
        if self.shared_rap_folder:
            general['shared_rap_folder'] = str(self.shared_rap_folder)
        if general:
            config['General'] = general

        # Save course configurations
        for course_id, course in self.courses.items():
            section = f'Course.{course_id}'
            config[section] = {
                'name': course.course_name,
                'end_at': course.end_at if course.end_at else '',
                'csv_file': str(course.csv_file)
            }

        with open(self.config_file, 'w') as f:
            config.write(f)
        
    def add_course(self, course_id: str, course_name: str,
                  end_at: Optional[str] = None) -> CourseConfig:
        """Add a new course configuration"""
        self.logger.debug(f"Adding course {course_id}: {course_name} with end_at: {end_at}")
        course = CourseConfig(
            course_id=course_id,
            course_name=course_name,
            end_at=end_at,
        )
        self.courses[course_id] = course
        self.save_config()
        return course

    def get_rap_pdf_files(self, folder: Path) -> list:
        """Get all RAP PDF files from a folder (legacy PDF support)"""
        if folder and folder.exists():
            return list(folder.glob('*.pdf'))
        return []

@dataclass
class Student:
    name: str
    surname: str
    student_number: str
    extra_time_per_hour: int
    canvas_id: Optional[str] = None

class RAPReader:
    def __init__(self, log_level=logging.INFO):
        # Set up logging
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)  # Capture all logs at logger level
        
        # Create handlers if none exist
        if not self.logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.WARNING)  # Only WARNING and above for console
            formatter = logging.Formatter('%(levelname)s: %(message)s')
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
            
            # File handler
            file_handler = logging.FileHandler('rapydity.log')
            file_handler.setLevel(logging.DEBUG)  # Log everything to file
            file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)

        # Load configuration
        config = configparser.ConfigParser()
        if Path('config.ini').exists():
            config.read('config.ini')
            
            # Initialize Canvas API
            self.canvas_api = CanvasAPI(
                access_token=config['canvas']['access_token'],
                base_url=config['canvas']['base_url'],
                logger=self.logger
            )
            
            # Initialize course manager
            self.course_manager = CourseManager(self.canvas_api, self.logger)
            self.current_course: Optional[CourseConfig] = None
            
            # If no courses configured, fetch available courses
            if not self.course_manager.courses:
                self.fetch_available_courses()
        else:
            self.logger.debug("No config.ini found - waiting for GUI setup")
            # Initialize with empty values - will be set up by GUI
            self.canvas_api = None
            self.course_manager = None
            self.current_course = None

    def initialize_from_config(self):
        """Initialize API and course manager after config is created by GUI"""
        if not Path('config.ini').exists():
            self.logger.error("Cannot initialize - no config.ini found")
            return False
            
        config = configparser.ConfigParser()
        config.read('config.ini')
        
        # Initialize Canvas API
        self.canvas_api = CanvasAPI(
            access_token=config['canvas']['access_token'],
            base_url=config['canvas']['base_url'],
            logger=self.logger
        )
        
        # Initialize course manager
        self.course_manager = CourseManager(self.canvas_api, self.logger)
        self.current_course = None
        
        # Fetch courses with proper end dates
        courses = self.canvas_api.list_courses()
        if courses:
            self.logger.info(f"Found {len(courses)} courses in Canvas")
            for course in courses:
                course_id = str(course['id'])
                course_name = course['name']
                end_at = course.get('effective_end_at')  # Use the effective end date
                self.logger.debug(f"Adding course {course_id}: {course_name} (ends: {end_at})")
                self.course_manager.add_course(course_id, course_name, end_at)
            
        return True

    def fetch_available_courses(self):
        """Fetch available courses from Canvas and prompt user to select which to manage"""
        courses = self.canvas_api.list_courses()
        if not courses:
            self.logger.warning("No courses found in Canvas")
            return
        
        self.logger.info(f"Found {len(courses)} courses in Canvas")
        for course in courses:
            course_id = str(course['id'])
            course_name = course['name']
            end_at = course.get('end_at')
            self.logger.debug(f"Course {course_id}: {course_name} (ends: {end_at})")
            self.course_manager.add_course(course_id, course_name, end_at)

    def extract_student_info_from_pdf(self, pdf_path: Path) -> Optional[Student]:
        """Extract student info from a RAP PDF file"""
        try:
            with open(pdf_path, 'rb') as f:
                pdf = PyPDF2.PdfReader(f)
                # Extract text and normalize whitespace
                text = ' '.join(
                    ' '.join(page.extract_text().split()) 
                    for page in pdf.pages
                )
                self.logger.info(f"\nProcessing {pdf_path.name}:")
                self.logger.debug(f"Extracted text: {text[:200]}...") # Log first 200 chars for debugging
                
                # Extract name and surname - get first match of name followed by uppercase surname and exactly 7 digits
                name_matches = list(re.finditer(r'(\w+)\s*([A-Z][-A-Z]{1,}?)\s*(\d{7})', text))
                name_match = name_matches[0] if name_matches else None
                
                # Extract extra time - format is "Extra time 30 mins per hour"
                extra_time_match = re.search(r'Extra time (\d+) mins? per hour', text)
                
                if name_match and extra_time_match:
                    student = Student(
                        name=name_match.group(1),          # First name (John)
                        surname=name_match.group(2),        # Surname (DOE or ADAMS-WILSON)
                        student_number=name_match.group(3), # Student number (3472571)
                        extra_time_per_hour=int(extra_time_match.group(1)) # 30
                    )
                    self.logger.info(f"Found student in {pdf_path.name}: {student}")
                    return student
                else:
                    self.logger.warning("Could not find all required information:")
                    
                    if not name_match:
                        self.logger.warning(f"Student name/number not found in {pdf_path.name}")
                    
                    if not extra_time_match:
                        self.logger.warning(
                            f"No extra time information found in {pdf_path.name}. "
                            f"Please check this PDF manually to verify if extra time accommodation is specified."
                        )
                    
                    self.logger.debug(f"Name match: {name_match}")
                    self.logger.debug(f"Extra time match: {extra_time_match}")

        except Exception as e:
            self.logger.error(f"Error processing {pdf_path}: {e}")

        return None

    def extract_students_from_rap_csv(self, csv_path: Path) -> Dict[str, Student]:
        """Extract student info from a RAP CSV file.

        Reads columns: u_student_id, u_exam_time, u_requested_for1
        Returns dict keyed by student number (only students with extra time).
        """
        students = {}
        no_time_phrases = {'no additional time required', 'no additional time required.', ''}

        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    raw_id = row.get('u_student_id', '').strip()
                    raw_time = row.get('u_exam_time', '').strip()
                    raw_name = row.get('u_requested_for1', '').strip()

                    # Parse student number: strip leading C/c
                    student_number = raw_id.lstrip('Cc')
                    if not student_number or not student_number.isdigit():
                        self.logger.debug(f"Skipping row with invalid student ID: '{raw_id}'")
                        continue

                    # Parse extra time
                    if raw_time.lower() in no_time_phrases:
                        continue  # No extra time needed

                    time_match = re.search(r'(\d+)', raw_time)
                    if time_match:
                        extra_time = int(time_match.group(1))
                    else:
                        self.logger.warning(
                            f"Unrecognized exam time format for student {raw_id}: '{raw_time}'"
                        )
                        continue

                    # Parse name
                    parts = raw_name.split()
                    name = parts[0] if parts else ""
                    surname = " ".join(parts[1:]) if len(parts) > 1 else ""

                    students[student_number] = Student(
                        name=name,
                        surname=surname,
                        student_number=student_number,
                        extra_time_per_hour=extra_time
                    )

            self.logger.info(f"Read {len(students)} students with extra time from RAP CSV")
        except Exception as e:
            self.logger.error(f"Error reading RAP CSV {csv_path}: {e}")

        return students

    def _read_existing_csv(self, csv_path: Path) -> Dict[str, Student]:
        """Read existing CSV file into dictionary keyed by student number"""
        students = {}
        if csv_path.exists():
            with open(csv_path, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    students[row['student_number']] = Student(
                        name=row['name'],
                        surname=row['surname'],
                        student_number=row['student_number'],
                        extra_time_per_hour=int(row['extra_time_per_hour']),
                        canvas_id=row.get('canvas_id')
                    )
            self.logger.info(f"Read {len(students)} existing students from {csv_path}")
        return students


    def _write_csv(self, students: List[Student], csv_path: Path):
        """Write students to CSV file"""
        with open(csv_path, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['name', 'surname', 'student_number', 'extra_time_per_hour', 'canvas_id'])
            writer.writeheader()
            for student in students:
                writer.writerow(vars(student))
        self.logger.info(f"Wrote {len(students)} students to {csv_path}")


    def update_csv_from_raps(self, source="csv"):
        """Update extra_time.csv with information from RAP data.

        Args:
            source: "csv" to read from RAP CSV file, "pdf" to read from PDF folder.
        """
        if not self.current_course:
            self.logger.error("No course selected")
            return

        course = self.current_course
        self.logger.info(f"Starting RAP processing (source: {source})...")

        # Initialize counters for warnings and errors
        warning_count = 0
        error_count = 0

        # Create a custom handler to count warnings and errors
        class CountingHandler(logging.Handler):
            def emit(self, record):
                nonlocal warning_count, error_count
                if record.levelno == logging.WARNING:
                    warning_count += 1
                elif record.levelno >= logging.ERROR:
                    error_count += 1

        # Add the counting handler temporarily
        counting_handler = CountingHandler()
        self.logger.addHandler(counting_handler)

        try:
            students_by_number = self._read_existing_csv(course.csv_file)
            new_count = 0

            if source == "csv":
                # CSV mode: read RAP CSV, filter by enrollment
                rap_csv = self.course_manager.rap_csv_file
                if not rap_csv or not rap_csv.exists():
                    self.logger.error("No RAP CSV file configured or file not found")
                    return

                all_rap_students = self.extract_students_from_rap_csv(rap_csv)
                self.logger.info(f"Found {len(all_rap_students)} students with extra time in RAP CSV")

                # Filter to only students enrolled in this course
                for student_number, student in all_rap_students.items():
                    if student_number not in students_by_number:
                        canvas_id = self.canvas_api.find_student_canvas_id(
                            course.course_id, student_number
                        )
                        if canvas_id:
                            student.canvas_id = canvas_id
                            students_by_number[student_number] = student
                            new_count += 1
                            self.logger.info(f"Added new student: {student.name} {student.surname}")
                        # No warning here - most RAP CSV students won't be in this course
                    else:
                        self.logger.info(f"Student already exists: {students_by_number[student_number].name} {students_by_number[student_number].surname}")

                self.logger.info(f"Added {new_count} new students from RAP CSV")

            elif source == "pdf":
                # PDF mode: legacy fallback
                rap_folder = self.course_manager.shared_rap_folder
                if not rap_folder:
                    self.logger.error("No shared RAP folder configured for PDF processing")
                    return

                pdf_files = self.course_manager.get_rap_pdf_files(rap_folder)
                pdf_count = 0

                for pdf_path in pdf_files:
                    pdf_count += 1
                    student = self.extract_student_info_from_pdf(pdf_path)
                    if student:
                        if student.student_number not in students_by_number:
                            canvas_id = self.canvas_api.find_student_canvas_id(
                                course.course_id, student.student_number
                            )
                            if canvas_id:
                                student.canvas_id = canvas_id
                                students_by_number[student.student_number] = student
                                new_count += 1
                                self.logger.info(f"Added new student: {student.name} {student.surname}")
                            else:
                                self.logger.warning(
                                    f"Could not find Canvas ID for student: {student.name} {student.surname} "
                                    f"(probably not enrolled in this course)"
                                )
                        else:
                            self.logger.info(f"Student already exists: {student.name} {student.surname}")

                self.logger.info(f"Processed {pdf_count} PDFs")
                self.logger.info(f"Added {new_count} new students")

            # Write updated CSV
            self._write_csv(list(students_by_number.values()), course.csv_file)

            # Show alert if there were warnings or errors
            if warning_count > 0 or error_count > 0:
                import tkinter.messagebox as messagebox

                message = "Issues occurred during RAP processing:\n\n"
                if warning_count > 0:
                    message += f"• {warning_count} warning{'s' if warning_count > 1 else ''}\n"
                if error_count > 0:
                    message += f"• {error_count} error{'s' if error_count > 1 else ''}\n"

                message += "\nPlease review the log for details."

                if error_count > 0:
                    messagebox.showerror("Processing Errors", message)
                else:
                    messagebox.showwarning("Processing Warnings", message)

        finally:
            # Remove the counting handler
            self.logger.removeHandler(counting_handler)

class CanvasAPI:
    def __init__(self, 
                access_token=None, 
                base_url=None, 
                course_id=None,
                logger=None):
        self.access_token = access_token
        self.base_url = base_url
        self.course_id = course_id
        self.logger = logger or logging.getLogger(__name__)
        self.header = {'Authorization': f'Bearer {self.access_token}'}
        self._enrollments_cache = {}  # Cache enrollments per course
        
    def list_courses(self):
        """Return a list of active courses where the user is a teacher"""
        url = f'{self.base_url}/api/v1/courses'
        params = {
            'enrollment_type': 'teacher',
            'state[]': ['available', 'completed'],  # Include current and past courses
            'include[]': ['term', 'concluded', 'enrollment_term'],  # Include all term info
            'per_page': 100
        }
        
        try:
            courses = self.get_paginated_results(url, params)
            self.logger.debug("Fetching courses from Canvas")
            for course in courses:
                # Try different ways Canvas might indicate course end
                course_end = (
                    course.get('end_at') or  # Course-specific end date
                    course.get('term', {}).get('end_at') or  # Term end date
                    course.get('enrollment_term', {}).get('end_at') or  # Enrollment term end date
                    (course.get('concluded') and course.get('created_at'))  # If concluded, use creation date
                )
                course['effective_end_at'] = course_end
                self.logger.debug(  # Keep detailed course data at DEBUG level
                    f"Course {course['id']}: {course['name']}\n"
                    f"  course end_at: {course.get('end_at')}\n"
                    f"  term end_at: {course.get('term', {}).get('end_at')}\n"
                    f"  enrollment_term end_at: {course.get('enrollment_term', {}).get('end_at')}\n"
                    f"  concluded: {course.get('concluded')}\n"
                    f"  effective end date: {course_end}"
                )
            
            # Sort by term and name
            sorted_courses = sorted(
                courses,
                key=lambda c: (
                    c.get('term', {}).get('start_at', ''),
                    c.get('name', '')
                ),
                reverse=True  # Most recent first
            )
            return sorted_courses
        except Exception as e:
            self.logger.error(f"Failed to fetch courses: {e}")
            return []

    def get_paginated_results(self, url, params=None):
        """Return full list of responses from requests.get(url), walking through pagination"""
        results = []
        try:
            r = requests.get(url, headers=self.header, params=params)
            if r.status_code == requests.codes.ok:
                response = r.json()
                results.extend(response)
                while 'next' in r.links.keys():
                    r = requests.get(r.links['next']['url'], headers=self.header)
                    response = r.json()
                    results.extend(response)
            self.logger.debug(f"Retrieved {len(results)} results from API")
        except Exception as e:
            self.logger.error(f"API request failed: {e}")
        return results

    def get_enrollments(self, course_id: str) -> List[dict]:
        """Get enrollments for a specific course, using cache if available"""
        if course_id not in self._enrollments_cache:
            url = f"{self.base_url}/api/v1/courses/{course_id}/enrollments"
            self.logger.debug(f"Fetching enrollments for course {course_id}")  # Keep as DEBUG
            self._enrollments_cache[course_id] = self.get_paginated_results(url)
            self.logger.debug(f"Found {len(self._enrollments_cache[course_id])} enrollments")  # Move to DEBUG
        return self._enrollments_cache[course_id]

    def find_student_canvas_id(self, course_id: str, student_number: str) -> Optional[str]:
        """Find Canvas user ID for a student by their student number in a specific course"""
        self.logger.debug(f"Looking up Canvas ID for student number: {student_number}")  # Keep as DEBUG
        
        for enrollment in self.get_enrollments(course_id):
            user = enrollment.get('user', {})
            self.logger.debug(f"Checking user: {user.get('name')} - sis_user_id: {user.get('sis_user_id')}")  # Keep as DEBUG
            if user.get('sis_user_id') == 'c'+str(student_number):
                self.logger.info(f"Found Canvas ID for student {user.get('name')} ({student_number})")  # Add INFO for success
                return str(user['id'])
        
        self.logger.warning(f"No matching user found for student number: {student_number}")  # Keep as WARNING
        return None

    def list_assignments(self, published_only=True) -> List[dict]:
        """Return a list of assignments for this course from Canvas"""
        url = f"{self.base_url}/api/v1/courses/{self.course_id}/assignments"
        self.logger.debug(f"Fetching assignments for course {self.course_id}")
        
        assignments = self.get_paginated_results(url)
        self.logger.debug(f"Found {len(assignments)} total assignments")
        
        if published_only:
            # Return only published assignments
            published = [a for a in assignments if a.get('published')]
            self.logger.info(f"Found {len(published)} published assignments")
            
        return assignments

    def get_assignment_time_limit(self, assignment_id: str) -> Optional[int]:
        """Get time limit in minutes for an assignment, or None if no limit"""
        url = f"{self.base_url}/api/v1/courses/{self.course_id}/assignments/{assignment_id}"
        r = requests.get(url, headers=self.header)
        if r.status_code == requests.codes.ok:
            assignment = r.json()
            # For quizzes, we need to check the quiz settings
            if assignment.get('is_quiz_assignment'):
                quiz_id = assignment.get('quiz_id')
                quiz_url = f"{self.base_url}/api/v1/courses/{self.course_id}/quizzes/{quiz_id}"
                r = requests.get(quiz_url, headers=self.header)
                if r.status_code == requests.codes.ok:
                    quiz = r.json()
                    time_limit = quiz.get('time_limit')
                    self.logger.debug(f"Quiz time limit: {time_limit} minutes")
                    return time_limit
            return None
        return None

    def post_extra_time(self, assignment_id: str, student_adjustments: List[dict]) -> bool:
        """Post extra time for quiz. Each adjustment needs user_id and extra_time"""
        # First get the quiz_id from the assignment
        assignment_url = f"{self.base_url}/api/v1/courses/{self.course_id}/assignments/{assignment_id}"
        r = requests.get(assignment_url, headers=self.header)
        if r.status_code != requests.codes.ok:
            self.logger.error(f"Failed to get assignment info: {r.status_code} - {r.text}")
            return False
        
        assignment = r.json()
        if not assignment.get('is_quiz_assignment'):
            self.logger.error("This is not a quiz assignment")
            return False
        
        quiz_id = assignment.get('quiz_id')
        if not quiz_id:
            self.logger.error("Could not find quiz_id")
            return False
        
        # Now post the extensions to the quiz endpoint
        url = f"{self.base_url}/api/v1/courses/{self.course_id}/quizzes/{quiz_id}/extensions"
        
        # Format adjustments for the quiz extension endpoint
        extensions = {
            'quiz_extensions': [
                {
                    'user_id': adj['user_id'],
                    'extra_time': adj['extra_time_mins']
                }
                for adj in student_adjustments
            ]
        }
        
        try:
            r = requests.post(url, headers=self.header, json=extensions)
            if r.status_code == requests.codes.ok:
                self.logger.info(f"Successfully posted extra time for {len(student_adjustments)} students")
                return True
            else:
                self.logger.error(f"Failed to post extra time: {r.status_code} - {r.text}")
                return False
        except Exception as e:
            self.logger.error(f"Error posting extra time: {e}")
            return False

    def verify_student_enrollments(self, student_ids: List[str]) -> List[str]:
        """Return list of student IDs that are still enrolled in the course"""
        enrollments = self.get_enrollments(self.course_id)
        enrolled_ids = {str(e['user']['id']) for e in enrollments}
        return [sid for sid in student_ids if sid in enrolled_ids]

class RAPReaderGUI:
    def __init__(self, reader):
        self.reader = reader
        self.logger = reader.logger
        
        # Check if initial setup is needed
        if not Path('config.ini').exists():
            self.show_setup_dialog()
            # Initialize reader with new config
            if not self.reader.initialize_from_config():
                messagebox.showerror("Error", "Failed to initialize application")
                sys.exit(1)

        # Check RAP CSV file is configured and exists
        if self.reader.course_manager:
            rap_csv = self.reader.course_manager.rap_csv_file
            if rap_csv and rap_csv.exists():
                self.logger.info(f"Using RAP CSV file: {rap_csv}")
            else:
                # Show file picker to select RAP CSV
                filepath = filedialog.askopenfilename(
                    title="Select RAP CSV File",
                    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
                )
                if filepath:
                    self.reader.course_manager.rap_csv_file = Path(filepath)
                    self.reader.course_manager.save_config()
                    self.logger.info(f"RAP CSV file set to: {filepath}")
                else:
                    self.logger.warning("No RAP CSV file selected. You can set one later via 'Change RAP File...'")

        self.root = tk.Tk()
        self.root.title(f"RAPydity v{__version__}")
        self.root.geometry("800x600")
        
        # Set window icon
        self._set_window_icon(self.root)
        
        # Create menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Instructions", command=self.show_instructions)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Course selector frame
        course_frame = ttk.LabelFrame(main_frame, text="Course Selection", padding="5")
        course_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Course selector in its own frame
        selector_frame = ttk.Frame(course_frame)
        selector_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Course selector
        self.course_var = tk.StringVar()
        self.course_selector = ttk.Combobox(
            selector_frame, 
            textvariable=self.course_var,
            state='readonly',
            width=50
        )
        self.course_selector.pack(fill=tk.X, padx=5)
        
        # Options frame for checkbox and buttons
        options_frame = ttk.Frame(course_frame)
        options_frame.pack(fill=tk.X)
        
        # Show current courses only checkbox
        self.show_current_only = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Show current courses only",
            variable=self.show_current_only,
            command=self._update_course_list
        ).pack(side=tk.LEFT, padx=5)
        
        # Action buttons frame
        action_frame = ttk.Frame(options_frame)
        action_frame.pack(side=tk.RIGHT)

        # Add View Extra Time Data button
        ttk.Button(
            action_frame,
            text="View Extra Time Data",
            command=self.view_extra_time_data
        ).pack(side=tk.LEFT, padx=5)

        # Update from RAP CSV button (primary)
        ttk.Button(
            action_frame,
            text="Update from RAP CSV",
            command=self.update_raps_csv
        ).pack(side=tk.LEFT, padx=5)

        # Update from RAP PDFs button (legacy fallback)
        ttk.Button(
            action_frame,
            text="Update from RAP PDFs",
            command=self.update_raps_pdf
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            action_frame,
            text="Apply Extra Time",
            command=self.apply_extra_time
        ).pack(side=tk.LEFT, padx=5)

        # Create the accent button style
        style = ttk.Style()
        style.configure("Accent.TButton",
                        font=('', 9, 'bold'))  # Make the font bold

        # Add "Just Do It" button with distinctive styling and checkmark
        just_do_it_button = ttk.Button(
            action_frame,
            text="✓ Just Do It!",  # Unicode checkmark
            command=self.just_do_it,
            style="Accent.TButton"  # Custom style for emphasis
        )
        just_do_it_button.pack(side=tk.LEFT, padx=5)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add log text widget
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Add buttons at the bottom of log frame
        bottom_buttons_frame = ttk.Frame(log_frame)
        bottom_buttons_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5,0))
        
        # Add Clear Log button
        ttk.Button(
            bottom_buttons_frame,
            text="Clear Log",
            command=self.clear_log
        ).pack(side=tk.LEFT, padx=(0,5))

        # Add Manage Courses button
        ttk.Button(
            bottom_buttons_frame,
            text="Manage Courses",
            command=self.show_course_manager
        ).pack(side=tk.LEFT, padx=(0,5))

        # Add Change RAP File button
        ttk.Button(
            bottom_buttons_frame,
            text="Change RAP File...",
            command=self.change_rap_file
        ).pack(side=tk.LEFT)
        
        # Add custom handler for logging to text widget
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
        self.reader.logger.addHandler(text_handler)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, pady=(5,0))
        
        # Now that everything is set up, update course list and bind selection change
        self._update_course_list()
        self.course_selector.bind('<<ComboboxSelected>>', self._on_course_selected)

    def _set_window_icon(self, window):
        """Set the window icon based on the operating system"""
        try:
            if platform.system() == 'Windows':
                window.iconbitmap('R_logo3.ico')
            else:
                icon = tk.PhotoImage(file='R_logo3.gif')
                window.iconphoto(True, icon)
        except Exception as e:
            self.logger.warning(f"Could not load icon: {e}")

    def update_raps_csv(self):
        """Handle updating from RAP CSV file"""
        if not self.reader.course_manager.rap_csv_file or not self.reader.course_manager.rap_csv_file.exists():
            messagebox.showerror("Error", "No RAP CSV file configured. Use 'Change RAP File...' to select one.")
            return
        try:
            self.status_var.set("Processing RAP CSV...")
            self.root.config(cursor="watch")
            self.root.update()
            self.reader.update_csv_from_raps(source="csv")
            self.status_var.set("Ready")
        except Exception as e:
            error_msg = f"Failed to process RAP CSV: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Error occurred")
        finally:
            self.root.config(cursor="")

    def update_raps_pdf(self):
        """Handle updating from RAP PDFs (legacy fallback)"""
        if not self.reader.course_manager.shared_rap_folder:
            # Ask user to select a folder
            folder = filedialog.askdirectory(title="Select folder containing RAP PDFs")
            if not folder:
                return
            self.reader.course_manager.shared_rap_folder = Path(folder)
            self.reader.course_manager.save_config()
        try:
            self.status_var.set("Processing RAP PDFs...")
            self.root.config(cursor="watch")
            self.root.update()
            self.reader.update_csv_from_raps(source="pdf")
            self.status_var.set("Ready")
        except Exception as e:
            error_msg = f"Failed to process RAP PDFs: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Error occurred")
        finally:
            self.root.config(cursor="")

    def change_rap_file(self):
        """Open file picker to change the RAP CSV file"""
        initial_dir = None
        if self.reader.course_manager.rap_csv_file:
            initial_dir = str(self.reader.course_manager.rap_csv_file.parent)
        filepath = filedialog.askopenfilename(
            title="Select RAP CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if filepath:
            self.reader.course_manager.rap_csv_file = Path(filepath)
            self.reader.course_manager.save_config()
            self.logger.info(f"RAP CSV file changed to: {filepath}")
    
    def apply_extra_time(self):
        """Show dialog for selecting assignments to apply extra time to"""
        # Check if course is selected
        if not self.reader.current_course:
            messagebox.showerror("Error", "Please select a course first")
            return
        
        self.status_var.set("Checking student data...")
        # Check if we have a CSV file with student data
        csv_path = self.reader.current_course.csv_file
        if not csv_path.exists():
            messagebox.showerror(
                "Error", 
                "No student data found. Please update from RAP CSV first."
            )
            self.status_var.set("Ready")
            return
        
        # Read student data
        students = self.reader._read_existing_csv(csv_path)
        if not students:
            messagebox.showerror(
                "Error", 
                "No students found in CSV file. Please update from RAP CSV first."
            )
            self.status_var.set("Ready")
            return
        
        try:
            self.status_var.set("Fetching assignments from Canvas...")
            self.root.config(cursor="watch")
            self.root.update()
            # Create assignment selector dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Select Assignments")
            dialog.geometry("550x400")
            
            # Set dialog icon
            self._set_window_icon(dialog)
            
            # Get assignments
            assignments = self.reader.canvas_api.list_assignments(published_only=True)
            self.status_var.set("Ready")
            self.root.config(cursor="")

            # Main frame with padding
            main_frame = ttk.Frame(dialog)
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Configure grid weights for main_frame
            main_frame.grid_columnconfigure(0, weight=1)
            main_frame.grid_rowconfigure(0, weight=1)
            
            # Create treeview for assignments
            tree = ttk.Treeview(main_frame, columns=('id', 'name'), show='headings')
            tree.heading('id', text='ID')
            tree.heading('name', text='Assignment Name')
            tree.column('id', width=100)
            tree.column('name', width=400, stretch=True)
            tree.configure(selectmode='extended')
            
            # Add scrollbar
            scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # Add assignments to treeview
            for assignment in assignments:
                tree.insert('', 'end', values=(assignment['id'], assignment['name']))
            
            # Layout
            tree.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
            scrollbar.grid(row=0, column=1, sticky='ns', padx=(0,5), pady=5)
            
            # Buttons frame
            btn_frame = ttk.Frame(main_frame)
            btn_frame.grid(row=1, column=0, columnspan=2, pady=5)
            
            def apply_extra_time():
                selected = tree.selection()
                if not selected:
                    self.logger.warning("No assignments selected for extra time application")
                    messagebox.showwarning("Warning", "Please select at least one assignment")
                    return
                
                selected_assignments = [
                    (tree.item(item)['values'][0], tree.item(item)['values'][1])
                    for item in selected
                ]
                
                # Show confirmation dialog
                assignment_names = "\n".join(f"- {name}" for _, name in selected_assignments)
                self.logger.info(f"Preparing to apply extra time to:\n{assignment_names}")
                confirm = messagebox.askokcancel(
                    "Confirm Extra Time Application",
                    f"Are you sure you want to apply extra time to these assignments?\n\n{assignment_names}",
                    icon='warning'
                )
                
                if not confirm:
                    self.logger.debug("Extra time application cancelled by user")
                    self.status_var.set("Ready")
                    dialog.config(cursor="")
                    return
                
                self.status_var.set("Verifying student enrollments...")
                dialog.config(cursor="watch")
                dialog.update()
                # Verify student enrollments
                student_ids = [s.canvas_id for s in students.values() if s.canvas_id]
                active_ids = self.reader.canvas_api.verify_student_enrollments(student_ids)
                
                # Process each assignment
                for assignment_id, assignment_name in selected_assignments:
                    self.status_var.set(f"Processing assignment: {assignment_name}")
                    dialog.update()
                    # Get time limit
                    time_limit = self.reader.canvas_api.get_assignment_time_limit(assignment_id)
                    if time_limit is None:  # Check explicitly for None since 0 is a valid time limit
                        self.logger.warning(f"Assignment '{assignment_name}' has no time limit, skipping")
                        continue
                    
                    # Calculate adjustments for each student
                    adjustments = []
                    for student in students.values():
                        if student.canvas_id in active_ids:
                            # Calculate extra time (round up)
                            extra_mins = int((student.extra_time_per_hour * time_limit) / 60 + 0.5)
                            adjustments.append({
                                'user_id': student.canvas_id,
                                'extra_time_mins': extra_mins
                            })
                    
                    # Post adjustments
                    if adjustments:
                        self.status_var.set(f"Applying extra time for: {assignment_name}")
                        dialog.update()
                        success = self.reader.canvas_api.post_extra_time(assignment_id, adjustments)
                        if success:
                            self.logger.info(
                                f"Applied extra time to {len(adjustments)} students "
                                f"for assignment '{assignment_name}'"
                            )
                        else:
                            self.logger.error(
                                f"Failed to apply extra time for assignment '{assignment_name}'"
                            )
                    else:
                        self.logger.warning(
                            f"No active students found for assignment '{assignment_name}'"
                        )
                
                self.status_var.set("Extra time application complete")
                messagebox.showinfo("Success", "Extra time application complete. Check logs for details.")
                dialog.destroy()
            
            # Add buttons
            ttk.Button(btn_frame, text="Apply", command=apply_extra_time).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        except Exception as e:
            error_msg = f"Failed to apply extra time: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Error occurred")
    
    def _update_course_list(self):
        """Update the course selector dropdown"""
        courses = self.reader.course_manager.courses
        if not courses:
            self.course_selector['values'] = ['No courses configured']
            self.course_selector.set('No courses configured')
            return
        
        # Filter courses if show_current_only is checked
        if self.show_current_only.get():
            from datetime import datetime
            now = datetime.now().isoformat()
            self.logger.debug(f"Filtering courses. Current time: {now}")
            current_courses = {
                cid: course for cid, course in courses.items()
                if not course.end_at or course.end_at > now
            }
            self.logger.debug(f"Found {len(current_courses)} current courses out of {len(courses)} total")
            for cid, course in courses.items():
                self.logger.debug(f"Course {cid}: {course.course_name} (ends: {course.end_at})")
            courses = current_courses
        
        # Format: "COURSE_NAME (ID: COURSE_ID)"
        course_list = [
            f"{course.course_name} (ID: {course.course_id})"
            for course in courses.values()
        ]
        
        # Sort courses alphabetically
        course_list.sort()
        
        self.course_selector['values'] = course_list
        
        # Select first course if none selected
        if not self.course_var.get() and course_list:
            self.course_selector.set(course_list[0])
            self._on_course_selected(None)

    def _on_course_selected(self, event):
        """Handle course selection change"""
        selection = self.course_var.get()
        if not selection or selection == 'No courses configured':
            return
        
        # Extract course ID from selection string
        course_id = selection.split('ID: ')[-1].rstrip(')')
        course = self.reader.course_manager.courses.get(course_id)
        if course:
            self.reader.current_course = course
            self.reader.canvas_api.course_id = course_id
            self.status_var.set(f"Selected course: {course.course_name}")
            self.logger.info(f"Selected course: {course.course_name} (ID: {course_id})")

    def show_course_manager(self):
        """Show dialog for managing courses"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Manage Courses")
        dialog.geometry("700x600")

        # Set dialog icon
        self._set_window_icon(dialog)

        # Create treeview for courses
        tree = ttk.Treeview(dialog, columns=('id', 'name'), show='headings')
        tree.heading('id', text='Course ID')
        tree.heading('name', text='Course Name')
        tree.column('id', width=100)
        tree.column('name', width=500, stretch=True)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(dialog, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        # Add current courses to treeview
        for course in self.reader.course_manager.courses.values():
            tree.insert('', 'end', values=(course.course_id, course.course_name))

        # Layout
        tree.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        scrollbar.grid(row=0, column=1, sticky='ns')

        # Buttons frame
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)

        def refresh_courses():
            """Fetch and add new courses from Canvas"""
            self.status_var.set("Fetching courses from Canvas...")
            dialog.config(cursor="watch")
            dialog.update()

            try:
                courses = self.reader.canvas_api.list_courses()
                self.logger.debug("Raw course data received from Canvas:")
                for course in courses:
                    self.logger.debug(
                        f"Course {course['id']}: {course['name']}\n"
                        f"  course end_at: {course.get('end_at')}\n"
                        f"  term end_at: {course.get('term', {}).get('end_at')}\n"
                        f"  enrollment_term end_at: {course.get('enrollment_term', {}).get('end_at')}\n"
                        f"  concluded: {course.get('concluded')}\n"
                        f"  effective end date: {course['effective_end_at']}"
                    )

                if not courses:
                    messagebox.showwarning("No Courses", "No courses found in Canvas")
                    return

                # Clear existing items
                for item in tree.get_children():
                    tree.delete(item)

                # Add all courses
                for course in courses:
                    course_id = str(course['id'])
                    course_name = course['name']
                    end_at = course.get('effective_end_at')
                    self.logger.debug(f"Adding course: {course_id} - {course_name} (ends: {end_at})")
                    self.reader.course_manager.add_course(course_id, course_name, end_at)
                    tree.insert('', 'end', values=(course_id, course_name))

                self._update_course_list()
                self.logger.info(f"Successfully added {len(courses)} courses from Canvas")
                messagebox.showinfo("Success", f"Added {len(courses)} courses")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to fetch courses: {str(e)}")
            finally:
                dialog.config(cursor="")
                self.status_var.set("Ready")

        ttk.Button(btn_frame, text="Refresh from Canvas", command=refresh_courses).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        # Configure grid weights
        dialog.grid_rowconfigure(0, weight=1)
        dialog.grid_columnconfigure(0, weight=1)

    def view_extra_time_data(self):
        """Display the extra time data from the CSV file"""
        if not self.reader.current_course:
            messagebox.showerror("Error", "Please select a course first")
            return
            
        csv_path = self.reader.current_course.csv_file
        if not csv_path.exists():
            self.logger.info("No extra time data found. Please run 'Update from RAP CSV' first to process RAP data.")
            return
            
        # Create data viewer window
        viewer = tk.Toplevel(self.root)
        viewer.title(f"Extra Time Data - {self.reader.current_course.course_name}")
        viewer.geometry("800x600")
        
        # Set window icon
        self._set_window_icon(viewer)
        
        # Create main frame with padding
        main_frame = ttk.Frame(viewer, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview
        tree = ttk.Treeview(main_frame, columns=('name', 'surname', 'student_number', 'extra_time'), show='headings')
        tree.heading('name', text='First Name')
        tree.heading('surname', text='Surname')
        tree.heading('student_number', text='Student Number')
        tree.heading('extra_time', text='Extra Time (mins/hour)')
        
        # Set column widths
        tree.column('name', width=150)
        tree.column('surname', width=150)
        tree.column('student_number', width=150)
        tree.column('extra_time', width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add button frame at the bottom
        button_frame = ttk.Frame(viewer, padding="10")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Add Apply Extra Time button
        ttk.Button(
            button_frame,
            text="Apply Extra Time",
            command=self.apply_extra_time
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Add Close button
        ttk.Button(
            button_frame,
            text="Close",
            command=viewer.destroy
        ).pack(side=tk.LEFT)
        
        # Read and display data
        try:
            students_data = []
            students = self.reader._read_existing_csv(csv_path)
            for student in students.values():
                students_data.append((
                    student.name,
                    student.surname,
                    student.student_number,
                    student.extra_time_per_hour
                ))
            
            # Variables to track sorting
            sort_column = 'surname'  # Default sort by surname
            sort_reverse = False
            
            def populate_tree(data):
                """Clear and repopulate the treeview with the given data"""
                # Clear existing items
                for item in tree.get_children():
                    tree.delete(item)
                
                # Add data to treeview
                for values in data:
                    tree.insert('', 'end', values=values)
            
            def sort_treeview():
                """Sort the treeview based on current sort_column and sort_reverse"""
                # Get column index for sorting
                col_index = {'name': 0, 'surname': 1, 'student_number': 2, 'extra_time': 3}[sort_column]
                
                # Sort the data
                sorted_data = sorted(
                    students_data,
                    key=lambda x: (
                        # Handle numeric sorting for extra_time column
                        int(x[col_index]) if sort_column == 'extra_time' else x[col_index].lower()
                    ),
                    reverse=sort_reverse
                )
                
                # Update column headings to show sort indicators
                for col in ('name', 'surname', 'student_number', 'extra_time'):
                    if col == sort_column:
                        indicator = " ▼" if sort_reverse else " ▲"
                        tree.heading(col, text=f"{col.title().replace('_', ' ')}{indicator}", 
                                    command=lambda c=col: sort_by_column(c))
                    else:
                        tree.heading(col, text=col.title().replace('_', ' '), 
                                    command=lambda c=col: sort_by_column(c))
                
                # Repopulate the tree with sorted data
                populate_tree(sorted_data)
            
            def sort_by_column(column):
                """Handle column header click for sorting"""
                nonlocal sort_column, sort_reverse
                
                if sort_column == column:
                    # If already sorting by this column, reverse the order
                    sort_reverse = not sort_reverse
                else:
                    # New sort column, default to ascending
                    sort_column = column
                    sort_reverse = False
                
                sort_treeview()
            
            # Set up column heading click events
            for col in ('name', 'surname', 'student_number', 'extra_time'):
                tree.heading(col, command=lambda c=col: sort_by_column(c))
            
            # Initial sort and display
            sort_treeview()
            
            self.logger.info(f"Displaying extra time data for {len(students)} students")
        except Exception as e:
            self.logger.error(f"Error reading CSV file: {e}")
            messagebox.showerror("Error", f"Failed to read extra time data: {str(e)}")
            viewer.destroy()

    def clear_log(self):
        """Clear the contents of the log text widget"""
        self.log_text.delete('1.0', tk.END)
        self.logger.info("Log cleared")

    def show_about(self):
        """Show About dialog"""
        about = tk.Toplevel(self.root)
        about.title("About RAPydity")
        about.geometry("400x300")
        
        # Set window icon
        self._set_window_icon(about)
        
        # Main frame with padding
        main_frame = ttk.Frame(about, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # App name and version
        ttk.Label(
            main_frame,
            text=f"RAPydity v{__version__}",
            font=('', 14, 'bold')
        ).pack(pady=(0, 10))
        
        # Description
        ttk.Label(
            main_frame,
            text="A tool for processing Reasonable Adjustment Plans (RAPs)\n"
                 "and applying extra time accommodations to Canvas assignments.",
            justify=tk.CENTER,
            wraplength=350
        ).pack(pady=(0, 20))
        
        # Copyright
        ttk.Label(
            main_frame,
            text="© 2024 Florian Breuer",
            font=('', 9)
        ).pack(pady=(0, 5))
        
        # License
        license_frame = ttk.Frame(main_frame)
        license_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            license_frame,
            text="Released under the MIT License",
            font=('', 9)
        ).pack(side=tk.LEFT)
        
        def open_license():
            import webbrowser
            webbrowser.open('https://opensource.org/licenses/MIT')
        
        ttk.Button(
            license_frame,
            text="View License",
            command=open_license,
            style='Small.TButton'
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # GitHub link
        def open_github():
            import webbrowser
            webbrowser.open('https://github.com/florianbreuer/rapydity')
        
        ttk.Button(
            main_frame,
            text="View on GitHub",
            command=open_github
        ).pack()
        
        # Close button
        ttk.Button(
            main_frame,
            text="Close",
            command=about.destroy
        ).pack(pady=(20, 0))

    def show_instructions(self):
        """Open instructions in default web browser"""
        import webbrowser
        webbrowser.open('instructions.html')

    def show_setup_dialog(self):
        """Show initial setup dialog for configuring RAP folder and Canvas token"""
        setup = tk.Tk()
        setup.title(f"RAPydity v{__version__} Setup")
        setup.geometry("600x300")
        
        # Set window icon
        self._set_window_icon(setup)
        
        main_frame = ttk.Frame(setup, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Welcome message
        ttk.Label(
            main_frame,
            text="Welcome to RAPydity!\nPlease complete the initial setup.",
            font=('', 12, 'bold')
        ).pack(pady=(0, 20))
        
        # Canvas token frame
        token_frame = ttk.LabelFrame(main_frame, text="Canvas API Token", padding="10")
        token_frame.pack(fill=tk.X, pady=(0, 10))
        
        token_var = tk.StringVar()
        token_entry = ttk.Entry(token_frame, textvariable=token_var, width=50)
        token_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        def open_canvas_help():
            import webbrowser
            webbrowser.open('https://community.canvaslms.com/t5/Student-Guide/How-do-I-manage-API-access-tokens-as-a-student/ta-p/273')
        
        ttk.Button(
            token_frame,
            text="How to get token?",
            command=open_canvas_help
        ).pack(side=tk.LEFT)
        
        # RAP CSV file frame
        csv_frame = ttk.LabelFrame(main_frame, text="RAP CSV File", padding="10")
        csv_frame.pack(fill=tk.X, pady=(0, 20))

        csv_var = tk.StringVar()
        csv_entry = ttk.Entry(csv_frame, textvariable=csv_var, width=50)
        csv_entry.pack(side=tk.LEFT, padx=(0, 10))

        def browse_csv():
            filepath = filedialog.askopenfilename(
                title="Select RAP CSV File",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if filepath:
                csv_var.set(filepath)

        ttk.Button(
            csv_frame,
            text="Browse...",
            command=browse_csv
        ).pack(side=tk.LEFT)

        def save_config():
            if not token_var.get().strip():
                messagebox.showerror("Error", "Please enter your Canvas API token")
                return

            # Save configuration
            config = configparser.ConfigParser()
            config['canvas'] = {
                'access_token': token_var.get().strip(),
                'base_url': 'https://canvas.newcastle.edu.au/'
            }
            general = {}
            if csv_var.get().strip():
                general['rap_csv_file'] = csv_var.get().strip()
            if general:
                config['General'] = general

            try:
                with open('config.ini', 'w') as f:
                    config.write(f)
                self.logger.info("Created config.ini")
                setup.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not save configuration: {e}")
        
        # Save button
        ttk.Button(
            main_frame,
            text="Save and Continue",
            command=save_config
        ).pack()
        
        # Make dialog modal
        setup.grab_set()
        setup.focus_set()
        setup.wait_window()

    def just_do_it(self):
        """Combined action to update RAPs and apply extra time to all quizzes"""
        # Check if course is selected
        if not self.reader.current_course:
            messagebox.showerror("Error", "Please select a course first")
            return
        
        # Show confirmation dialog
        course_name = self.reader.current_course.course_name
        confirm = messagebox.askokcancel(
            "Confirm Action",
            f"This will perform the following actions for {course_name}:\n\n"
            f"1. Update student data from RAP CSV\n"
            f"2. Apply extra time to ALL quizzes in this course\n\n"
            f"Are you sure you want to proceed?",
            icon='warning'
        )

        if not confirm:
            self.logger.info("Just Do It action cancelled by user")
            return

        # Step 1: Update RAPs from CSV
        self.logger.info("STEP 1: Updating from RAP CSV...")
        self.status_var.set("Updating from RAP CSV...")
        self.root.update()

        # Run the RAP update process
        self.reader.update_csv_from_raps(source="csv")

        # Check if we have student data after update
        csv_path = self.reader.current_course.csv_file
        if not csv_path.exists():
            messagebox.showerror(
                "Error",
                "No student data was created. Please check that the RAP CSV file is configured."
            )
            self.status_var.set("Ready")
            return

        students = self.reader._read_existing_csv(csv_path)
        if not students:
            messagebox.showerror(
                "Error",
                "No students found in the data. Please check your RAP CSV file."
            )
            self.status_var.set("Ready")
            return
        
        # Step 2: Apply extra time to all quizzes
        self.logger.info("STEP 2: Applying extra time to all quizzes...")
        self.status_var.set("Fetching quizzes from Canvas...")
        self.root.config(cursor="watch")
        self.root.update()
        
        try:
            # Get all quizzes (assignments with time limits)
            assignments = self.reader.canvas_api.list_assignments(published_only=True)
            
            # Filter to only include quizzes (assignments with time limits)
            quizzes = []
            for assignment in assignments:
                time_limit = self.reader.canvas_api.get_assignment_time_limit(assignment['id'])
                if time_limit is not None:  # Has a time limit
                    quizzes.append(assignment)
            
            if not quizzes:
                messagebox.showinfo(
                    "No Quizzes Found", 
                    "No quizzes with time limits were found in this course."
                )
                self.status_var.set("Ready")
                self.root.config(cursor="")
                return
            
            # Show confirmation with list of quizzes
            quiz_names = "\n".join(f"• {q['name']}" for q in quizzes)
            confirm_quizzes = messagebox.askokcancel(
                "Confirm Quiz Selection",
                f"Extra time will be applied to these {len(quizzes)} quizzes:\n\n{quiz_names}",
                icon='info'
            )
            
            if not confirm_quizzes:
                self.logger.info("Quiz selection cancelled by user")
                self.status_var.set("Ready")
                self.root.config(cursor="")
                return
            
            # Verify student enrollments
            self.status_var.set("Verifying student enrollments...")
            self.root.update()
            student_ids = [s.canvas_id for s in students.values() if s.canvas_id]
            active_ids = self.reader.canvas_api.verify_student_enrollments(student_ids)
            
            # Apply extra time to each quiz
            success_count = 0
            for quiz in quizzes:
                self.status_var.set(f"Processing: {quiz['name']}")
                self.root.update()
                
                # Get time limit
                time_limit = self.reader.canvas_api.get_assignment_time_limit(quiz['id'])
                
                # Calculate adjustments for each student
                adjustments = []
                for student in students.values():
                    if student.canvas_id in active_ids:
                        # Calculate extra time (round up)
                        extra_mins = int((student.extra_time_per_hour * time_limit) / 60 + 0.5)
                        adjustments.append({
                            'user_id': student.canvas_id,
                            'extra_time_mins': extra_mins
                        })
                
                # Apply the adjustments
                if adjustments:
                    try:
                        self.reader.canvas_api.post_extra_time(quiz['id'], adjustments)
                        self.logger.info(f"Applied extra time to {quiz['name']} for {len(adjustments)} students")
                        success_count += 1
                    except Exception as e:
                        self.logger.error(f"Failed to apply extra time to {quiz['name']}: {e}")
                else:
                    self.logger.warning(f"No eligible students found for {quiz['name']}")
            
            # Show completion message
            messagebox.showinfo(
                "Process Complete",
                f"The Just Do It process has completed:\n\n"
                f"• Updated student data from RAP CSV\n"
                f"• Applied extra time to {success_count} of {len(quizzes)} quizzes"
            )
            
        except Exception as e:
            self.logger.error(f"Error during Just Do It process: {e}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            self.status_var.set("Ready")
            self.root.config(cursor="")

    def run(self):
        """Start the GUI"""
        self.root.mainloop()

class TextHandler(logging.Handler):
    """Handler for redirecting logging to tkinter text widget with colored messages"""
    def __init__(self, text_widget):
        super().__init__()
        self.setLevel(logging.INFO)  # Changed from WARNING to INFO
        self.text_widget = text_widget
        
        # Configure tags for different log levels
        self.text_widget.tag_configure('INFO', foreground='black')
        self.text_widget.tag_configure('DEBUG', foreground='gray')
        self.text_widget.tag_configure('WARNING', background='#FFCC00', foreground='black')  # Yellow background
        self.text_widget.tag_configure('ERROR', background='#FF3333', foreground='white')    # Red background
        self.text_widget.tag_configure('CRITICAL', background='#CC0000', foreground='white') # Darker red for critical
        
        # Map log levels to tag names
        self.level_tags = {
            logging.DEBUG: 'DEBUG',
            logging.INFO: 'INFO',
            logging.WARNING: 'WARNING',
            logging.ERROR: 'ERROR',
            logging.CRITICAL: 'CRITICAL'
        }
        
    def emit(self, record):
        msg = self.format(record)
        
        def append():
            # Get the appropriate tag for this log level
            tag = self.level_tags.get(record.levelno, 'INFO')
            
            # Insert the message with the tag
            self.text_widget.insert(tk.END, msg + '\n', tag)
            self.text_widget.see(tk.END)
            
        # Schedule the update on the main thread
        self.text_widget.after(0, append)

def main():
    # Set default icon for all windows
    if platform.system() == 'Windows':
        import ctypes
        myappid = 'newcastle.rapydity.1.0'  # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    reader = RAPReader()
    reader.logger.info(f"Starting RAPydity v{__version__}")
    gui = RAPReaderGUI(reader)
    gui.run()

if __name__ == "__main__":
    main() 