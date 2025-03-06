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

__version__ = "0.9.0"

@dataclass
class CourseConfig:
    course_id: str
    course_name: str
    end_at: Optional[str] = None
    course_rap_folder: Optional[Path] = None
    csv_file: Optional[Path] = None
    
    def __post_init__(self):
        # Set default CSV filename if not provided
        if self.csv_file is None:
            self.csv_file = Path(f'extra_time_{self.course_id}.csv')
        # Convert Path strings to Path objects if needed
        if isinstance(self.csv_file, str):
            self.csv_file = Path(self.csv_file)
        if isinstance(self.course_rap_folder, str):
            self.course_rap_folder = Path(self.course_rap_folder)

class CourseManager:
    def __init__(self, canvas_api, logger=None):
        self.canvas_api = canvas_api
        self.logger = logger or logging.getLogger(__name__)
        self.config_file = Path('courses.ini')
        
        # Load shared RAP folder from config
        config = configparser.ConfigParser()
        if self.config_file.exists():
            config.read(self.config_file)
            if 'General' in config:
                self.shared_rap_folder = Path(config['General']['shared_rap_folder'])
            else:
                self.shared_rap_folder = Path('RAP')
        else:
            self.shared_rap_folder = Path('RAP')
        
        self.courses: Dict[str, CourseConfig] = {}
        self._load_config()
        
    def _load_config(self):
        """Load course configurations from file"""
        config = configparser.ConfigParser()
        
        if not self.config_file.exists():
            # Create default config with shared RAP folder
            config['General'] = {
                'shared_rap_folder': str(self.shared_rap_folder)
            }
            with open(self.config_file, 'w') as f:
                config.write(f)
        else:
            config.read(self.config_file)
            # Load shared RAP folder
            if 'General' in config:
                self.shared_rap_folder = Path(config['General']['shared_rap_folder'])
            
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
                        course_rap_folder=config[section].get('course_rap_folder'),
                        csv_file=config[section].get('csv_file')
                    )
                    self.courses[course_id] = course_config
        
    def save_config(self):
        """Save current configuration to file"""
        config = configparser.ConfigParser()
        
        # Save general settings
        config['General'] = {
            'shared_rap_folder': str(self.shared_rap_folder)
        }
        
        # Save course configurations
        for course_id, course in self.courses.items():
            section = f'Course.{course_id}'
            config[section] = {
                'name': course.course_name,
                'end_at': course.end_at if course.end_at else '',
                'csv_file': str(course.csv_file)
            }
            if course.course_rap_folder:
                config[section]['course_rap_folder'] = str(course.course_rap_folder)
        
        with open(self.config_file, 'w') as f:
            config.write(f)
        
    def add_course(self, course_id: str, course_name: str, 
                  end_at: Optional[str] = None,
                  course_rap_folder: Optional[Path] = None) -> CourseConfig:
        """Add a new course configuration"""
        self.logger.debug(f"Adding course {course_id}: {course_name} with end_at: {end_at}")
        course = CourseConfig(
            course_id=course_id,
            course_name=course_name,
            end_at=end_at,
            course_rap_folder=course_rap_folder
        )
        self.courses[course_id] = course
        self.save_config()
        return course
    
    def get_rap_files(self, course_id: str) -> list[Path]:
        """Get all RAP files for a course (from both shared and course-specific folders)"""
        pdf_files = []
        
        # Get PDFs from shared folder
        if self.shared_rap_folder.exists():
            pdf_files.extend(self.shared_rap_folder.glob('*.pdf'))
        
        # Get PDFs from course-specific folder if it exists
        course = self.courses.get(course_id)
        if course and course.course_rap_folder and course.course_rap_folder.exists():
            pdf_files.extend(course.course_rap_folder.glob('*.pdf'))
        
        return pdf_files

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
                    self.logger.warning(f"Name match: {name_match}")
                    self.logger.warning(f"Extra time match: {extra_time_match}")
                    

        except Exception as e:
            self.logger.error(f"Error processing {pdf_path}: {e}")
        

        return None

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


    def update_csv_from_raps(self):
        """Update extra_time.csv with information from RAP PDFs"""
        if not self.current_course:
            self.logger.error("No course selected")
            return
        
        course = self.current_course
        self.logger.info("Starting RAP processing...")
        
        students_by_number = self._read_existing_csv(course.csv_file)
        
        # Process each PDF in RAP folder
        pdf_count = 0
        new_count = 0
        
        # Get PDFs from both shared and course-specific folders
        pdf_files = self.course_manager.get_rap_files(course.course_id)
        for pdf_path in pdf_files:
            pdf_count += 1
            student = self.extract_student_info_from_pdf(pdf_path)
            if student:
                if student.student_number not in students_by_number:
                    # Look up Canvas ID for new student
                    canvas_id = self.canvas_api.find_student_canvas_id(course.course_id, student.student_number)
                    if canvas_id:
                        student.canvas_id = canvas_id
                        students_by_number[student.student_number] = student
                        new_count += 1
                        msg = f"Added new student: {student.name} {student.surname}"
                        self.logger.info(msg)
                    else:
                        msg = f"Could not find Canvas ID for student: {student.name} {student.surname} (probably not enrolled in this course)"
                        self.logger.warning(msg)
                else:
                    msg = f"Student already exists: {student.name} {student.surname}"
                    self.logger.info(msg)
        
        # Write updated CSV
        self._write_csv(list(students_by_number.values()), course.csv_file)
        
        summary = f"\nProcessed {pdf_count} PDFs"
        self.logger.info(summary)
        
        summary = f"Added {new_count} new students"
        self.logger.info(summary)

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
        
        # Manage courses button
        ttk.Button(
            options_frame,
            text="Manage Courses",
            command=self.show_course_manager
        ).pack(side=tk.LEFT, padx=5)
        
        # Action buttons frame
        action_frame = ttk.Frame(options_frame)
        action_frame.pack(side=tk.RIGHT)
        
        # Add Configure Course button
        ttk.Button(
            action_frame,
            text="Configure Course",
            command=self.configure_selected_course
        ).pack(side=tk.LEFT, padx=5)
        
        # Add View Extra Time Data button
        ttk.Button(
            action_frame,
            text="View Extra Time Data",
            command=self.view_extra_time_data
        ).pack(side=tk.LEFT, padx=5)
        
        # Move RAP and Extra Time buttons here
        ttk.Button(
            action_frame,
            text="Update from RAP PDFs",
            command=self.update_raps
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            action_frame,
            text="Apply Extra Time",
            command=self.apply_extra_time
        ).pack(side=tk.LEFT, padx=5)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add log text widget
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Add Clear Log button
        ttk.Button(
            log_frame,
            text="Clear Log",
            command=self.clear_log
        ).pack(side=tk.BOTTOM, pady=(5,0))
        
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

    def update_raps(self):
        """Handle updating from RAP PDFs"""
        try:
            self.status_var.set("Processing RAP files...")
            self.root.config(cursor="watch")  # Change cursor to hourglass/watch
            self.root.update()
            self.reader.update_csv_from_raps()
            self.status_var.set("Ready")
        except Exception as e:
            error_msg = f"Failed to process RAP files: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Error occurred")
        finally:
            self.root.config(cursor="")  # Restore default cursor
    
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
                "No student data found. Please process RAP files first."
            )
            self.status_var.set("Ready")
            return
        
        # Read student data
        students = self.reader._read_existing_csv(csv_path)
        if not students:
            messagebox.showerror(
                "Error", 
                "No students found in CSV file. Please process RAP files first."
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
            dialog.transient(self.root)
            
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
            
            # Configure dialog grid weights
            dialog.grid_columnconfigure(0, weight=1)
            dialog.grid_rowconfigure(0, weight=1)
            
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
        dialog.geometry("800x600")
        dialog.transient(self.root)
        
        # Set dialog icon
        self._set_window_icon(dialog)
        
        # Create treeview for courses
        tree = ttk.Treeview(dialog, columns=('id', 'name', 'folder'), show='headings')
        tree.heading('id', text='Course ID')
        tree.heading('name', text='Course Name')
        tree.heading('folder', text='Course RAP Folder')
        tree.column('id', width=100)
        tree.column('name', width=400)
        tree.column('folder', width=200)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(dialog, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Add current courses to treeview
        for course in self.reader.course_manager.courses.values():
            folder = str(course.course_rap_folder) if course.course_rap_folder else "Using shared folder"
            tree.insert('', 'end', values=(course.course_id, course.course_name, folder))
        
        # Layout
        tree.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # Buttons frame
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        def configure_course_from_manager():
            """Configure selected course settings from the course manager"""
            selected = tree.selection()
            if not selected:
                msg = "Please select a course to configure"
                self.logger.warning(msg)
                messagebox.showwarning("Warning", msg)
                return
            
            course_id = str(tree.item(selected[0])['values'][0])  # Ensure course_id is string
            self.configure_course(course_id)
            
            # Update treeview after configuration
            folder = str(self.reader.course_manager.courses[course_id].course_rap_folder) if self.reader.course_manager.courses[course_id].course_rap_folder else "Using shared folder"
            tree.set(selected[0], 'folder', folder)
        
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
                    tree.insert('', 'end', values=(course_id, course_name, "Using shared folder"))
                
                self._update_course_list()
                self.logger.info(f"Successfully added {len(courses)} courses from Canvas")
                messagebox.showinfo("Success", f"Added {len(courses)} courses")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to fetch courses: {str(e)}")
            finally:
                dialog.config(cursor="")
                self.status_var.set("Ready")
        
        ttk.Button(btn_frame, text="Configure", command=configure_course_from_manager).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Refresh from Canvas", command=refresh_courses).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        dialog.grid_rowconfigure(0, weight=1)
        dialog.grid_columnconfigure(0, weight=1)
        
        # Make dialog modal
        dialog.grab_set()
        dialog.focus_set()

    def view_extra_time_data(self):
        """Display the extra time data from the CSV file"""
        if not self.reader.current_course:
            messagebox.showerror("Error", "Please select a course first")
            return
            
        csv_path = self.reader.current_course.csv_file
        if not csv_path.exists():
            self.logger.info("No extra time data found. Please run 'Update from RAP PDFs' first to process RAP files.")
            return
            
        # Create data viewer window
        viewer = tk.Toplevel(self.root)
        viewer.title(f"Extra Time Data - {self.reader.current_course.course_name}")
        viewer.geometry("800x600")
        viewer.transient(self.root)
        
        # Set window icon
        self._set_window_icon(viewer)
        
        # Create treeview
        tree = ttk.Treeview(viewer, columns=('name', 'surname', 'student_number', 'extra_time'), show='headings')
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
        scrollbar = ttk.Scrollbar(viewer, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # Read and display data
        try:
            students = self.reader._read_existing_csv(csv_path)
            for student in students.values():
                tree.insert('', 'end', values=(
                    student.name,
                    student.surname,
                    student.student_number,
                    student.extra_time_per_hour
                ))
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
        about.transient(self.root)
        
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
        
        # Make dialog modal
        about.grab_set()
        about.focus_set()

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
        
        # RAP folder frame
        folder_frame = ttk.LabelFrame(main_frame, text="Shared RAP Folder", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 20))
        
        folder_var = tk.StringVar(value=str(Path.home() / 'RAP'))
        folder_entry = ttk.Entry(folder_frame, textvariable=folder_var, width=50)
        folder_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        def browse_folder():
            folder = filedialog.askdirectory(
                title="Select Shared RAP Folder",
                initialdir=folder_var.get()
            )
            if folder:
                folder_var.set(folder)
        
        ttk.Button(
            folder_frame,
            text="Browse...",
            command=browse_folder
        ).pack(side=tk.LEFT)
        
        def save_config():
            if not token_var.get().strip():
                messagebox.showerror("Error", "Please enter your Canvas API token")
                return
                
            # Create folders if they don't exist
            rap_folder = Path(folder_var.get())
            try:
                rap_folder.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Could not create RAP folder: {e}")
                return
            
            # Save configuration
            config = configparser.ConfigParser()
            config['canvas'] = {
                'access_token': token_var.get().strip(),
                'base_url': 'https://canvas.newcastle.edu.au/'
            }
            config['General'] = {
                'shared_rap_folder': str(rap_folder)
            }
            
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
        setup.transient()
        setup.grab_set()
        setup.focus_set()
        setup.wait_window()

    def configure_selected_course(self):
        """Configure the currently selected course"""
        if not self.reader.current_course:
            messagebox.showerror("Error", "Please select a course first")
            return
        
        course = self.reader.current_course
        self.configure_course(course.course_id)
    
    def configure_course(self, course_id):
        """Configure settings for a specific course"""
        course = self.reader.course_manager.courses[course_id]
        self.logger.debug(f"Configuring course: {course.course_name} (ID: {course_id})")
        
        # Create configuration dialog
        config_dialog = tk.Toplevel(self.root)
        config_dialog.title(f"Configure {course.course_name}")
        config_dialog.geometry("500x200")
        config_dialog.transient(self.root)
        
        # Set dialog icon
        self._set_window_icon(config_dialog)
        
        # Course folder frame
        folder_frame = ttk.LabelFrame(config_dialog, text="Course RAP Folder", padding="5")
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Radio buttons for folder choice
        folder_var = tk.StringVar(value="shared" if not course.course_rap_folder else "specific")
        
        def update_folder_state():
            folder_entry.configure(state='normal' if folder_var.get() == "specific" else 'disabled')
            browse_btn.configure(state='normal' if folder_var.get() == "specific" else 'disabled')
        
        ttk.Radiobutton(
            folder_frame,
            text="Use shared RAP folder",
            variable=folder_var,
            value="shared",
            command=update_folder_state
        ).pack(anchor=tk.W)
        
        specific_frame = ttk.Frame(folder_frame)
        specific_frame.pack(fill=tk.X, pady=5)
        
        ttk.Radiobutton(
            specific_frame,
            text="Use course-specific folder:",
            variable=folder_var,
            value="specific",
            command=update_folder_state
        ).pack(side=tk.LEFT)
        
        folder_entry = ttk.Entry(specific_frame)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        if course.course_rap_folder:
            folder_entry.insert(0, str(course.course_rap_folder))
        
        def browse_folder():
            folder = filedialog.askdirectory(
                title=f"Select RAP folder for {course.course_name}",
                initialdir=str(course.course_rap_folder or self.reader.course_manager.shared_rap_folder)
            )
            if folder:
                folder_entry.delete(0, tk.END)
                folder_entry.insert(0, folder)
        
        browse_btn = ttk.Button(specific_frame, text="Browse...", command=browse_folder)
        browse_btn.pack(side=tk.LEFT)
        
        update_folder_state()
        
        # Buttons
        btn_frame = ttk.Frame(config_dialog)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        
        def save_config():
            if folder_var.get() == "specific":
                folder = Path(folder_entry.get())
                course.course_rap_folder = folder
                # Create folder if it doesn't exist
                if not folder.exists():
                    try:
                        folder.mkdir(parents=True)
                        self.logger.info(f"Created course-specific RAP folder: {folder}")
                    except Exception as e:
                        self.logger.error(f"Failed to create RAP folder: {e}")
                        raise
                else:
                    self.logger.debug(f"Using existing RAP folder: {folder}")
            else:
                course.course_rap_folder = None
                self.logger.info(f"Course {course.course_name} set to use shared RAP folder")
            
            self.reader.course_manager.save_config()
            self.logger.info(f"Updated configuration for course: {course.course_name}")
            config_dialog.destroy()
        
        ttk.Button(btn_frame, text="Save", command=save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=config_dialog.destroy).pack(side=tk.LEFT, padx=5)

    def run(self):
        """Start the GUI"""
        self.root.mainloop()

class TextHandler(logging.Handler):
    """Handler for redirecting logging to tkinter text widget"""
    def __init__(self, text_widget):
        super().__init__()
        self.setLevel(logging.INFO)  # Changed from WARNING to INFO
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
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