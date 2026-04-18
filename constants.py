import os
import threading
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(dotenv_path=Path(__file__).parent / ".env")

SHAREPOINT_SITE_URL = "https://postidcac.sharepoint.com/sites/ComputerScienceLibrary-StudentsTeam2"
SHAREPOINT_BASE_FOLDER = "/sites/ComputerScienceLibrary-StudentsTeam2/Shared Documents"
STUDENT_EMAIL = os.getenv("SHAREPOINT_EMAIL")
STUDENT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")

_RETRY_STATUS = ("429", "503", "Too Many Requests", "Service Unavailable")

# Concurrency cap across all SharePoint-hitting tools. Prevents cascade 429s
# when Claude fires parallel tool calls (e.g. multiple read_pdf).
_SP_CONCURRENCY = int(os.getenv("SHAREPOINT_MAX_CONCURRENCY", "4"))
_sp_semaphore = threading.Semaphore(_SP_CONCURRENCY)

_REGISTRY: list[dict] = [
    # CS Core — Year A
    {"name": "Intro To Computer Science",        "year": "A", "semester": "A", "path": "Year A/Semester A/Intro To Computer Science"},
    {"name": "Calculus 1",                        "year": "A", "semester": "A", "path": "Year A/Semester A/Calculus 1"},
    {"name": "Linear Algebra 1",                  "year": "A", "semester": "A", "path": "Year A/Semester A/Linear Algebra 1"},
    {"name": "Discrete Math",                     "year": "A", "semester": "A", "path": "Year A/Semester A/Discrete Math"},
    {"name": "Calculus 2",                        "year": "A", "semester": "B", "path": "Year A/Semester B/Calculus 2"},
    {"name": "Linear Algebra 2",                  "year": "A", "semester": "B", "path": "Year A/Semester B/Linear Algebra 2"},
    {"name": "Logic & Set Theory",                "year": "A", "semester": "B", "path": "Year A/Semester B/Logic & Set Theory"},
    {"name": "System Programming in C",           "year": "A", "semester": "B", "path": "Year A/Semester B/System Programming in C"},
    {"name": "Data Structures",                   "year": "A", "semester": "B", "path": "Year A/Semester B/Data Structures"},
    # CS Core — Year B
    {"name": "Algorithms",                        "year": "B", "semester": "A", "path": "Year B/Semester A/Algorithms"},
    {"name": "Intro To Probability",              "year": "B", "semester": "A", "path": "Year B/Semester A/Intro To Probability"},
    {"name": "Advanced Programming",              "year": "B", "semester": "A", "path": "Year B/Semester A/Advanced Programming"},
    {"name": "Digital Architecture",              "year": "B", "semester": "A", "path": "Year B/Semester A/Digital Architecture"},
    {"name": "Operating Systems",                 "year": "B", "semester": "B", "path": "Year B/Semester B/Operating Systems"},
    {"name": "Computational Models",              "year": "B", "semester": "B", "path": "Year B/Semester B/Computational Models"},
    {"name": "Machine Learning",                  "year": "B", "semester": "B", "path": "Year B/Semester B/Machine Learning"},
    # CS Core — Year C
    {"name": "Communication Networks",            "year": "C", "semester": "A", "path": "Year C/Semester A/Communication Networks"},
    {"name": "Computability & Complexity",        "year": "C", "semester": "A", "path": "Year C/Semester A/Computability & Complexity"},
    {"name": "Functional & Logic Programming",    "year": "C", "semester": "B", "path": "Year C/Semester B/Functional & Logic Programming"},
    {"name": "Intro To Computer Graphics",        "year": "C", "semester": "B", "path": "Year C/Semester B/Intro To Computer Graphics"},
    # Entrepreneurship — Year A
    {"name": "Entrepreneurial Identity",                    "year": "A", "semester": None, "path": "Entrepreneurship/Year A/Entrepreneurial Identity"},
    {"name": "Creativity and Innovation - 0 to 1",          "year": "A", "semester": None, "path": "Entrepreneurship/Year A/Creativity and Innovation - 0 to 1"},
    {"name": "Tools and Methods for Venture Design",        "year": "A", "semester": None, "path": "Entrepreneurship/Year A/Tools and Methods for Venture Design"},
    {"name": "Wind and Entrepreneurship",                   "year": "A", "semester": None, "path": "Entrepreneurship/Year A/Wind and Entrepreneurship"},
    {"name": "Research Methods in Entrepreneurship",        "year": "A", "semester": None, "path": "Entrepreneurship/Year A/Research Methods in Entrepreneurship"},
    # Entrepreneurship — Year B
    {"name": "Product Management",                          "year": "B", "semester": None, "path": "Entrepreneurship/Year B/Product Management"},
    {"name": "Product Design, User Experience and User Research", "year": "B", "semester": None, "path": "Entrepreneurship/Year B/Product Design, User Experience and User Researc"},
    {"name": "Finance for Entrepreneurs",                   "year": "B", "semester": None, "path": "Entrepreneurship/Year B/Finance for Entrepreneurs"},
    {"name": "Marketing for Entrepreneurs",                 "year": "B", "semester": None, "path": "Entrepreneurship/Year B/Marketing for Entrepreneurs"},
    {"name": "Legal Aspects and Business Models",           "year": "B", "semester": None, "path": "Entrepreneurship/Year B/Legal Aspects and Business Models"},
    {"name": "Organizational Behavior for Entrepreneurs",   "year": "B", "semester": None, "path": "Entrepreneurship/Year B/Organizational Behavior for Entrepreneurs"},
    # Entrepreneurship — Year C
    {"name": "Human Capital Management for Entrepreneurs",  "year": "C", "semester": None, "path": "Entrepreneurship/Year C/Human Capital Management for Entrepreneurs"},
    {"name": "Venture Decisions",                           "year": "C", "semester": None, "path": "Entrepreneurship/Year C/Venture Decisions"},
    {"name": "Practicing Entrepreneurship",                 "year": "C", "semester": None, "path": "Entrepreneurship/Year C/Practicing Entrepreneurship"},
    # Electives
    {"name": "Elective Course 1", "year": None, "semester": None, "path": "Elective Courses - Documents/Elective Course 1"},
    {"name": "Elective Course 2", "year": None, "semester": None, "path": "Elective Courses - Documents/Elective Course 2"},
    {"name": "Elective Course 3", "year": None, "semester": None, "path": "Elective Courses - Documents/Elective Course 3"},
    {"name": "Elective Course 4", "year": None, "semester": None, "path": "Elective Courses - Documents/Elective Course 4"},
]
