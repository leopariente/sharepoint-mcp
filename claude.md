# SharePoint Course Registry

You have access to a SharePoint library for a Computer Science program. When a user asks about a course, you already know exactly where it lives — no browsing needed.

**SharePoint base path:**
`/sites/ComputerScienceLibrary-StudentsTeam2/Shared Documents`

All course paths below are relative to this base.

---

## CS Core Courses

### Year A — Semester A
| Course | Path |
|---|---|
| Intro To Computer Science | `Year A/Semester A/Intro To Computer Science` |
| Calculus 1 | `Year A/Semester A/Calculus 1` |
| Linear Algebra 1 | `Year A/Semester A/Linear Algebra 1` |
| Discrete Math | `Year A/Semester A/Discrete Math` |

### Year A — Semester B
| Course | Path |
|---|---|
| Calculus 2 | `Year A/Semester B/Calculus 2` |
| Linear Algebra 2 | `Year A/Semester B/Linear Algebra 2` |
| Logic & Set Theory | `Year A/Semester B/Logic & Set Theory` |
| System Programming in C | `Year A/Semester B/System Programming in C` |
| Data Structures | `Year A/Semester B/Data Structures` |

### Year B — Semester A
| Course | Path |
|---|---|
| Algorithms | `Year B/Semester A/Algorithms` |
| Intro To Probability | `Year B/Semester A/Intro To Probability` |
| Advanced Programming | `Year B/Semester A/Advanced Programming` |
| Digital Architecture | `Year B/Semester A/Digital Architecture` |

### Year B — Semester B
| Course | Path |
|---|---|
| Operating Systems | `Year B/Semester B/Operating Systems` |
| Computational Models | `Year B/Semester B/Computational Models` |
| Machine Learning | `Year B/Semester B/Machine Learning` |

### Year C — Semester A
| Course | Path |
|---|---|
| Communication Networks | `Year C/Semester A/Communication Networks` |
| Computability & Complexity | `Year C/Semester A/Computability & Complexity` |

### Year C — Semester B
| Course | Path |
|---|---|
| Functional & Logic Programming | `Year C/Semester B/Functional & Logic Programming` |
| Intro To Computer Graphics | `Year C/Semester B/Intro To Computer Graphics` |

---

## Entrepreneurship Courses

### Year A
| Course | Path |
|---|---|
| Entrepreneurial Identity | `Entrepreneurship/Year A/Entrepreneurial Identity` |
| Creativity and Innovation - 0 to 1 | `Entrepreneurship/Year A/Creativity and Innovation - 0 to 1` |
| Tools and Methods for Venture Design | `Entrepreneurship/Year A/Tools and Methods for Venture Design` |
| Wind and Entrepreneurship | `Entrepreneurship/Year A/Wind and Entrepreneurship` |
| Research Methods in Entrepreneurship | `Entrepreneurship/Year A/Research Methods in Entrepreneurship` |

### Year B
| Course | Path |
|---|---|
| Product Management | `Entrepreneurship/Year B/Product Management` |
| Product Design, User Experience and User Research | `Entrepreneurship/Year B/Product Design, User Experience and User Researc` |
| Finance for Entrepreneurs | `Entrepreneurship/Year B/Finance for Entrepreneurs` |
| Marketing for Entrepreneurs | `Entrepreneurship/Year B/Marketing for Entrepreneurs` |
| Legal Aspects and Business Models | `Entrepreneurship/Year B/Legal Aspects and Business Models` |
| Organizational Behavior for Entrepreneurs | `Entrepreneurship/Year B/Organizational Behavior for Entrepreneurs` |

### Year C
| Course | Path |
|---|---|
| Human Capital Management for Entrepreneurs | `Entrepreneurship/Year C/Human Capital Management for Entrepreneurs` |
| Venture Decisions | `Entrepreneurship/Year C/Venture Decisions` |
| Practicing Entrepreneurship | `Entrepreneurship/Year C/Practicing Entrepreneurship` |

---

## Elective Courses
| Course | Path |
|---|---|
| Elective Course 1 | `Elective Courses - Documents/Elective Course 1` |
| Elective Course 2 | `Elective Courses - Documents/Elective Course 2` |
| Elective Course 3 | `Elective Courses - Documents/Elective Course 3` |
| Elective Course 4 | `Elective Courses - Documents/Elective Course 4` |

---

## Instructions

- When a user mentions a course by name (or a close match), look it up in the tables above and use the full SharePoint path directly — do not browse SharePoint to find it.
- If a course is not in the registry, say so and offer to browse SharePoint to locate it.
- To add a new course to this registry, append a row to the appropriate table above with the course name and its SharePoint path.
- Elective course folders are generic (`Elective Course 1`, etc.) — if the user tells you which elective maps to which slot, update the table accordingly.

## Finding Homework Questions

When a user asks to find a specific homework question, exercise, or problem in the SharePoint course library — for example "find the question about X in exercise/homework Y", "find exercise 7 about graphs", "look for a problem about dynamic programming in algorithms" — follow the full search flow below.

---

### Step 1 — Identify What You're Looking For

From the user's request, extract:
- **Course name** (e.g. "Algorithms", "Advanced Programming", "Intro to Probability")
- **Homework / Exercise number** (e.g. HW7, Exercise 7, Assignment 7)
- **Topic hint** (optional, e.g. "about kids and cities", "network flow", "dynamic programming")

If the course is ambiguous, ask the user before proceeding.

---

### Step 2 — Find the Right Folder

Use the Course Registry above to get the SharePoint path for the course directly — no browsing needed. Inside the course folder, there are year subfolders: `2021/`, `2022/`, `2023/`, `2024/`, `2025/`, `2026/`. Inside each year, look for an `Exercises/` subfolder.

Folder path pattern:
```
/sites/ComputerScienceLibrary-StudentsTeam2/Shared Documents/
  {course path} / {year} / Exercises / {filename}.pdf
```

---

### Step 3 — Collect Candidate PDF Files

Use `list_files` on the `Exercises/` folder for the most recent years first (2025, 2024, 2023). Filter to files whose name contains the homework number (e.g. `HW7`, `hw7`, `Assignment 7`, `HW_7`).

Example matching files:
```
ElizavetaKhanan_HW7_82.pdf
OhadSwissa_HW7_97.pdf
Assignment 7 - Rotem Haim 100.pdf
```

---

### Step 4 — Read PDFs Smartly

**CRITICAL: Most PDFs are large (>1MB) scanned files. Never read all pages.**

Always call:
```
read_pdf(file_url=..., pages="1-5", max_chars=8000)
```

- The question statement is on the first 1–3 pages; student answers follow after
- `search_content` will not work for scanned PDFs — always use `read_pdf` directly
- **Stop as soon as you find the question** — do not keep reading more files

---

### Step 5 — Extract and Present the Question

- Quote the exact question text, including all sub-parts
- Mention which file and year it came from
- If not found in one year, try the next (2025 → 2024 → 2023)