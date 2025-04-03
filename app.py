import os
import glob
from openpyxl import load_workbook
import pulp
from collections import defaultdict


def read_excel_files(folder_path, data_cells, data_descs):
    """
    Reads specified cells from all Excel files in a given folder.

    Parameters:
    folder_path (str): Path to the folder containing Excel files.
    data_cells (list): List of cell addresses to read.
    data_descs (list): List of cell addresses for descriptions.

    Returns:
    list: A list of dictionaries containing the extracted data.
    """
    # List to store the output
    students = {}

    # Find all Excel files
    xlsx_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

    for file_path in xlsx_files:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        try:
            wb = load_workbook(filename=file_path, data_only=True)
            sheet = wb.active  # or use wb["SheetName"] if you know it

            # Extract values into a dictionary
            if sheet[data_cells[0]].value is not None:
                students[sheet[data_cells[0]].value] = {}
                for cell, desc in zip(data_cells, data_descs):
                    # Skip the first cell (student ID) as it's already used as the key
                    if cell == data_cells[0]:
                        continue
                    # Check if the cell is empty
                    if sheet[cell].value is None:
                        continue
                    # Add the value to the dictionary with the description as the key
                    students[sheet[data_cells[0]].value][sheet[desc].value] = sheet[
                        cell
                    ].value

        except Exception as e:
            print(f"Error reading {file_path}: {e}")

    return students


if __name__ == "__main__":
    # Example usage
    # Folder containing the Excel files
    folder_path = ""  # <- change this!

    # List of cell addresses to read
    data_cells = [
        "B1",
        "B4",
        "B5",
        "B6",
        "B7",
        "B8",
        "B9",
        "B10",
        "B11",
    ]
    data_descs = [
        "A1",
        "A4",
        "A5",
        "A6",
        "A7",
        "A8",
        "A9",
        "A10",
        "A11",
    ]

    print("Reading Excel files...")
    students = read_excel_files(folder_path, data_cells, data_descs)

    print("Done.")
    for student, preferences in students.items():
        print(f"{student}: {preferences}")

    # students = {
    #     "Alice": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Bob": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Charlie": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "David": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Eve": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Frank": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Grace": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Heidi": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Ivan": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Judy": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Karl": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    #     "Liam": {
    #         "Hip hop": 1,
    #         "Street furniture": 2,
    #         "Rap": 3,
    #         "Upcycling": 4,
    #         "Street food": 5,
    #         "Ultimate": 6,
    #         "Beatz": 7,
    #         "Graffiti": 8,
    #     },
    # }

    # Courses = specific scheduled classes that map to subjects
    courses = {
        "C1": {
            "subject": "Hip hop",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C2": {
            "subject": "Street furniture",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C3": {
            "subject": "Rap",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C4": {
            "subject": "Upcycling",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C5": {
            "subject": "Street food",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C6": {
            "subject": "Ultimate",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C7": {
            "subject": "Beatz",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 20,
        },
        "C8": {
            "subject": "Graffiti",
            "day": "Mon",
            "start": 8,
            "duration": 8,
            "capacity": 20,
        },
        "C11": {
            "subject": "Hip hop",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C12": {
            "subject": "Street furniture",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C13": {
            "subject": "Rap",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C14": {
            "subject": "Upcycling",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C15": {
            "subject": "Street food",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C16": {
            "subject": "Ultimate",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 10,
        },
        "C17": {
            "subject": "Beatz",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 20,
        },
        "C18": {
            "subject": "Graffiti",
            "day": "Tue",
            "start": 8,
            "duration": 8,
            "capacity": 20,
        },
    }

    # Students must be assigned to exactly 6 hours (2 blocks/day × 3 days)
    required_hours_per_student = 16

    # === HELPER FUNCTION ===

    def course_times(course):
        """Return list of (day, hour) slots the course occupies."""
        return [(course["day"], course["start"] + i) for i in range(course["duration"])]

    # === BUILD OPTIMIZATION MODEL ===

    # Decision variables: x[student][course] = 1 if assigned
    x = {
        student: {
            course: pulp.LpVariable(f"x_{student}_{course}", cat="Binary")
            for course in courses
        }
        for student in students
    }

    # Define the problem
    prob = pulp.LpProblem("StudentCourseAssignment", pulp.LpMinimize)

    # === OBJECTIVE FUNCTION ===
    # Minimize total preference score (based on subject ranking)
    prob += pulp.lpSum(
        students[student][courses[course]["subject"]] * x[student][course]
        for student in students
        for course in courses
    )

    # === CONSTRAINTS ===

    # 1. Each course's capacity must not be exceeded
    for course in courses:
        prob += (
            pulp.lpSum(x[student][course] for student in students)
            <= courses[course]["capacity"]
        )

    # 2. Each student must be assigned to exactly 6 hours of courses
    for student in students:
        prob += (
            pulp.lpSum(
                courses[course]["duration"] * x[student][course] for course in courses
            )
            == required_hours_per_student
        )

    # 3. No time slot conflicts (1 course max per (day, hour) per student)
    # Build: (day, hour) → courses
    slot_courses = defaultdict(list)
    for course_name, course_info in courses.items():
        for slot in course_times(course_info):
            slot_courses[slot].append(course_name)

    for student in students:
        for slot, overlapping_courses in slot_courses.items():
            prob += (
                pulp.lpSum(x[student][course] for course in overlapping_courses) <= 1
            )

    # 4. Only one course per subject per student
    # Build: subject → courses
    subject_courses = defaultdict(list)
    for course_name, course_info in courses.items():
        subject_courses[course_info["subject"]].append(course_name)

    for student in students:
        for subject, related_courses in subject_courses.items():
            prob += pulp.lpSum(x[student][course] for course in related_courses) <= 1

    # === SOLVE ===

    prob.solve()

    # === OUTPUT ===

    print("\n📋 Assignment Results:\n")
    for student in students:
        assigned = [course for course in courses if x[student][course].value() == 1]
        total_pref = sum(students[student][courses[c]["subject"]] for c in assigned)
        total_hours = sum(courses[c]["duration"] for c in assigned)
        subjects = set(courses[c]["subject"] for c in assigned)
        print(
            f"{student}: {assigned} | Subjects: {sorted(subjects)} | Total pref = {total_pref}, Hours = {total_hours}"
        )
