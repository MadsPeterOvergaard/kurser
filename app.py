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
    print("Done!")

    # Courses: each with fixed day, start time, duration (in hours), and capacity
    courses = {
        "Kursus 1": {"day": "Mon", "start": 9, "duration": 2, "capacity": 2},
        "Kursus 2": {"day": "Mon", "start": 11, "duration": 2, "capacity": 2},
        "Kursus 3": {"day": "Mon", "start": 13, "duration": 2, "capacity": 2},
        "Kursus 4": {"day": "Tue", "start": 9, "duration": 1, "capacity": 2},
        "Kursus 5": {"day": "Tue", "start": 10, "duration": 2, "capacity": 2},
        "Kursus 6": {"day": "Wed", "start": 9, "duration": 1, "capacity": 2},
        "Kursus 7": {"day": "Wed", "start": 10, "duration": 2, "capacity": 2},
        "Kursus 8": {"day": "Wed", "start": 12, "duration": 1, "capacity": 2},
    }

    # Each student must attend exactly 6 hours total (2 blocks/day * 3 days)
    required_hours_per_student = 6

    # === HELPER FUNCTION ===

    def course_times(course):
        """Return a list of (day, hour) slots this course occupies."""
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

    # Objective: Minimize total preference score
    prob += pulp.lpSum(
        students[student][course] * x[student][course]
        for student in students
        for course in courses
    )

    # Constraint: Each course's capacity must not be exceeded
    for course in courses:
        prob += (
            pulp.lpSum(x[student][course] for student in students)
            <= courses[course]["capacity"]
        )

    # Constraint: Each student must be assigned exactly 6 hours of courses
    for student in students:
        prob += (
            pulp.lpSum(
                courses[course]["duration"] * x[student][course] for course in courses
            )
            == required_hours_per_student
        )

    # Constraint: No overlapping courses (account for multi-hour durations)
    # Step 1: Build a mapping from (day, hour) → courses active during that slot
    slot_courses = defaultdict(list)
    for course_name, course_info in courses.items():
        for slot in course_times(course_info):
            slot_courses[slot].append(course_name)

    # Step 2: For each student and time slot, ensure they take at most one course at that time
    for student in students:
        for slot, course_list in slot_courses.items():
            prob += pulp.lpSum(x[student][course] for course in course_list) <= 1

    # === SOLVE ===

    prob.solve()

    # === OUTPUT ===

    print("\n📋 Assignment Results:\n")
    for student in students:
        assigned = [course for course in courses if x[student][course].value() == 1]
        total_pref = sum(students[student][c] for c in assigned)
        total_hours = sum(courses[c]["duration"] for c in assigned)
        print(
            f"{student}: {assigned} | Total pref = {total_pref}, Total hours = {total_hours}"
        )
