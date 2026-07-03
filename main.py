# ==========================================================
# AI TIMETABLE DRAFTER
# ==========================================================

import pandas as pd
import random
from openpyxl import Workbook

# ==========================================================
# LOAD DATASETS
# ==========================================================

faculty = pd.read_csv("datasets/faculty.csv")
courses = pd.read_csv("datasets/courses.csv")
student_groups = pd.read_csv("datasets/student_groups.csv")
faculty_course_map = pd.read_csv("datasets/faculty_course_map.csv")

print("✅ AI Timetable System Ready")

# ==========================================================
# BASIC DETAILS
# ==========================================================

sections = student_groups["Section"].unique().tolist()

days = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday"
]

periods = [
    "P1","P2","P3","P4",
    "P5","P6","P7","P8"
]

# ==========================================================
# COURSE & FACULTY MAPS
# ==========================================================

course_name_map = dict(
    zip(courses["Course_ID"], courses["Course_Name"])
)

faculty_name_map = dict(
    zip(faculty["Faculty_ID"], faculty["Faculty_Name"])
)

faculty_map = {
    (r["Course_ID"], r["Section"]): r["Faculty_ID"]
    for _, r in faculty_course_map.iterrows()
}

# ==========================================================
# SHORT NAMES
# ==========================================================

short_names = {

    "Software Engineering": "SE",
    "Operating Systems": "OS",
    "Database Management Systems": "DBMS",
    "C++ Programming": "C++",
    "Digital Principles & CO": "Digi",
    "Discrete Mathematics": "DM",
    "Engineering Graphics": "EG",
    "Programming Lab": "PLAB",
    "DBMS Lab": "DBLAB"

}

def short(cid):

    name = course_name_map.get(cid, cid)

    return short_names.get(name, name)

# ==========================================================
# GENERATE TIMETABLE
# ==========================================================

def generate_timetable():

    # Empty timetable

    timetable = {
        sec: {
            day: [None] * 8
            for day in days
        }
        for sec in sections
    }

    # Separate theory and labs

    theory_courses = courses[
        courses["Course_Type"] == "Theory"
    ]

    lab_courses = courses[
        courses["Course_Type"] == "Lab"
    ]

    # ------------------------------------------------------
    # Continue from Part 2...
    # ------------------------------------------------------
        # ==========================================================
    # ASSIGN LABS (3 CONTINUOUS PERIODS)
    # ==========================================================

    lab_slots = [
        [0, 1, 2],
        [3, 4, 5],
        [5, 6, 7]
    ]

    available_days = days.copy()

    random.shuffle(available_days)

    for sec in sections:

        random.shuffle(available_days)

        for i, (_, row) in enumerate(lab_courses.iterrows()):

            cid = row["Course_ID"]

            day = available_days[i % len(available_days)]

            slots = random.choice(lab_slots)

            if all(timetable[sec][day][p] is None for p in slots):

                for p in slots:
                    timetable[sec][day][p] = cid

    # ==========================================================
    # BUILD THEORY SUBJECT POOL
    # ==========================================================

    for sec in sections:

        subject_pool = []

        for _, row in theory_courses.iterrows():

            cid = row["Course_ID"]

            hours = int(row["Weekly_Hours"])

            subject_pool.extend([cid] * hours)

        random.shuffle(subject_pool)

        index = 0

        # ======================================================
        # PLACE THEORY SUBJECTS
        # ======================================================

        for day in days:

            daily_count = {}

            for p in range(8):

                if timetable[sec][day][p] is not None:
                    continue

                attempts = 0

                while attempts < len(subject_pool):

                    cid = subject_pool[index % len(subject_pool)]

                    if daily_count.get(cid, 0) < 2:

                        timetable[sec][day][p] = cid

                        daily_count[cid] = daily_count.get(cid, 0) + 1

                        index += 1

                        break

                    index += 1
                    attempts += 1
                        # ==========================================================
    # FILL REMAINING EMPTY SLOTS
    # ==========================================================

    theory_list = theory_courses["Course_ID"].tolist()

    for sec in sections:

        for day in days:

            for p in range(8):

                if timetable[sec][day][p] is None:

                    timetable[sec][day][p] = random.choice(theory_list)

    print("\n✅ AI Scheduling Completed")

    return timetable

# ==========================================================
# PRINT TIMETABLE
# ==========================================================

def print_timetable(timetable, section):

    print(f"\n\n📅 Section {section}\n")

    for day in days:

        print(f"{day:<10}: ", end="")

        for cid in timetable[section][day]:

            sname = short(cid)

            fid = faculty_map.get((cid, section), "")

            fname = faculty_name_map.get(fid, "")

            if fname != "":
                display = f"{sname}({fname})"
            else:
                display = sname

            print(f"{display:<22}", end=" | ")

        print()

# ==========================================================
# EXPORT TO EXCEL
# ==========================================================

def export_to_excel(timetable, filename="AI_Timetable_Output.xlsx"):

    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    for section in sections:

        ws = wb.create_sheet(title=f"Section {section}")

        ws.append(["Day"] + periods)

        for day in days:

            row = [day]

            for cid in timetable[section][day]:

                sname = short(cid)

                fid = faculty_map.get((cid, section), "")

                fname = faculty_name_map.get(fid, "")

                if fname != "":
                    display = f"{sname} ({fname})"
                else:
                    display = sname

                row.append(display)

            ws.append(row)

    filename = "AI_Timetable_Output.xlsx"

    try:
        wb.save(filename)

        print("\n✅ Excel file generated successfully.")
        print(f"📄 File Name : {filename}")

    except PermissionError:

        print("\n❌ Cannot save Excel file.")
        print("Please close 'AI_Timetable_Output.xlsx' if it is already open in Excel and run the program again.")

if __name__ == "__main__":

    timetable = generate_timetable()

    for sec in sections:
        print_timetable(timetable, sec)

    export_to_excel(timetable)

    print("\n==========================================")
    print(" AI TIMETABLE GENERATED SUCCESSFULLY ")
    print("==========================================")