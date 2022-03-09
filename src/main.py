import os
import xlrd

# Constants
SEASON_FALL = '7'
SEASON_SPRING = '3'
SEASON_SUMMER = '5'

CELLKEY_COURSE_TERM = 0
CELLKEY_COURSE_ID = 4
CELLKEY_COURSE_SUBJECT = 7
CELLKEY_COURSE_CATALOG = 8
CELLKEY_COURSE_DESC = 10
CELLKEY_SECTION_NUMBER = 9
CELLKEY_SECTION_FACILITY = 19
CELLKEY_SECTION_WEEKDAYS = 20
CELLKEY_SECTION_ENROLLED_CAP = 14
CELLKEY_SECTION_ENROLLED_TOTAL = 15
CELLKEY_SECTION_WAITLIST_CAP = 16
CELLKEY_SECTION_WAITLIST_TOTAL = 17
CELLKEY_INSTRUCTOR_LAST = 23
CELLKEY_INSTRUCTOR_FIRST = 24

# Classes
class AcademicYear:
  def __init__(self, first_half_year, second_half_year):
    self.first_half_year = str(first_half_year)
    self.second_half_year = str(second_half_year)
    self.courses = {
        "Fall": [],
        "Spring": [],
        "Summer": []
    }

class Section:
    def __init__(self, course, instructor, number, facility, weekdays, enrolled_total):
        self.course = course
        self.instructor = instructor
        self.number = number.strip()
        self.facility = facility.strip()
        self.weekdays = weekdays.strip()
        self.enrolled_total = enrolled_total

    def __eq__(self, other):
        if isinstance(other, Section):
            return self.number == other.number and self.course.id == other.course.id
        else:
            return False

class Course:
    def __init__(self, id, subject, catalog, desc):
        self.id = id.strip()
        self.subject = subject.strip()
        self.catalog = catalog.strip()
        self.desc = desc.strip()
        self.sections = []

    def __eq__(self, other):
        if isinstance(other, Course):
            return self.id == other.id
        else:
            return False

class Instructor:
    def __init__(self, name_last, name_first):
        self.name_last = name_last.strip()
        self.name_first = name_first.strip()

    def __eq__(self, other):
        if isinstance(other, Instructor):
            return self.name_last == other.name_last and self.name_first == other.name_first
        else:
            return False

# Main
academic_years = [
    AcademicYear(2018, 2019),
    AcademicYear(2019, 2020),
    AcademicYear(2020, 2021)
]

# Per File
for filename in os.listdir(os.getcwd() + ".\input"):

    # Skip non xlsx files.
    if not filename.endswith(".xlsx"): continue

    workbook = xlrd.open_workbook(os.path.join(".\input", filename))
    worksheet = workbook.sheet_by_index(0)

    year_and_season = worksheet.cell(2, CELLKEY_COURSE_TERM).value
    year_value = year_and_season[0] + "0" + year_and_season[1:3]

    current_academic_year = None
    current_season = None

    # Find Academic Year instance.
    for year in academic_years:
        if year.first_half_year == year_value and year_and_season[3] == SEASON_FALL:
            # print("1. " + filename)
            current_academic_year = year
        elif year.second_half_year == year_value:
            # print("2. " + filename)
            current_academic_year = year

    # Find the semester season.
    if year_and_season[3] == SEASON_FALL:
        current_season = "Fall"
    elif year_and_season[3] == SEASON_SPRING:
        current_season = "Spring"
    else:
        current_season = "Summer"

    # Something is causing an invalid year not sure what
    if current_academic_year is None: continue

    # For Row, Skips the first 2 rows.
    for rx in range(2, worksheet.nrows):

        # Extract data from cells
        CELLDATA_Course_ID = worksheet.cell(rx, CELLKEY_COURSE_ID).value
        CELLDATA_Course_SUBJECT = worksheet.cell(rx, CELLKEY_COURSE_SUBJECT).value
        CELLDATA_Course_CATALOG = worksheet.cell(rx, CELLKEY_COURSE_CATALOG).value
        CELLDATA_Course_DESC = worksheet.cell(rx, CELLKEY_COURSE_DESC).value
        CELLDATA_Section_NUMBER = worksheet.cell(rx, CELLKEY_SECTION_NUMBER).value
        CELLDATA_Section_FACILITY = worksheet.cell(rx, CELLKEY_SECTION_FACILITY).value
        CELLDATA_Section_WEEKDAYS = worksheet.cell(rx, CELLKEY_SECTION_WEEKDAYS).value
        CELLDATA_Section_WAITLIST_CAP = worksheet.cell(rx, CELLKEY_SECTION_WAITLIST_CAP).value
        CELLDATA_Section_WAITLIST_TOTAL = worksheet.cell(rx, CELLKEY_SECTION_WAITLIST_TOTAL).value
        CELLDATA_Section_ENROLLED_CAP = worksheet.cell(rx, CELLKEY_SECTION_ENROLLED_CAP).value
        CELLDATA_Section_ENROLLED_TOTAL = worksheet.cell(rx, CELLKEY_SECTION_ENROLLED_TOTAL).value
        CELLDATA_Instructor_LAST = worksheet.cell(rx, CELLKEY_INSTRUCTOR_LAST).value
        CELLDATA_Instructor_FIRST = worksheet.cell(rx, CELLKEY_INSTRUCTOR_FIRST).value

        instructor_instance = Instructor(CELLDATA_Instructor_LAST, CELLDATA_Instructor_FIRST)
        course_instance = Course(CELLDATA_Course_ID, CELLDATA_Course_SUBJECT, CELLDATA_Course_CATALOG, CELLDATA_Course_DESC)
        section_instance = Section(course_instance, instructor_instance, CELLDATA_Section_NUMBER, CELLDATA_Section_FACILITY, CELLDATA_Section_WEEKDAYS, CELLDATA_Section_ENROLLED_TOTAL)

        if course_instance in current_academic_year.courses[current_season]:
            index = current_academic_year.courses[current_season].index(course_instance)
            found_course = current_academic_year.courses[current_season][index]

            if (section_instance in found_course.sections):
                index = found_course.sections.index(section_instance)
                found_section = found_course.sections[index]

                found_section.enrolled_total = found_section.enrolled_total
            else:
                found_course.sections.append(section_instance)
        else:
            course_instance.sections.append(section_instance)
            current_academic_year.courses[current_season].append(course_instance)

# Output
for academic_year in academic_years:
    fname = str(academic_year.first_half_year + "_" + academic_year.second_half_year)
    output = open(os.getcwd() + ".\output\%s.csv" % fname, "w")

    for course in academic_year.courses["Fall"]:
        for section in course.sections:
            output.write(course.subject + course.catalog + ", ")
            output.write(course.desc + ", ")
            output.write(section.number + ", ")
            output.write(academic_year.first_half_year[2:4] + "-" + academic_year.second_half_year[2:4] + ", ")
            output.write("F" + academic_year.first_half_year + ", ")
            output.write(section.weekdays + ", ")
            output.write(section.facility + ", ")
            output.write(str(round(section.enrolled_total)) + ", ")
            output.write(section.instructor.name_first + ", ")
            output.write(section.instructor.name_last)
            output.write("\n")

    for course in academic_year.courses["Spring"]:
        for section in course.sections:
            output.write(course.subject + course.catalog + ", ")
            output.write(course.desc + ", ")
            output.write(section.number + ", ")
            output.write(academic_year.first_half_year[2:4] + "-" + academic_year.second_half_year[2:4] + ", ")
            output.write("S" + academic_year.second_half_year + ", ")
            output.write(section.weekdays + ", ")
            output.write(section.facility + ", ")
            output.write(str(round(section.enrolled_total)) + ", ")
            output.write(section.instructor.name_first + ", ")
            output.write(section.instructor.name_last)
            output.write("\n")

    for course in academic_year.courses["Summer"]:
        for section in course.sections:
            output.write(course.subject + course.catalog + ", ")
            output.write(course.desc + ", ")
            output.write(section.number + ", ")
            output.write(academic_year.first_half_year[2:4] + "-" + academic_year.second_half_year[2:4] + ", ")
            output.write("SUM" + academic_year.second_half_year + ", ")
            output.write(section.weekdays + ", ")
            output.write(section.facility + ", ")
            output.write(str(round(section.enrolled_total)) + ", ")
            output.write(section.instructor.name_first + ", ")
            output.write(section.instructor.name_last)
            output.write("\n")

    output.close()