class Student:
    def __init__(self, name , age , student_id ,student_email , department):
        self.name = name
        self.age = age
        self.student_id = student_id
        self.email = student_email
        self.department = department
        self.grades = {}
        self.attendance = []
        

    def add_grade(self, subject, score):
        self.grades[subject] = score

    def get_total_score(self):
        return sum(self.grades.values())

    def get_average_grade(self):
        if not self.grades:
            return 0
        return self.get_total_score() / len(self.grades)

    def mark_attendance(self, status):
        self.attendance.append(status)

    def get_attendance_percentage(self):
        if not self.attendance:
            return 0
        present_count = self.attendance.count("present")
        return (present_count / len(self.attendance)) * 100

    def get_status(self):
        status_dict = {}
        for subject, score in self.grades.items():
           status_dict[subject] = "PASS" if score >= 50  else "FAIL"
        return status_dict


class Course:
    def __init__(self, course_name):
        self.course_name = course_name
        self.students = []

    def add_student(self, student):
        self.students.append(student)

    def get_course_average(self):
        if not self.students:
            return 0
        total = sum([self.get_average_grade() for s in self.students])
        return total / len(self.students)


class SchoolManagementSystem:
    def __init__(self):
        self.students = {}
        self.courses = {}

    def add_student(self, student):
        self.students[student.student_id] = student

    def add_course(self, course):
        self.courses[course.course_name] = course

    def enroll_student(self, student_id, course_name):
        student = self.students.get(student_id)
        course = self.courses.get(course_name)
        if student and course:
            course.add_student(student)

    def generate_report(self, student_id):
        student = self.students.get(student_id)
        if not student:
            return "Student not found."
        return {
            "Name": student.name,
            "Age": student.age,
            "Department": student.department,
            "Email": student.email,
            "Grades": student.grades,
            "Total Score": student.get_total_score(),
            "Average Grade": student.get_average_grade(),
            "Attendance %": student.get_attendance_percentage(),
            "Status": student.get_status()  }
    def save_reports_to_excel(self, filename="student_reports.xlsx"):
        import openpyxl # type: ignore
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reports"
        ws.append(["Student ID","Name","Age","Department","Email","Grades",
            "Total Score","Average Grade","Attendance %","Status"])

        for student_id, student in self.students.items():
            report = self.generate_report(student_id)
            
            ws.append([
                student_id,
                report["Name"],
                report["Age"],
                report["Department"],
                report["Email"],
                str(report["Grades"]),
                report["Total Score"],
                report["Average Grade"],
                report["Attendance %"],
                str(report["Status"]) ])

        wb.save(filename)
        print(f"Report saved to {filename}")
        


if __name__ == "__main__":
    system = SchoolManagementSystem()
    
    num_courses = int(input("How many courses do you want to add? "))

    for i in range(num_courses):
        course_name = input("Course name: ")
        course = Course(course_name)
        system.add_course(course)

    print("Enter student (type 'Q' for name to quit) : \n")
    while True:
        name = input("Name : ")
        if name == "Q":
            break
        age = int(input("Age : "))
        student_id = int(input("Student_id : "))
        email = f"{student_id}@std.neu.edu.tr"
        print("Email:" , {email})
        department = input("department : ")
        
        
        s = Student(name , age , student_id , email , department)
        system.add_student(s)
        
        

    
        
        for student in system.students:
            system.enroll_student(student_id, course_name)

        print("\n    Mark Attendance   ")
        for student in system.students.values():
            while True:
                status = input(f"Attendance for {student.name}  (present/absent): ")
                student.mark_attendance(status)
                
                more_attendance =input("Add another attendance for this student? (yes / no): ")
                if more_attendance == "no":
                    break

                    
            
        print("\n    Enter Grade   ")
        for course_name in system.courses:
            grade = int(input(f"{course_name} grade for {student.name}: "))
            student.add_grade(course_name, grade)

            
            
        print("\n   Student Report   ")
        for student_id in system.students:
            print(system.generate_report(student_id))
        system.save_reports_to_excel("student_reports.xlsx")
    