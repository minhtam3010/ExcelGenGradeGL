import pandas as pd
import xlsxwriter

sheet_names_for7_8 = ["7A", "7B", "7C", "7D", "7E", "8A", "8B", "8C"]
sheet_names_for9_10 = ["9A", "9B", "9C", "9D", "9E", "9F", "10D"]
sheet_names_for10 = ["10A", "10B", "10C"]


attendance_export = "attendance-export_%s.xlsx"

initialSentence = "Giáo Làng thông báo kết quả học tập của bạn %s tại lớp %s như sau:\n"
exellentStudent = """Đánh giá chung: Giỏi
- Điểm yếu: từ vựng (thuộc 80%), nghe (70%). 
- Khuyến nghị: học từ vựng đều đặn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Xuất sắc
    - Nói: Khá
    - Nghe: Khá 
    - Viết: Tốt
    - Ngữ pháp: Tốt
"""

goodStudent = """Đánh giá chung: Khá
- Điểm yếu: từ vựng (thuộc 70%), nghe (60%). 
- Khuyến nghị: học từ vựng đều đặn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Tốt
    - Nói: Khá
    - Nghe: Trung bình-Khá
    - Viết: Khá
    - Ngữ pháp: Tốt
"""

averageStudent = """Đánh giá chung: Trung bình-Khá
- Điểm yếu: từ vựng (thuộc chỉ 60%), nghe (khoảng 50%). 
- Khuyến nghị: Học kỹ từ vựng hơn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Khá 
    - Nói: Trung bình
    - Nghe: Trung bình
    - Viết: Trung bình-Khá
    - Ngữ pháp: Khá
"""

badStudent = """Đánh giá chung: Trung bình-Yếu
- Điểm yếu: từ vựng (thuộc chỉ 50%), nghe (dưới 50%). 
- Khuyến nghị: Học kỹ từ vựng hơn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Trung bình
    - Nói: Trung bình-Yếu
    - Nghe: Yếu 
    - Viết: Trung bình
    - Ngữ pháp: Trung bình
"""

exellentStudent2 = """Đánh giá chung: Xuất Sắc
- Điểm yếu: từ vựng (thuộc 80%), cấu trúc từ vựng (75%), gốc của từ (80%). 
- Khuyến nghị: học từ vựng đều đặn, xem lại toàn bộ bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc & Điền khuyết: Tốt
    - Từ vựng: Xuất sắc
    - Ngữ pháp: Tốt
    - Viết lại câu: Xuất sắc 
    - Phát âm & dấu nhấn: Khá
"""

goodStudent2 = """Đánh giá chung: Tốt
- Điểm yếu: từ vựng (thuộc 75%), cấu trúc từ vựng (65%), gốc của từ (70%). 
- Khuyến nghị: học từ vựng đều đặn, xem lại toàn bộ bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc & Điền khuyết: Khá
    - Từ vựng: Khá-Tốt
    - Ngữ pháp: Tốt
    - Viết lại câu: Khá
    - Phát âm & dấu nhấn: Khá
"""

averageStudent2 = """Đánh giá chung: Trung bình-Khá
- Điểm yếu: từ vựng (thuộc chỉ 65%), cấu trúc từ vựng (55%), gốc của từ ( thuộc 60%). 
- Khuyến nghị: học từ vựng đều đặn, xem lại toàn bộ bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc & Điền khuyết: Trung bình-Khá
    - Từ vựng: Khá
    - Ngữ pháp: Trung bình-Khá
    - Viết lại câu: Trung bình
    - Phát âm & dấu nhấn: Trung bình-Khá
"""

badStudent2 = """Đánh giá chung: Trung bình-Yếu
- Điểm yếu: từ vựng (thuộc chỉ 55%), cấu trúc từ vựng (dưới 50%), gốc của từ ( thuộc chỉ 50%). 
- Khuyến nghị: học từ vựng đều đặn, xem lại toàn bộ bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc & Điền khuyết: Trung bình
    - Từ vựng: Trung bình-Khá
    - Ngữ pháp: Trung bình
    - Viết lại câu: Trung bình-Yếu
    - Phát âm & dấu nhấn: Trung bình-Yếu
"""

exellentStudent3 = """Đánh giá chung: Giỏi
- Điểm yếu: từ vựng (thuộc 80%), nghe (70%). 
- Khuyến nghị: học từ vựng đều đặn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Tốt
    - Nói: Khá
    - Nghe: Trung bình-Khá
    - Viết: Khá
    - Ngữ pháp: Tốt
"""

goodStudent3 = """Đánh giá chung: Khá
- Điểm yếu: từ vựng (thuộc 70%), nghe (60%). 
- Khuyến nghị: học từ vựng đều đặn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Khá 
    - Nói: Trung bình
    - Nghe: Trung bình
    - Viết: Trung bình-Khá
    - Ngữ pháp: Khá
"""

averageStudent3 = """Đánh giá chung: Trung bình
- Điểm yếu: từ vựng (thuộc chỉ 60%), nghe (khoảng 50%). 
- Khuyến nghị: Học kỹ từ vựng hơn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Trung bình-Khá
    - Nói: Trung bình
    - Nghe: Yếu 
    - Viết: Trung bình
    - Ngữ pháp: Trung bình-Khá
"""

badStudent3 = """Đánh giá chung: Yếu
- Điểm yếu: từ vựng (thuộc chỉ 50%), nghe (dưới 50%). 
- Khuyến nghị: Học kỹ từ vựng hơn, luyện nghe thêm tại nhà, xem lại các bài tập đã được sửa ở lớp.

Đánh giá chi tiết: 
    - Đọc: Trung bình
    - Nói: Trung bình-Yếu
    - Nghe: Yếu 
    - Viết: Trung bình-Yếu
    - Ngữ pháp: Trung bình
"""

path = "/Users/minhhtamm/Working/Excel/new_grade.xlsx"
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')

# writer = pd.ExcelWriter("/Users/minhhtamm/Working/Excel/new_grade.xlsx", engine = 'xlsxwriter')

attendanceExclusive = []

def checkGrade(grade, studentName, className, new_dict):
    try:
        attendance =  "Chuyên cần:\n" + new_dict[studentName]
    except:
        attendanceExclusive.append(studentName)
        attendance = "Chuyên cần:\n"
    result = initialSentence % (studentName, className)
    try:
        grade = str(grade).replace(",", ".")
        averageGrade = "Điểm kiểm tra chất lượng khoá hè: %s\n" % grade
        grade = float(grade)
        if grade >= 8:
            return result + attendance + "\n\n" + averageGrade + exellentStudent
        elif grade >= 6.5 and grade < 8:
            return result + attendance + "\n\n" + averageGrade + goodStudent
        elif grade >= 5 and grade < 6.5:
            return result + attendance + "\n\n" + averageGrade + averageStudent
        else:
            return result + attendance + "\n\n" + averageGrade + badStudent
    except:
        return result + attendance + "\n\n" + "No data"
    
def checkGrade2(grade, studentName, className, new_dict):
    try:
        attendance =  "Chuyên cần:\n" + new_dict[studentName]
    except:
        attendanceExclusive.append(studentName)
        attendance = "Chuyên cần:\n"
    result = initialSentence % (studentName, className)
    try:
        grade = str(grade).replace(",", ".")
        averageGrade = "Điểm kiểm tra chất lượng khoá hè: %s\n" % grade
        grade = float(grade)
        if grade >= 8.5:
            return result + attendance + "\n\n" + averageGrade + exellentStudent2
        elif grade >= 7 and grade < 8.5:
            return result + attendance + "\n\n" + averageGrade + goodStudent2
        elif grade >= 5.5 and grade < 7:
            return result + attendance + "\n\n" + averageGrade + averageStudent2
        else:
            return result + attendance + "\n\n" + averageGrade + badStudent2
    except:
        return result + attendance + "\n\n" + "No data"

def checkGrade3(grade, studentName, className, new_dict):
    try:
        attendance =  "Chuyên cần:\n" + new_dict[studentName]
    except:
        attendanceExclusive.append(studentName)
        attendance = "Chuyên cần:\n"
    result = initialSentence % (studentName, className)
    try:
        grade = str(grade).replace(",", ".")
        averageGrade = "Điểm kiểm tra chất lượng khoá hè: %s\n" % grade
        grade = float(grade)
        if grade >= 8:
            return result + attendance + "\n\n" + averageGrade + exellentStudent3
        elif grade >= 6.5 and grade < 8:
            return result + attendance + "\n\n" + averageGrade + goodStudent3
        elif grade >= 5 and grade < 6.5:
            return result + attendance + "\n\n" + averageGrade + averageStudent3
        else:
            return result + attendance + "\n\n" + averageGrade + badStudent3
    except:
        return result + attendance + "\n\n" + "No data"
    

def getPhone(phone_dict, fullname):
    try:
        return phone_dict[fullname]
    except:
        return "Ko tra đc số điện thoại"
    
for s in sheet_names_for7_8:
    df = pd.read_excel('grade.xlsx', sheet_name=s)
    df_attendance = pd.read_excel(attendance_export % s)
    df_attendance.rename(columns={df_attendance.columns[4]: "Message"}, inplace=True)
    df_attendance.rename(columns={df_attendance.columns[1]: "StudentName"}, inplace=True)
    df_attendance.rename(columns={df_attendance.columns[2]: "Phone"}, inplace=True)

    new_dict = {}
    phone_dict = {}
    df_attendance["StudentName"] = df_attendance["StudentName"].str.strip()

    for index, row in df_attendance.iterrows():
        result = row["Message"].split("\n")[1:]
        # result = "\n".join(result)
        newResult = ""
        absence = 0
        for i in range(len(result)):
            split = result[i].split(":")
            if "vắng" in split[0]:
                absence = int(split[1].strip())
            else:
                newResult += result[i] + "\n"
        
        if absence > 0:
            newResult += "- Tổng số buổi vắng: %s\n" % absence
        new_dict[row["StudentName"]] = newResult
        phone_dict[row["StudentName"]] = "0" + str(row["Phone"])
    if s == "7D":
        new_dict["Bùi Thị Quỳnh Chi"] = """Giáo Làng thông báo kết quả điểm chuyên cần của học sinh Bùi Thị Quỳnh Chi trong lớp LOP7-E:
	- Tổng số buổi vắng không phép: 1
	- Tổng số buổi đi học: 10"""
    
    # Rename the third column to "Grade"
    df.rename(columns={df.columns[3]: "Grade"}, inplace=True)
    df.rename(columns={df.columns[2]: "FirstName"}, inplace=True)
    df.rename(columns={df.columns[1]: "LastName"}, inplace=True)
    className = s
    df["phone"] = df.apply(lambda x: getPhone(phone_dict, str(x["LastName"]) + " " + str(x["FirstName"])), axis=1)

    df["feedback"] = df.apply(lambda x: checkGrade(x["Grade"], str(x["LastName"]) + " " + str(x["FirstName"]), s, new_dict=new_dict), axis=1)
    df.to_excel(writer, sheet_name=s, index=False)

for s in sheet_names_for9_10:
    df = pd.read_excel('grade.xlsx', sheet_name=s)
    df_attendance = pd.read_excel(attendance_export % s)
    df_attendance.rename(columns={df_attendance.columns[4]: "Message"}, inplace=True)
    df_attendance.rename(columns={df_attendance.columns[1]: "StudentName"}, inplace=True)
    df_attendance.rename(columns={df_attendance.columns[2]: "Phone"}, inplace=True)

    new_dict = {}
    phone_dict = {}
    df_attendance["StudentName"] = df_attendance["StudentName"].str.strip()

    for index, row in df_attendance.iterrows():
        result = row["Message"].split("\n")[1:]
                # result = "\n".join(result)
        newResult = ""
        absence = 0
        for i in range(len(result)):
            split = result[i].split(":")
            if "vắng" in split[0]:
                absence = int(split[1].strip())
            else:
                newResult += result[i] + "\n"
        
        if absence > 0:
            newResult += "- Tổng số buổi vắng: %s\n" % absence
        # result = "\n".join(result)
        new_dict[row["StudentName"]] = newResult
        phone_dict[row["StudentName"]] = "0" + str(row["Phone"])
    
    if s == "7D":
        new_dict["Bùi Thị Quỳnh Chi"] = """
	- Tổng số buổi vắng: 1
	- Tổng số buổi đi học: 10"""
    

    # Rename the third column to "Grade"
    if "10" in s:
        df.rename(columns={df.columns[3]: "Grade"}, inplace=True)
        df.rename(columns={df.columns[2]: "FirstName"}, inplace=True)
        df.rename(columns={df.columns[1]: "LastName"}, inplace=True)
        className = s
        df["phone"] = df.apply(lambda x: getPhone(phone_dict, str(x["LastName"]) + " " + str(x["FirstName"])), axis=1)
        df["feedback"] = df.apply(lambda x: checkGrade2(x["Grade"], str(x["LastName"]) + " " + str(x["FirstName"]), s, new_dict=new_dict), axis=1)
    else:
        df.rename(columns={df.columns[2]: "Grade"}, inplace=True)
        df.rename(columns={df.columns[1]: "FullName"}, inplace=True)
        className = s
        df["phone"] = df.apply(lambda x: getPhone(phone_dict, str(x["FullName"])), axis=1)
        df["feedback"] = df.apply(lambda x: checkGrade2(x["Grade"], str(x["FullName"]), s, new_dict=new_dict), axis=1)

    df.to_excel(writer, sheet_name=s, index=False)

for s in sheet_names_for10:
    df = pd.read_excel('grade.xlsx', sheet_name=s)
    df_attendance = pd.read_excel(attendance_export % s)
    df_attendance.rename(columns={df_attendance.columns[4]: "Message"}, inplace=True)
    df_attendance.rename(columns={df_attendance.columns[1]: "StudentName"}, inplace=True)
    df_attendance.rename(columns={df_attendance.columns[2]: "Phone"}, inplace=True)

    new_dict = {}
    phone_dict = {}
    df_attendance["StudentName"] = df_attendance["StudentName"].str.strip()

    for index, row in df_attendance.iterrows():
        result = row["Message"].split("\n")[1:]
                # result = "\n".join(result)
        newResult = ""
        absence = 0
        for i in range(len(result)):
            split = result[i].split(":")
            if "vắng" in split[0]:
                absence = int(split[1].strip())
            else:
                newResult += result[i] + "\n"
        
        if absence > 0:
            newResult += "- Tổng số buổi vắng: %s\n" % absence
        # result = "\n".join(result)
        new_dict[row["StudentName"]] = newResult
        phone_dict[row["StudentName"]] = "0" + str(row["Phone"])

    if s == "7D":
        new_dict["Bùi Thị Quỳnh Chi"] = """Giáo Làng thông báo kết quả điểm chuyên cần của học sinh Bùi Thị Quỳnh Chi trong lớp LOP7-E:
	- Tổng số buổi vắng không phép: 1
	- Tổng số buổi đi học: 10"""
    
    # Rename the third column to "Grade"
    if "10" in s:
        df.rename(columns={df.columns[3]: "Grade"}, inplace=True)
        df.rename(columns={df.columns[2]: "FirstName"}, inplace=True)
        df.rename(columns={df.columns[1]: "LastName"}, inplace=True)
        className = s
        df["phone"] = df.apply(lambda x: getPhone(phone_dict, str(x["LastName"]) + " " + str(x["FirstName"])), axis=1)
        df["feedback"] = df.apply(lambda x: checkGrade3(x["Grade"], str(x["LastName"]) + " " + str(x["FirstName"]), s, new_dict=new_dict), axis=1)
    else:
        df.rename(columns={df.columns[2]: "Grade"}, inplace=True)
        df.rename(columns={df.columns[1]: "FullName"}, inplace=True)
        className = s
        df["phone"] = df.apply(lambda x: getPhone(phone_dict, str(x["FullName"])), axis=1)
        df["feedback"] = df.apply(lambda x: checkGrade3(x["Grade"], str(x["FullName"]), s, new_dict=new_dict), axis=1)

    df.to_excel(writer, sheet_name=s, index=False)

writer.close()

print(attendanceExclusive)
# How to write to excel file with separating sheets
# https://stackoverflow.com/questions/42370977/how-to-write-to-excel-file-with-separating-sheets
