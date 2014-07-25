using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace JH.IBSH.Report.PeriodicalExam
{
    
    public class GradePeriodicalExamGPARecord 
    {
        public GradePeriodicalExamGPARecord(int SchoolYear, int Semester, int Grade, int Exam, Dictionary<string, decimal> StudentGPA)
        {
            this.SchoolYear = SchoolYear;
            this.Semester = Semester;
            this.Grade = Grade;
            this.Exam = Exam;
            this.StudentGPA = StudentGPA;
        }
        public int SchoolYear { get; private set; }
        public int Semester { get; private set; }
        public int Exam { get; private set; }
        public int Grade { get; private set; }
        public Dictionary<string, decimal> StudentGPA { get; private set; }
    }
    public class GPADistributionPart
    {
        public decimal GPACeiling { get; set; }
        public decimal GPAFloor { get; set; }
        public int Count { get; set; }
    }
    public class GradePeriodicalExamGPA
    {
        public static CourseGradeB.Tool.SubjectType StringToSubjectType(string type)
        {
            switch (type)
            {
                case "Honor":
                    return CourseGradeB.Tool.SubjectType.Honor;
                case "Regular":
                    return CourseGradeB.Tool.SubjectType.Regular;
                default:
                    return CourseGradeB.Tool.SubjectType.Regular;
            }
        }
        private GradePeriodicalExamGPA() { }
        private static Dictionary<string, GradePeriodicalExamGPARecord> dcpegpar = new Dictionary<string, GradePeriodicalExamGPARecord>();
        public static GradePeriodicalExamGPARecord GetGradePeriodicalExamGPARecord(int SchoolYear, int Semester, int Grade, int Exam)
        {
            string key = SchoolYear + "#" + Semester + "#" + Grade + "#" + Exam;

            if (dcpegpar.ContainsKey(key))
            {
                return dcpegpar[key];
            }
            DataTable dt = tool._Q.Select(@"select student.id,student.english_name,student.name,student.student_number,student.seat_no,class.class_name,teacher.teacher_name,class.grade_year,course.id as course_id,course.period as period,course.credit as credit,course.subject as subject,$ischool.subject.list.group as group,$ischool.subject.list.english_name as subject_english_name,$ischool.subject.list.type as type,sc_attend.ref_student_id as student_id,sce_take.ref_sc_attend_id as sc_attend_id,sce_take.ref_exam_id as exam_id,xpath_string(sce_take.extension,'//Score') as score 
from sc_attend
join sce_take on sce_take.ref_sc_attend_id=sc_attend.id
join course on course.id=sc_attend.ref_course_id
join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id
left join student on student.id=sc_attend.ref_student_id
left join class on student.ref_class_id=class.id
left join $ischool.subject.list on course.subject=$ischool.subject.list.name 
left join teacher on teacher.id=class.ref_teacher_id
where  class.grade_year =" + Grade + " and course.school_year=" + SchoolYear + " and course.semester=" + Semester + " and sce_take.ref_exam_id = " + Exam);
            Dictionary<string, List<CustomSCETakeRecord>> dscetr = new Dictionary<string, List<CustomSCETakeRecord>>();
            foreach (DataRow row in dt.Rows)
            {
                string id = "" + row["id"];
                if (!dscetr.ContainsKey(id))
                    dscetr.Add(id, new List<CustomSCETakeRecord>());
                dscetr[id].Add(new CustomSCETakeRecord()
                {
                    RefStudentID = id,
                    Name = "" + row["name"],
                    EnglishName = "" + row["english_name"],
                    StudentNumber = "" + row["student_number"],
                    SeatNo = "" + row["seat_no"],
                    ClassName = "" + row["class_name"],
                    TeacherName = "" + row["teacher_name"],
                    GradeYear = "" + row["grade_year"],
                    Subject = "" + row["subject"],
                    Score = decimal.Parse("" + row["score"]),
                    CourseId = "" + row["course_id"],
                    CoursePeriod = int.Parse("" + row["period"]),
                    CourseCredit = int.Parse("" + row["credit"]),
                    SubjectEnglishName = "" + row["subject_english_name"],
                    CourseGroup = "" + row["group"],
                    CourseType = JH.IBSH.Report.PeriodicalExam.GradePeriodicalExamGPA.StringToSubjectType("" + row["type"]),
                    ExamId = "" + row["exam_id"]
                });
            }
            Dictionary<string, decimal> StudentGPA = new Dictionary<string, decimal>();
            foreach (KeyValuePair<string, List<CustomSCETakeRecord>> row in dscetr)
            {

                int personalCreditCount = 0;
                decimal personalGPACount = 0;
                foreach (CustomSCETakeRecord item in row.Value)
                {
                    switch (item.GradeYear)
                    {
                        case "9":
                        case "10":
                        case "11":
                        case "12":
                            personalCreditCount += item.CourseCredit;
                            CourseGradeB.Tool.GPA _gpa =  CourseGradeB.Tool.GPA.Eval(item.Score);
                            if ( CourseGradeB.Tool.SubjectType.Honor == item.CourseType )
                                personalGPACount += item.CourseCredit * _gpa.Honors ;
                            else if (CourseGradeB.Tool.SubjectType.Regular == item.CourseType)
                                personalGPACount += item.CourseCredit * _gpa.Regular ;
                            break;
                    }
                }
                StudentGPA.Add(row.Key, personalCreditCount > 0 ? personalGPACount / personalCreditCount : 0);
            }
            dcpegpar[key] = new GradePeriodicalExamGPARecord(SchoolYear, Semester, Grade, Exam, StudentGPA);
            return dcpegpar[key];
        }
        public static List<GPADistributionPart> toGPADistribution(GradePeriodicalExamGPARecord gpegpar)
        {
            int fg = 5;//part
            decimal Max = gpegpar.StudentGPA.Values.Max(),
                Min = gpegpar.StudentGPA.Values.Min(),
                Spacing = (Max - Min) / fg;

            List<GPADistributionPart> gpadpl = new List<GPADistributionPart>();
            if (Spacing == 0)
            {
                gpadpl.Add(new GPADistributionPart()
                {
                    Count = gpegpar.StudentGPA.Count,
                    GPAFloor = Min,
                    GPACeiling = Max
                });
                return gpadpl;
            }
            for (int i = 0; i < fg; i++)
			{
			    gpadpl.Add(new GPADistributionPart(){
                    Count = 0,
                    GPAFloor = Min + i * Spacing ,
                    GPACeiling = Max - (fg-i-1) * Spacing 
                });
			}
            foreach (KeyValuePair<string, decimal> item in gpegpar.StudentGPA)
            {
                if (item.Value == Min)
                    gpadpl[0].Count++;
                else
                {
                    gpadpl[int.Parse(""+Math.Ceiling((item.Value - Min) / Spacing))-1].Count++;
                }
            }
            return gpadpl;
        }
    }
}
