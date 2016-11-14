using Aspose.Words;
using Aspose.Words.Tables;
using CourseGradeB;
using FISCA.Data;
using FISCA.Permission;
using FISCA.UDT;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;


namespace ibshGradeYearReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];

            //權限設定
            Catalog permission2 = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission2.Add(new RibbonFeature("{A5A44C03-6835-42C4-BAE6-6C932AC3B4BB}", "ConductGradeReport(for Grade 5-12)"));

            var btnReport2 = item1["報表"]["成績相關報表"]["ConductGradeReport(for Gr.5-12;2015 年以後適用)"];
            btnReport2.Image = Properties.Resources.report_3_xxl;
            btnReport2.Enable = false;
            btnReport2.Click += delegate
            {
                new SemesterReportCard().ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0
                    && FISCA.Permission.UserAcl.Current["{A5A44C03-6835-42C4-BAE6-6C932AC3B4BB}"].Executable
                    )
                {
                    btnReport2.Enable = true;
                }
                else
                {
                    btnReport2.Enable = false;
                }
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature("{03880A25-AE37-468E-BEC9-C0BC51C3D269}", "質性評量成績單(for Grade 1-4)"));

            var btnReport = item1["報表"]["成績相關報表"]["質性評量成績單(for Gr.1-4;2015 年以後適用)"];
            btnReport.Image = Properties.Resources.report_3_xxl;
            btnReport.Enable = false;
            btnReport.Click += delegate
            {
                new GradeYearReportCard().ShowDialog();
            };
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0
                    && FISCA.Permission.UserAcl.Current["{03880A25-AE37-468E-BEC9-C0BC51C3D269}"].Executable
                    )
                {
                    btnReport.Enable = true;
                }
                else
                {
                    btnReport.Enable = false;
                }
            };
        }

    }





    public class StudentTeacherObj
    {
        public string studentId, courseId, TeacherName, TeacherNickName, SubjectName;
        public int Sequence;

        public StudentTeacherObj(DataRow row)
        {
            studentId = row["ref_student_id"] + "";
            courseId = row["ref_course_id"] + "";
            TeacherName = row["teacher_name"] + "";
            TeacherNickName = "" + row["nickname"];
            SubjectName = row["subject"] + "";

            int i = 0;
            Sequence = int.TryParse(row["sequence"] + "", out i) ? i : 4;
        }
    }

    public class ConductObj
    {
        public static XmlDocument _xdoc;
        public Dictionary<string, string> term1 = new Dictionary<string, string>();
        public Dictionary<string, string> term2 = new Dictionary<string, string>();
        public string Comment1, Comment2;
        public string StudentID;
        public StudentRecord Student;
        public ClassRecord Class;
        public string PersonalDays, SickDays, SchoolDays;
        public List<string> SubjectList;

        public Dictionary<string, Dictionary<string, List<string>>> Template = new Dictionary<string, Dictionary<string, List<string>>>();

        public ConductObj(ConductRecord record)
        {
            StudentID = record.RefStudentId + "";

            Student = K12.Data.Student.SelectByID(StudentID);
            Class = Student.Class;

            if (Student == null)
                Student = new StudentRecord();

            if (Class == null)
                Class = new ClassRecord();

            SubjectList = new List<string>();
        }

        public void LoadRecord(ConductRecord record)
        {
            string subj = record.Subject;
            if (string.IsNullOrWhiteSpace(subj))
                subj = "Homeroom";

            string term = record.Term;

            //Comment
            if (subj == "Homeroom")
            {
                if (term == "1")
                    Comment1 = record.Comment;

                if (term == "2")
                    Comment2 = record.Comment;
            }

            if (!SubjectList.Contains(subj))
                SubjectList.Add(subj);

            //XML
            if (_xdoc == null)
                _xdoc = new XmlDocument();

            _xdoc.RemoveAll();
            if (!string.IsNullOrWhiteSpace(record.Conduct))
                _xdoc.LoadXml(record.Conduct);

            foreach (XmlElement elem in _xdoc.SelectNodes("//Conduct"))
            {
                string group = elem.GetAttribute("Group");

                foreach (XmlElement item in elem.SelectNodes("Item"))
                {
                    string title = item.GetAttribute("Title");
                    string grade = item.GetAttribute("Grade");

                    if (term == "1")
                    {
                        if (!term1.ContainsKey(subj + "_" + group + "_" + title))
                            term1.Add(subj + "_" + group + "_" + title, grade);
                    }

                    if (term == "2")
                    {
                        if (!term2.ContainsKey(subj + "_" + group + "_" + title))
                            term2.Add(subj + "_" + group + "_" + title, grade);
                    }

                    if (!Template.ContainsKey(subj))
                        Template.Add(subj, new Dictionary<string, List<string>>());
                    if (!Template[subj].ContainsKey(group))
                        Template[subj].Add(group, new List<string>());
                    if (!Template[subj][group].Contains(title))
                        Template[subj][group].Add(title);
                }
            }
        }

        public int GetTotalAbsence()
        {
            int p = 0;
            int s = 0;
            int.TryParse(PersonalDays, out p);
            int.TryParse(SickDays, out s);

            return p + s;
        }

        public int? GetAttendanceCount()
        {
            int p = 0;
            int s = 0;
            int all = 0;
            int.TryParse(PersonalDays, out p);
            int.TryParse(SickDays, out s);
            if (int.TryParse(SchoolDays, out all))
            {
                return all - p - s;
            }
            else
                return null;
        }
    }
    class CustomSCETakeRecord
    {
        public string RefStudentID;
        public string Name;
        public string EnglishName;
        public string StudentNumber;
        public string SeatNo;
        public string ClassName;
        public string TeacherName;
        public string GradeYear;
        public string Subject;
        public decimal Score;
        public string CourseId;
        public int CoursePeriod;
        public int CourseCredit;
        public string SubjectEnglishName;
        public string CourseGroup;
        public CourseGradeB.Tool.SubjectType CourseType;
        public string ExamId;
    }
}
