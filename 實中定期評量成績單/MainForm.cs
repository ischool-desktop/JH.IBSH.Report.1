using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Aspose.Words;
using Aspose.Words.Drawing;
using Campus.Report;
using JHSchool.Data;
using K12.Data;
//using Campus.ePaper;

namespace JH.IBSH.Report.PeriodicalExam
{
    public partial class MainForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bgw = new BackgroundWorker();

        private const string config3_6 = "JH.IBSH.Report.PeriodicalExam.Config.3_6";
        private const string config7_8 = "JH.IBSH.Report.PeriodicalExam.Config.7_8";
        private const string config9_12 = "JH.IBSH.Report.PeriodicalExam.Config.9_12";

        public static ReportConfiguration ReportConfiguration3_6 = new Campus.Report.ReportConfiguration(config3_6);
        public static ReportConfiguration ReportConfiguration7_8 = new Campus.Report.ReportConfiguration(config7_8);
        public static ReportConfiguration ReportConfiguration9_12 = new Campus.Report.ReportConfiguration(config9_12);
        public string current = "評量通知單";
        class filter {
            public SchoolYearSemester sys;
            public int exam;
            public string gradeSection;
            public string examText;
        }
        public MainForm()
        {
            InitializeComponent();
            #region 設定comboBox選單
            foreach (int item in Enumerable.Range(int.Parse(School.DefaultSchoolYear)-11, 13) )
            {
                comboBoxEx2.Items.Add(item);
            }
            foreach (string item in new string[] { "1", "2" })
            {
                comboBoxEx3.Items.Add(item);
            }
            foreach (string item in new string[] { "Midterm", "Final" })
            {
                comboBoxEx4.Items.Add(item);
            }
            foreach (string item in new string[] { "3~6", "7~8","9~12" })
            {
                comboBoxEx1.Items.Add(item);
            }
            #endregion
            comboBoxEx2.Text = School.DefaultSchoolYear;
            comboBoxEx3.Text = School.DefaultSemester;
            comboBoxEx4.SelectedIndex = 0;
            comboBoxEx1.SelectedIndex = 0;
            _bgw.DoWork += new DoWorkEventHandler(_bgw_DoWork);
            _bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgw_RunWorkerCompleted);
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Config().ShowDialog();
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生");
                return;
            }
            btnPrint.Enabled = false;
            _bgw.RunWorkerAsync(new filter
                {
                    sys = new SchoolYearSemester(int.Parse(comboBoxEx2.Text), int.Parse(comboBoxEx3.Text)),
                    exam = comboBoxEx4.Text=="Midterm"?1:2,
                    examText = comboBoxEx4.Text ,
                    gradeSection = comboBoxEx1.Text
                }
            );
        }
        void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            filter f = (filter)e.Argument;
            Document document = new Document();
            Document template3_6 = (ReportConfiguration3_6.Template != null) //3-6單頁範本
                 ? ReportConfiguration3_6.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word).ToDocument();
            Document template7_8 = (ReportConfiguration7_8.Template != null) //7-8單頁範本
                 ? ReportConfiguration7_8.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word).ToDocument();
            Document template9_12 = (ReportConfiguration9_12.Template != null) //9-12單頁範本
                 ? ReportConfiguration9_12.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources._9樣版, Campus.Report.TemplateType.Word).ToDocument();
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                return;
            List<string> sids = K12.Presentation.NLDPanels.Student.SelectedSource;

            int SchoolYear = f.sys.SchoolYear;
            int Semester = f.sys.Semester;
            int Exam = f.exam;
            string ExamText = f.examText;
            string gradeSection = f.gradeSection;
            string sql = @"select student.id,student.english_name,student.name,student.student_number,student.seat_no,class.class_name,teacher.teacher_name,class.grade_year,course.id as course_id,course.period as period,course.credit as credit,course.subject as subject,$ischool.subject.list.group as group,$ischool.subject.list.type as type,$ischool.subject.list.english_name as subject_english_name,sc_attend.ref_student_id as student_id,sce_take.ref_sc_attend_id as sc_attend_id,sce_take.ref_exam_id as exam_id,xpath_string(sce_take.extension,'//Score') as score 
from sc_attend
join sce_take on sce_take.ref_sc_attend_id=sc_attend.id
join course on course.id=sc_attend.ref_course_id
join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id
left join student on student.id=sc_attend.ref_student_id
left join class on student.ref_class_id=class.id
left join $ischool.subject.list on course.subject=$ischool.subject.list.name 
left join teacher on teacher.id=class.ref_teacher_id
where sc_attend.ref_student_id in (" +string.Join("," ,sids) + ") and course.school_year=" + SchoolYear + " and course.semester=" + Semester + " and sce_take.ref_exam_id = " + Exam + "";
            DataTable dt = tool._Q.Select(sql);
            //return;
            Dictionary<string, List<CustomSCETakeRecord>> dscetr = new Dictionary<string, List<CustomSCETakeRecord>>();
            foreach (DataRow row in dt.Rows)
            {
                string id = ""+row["id"] ;
                if (!dscetr.ContainsKey(id))
                    dscetr.Add(id, new List<CustomSCETakeRecord>());
                decimal tmp_score;
                int tmp_period, tmp_credit;
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
                    Score = decimal.TryParse("" + row["score"],out tmp_score)?tmp_score:tmp_score,
                    CourseId = "" + row["course_id"],
                    CoursePeriod = int.TryParse("" + row["period"],out tmp_period)?tmp_period:tmp_period,
                    CourseCredit = int.TryParse("" + row["credit"],out tmp_credit)?tmp_credit:tmp_credit,
                    SubjectEnglishName = "" + row["subject_english_name"],
                    CourseGroup = "" + row["group"],
                    CourseType = JH.IBSH.Report.PeriodicalExam.GradePeriodicalExamGPA.StringToSubjectType("" + row["type"]),
                    ExamId = "" + row["exam_id"]
                });
            }

            Dictionary<string, object> mailmerge = new Dictionary<string, object>();
            List<StudentRecord> srl = K12.Data.Student.SelectByIDs(sids);
            srl.Sort(delegate(StudentRecord a, StudentRecord b) {
                StudentRecord aStudent = a;
                ClassRecord aClass = a.Class;
                StudentRecord bStudent = b;
                ClassRecord bClass = b.Class;

                string aa = aClass == null ? (string.Empty).PadLeft(10, '0') : (aClass.Name).PadLeft(10, '0');
                aa += aStudent == null ? (string.Empty).PadLeft(3, '0') : (aStudent.SeatNo + "").PadLeft(3, '0');
                aa += aStudent == null ? (string.Empty).PadLeft(10, '0') : (aStudent.StudentNumber).PadLeft(10, '0');

                string bb = bClass == null ? (string.Empty).PadLeft(10, '0') : (bClass.Name).PadLeft(10, '0');
                bb += bStudent == null ? (string.Empty).PadLeft(3, '0') : (bStudent.SeatNo + "").PadLeft(3, '0');
                bb += bStudent == null ? (string.Empty).PadLeft(10, '0') : (bStudent.StudentNumber).PadLeft(10, '0');

                return aa.CompareTo(bb);
            });
            foreach (StudentRecord sr in srl)
            {
                mailmerge.Clear();

                mailmerge.Add("學年", SchoolYear);
                mailmerge.Add("學期", Semester);
                mailmerge.Add("學段", ExamText);
                mailmerge.Add("班級", sr.Class != null ? sr.Class.Name : "");
                mailmerge.Add("級別", "");
                if (sr.Class != null)
                    mailmerge["級別"] = sr.Class.GradeYear;
                mailmerge.Add("座號", sr.SeatNo);
                mailmerge.Add("姓名", sr.Name);
                mailmerge.Add("英文名", sr.EnglishName);

                mailmerge.Add("學校名稱", School.ChineseName);
                mailmerge.Add("學校英文名稱", School.EnglishName);

                Document each;
                each = (Document)template3_6.Clone(true);
                int subjecti = 1;
                int parti = 1;
                if (dscetr.ContainsKey(sr.ID) && sr.Class != null && sr.Class.GradeYear.HasValue)
                {
                    #region 學生成績
                    switch (gradeSection)
                    {
                        case "3~6":
                            each = (Document)template3_6.Clone(true);
                            foreach (CustomSCETakeRecord item in dscetr[sr.ID])
                            {
                                mailmerge.Add(string.Format("科目{0}", subjecti), item.Subject + " " + item.SubjectEnglishName);
                                mailmerge.Add(string.Format("成績{0}", subjecti), CourseGradeB.Tool.GPA.Eval(item.Score).Letter);
                                subjecti++;
                            }
                            break;
                        case "7~8":
                            each = (Document)template7_8.Clone(true);
                            foreach (CustomSCETakeRecord item in dscetr[sr.ID])
                            {
                                mailmerge.Add(string.Format("科目{0}", subjecti), item.Subject + " " + item.SubjectEnglishName);
                                mailmerge.Add(string.Format("成績{0}", subjecti), CourseGradeB.Tool.GPA.Eval(item.Score).Letter);
                                subjecti++;
                            }
                            break;
                        case "9~12":
                            each = (Document)template9_12.Clone(true);
                            int personalCreditCount = 0;
                            decimal personalGPACount = 0, personalAverageCount = 0;
                            foreach (CustomSCETakeRecord item in dscetr[sr.ID])
                            {
                                mailmerge.Add(string.Format("科目{0}", subjecti), item.Subject + " " + item.SubjectEnglishName);
                                mailmerge.Add(string.Format("成績{0}", subjecti), item.Score);
                                personalCreditCount += item.CourseCredit;
                                CourseGradeB.Tool.GPA _gpa = CourseGradeB.Tool.GPA.Eval(item.Score);
                                if (CourseGradeB.Tool.SubjectType.Honor == item.CourseType)
                                    personalGPACount += item.CourseCredit * _gpa.Honors;
                                else if (CourseGradeB.Tool.SubjectType.Regular == item.CourseType)
                                    personalGPACount += item.CourseCredit * _gpa.Regular;

                                personalAverageCount += item.CourseCredit * item.Score;
                                subjecti++;
                            }
                            if (personalCreditCount > 0)
                            {
                                mailmerge.Add("科目平均", decimal.Round(personalAverageCount / personalCreditCount, 2, MidpointRounding.AwayFromZero));
                                mailmerge.Add("GPA", decimal.Round(personalGPACount / personalCreditCount, 2, MidpointRounding.AwayFromZero));
                            }
                            if (sr.Class.GradeYear.HasValue)
                            {
                                GradePeriodicalExamGPARecord gpegpar = GradePeriodicalExamGPA.GetGradePeriodicalExamGPARecord(SchoolYear, Semester, sr.Class.GradeYear.Value, Exam);
                                List<GPADistributionPart> gpad = GradePeriodicalExamGPA.toGPADistribution(gpegpar);
                                gpad.Reverse();
                                foreach (GPADistributionPart gpadp in gpad)
                                {
                                    mailmerge.Add(string.Format("GPA分段{0}", parti), decimal.Round(gpadp.GPACeiling, 2, MidpointRounding.AwayFromZero) + "~" + decimal.Round(gpadp.GPAFloor, 2, MidpointRounding.AwayFromZero));
                                    mailmerge.Add(string.Format("GPA計數{0}", parti), gpadp.Count);
                                    parti++;
                                }
                            }
                            break;
                    }
                    #endregion
                }
                for (; parti <= 5; parti++)
                {
                    mailmerge.Add(string.Format("GPA分段{0}", parti), "");
                    mailmerge.Add(string.Format("GPA計數{0}", parti), "");
                }
                mailmerge.Add(string.Format("科目{0}", subjecti), "(以下空白)");
                mailmerge.Add(string.Format("成績{0}", subjecti), "");
                for (subjecti++; subjecti <= 20; subjecti++)
                {
                    mailmerge.Add(string.Format("科目{0}", subjecti), "");
                    mailmerge.Add(string.Format("成績{0}", subjecti), "");
                }
                each.MailMerge.Execute(mailmerge.Keys.ToArray(), mailmerge.Values.ToArray());

                document.Sections.Add(document.ImportNode(each.FirstSection, true));
            }
            
            document.Sections.RemoveAt(0);
            e.Result = document;
        }
        
        void _bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document inResult = (Document)e.Result;
            btnPrint.Enabled = true;
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = current;

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    inResult.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(SaveFileDialog1.FileName + ",列印完成!!");
                    //Update_ePaper ue = new Update_ePaper(new List<Document> { inResult }, current, PrefixStudent.學號);
                    //ue.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                    return;
                }
            }
            catch (Exception exp)
            {
                string msg = "檔案儲存錯誤,請檢查檔案是否開啟中!!";
                FISCA.Presentation.Controls.MsgBox.Show(msg + "\n" + exp.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(msg + "\n" + exp.Message);
            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
    
}
