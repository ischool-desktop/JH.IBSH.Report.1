using Aspose.Words;
using Aspose.Words.Tables;
using Campus.Report;
using CourseGradeB;
using FISCA.Data;
using FISCA.Permission;
using FISCA.Presentation.Controls;
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
    public partial class SemesterReportCard : BaseForm
    {
        private int _SchoolYear = 104;
        private int _Semester = 1;

        public SemesterReportCard()
        {
            InitializeComponent();

            for (int i = int.Parse(K12.Data.School.DefaultSchoolYear); i >= 104; i--)
            {
                cboSY.Items.Add("" + i);
                cboSY.SelectedIndex = 0;
            }

            cboSem.Items.Add("1");
            cboSem.Items.Add("2");
            if (K12.Data.School.DefaultSemester == "1")
                cboSem.SelectedIndex = 0;
            else
                cboSem.SelectedIndex = 1;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportConfiguration ReportConfigurationGrade1 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + 5);

            #region 第一次列印時複製前學期的樣板
            if (ReportConfigurationGrade1.Template == null)
            {
                int prevSy, prevSm;
                if (_Semester == 2)
                {
                    prevSy = _SchoolYear;
                    prevSm = 1;
                }
                else
                {
                    prevSy = _SchoolYear - 1;
                    prevSm = 2;
                }
                ReportConfiguration reportConfigurationGrade5 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + "S" + prevSm + ".G" + 5);
                if (reportConfigurationGrade5.Template != null)
                {
                    ReportConfigurationGrade1.Template = reportConfigurationGrade5.Template;
                }
            }
            #endregion

            Campus.Report.TemplateSettingForm TemplateForm;

            TemplateForm = new Campus.Report.TemplateSettingForm(ReportConfigurationGrade1.Template == null ? new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word) : ReportConfigurationGrade1.Template, new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "StudentReportCardTemplateForGrade56";

            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                foreach (var target in new int[] { 5, 6 })
                {
                    ReportConfiguration config = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + target);
                    config.Template = TemplateForm.Template;
                    config.Save();
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportConfiguration ReportConfigurationGrade2 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + 7);

            #region 第一次列印時複製前學期的樣板
            if (ReportConfigurationGrade2.Template == null)
            {
                int prevSy, prevSm;
                if (_Semester == 2)
                {
                    prevSy = _SchoolYear;
                    prevSm = 1;
                }
                else
                {
                    prevSy = _SchoolYear - 1;
                    prevSm = 2;
                }
                ReportConfiguration reportConfigurationGrade7 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + "S" + prevSm + ".G" + 7);
                if (reportConfigurationGrade7.Template != null)
                {
                    ReportConfigurationGrade2.Template = reportConfigurationGrade7.Template;
                }
            }
            #endregion

            Campus.Report.TemplateSettingForm TemplateForm;

            TemplateForm = new Campus.Report.TemplateSettingForm(ReportConfigurationGrade2.Template == null ? new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word) : ReportConfigurationGrade2.Template, new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "StudentReportCardTemplateForGrade7_12";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                foreach (var target in new int[] { 7, 8, 9, 10, 11, 12 })
                {
                    ReportConfiguration config = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + target);
                    config.Template = TemplateForm.Template;
                    config.Save();
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            #region 第一次列印時複製前學期的樣板
            int prevSy, prevSm;
            if (_Semester == 2)
            {
                prevSy = _SchoolYear;
                prevSm = 1;
            }
            else
            {
                prevSy = _SchoolYear - 1;
                prevSm = 2;
            }

            ReportConfiguration reportConfigurationGrade5 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + "S" + prevSm + ".G" + 5);
            foreach (var target in new int[] { 5, 6 })
            {
                ReportConfiguration reportConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + target);
                if (reportConfigurationGrade.Template == null)
                {
                    if (reportConfigurationGrade5.Template != null)
                    {
                        reportConfigurationGrade.Template = reportConfigurationGrade5.Template;
                        reportConfigurationGrade.Save();
                    }
                }
            }

            ReportConfiguration reportConfigurationGrade7 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + "S" + prevSm + ".G" + 7);
            foreach (var target in new int[] { 7, 8, 9, 10, 11, 12 })
            {
                ReportConfiguration reportConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + target);
                if (reportConfigurationGrade.Template == null)
                {
                    if (reportConfigurationGrade7.Template != null)
                    {
                        reportConfigurationGrade.Template = reportConfigurationGrade7.Template;
                        reportConfigurationGrade.Save();
                    }
                }
            }
            #endregion

            #region 列印成績單
            int _schoolYear = _SchoolYear;
            AccessHelper _A = new AccessHelper();
            QueryHelper _Q = new QueryHelper();
            List<string> _ids = new List<string>(K12.Presentation.NLDPanels.Student.SelectedSource);
            bool fieldMode = label1.Text != "";
            Dictionary<Document, string> documents = new Dictionary<Document, string>();

            BackgroundWorker bkw = new BackgroundWorker();
            bkw.RunWorkerCompleted += delegate
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("質性評量成績單產生完成。");
                List<string> files = new List<string>();
                foreach (var doc in documents.Keys)
                {

                    SaveFileDialog save = new SaveFileDialog();
                    save.Title = "另存新檔";
                    save.FileName = documents[doc];
                    save.Filter = "Word檔案 (*.docx)|*.docx|Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";

                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            doc.Save(save.FileName);
                            files.Add(save.FileName);
                        }
                        catch
                        {
                            MessageBox.Show("檔案儲存失敗");
                        }
                    }
                }
                foreach (var file in files)
                {
                    System.Diagnostics.Process.Start(file);
                }
            };

            bkw.DoWork += delegate
            {
                Dictionary<string, DataRow> dicStudentRow = new Dictionary<string, DataRow>();
                DataTable mergeDT = new DataTable();

                mergeDT.Columns.Add("系統編號");
                mergeDT.Columns.Add("學年");
                mergeDT.Columns.Add("學期");
                //每個學生增加一個ROW
                foreach (var id in _ids)
                {
                    var studentRow = mergeDT.Rows.Add();
                    studentRow[mergeDT.Columns.IndexOf("系統編號")] = id;
                    studentRow[mergeDT.Columns.IndexOf("學年")] = (_schoolYear + 1911) + "-" + (_schoolYear + 1912);
                    studentRow[mergeDT.Columns.IndexOf("學期")] = _Semester == 1 ? "1st" : "2nd";
                    dicStudentRow.Add(id, studentRow);
                }
                #region 取得學生基本資料(目前)
                {
                    mergeDT.Columns.Add("姓名");
                    mergeDT.Columns.Add("EnglishName");
                    mergeDT.Columns.Add("AKA");
                    mergeDT.Columns.Add("GivenName");
                    mergeDT.Columns.Add("MiddleName");
                    mergeDT.Columns.Add("FamilyName");
                    mergeDT.Columns.Add("學號");
                    mergeDT.Columns.Add("年級");
                    mergeDT.Columns.Add("班級");
                    mergeDT.Columns.Add("座號");
                    DataTable dt = _Q.Select(string.Format(@"
                        select student.id,
                                student.name,
                                student.english_name, 
                                sx.nick_name, 
                                sx.given_name, 
                                sx.middle_name, 
                                sx.family_name, 
                                student.student_number, 
                                class.class_name as classname,
                                class.grade_year,
                                student.seat_no
                        from student
                            left outer join class on student.ref_class_id = class.id
                            left outer join $jhcore_bilingual.studentrecordext as sx on sx.ref_student_id::int = student.id
                        where student.id in ({0});
                    ", string.Join(",", _ids)));
                    foreach (DataRow row in dt.Rows)
                    {
                        var studentRow = dicStudentRow["" + row["id"]];
                        studentRow[mergeDT.Columns.IndexOf("姓名")] = "" + row["name"];
                        studentRow[mergeDT.Columns.IndexOf("EnglishName")] = "" + row["english_name"];
                        studentRow[mergeDT.Columns.IndexOf("AKA")] = "" + row["nick_name"];
                        studentRow[mergeDT.Columns.IndexOf("GivenName")] = "" + row["given_name"];
                        studentRow[mergeDT.Columns.IndexOf("MiddleName")] = "" + row["middle_name"];
                        studentRow[mergeDT.Columns.IndexOf("FamilyName")] = "" + row["family_name"];
                        studentRow[mergeDT.Columns.IndexOf("學號")] = "" + row["student_number"];
                        studentRow[mergeDT.Columns.IndexOf("年級")] = "" + row["grade_year"];
                        studentRow[mergeDT.Columns.IndexOf("班級")] = "" + row["classname"];
                        studentRow[mergeDT.Columns.IndexOf("座號")] = "" + row["seat_no"];
                    }
                }
                #endregion

                {
                    #region 班導師資料
                    mergeDT.Columns.Add("Homeroom/teacherName");
                    mergeDT.Columns.Add("Homeroom/teacherEngName");
                    if ("" + _schoolYear == K12.Data.School.DefaultSchoolYear && "" + _Semester == K12.Data.School.DefaultSemester)
                    {
                        //當學期帶導師姓名
                        foreach (var stuRec in K12.Data.Student.SelectByIDs(_ids))
                        {
                            var studentRow = dicStudentRow[stuRec.ID];
                            if (stuRec.Class != null && stuRec.Class.Teacher != null)
                            {
                                studentRow[mergeDT.Columns.IndexOf("Homeroom/teacherName")] = "" + stuRec.Class.Teacher.Name;
                                studentRow[mergeDT.Columns.IndexOf("Homeroom/teacherEngName")] = "" + stuRec.Class.Teacher.Nickname;
                            }
                        }
                    }
                    #endregion
                    #region 取得學生學期歷程資料
                    if (!mergeDT.Columns.Contains("" + "SchoolDays"))
                        mergeDT.Columns.Add("" + "SchoolDays");//X學期上課天數
                    foreach (SemesterHistoryRecord r in K12.Data.SemesterHistory.SelectByStudentIDs(_ids))
                    {
                        foreach (SemesterHistoryItem item in r.SemesterHistoryItems)
                        {
                            if (item.SchoolYear == _schoolYear && item.Semester == _Semester)
                            {
                                var studentRow = dicStudentRow[item.RefStudentID];
                                studentRow[mergeDT.Columns.IndexOf("年級")] = "" + item.GradeYear;
                                studentRow[mergeDT.Columns.IndexOf("班級")] = item.ClassName;
                                studentRow[mergeDT.Columns.IndexOf("座號")] = "" + item.SeatNo;
                                studentRow[mergeDT.Columns.IndexOf("" + "SchoolDays")] = "" + item.SchoolDayCount;

                                if (item.Teacher.IndexOf(' ') > 0)
                                {
                                    studentRow[mergeDT.Columns.IndexOf("Homeroom/teacherName")] = item.Teacher.Substring(0, item.Teacher.IndexOf(' '));
                                    studentRow[mergeDT.Columns.IndexOf("Homeroom/teacherEngName")] = item.Teacher.Substring(item.Teacher.IndexOf(' ') + 1);
                                }
                                else
                                {
                                    studentRow[mergeDT.Columns.IndexOf("Homeroom/teacherName")] = item.Teacher;
                                    studentRow[mergeDT.Columns.IndexOf("Homeroom/teacherEngName")] = "";
                                }

                            }
                        }
                    }
                    #endregion
                    #region 取得缺課天數
                    {
                        if (!mergeDT.Columns.Contains("" + "AbsenceCount"))
                            mergeDT.Columns.Add("" + "AbsenceCount");//X學期缺曠天數
                        if (!mergeDT.Columns.Contains("" + "AttendanceCount"))
                            mergeDT.Columns.Add("" + "AttendanceCount");//X學期到校天數
                        //缺曠天數預設值給0
                        foreach (DataRow studentRow in dicStudentRow.Values)
                        {
                            studentRow[mergeDT.Columns.IndexOf("" + "AbsenceCount")] = "0";
                        }
                        /* 缺曠改抓學務系統
                        DataTable absence_dt = _Q.Select("select id from _udt_table where name='ischool.elementaryabsence'");
                        if (absence_dt.Rows.Count > 0)
                        {
                            string str = string.Format("select ref_student_id,personal_days + sick_days as absenceCount from $ischool.elementaryabsence where ref_student_id in ({0}) and school_year={1} and semester={2}", string.Join(",", _ids), _schoolYear, _Semester);
                            absence_dt = _Q.Select(str);

                            foreach (DataRow row in absence_dt.Rows)
                            {
                                var studentRow = dicStudentRow["" + row["ref_student_id"]];
                                string absenceCount = "" + row["absenceCount"];
                                studentRow[mergeDT.Columns.IndexOf("" + "AbsenceCount")] = absenceCount;
                                int sdc, absc;
                                if (int.TryParse("" + studentRow[mergeDT.Columns.IndexOf("" + "AbsenceCount")], out absc) &&
                                    int.TryParse("" + studentRow[mergeDT.Columns.IndexOf("" + "SchoolDays")], out sdc))
                                {
                                    studentRow[mergeDT.Columns.IndexOf("" + "AttendanceCount")] = "" + (sdc - absc);
                                }
                            }
                        }
                        */

                        #region 取得缺曠紀錄

                        #region 取得 Period List
                        Dictionary<string, string> dicPeriodType = new Dictionary<string, string>();

                        foreach (K12.Data.PeriodMappingInfo each in K12.Data.PeriodMapping.SelectAll())
                        {
                            if (!dicPeriodType.ContainsKey(each.Name)) //節次<-->類別
                                dicPeriodType.Add(each.Name, each.Type);
                        }
                        #endregion
                        foreach (var item in K12.BusinessLogic.AutoSummary.Select(
                            _ids
                            , new SchoolYearSemester[] { new SchoolYearSemester() { SchoolYear = _SchoolYear, Semester = _Semester } }
                            , SummaryType.Attendance
                            , true
                            )
                        )
                        {
                            //item.AbsenceCounts
                            int count = 0;
                            foreach (var absenceCount in item.AbsenceCounts)
                            {
                                if (absenceCount.PeriodType == "Study Hour")
                                {
                                    count += absenceCount.Count;
                                }
                                var studentRow = dicStudentRow["" + item.RefStudentID];
                                studentRow[mergeDT.Columns.IndexOf("" + "AbsenceCount")] = "" + (count / 7);
                            }

                        }

                        #endregion

                    }
                    #endregion
                    #region 取得修課及授課教師資訊
                    {
                        DataTable dt = _Q.Select(string.Format(@"
                        select sc_attend.id,
                            sc_attend.ref_student_id,
                            sc_attend.ref_course_id,
                            teacher.teacher_name,
                            teacher.nickname,
                            course.subject 
                        from sc_attend 
                            left join tc_instruct on tc_instruct.ref_course_id=sc_attend.ref_course_id
                            left join teacher on teacher.id=tc_instruct.ref_teacher_id
                            left join course on course.id=sc_attend.ref_course_id 
                        where ref_student_id in ({0}) and course.school_year={1} and course.semester={2} and tc_instruct.sequence = 1
                        order by sc_attend.ref_course_id
                    ", string.Join(",", _ids), _schoolYear, _Semester));
                        foreach (DataRow row in dt.Rows)
                        {

                            var studentRow = dicStudentRow["" + row["ref_student_id"]];

                            string subject_name = ("" + row["subject"]).Trim().Replace(' ', '_').Replace('"', '_');
                            //科目教師姓名欄位
                            string insertStrTeacherName = subject_name + "/teacherName";
                            if (!mergeDT.Columns.Contains(insertStrTeacherName))
                                mergeDT.Columns.Add(insertStrTeacherName);
                            string insertStrTeacherEngName = subject_name + "/teacherEngName";
                            if (!mergeDT.Columns.Contains(insertStrTeacherEngName))
                                mergeDT.Columns.Add(insertStrTeacherEngName);


                            studentRow[mergeDT.Columns.IndexOf(insertStrTeacherName)] = "" + row["teacher_name"];
                            studentRow[mergeDT.Columns.IndexOf(insertStrTeacherEngName)] = "" + row["nickname"];
                        }
                    }
                    #endregion
                    #region 取得定期評量成績
                    {
                        //                        string sql = @"select 
                        //                                        student.id,
                        //                                        student.english_name,
                        //                                        student.name,
                        //                                        student.student_number,
                        //                                        student.seat_no,
                        //                                        class.class_name,
                        //                                        teacher.teacher_name,
                        //                                        class.grade_year,
                        //                                        course.id as course_id,
                        //                                        course.period as period,
                        //                                        course.credit as credit,
                        //                                        course.subject as subject,
                        //                                        $ischool.subject.list.group as group,
                        //                                        $ischool.subject.list.type as type,
                        //                                        $ischool.subject.list.english_name as subject_english_name,
                        //                                        sc_attend.ref_student_id as student_id,
                        //                                        sce_take.ref_sc_attend_id as sc_attend_id,
                        //                                        sce_take.ref_exam_id as exam_id,
                        //                                        xpath_string(sce_take.extension,'//Score') as score 
                        //                                    from sc_attend
                        //                                        join sce_take on sce_take.ref_sc_attend_id=sc_attend.id
                        //                                        join course on course.id=sc_attend.ref_course_id
                        //                                        join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id
                        //                                        left join student on student.id=sc_attend.ref_student_id
                        //                                        left join class on student.ref_class_id=class.id
                        //                                        left join $ischool.subject.list on course.subject=$ischool.subject.list.name 
                        //                                        left join teacher on teacher.id=class.ref_teacher_id
                        //                                    where sc_attend.ref_student_id in ({0}) and course.school_year={1} and course.semester={2}";
                        //                        //sce_take.ref_exam_id 1 or 2
                        //                        DataTable dt = _Q.Select(string.Format(sql, string.Join(",", _ids), _schoolYear, _Semester));
                        //                        foreach (DataRow row in dt.Rows)
                        //                        {
                        //                            var studentRow = dicStudentRow["" + row["id"]];
                        //                            var subject = ("" + row["subject"]).Trim().Replace(' ', '_').Replace('"', '_');
                        //                            string insertStrExamScoreMark = (("" + row["exam_id"]) == "1" ? "MiddleTerm" : "FinalTrem") + "/" + subject + "/ExamScoreMark";
                        //                            if (!mergeDT.Columns.Contains(insertStrExamScoreMark))
                        //                                mergeDT.Columns.Add(insertStrExamScoreMark);
                        //                            decimal score = decimal.Parse("0" + row["score"]);
                        //                            studentRow[mergeDT.Columns.IndexOf(insertStrExamScoreMark)] = CourseGradeB.Tool.GPA.Eval(score).Letter;
                        //                        }
                    }
                    #endregion
                    #region 取得學期成績
                    {
                        //foreach (var semesterScoreRecord in K12.Data.SemesterScore.SelectBySchoolYearAndSemester(_ids, _schoolYear, _Semester))
                        //{
                        //    var studentRow = dicStudentRow[semesterScoreRecord.RefStudentID];
                        //    foreach (var subject in semesterScoreRecord.Subjects.Keys)
                        //    {
                        //        string insertStrSemsScore = "" + subject.Trim().Replace(' ', '_').Replace('"', '_') + "/SemsScore";
                        //        if (!mergeDT.Columns.Contains(insertStrSemsScore))
                        //            mergeDT.Columns.Add(insertStrSemsScore);
                        //        studentRow[mergeDT.Columns.IndexOf(insertStrSemsScore)] = "" + semesterScoreRecord.Subjects[subject].Score;
                        //    }
                        //    string insertStrSemsScoreAvgMark = "" + "/SemsScoreAvgMark";
                        //    if (!mergeDT.Columns.Contains(insertStrSemsScoreAvgMark))
                        //        mergeDT.Columns.Add(insertStrSemsScoreAvgMark);
                        //    if (semesterScoreRecord.AvgScore.HasValue)
                        //        studentRow[mergeDT.Columns.IndexOf(insertStrSemsScoreAvgMark)] = CourseGradeB.Tool.GPA.Eval(semesterScoreRecord.AvgScore.Value).Letter;
                        //}
                    }
                    #endregion
                    #region 取得Conduct資料
                    if (!mergeDT.Columns.Contains("MiddleTermComment"))
                        mergeDT.Columns.Add("MiddleTermComment");//Q1評語
                    if (!mergeDT.Columns.Contains("FinalTremComment"))
                        mergeDT.Columns.Add("FinalTremComment");//Q2評語
                    if (!mergeDT.Columns.Contains("Comment"))
                        mergeDT.Columns.Add("Comment");//S2評語
                    List<ConductRecord> records = _A.Select<ConductRecord>("ref_student_id in (" + string.Join(",", _ids) + ") and school_year=" + _schoolYear + " and semester=" + _Semester);
                    //and not term is null

                    //Dictionary<string, ConductObj> student_conduct = new Dictionary<string, ConductObj>();
                    foreach (ConductRecord record in records)
                    {
                        var studentRow = dicStudentRow["" + record.RefStudentId];


                        string subj = record.Subject;
                        if (string.IsNullOrWhiteSpace(subj))
                            subj = "Homeroom";

                        string term = record.Term;

                        //Comment
                        if (subj == "Homeroom")
                        {
                            if (term == "1")
                                studentRow[mergeDT.Columns.IndexOf("MiddleTermComment")] = record.Comment;
                            else if (term == "2")
                                studentRow[mergeDT.Columns.IndexOf("FinalTremComment")] = record.Comment;
                            else
                                studentRow[mergeDT.Columns.IndexOf("Comment")] = record.Comment;
                        }

                        var _xdoc = new XmlDocument();

                        if (!string.IsNullOrWhiteSpace(record.Conduct))
                            _xdoc.LoadXml(record.Conduct);

                        foreach (XmlElement elem in _xdoc.SelectNodes("//Conduct"))
                        {
                            string group = elem.GetAttribute("Group");

                            foreach (XmlElement item in elem.SelectNodes("Item"))
                            {
                                string title = item.GetAttribute("Title");
                                string grade = item.GetAttribute("Grade");

                                string insertStrC = "???????";
                                if (term == "1")
                                {
                                    insertStrC = "MiddleTerm/" + subj.Trim().Replace(' ', '_').Replace('"', '_') + "/" + group.Trim().Replace(' ', '_').Replace('"', '_') + "/" + title.Trim().Replace(' ', '_').Replace('"', '_');
                                }
                                else if (term == "2")
                                {
                                    insertStrC = "FinalTrem/" + subj.Trim().Replace(' ', '_').Replace('"', '_') + "/" + group.Trim().Replace(' ', '_').Replace('"', '_') + "/" + title.Trim().Replace(' ', '_').Replace('"', '_');
                                }
                                else
                                {
                                    insertStrC = "" + subj.Trim().Replace(' ', '_').Replace('"', '_') + "/" + group.Trim().Replace(' ', '_').Replace('"', '_') + "/" + title.Trim().Replace(' ', '_').Replace('"', '_');
                                }
                                if (!mergeDT.Columns.Contains(insertStrC))
                                    mergeDT.Columns.Add(insertStrC);

                                studentRow[mergeDT.Columns.IndexOf(insertStrC)] = grade;
                            }
                        }
                    }
                    #endregion
                }

                if (fieldMode)
                {
                    Document doc = new Document();
                    DocumentBuilder bu = new DocumentBuilder(doc);
                    bu.MoveToDocumentStart();
                    bu.CellFormat.Borders.LineStyle = LineStyle.Single;
                    bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    Table table = bu.StartTable();
                    foreach (DataColumn col in mergeDT.Columns)
                    {

                        bu.InsertCell();
                        bu.CellFormat.Width = 15;
                        bu.InsertField("MERGEFIELD " + col.Caption + @" \* MERGEFORMAT", "«.»");
                        bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                        bu.InsertCell();
                        bu.CellFormat.Width = 125;
                        bu.Write(col.Caption);
                        bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                        bu.EndRow();
                    }
                    table.AllowAutoFit = false;
                    bu.EndTable();
                    documents.Add(doc, "質性評量成績單合併欄位表.doc");
                }
                else
                {
                    Dictionary<string, DataTable> gradeDT = new Dictionary<string, DataTable>();
                    foreach (DataRow row in mergeDT.Rows)
                    {
                        var grade = "" + row["年級"];
                        if (!gradeDT.ContainsKey(grade))
                        {
                            var gDT = mergeDT.Clone();
                            gradeDT.Add(grade, gDT);
                        }
                        gradeDT[grade].ImportRow(row);
                    }

                    foreach (var grade in gradeDT.Keys)
                    {
                        ReportConfiguration reportConfiguration = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + "S" + _Semester + ".G" + grade);
                        if (reportConfiguration.Template != null)
                        {
                            var doc = new Document(new MemoryStream(reportConfiguration.Template.ToBinary()));
                            //合併，儲存
                            doc.MailMerge.Execute(gradeDT[grade]);
                            doc.MailMerge.DeleteFields();

                            documents.Add(doc, "質性評量成績單(" + grade + "年級) .doc");
                        }
                    }
                }
            };
            bkw.RunWorkerAsync();
            FISCA.Presentation.MotherForm.SetStatusBarMessage("質性評量成績單產生中...");
            this.Close();

            #endregion
        }

        private void label1_DoubleClick(object sender, EventArgs e)
        {
            label1.Text = (label1.Text == "" ? "產出合併欄位" : "");
        }

        private void SetSemester(object sender, EventArgs e)
        {
            if (cboSem.Text != "" && cboSY.Text != "")
            {
                _SchoolYear = int.Parse(cboSY.Text);
                _Semester = int.Parse(cboSem.Text);
            }
        }
    }
}
