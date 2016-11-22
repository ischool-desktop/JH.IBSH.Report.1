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
    public partial class GradeYearReportCard : BaseForm
    {
        private int _SchoolYear = 104;

        public GradeYearReportCard()
        {
            InitializeComponent();
            for (int i = int.Parse(K12.Data.School.DefaultSchoolYear); i >= 104; i--)
            {
                cboSY.Items.Add("" + i);
                cboSY.SelectedIndex = 0;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportConfiguration ReportConfigurationGrade1 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + ".G" + 1);


            #region 第一次列印時複製前學期的樣板
            if (ReportConfigurationGrade1.Template == null)
            {
                int prevSy = _SchoolYear - 1;

                ReportConfiguration prevConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + ".G" + 1);
                if (prevConfigurationGrade.Template != null)
                {
                    ReportConfigurationGrade1.Template = prevConfigurationGrade.Template;
                }
            }
            #endregion

            Campus.Report.TemplateSettingForm TemplateForm;

            TemplateForm = new Campus.Report.TemplateSettingForm(ReportConfigurationGrade1.Template == null ? new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word) : ReportConfigurationGrade1.Template, new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "StudentReportCardTemplateForGrade1";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                ReportConfigurationGrade1.Template = TemplateForm.Template;
                ReportConfigurationGrade1.Save();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportConfiguration ReportConfigurationGrade2 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + ".G" + 2);

            #region 第一次列印時複製前學期的樣板
            if (ReportConfigurationGrade2.Template == null)
            {
                int prevSy = _SchoolYear - 1;

                ReportConfiguration prevConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + ".G" + 2);
                if (prevConfigurationGrade.Template != null)
                {
                    ReportConfigurationGrade2.Template = prevConfigurationGrade.Template;
                }
            }
            #endregion

            Campus.Report.TemplateSettingForm TemplateForm;

            TemplateForm = new Campus.Report.TemplateSettingForm(ReportConfigurationGrade2.Template == null ? new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word) : ReportConfigurationGrade2.Template, new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "StudentReportCardTemplateForGrade2";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                ReportConfigurationGrade2.Template = TemplateForm.Template;
                ReportConfigurationGrade2.Save();
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportConfiguration ReportConfigurationGrade3 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + ".G" + 3);

            #region 第一次列印時複製前學期的樣板
            if (ReportConfigurationGrade3.Template == null)
            {
                int prevSy = _SchoolYear - 1;

                ReportConfiguration prevConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + ".G" + 3);
                if (prevConfigurationGrade.Template != null)
                {
                    ReportConfigurationGrade3.Template = prevConfigurationGrade.Template;
                }
            }
            #endregion

            Campus.Report.TemplateSettingForm TemplateForm;

            TemplateForm = new Campus.Report.TemplateSettingForm(ReportConfigurationGrade3.Template == null ? new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word) : ReportConfigurationGrade3.Template, new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "StudentReportCardTemplateForGrade3";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                ReportConfigurationGrade3.Template = TemplateForm.Template;
                ReportConfigurationGrade3.Save();
            }
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportConfiguration ReportConfigurationGrade4 = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + ".G" + 4);

            #region 第一次列印時複製前學期的樣板
            if (ReportConfigurationGrade4.Template == null)
            {
                int prevSy = _SchoolYear - 1;

                ReportConfiguration prevConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + ".G" + 4);
                if (prevConfigurationGrade.Template != null)
                {
                    ReportConfigurationGrade4.Template = prevConfigurationGrade.Template;
                }
            }
            #endregion

            Campus.Report.TemplateSettingForm TemplateForm;

            TemplateForm = new Campus.Report.TemplateSettingForm(ReportConfigurationGrade4.Template == null ? new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word) : ReportConfigurationGrade4.Template, new ReportTemplate(Properties.Resources.Doc1, TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "StudentReportCardTemplateForGrade4";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                ReportConfigurationGrade4.Template = TemplateForm.Template;
                ReportConfigurationGrade4.Save();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {

            #region 第一次列印時複製前學期的樣板
            for (int i = 1; i <= 4; i++)
            {
                ReportConfiguration reportConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + ".G" + i);

                if (reportConfigurationGrade.Template == null)
                {
                    int prevSy = _SchoolYear - 1;

                    ReportConfiguration prevConfigurationGrade = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + prevSy + ".G" + i);
                    if (prevConfigurationGrade.Template != null)
                    {
                        reportConfigurationGrade.Template = prevConfigurationGrade.Template;
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
                //每個學生增加一個ROW
                foreach (var id in _ids)
                {
                    var studentRow = mergeDT.Rows.Add();
                    studentRow[mergeDT.Columns.IndexOf("系統編號")] = id;
                    studentRow[mergeDT.Columns.IndexOf("學年")] = (_schoolYear + 1911) + "-" + (_schoolYear + 1912);
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

                foreach (var semester in new int[] { 1, 2 })
                {
                    #region 班導師資料
                    if (!mergeDT.Columns.Contains("Homeroom/teacherName"))
                        mergeDT.Columns.Add("Homeroom/teacherName");
                    if (!mergeDT.Columns.Contains("Homeroom/teacherEngName"))
                        mergeDT.Columns.Add("Homeroom/teacherEngName");
                    if ("" + _schoolYear == K12.Data.School.DefaultSchoolYear && "" + semester == K12.Data.School.DefaultSemester)
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
                    if (!mergeDT.Columns.Contains("S" + semester + "SchoolDays"))
                        mergeDT.Columns.Add("S" + semester + "SchoolDays");//X學期上課天數
                    foreach (SemesterHistoryRecord r in K12.Data.SemesterHistory.SelectByStudentIDs(_ids))
                    {
                        foreach (SemesterHistoryItem item in r.SemesterHistoryItems)
                        {
                            if (item.SchoolYear == _schoolYear && item.Semester == semester)
                            {
                                var studentRow = dicStudentRow[item.RefStudentID];
                                studentRow[mergeDT.Columns.IndexOf("年級")] = "" + item.GradeYear;
                                studentRow[mergeDT.Columns.IndexOf("班級")] = item.ClassName;
                                studentRow[mergeDT.Columns.IndexOf("座號")] = "" + item.SeatNo;
                                studentRow[mergeDT.Columns.IndexOf("S" + semester + "SchoolDays")] = "" + item.SchoolDayCount;

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
                        if (!mergeDT.Columns.Contains("S" + semester + "AbsenceCount"))
                            mergeDT.Columns.Add("S" + semester + "AbsenceCount");//X學期缺曠天數
                        if (!mergeDT.Columns.Contains("S" + semester + "AttendanceCount"))
                            mergeDT.Columns.Add("S" + semester + "AttendanceCount");//X學期到校天數

                        //缺曠天數預設值給0
                        foreach (DataRow studentRow in dicStudentRow.Values)
                        {
                            studentRow[mergeDT.Columns.IndexOf("S" + semester + "AbsenceCount")] = "0";
                        }

                        DataTable absence_dt = _Q.Select("select id from _udt_table where name='ischool.elementaryabsence'");
                        if (absence_dt.Rows.Count > 0)
                        {
                            string str = string.Format("select ref_student_id,personal_days + sick_days as absenceCount from $ischool.elementaryabsence where ref_student_id in ({0}) and school_year={1} and semester={2}", string.Join(",", _ids), _schoolYear, semester);
                            absence_dt = _Q.Select(str);

                            foreach (DataRow row in absence_dt.Rows)
                            {
                                var studentRow = dicStudentRow["" + row["ref_student_id"]];
                                string absenceCount = "" + row["absenceCount"];
                                studentRow[mergeDT.Columns.IndexOf("S" + semester + "AbsenceCount")] = absenceCount;
                                int sdc, absc;
                                if (int.TryParse("" + studentRow[mergeDT.Columns.IndexOf("S" + semester + "AbsenceCount")], out absc) &&
                                    int.TryParse("" + studentRow[mergeDT.Columns.IndexOf("S" + semester + "SchoolDays")], out sdc))
                                {
                                    studentRow[mergeDT.Columns.IndexOf("S" + semester + "AttendanceCount")] = "" + (sdc - absc);
                                }
                            }
                        }
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
                    ", string.Join(",", _ids), _schoolYear, semester));
                        foreach (DataRow row in dt.Rows)
                        {

                            var studentRow = dicStudentRow["" + row["ref_student_id"]];

                            string subject_name = ("" + row["subject"]).FixFieldName();
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
                        string sql = @"select 
                                        student.id,
                                        student.english_name,
                                        student.name,
                                        student.student_number,
                                        student.seat_no,
                                        class.class_name,
                                        teacher.teacher_name,
                                        class.grade_year,
                                        course.id as course_id,
                                        course.period as period,
                                        course.credit as credit,
                                        course.subject as subject,
                                        $ischool.subject.list.group as group,
                                        $ischool.subject.list.type as type,
                                        $ischool.subject.list.english_name as subject_english_name,
                                        sc_attend.ref_student_id as student_id,
                                        sce_take.ref_sc_attend_id as sc_attend_id,
                                        sce_take.ref_exam_id as exam_id,
                                        xpath_string(sce_take.extension,'//Score') as score 
                                    from sc_attend
                                        join sce_take on sce_take.ref_sc_attend_id=sc_attend.id
                                        join course on course.id=sc_attend.ref_course_id
                                        join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id
                                        left join student on student.id=sc_attend.ref_student_id
                                        left join class on student.ref_class_id=class.id
                                        left join $ischool.subject.list on course.subject=$ischool.subject.list.name 
                                        left join teacher on teacher.id=class.ref_teacher_id
                                    where sc_attend.ref_student_id in ({0}) and course.school_year={1} and course.semester={2}";
                        //sce_take.ref_exam_id 1 or 2
                        DataTable dt = _Q.Select(string.Format(sql, string.Join(",", _ids), _schoolYear, semester));
                        foreach (DataRow row in dt.Rows)
                        {
                            var studentRow = dicStudentRow["" + row["id"]];
                            var subject = ("" + row["subject"]).FixFieldName();
                            string insertStrExamScoreMark = "Q" + ((semester - 1) * 2 + int.Parse("0" + row["exam_id"])) + "/" + subject + "/ExamScoreMark";
                            if (!mergeDT.Columns.Contains(insertStrExamScoreMark))
                                mergeDT.Columns.Add(insertStrExamScoreMark);
                            decimal score = decimal.Parse("0" + row["score"]);
                            studentRow[mergeDT.Columns.IndexOf(insertStrExamScoreMark)] = CourseGradeB.Tool.GPA.Eval(score).Letter;
                        }
                    }
                    #endregion
                    #region 取得學期成績
                    {
                        foreach (var semesterScoreRecord in K12.Data.SemesterScore.SelectBySchoolYearAndSemester(_ids, _schoolYear, semester))
                        {
                            var studentRow = dicStudentRow[semesterScoreRecord.RefStudentID];
                            foreach (var subject in semesterScoreRecord.Subjects.Keys)
                            {
                                string insertStrSemsScore = "S" + semester + "/" + subject.FixFieldName() + "/SemsScore";
                                if (!mergeDT.Columns.Contains(insertStrSemsScore))
                                    mergeDT.Columns.Add(insertStrSemsScore);
                                studentRow[mergeDT.Columns.IndexOf(insertStrSemsScore)] = "" + semesterScoreRecord.Subjects[subject].Score;
                            }
                            string insertStrSemsScoreAvgMark = "S" + semester + "/SemsScoreAvgMark";
                            if (!mergeDT.Columns.Contains(insertStrSemsScoreAvgMark))
                                mergeDT.Columns.Add(insertStrSemsScoreAvgMark);
                            if (semesterScoreRecord.AvgScore.HasValue)
                                studentRow[mergeDT.Columns.IndexOf(insertStrSemsScoreAvgMark)] = CourseGradeB.Tool.GPA.Eval(semesterScoreRecord.AvgScore.Value).Letter;
                        }
                    }
                    #endregion
                    #region 取得Conduct資料
                    if (!mergeDT.Columns.Contains("Q1Comment"))
                        mergeDT.Columns.Add("Q1Comment");//Q1評語
                    if (!mergeDT.Columns.Contains("Q2Comment"))
                        mergeDT.Columns.Add("Q2Comment");//Q2評語
                    if (!mergeDT.Columns.Contains("Q3Comment"))
                        mergeDT.Columns.Add("Q3Comment");//Q3評語
                    if (!mergeDT.Columns.Contains("Q4Comment"))
                        mergeDT.Columns.Add("Q4Comment");//Q4評語
                    if (!mergeDT.Columns.Contains("S1Comment"))
                        mergeDT.Columns.Add("S1Comment");//S1評語
                    if (!mergeDT.Columns.Contains("S2Comment"))
                        mergeDT.Columns.Add("S2Comment");//S2評語
                    List<ConductRecord> records = _A.Select<ConductRecord>("ref_student_id in (" + string.Join(",", _ids) + ") and school_year=" + _schoolYear + " and semester=" + semester);
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
                                studentRow[mergeDT.Columns.IndexOf(semester == 1 ? "Q1Comment" : "Q3Comment")] = record.Comment;
                            else if (term == "2")
                                studentRow[mergeDT.Columns.IndexOf(semester == 1 ? "Q2Comment" : "Q4Comment")] = record.Comment;
                            else
                                studentRow[mergeDT.Columns.IndexOf(semester == 1 ? "S1Comment" : "S2Comment")] = record.Comment;
                        }

                        var _xdoc = new XmlDocument();

                        if (!string.IsNullOrWhiteSpace(record.Conduct))
                            _xdoc.LoadXml(record.Conduct);


                        #region 變動式合併欄位
                        {//檢查科目存在
                            string insertStrGroup = "DynamicCheck/Subject"
                                + "/" + subj.FixFieldName();
                            if (!mergeDT.Columns.Contains(insertStrGroup))
                                mergeDT.Columns.Add(insertStrGroup, typeof(string));

                            studentRow[mergeDT.Columns.IndexOf(insertStrGroup)] = "Y";
                        }
                        #endregion
                        int groupIndex = 0;
                        foreach (XmlElement elem in _xdoc.SelectNodes("//Conduct"))
                        {
                            string group = elem.GetAttribute("Group");
                            #region 變動式合併欄位
                            if (elem.GetAttribute("Group") != "Conduct")
                            {//group文字
                                groupIndex++;
                                string insertStrGroup = "DynamicField"
                                    + "/" + subj.FixFieldName()
                                    + "/Group/" + groupIndex;
                                if (!mergeDT.Columns.Contains(insertStrGroup))
                                    mergeDT.Columns.Add(insertStrGroup, typeof(string));

                                studentRow[mergeDT.Columns.IndexOf(insertStrGroup)] = group;
                            }
                            #endregion
                            int itemIndex = 0;
                            foreach (XmlElement item in elem.SelectNodes("Item"))
                            {
                                string title = item.GetAttribute("Title");
                                string grade = item.GetAttribute("Grade");

                                #region 全固定式合併欄位
                                {
                                    string insertStrC = "???????";
                                    if (term == "1")
                                    {
                                        insertStrC = "Q" + (semester * 2 - 1) + "/" + subj.FixFieldName() + "/" + group.FixFieldName() + "/" + title.FixFieldName();
                                    }
                                    else if (term == "2")
                                    {
                                        insertStrC = "Q" + (semester * 2) + "/" + subj.FixFieldName() + "/" + group.FixFieldName() + "/" + title.FixFieldName();
                                    }
                                    else
                                    {
                                        insertStrC = "S" + semester + "/" + subj.FixFieldName() + "/" + group.FixFieldName() + "/" + title.FixFieldName();
                                    }
                                    if (!mergeDT.Columns.Contains(insertStrC))
                                        mergeDT.Columns.Add(insertStrC);

                                    studentRow[mergeDT.Columns.IndexOf(insertStrC)] = grade;
                                }
                                #endregion

                                #region 變動式合併欄位
                                if (elem.GetAttribute("Group") != "Conduct")
                                {
                                    {//title文字
                                        itemIndex++;
                                        string insertStrGroupItem = "DynamicField"
                                           + "/" + subj.FixFieldName()
                                           + "/Group/" + groupIndex
                                           + "/Item/" + itemIndex;
                                        if (!mergeDT.Columns.Contains(insertStrGroupItem))
                                            mergeDT.Columns.Add(insertStrGroupItem, typeof(string));
                                        studentRow[mergeDT.Columns.IndexOf(insertStrGroupItem)] = title;
                                    }

                                    string insertStrC = "???????";

                                    if (term == "1")
                                    {
                                        insertStrC = "Dynamic"
                                            + "/" + "Q" + (semester * 2 - 1)
                                            + "/" + subj.FixFieldName()
                                            + "/Group/" + groupIndex
                                            + "/Item/" + itemIndex;
                                    }
                                    else if (term == "2")
                                    {
                                        insertStrC = "Dynamic"
                                            + "/" + "Q" + (semester * 2)
                                            + "/" + subj.FixFieldName()
                                            + "/Group/" + groupIndex
                                            + "/Item/" + itemIndex;
                                    }
                                    else
                                    {
                                        insertStrC = "Dynamic"
                                            + "/" + "S" + semester
                                            + "/" + subj.FixFieldName()
                                            + "/Group/" + groupIndex
                                            + "/Item/" + itemIndex;
                                    }
                                    if (!mergeDT.Columns.Contains(insertStrC))
                                        mergeDT.Columns.Add(insertStrC);

                                    studentRow[mergeDT.Columns.IndexOf(insertStrC)] = grade;
                                }

                                #endregion
                            }
                        }
                    }
                    #endregion
                }

                if (fieldMode)
                {
                    List<string> q1List = new List<string>();
                    List<string> s1List = new List<string>();
                    List<string> dynamicCheck = new List<string>();
                    List<string> dynamicField = new List<string>();
                    List<string> dynamicAnsQ1 = new List<string>();
                    List<string> dynamicAnsS1 = new List<string>();

                    foreach (DataColumn col in mergeDT.Columns)
                    {
                        if (col.Caption.StartsWith("Q1"))
                        {
                            q1List.Add(col.Caption);
                        }
                        if (col.Caption.StartsWith("S1"))
                        {
                            s1List.Add(col.Caption);
                        }
                        if (col.Caption.StartsWith("DynamicCheck/"))
                        {
                            dynamicCheck.Add(col.Caption);
                        }
                        if (col.Caption.StartsWith("DynamicField/"))
                        {
                            dynamicField.Add(col.Caption);
                        }
                        if (col.Caption.StartsWith("Dynamic/Q1"))
                        {
                            dynamicAnsQ1.Add(col.Caption);
                        }
                        if (col.Caption.StartsWith("Dynamic/S1"))
                        {
                            dynamicAnsS1.Add(col.Caption);
                        }
                    }
                    foreach (var q1 in q1List)
                    {
                        mergeDT.Columns.Remove(q1);
                        mergeDT.Columns.Add(q1);
                    }
                    foreach (var q1 in q1List)
                    {
                        var key = "Q2" + q1.Substring(2);
                        if (mergeDT.Columns.Contains(key))
                            mergeDT.Columns.Remove(key);
                        mergeDT.Columns.Add(key);
                    }
                    foreach (var q1 in q1List)
                    {
                        var key = "Q3" + q1.Substring(2);
                        if (mergeDT.Columns.Contains(key))
                            mergeDT.Columns.Remove(key);
                        mergeDT.Columns.Add(key);
                    }
                    foreach (var q1 in q1List)
                    {
                        var key = "Q4" + q1.Substring(2);
                        if (mergeDT.Columns.Contains(key))
                            mergeDT.Columns.Remove(key);
                        mergeDT.Columns.Add(key);
                    }
                    foreach (var s1 in s1List)
                    {
                        mergeDT.Columns.Remove(s1);
                        mergeDT.Columns.Add(s1);
                    }
                    foreach (var s1 in s1List)
                    {
                        var key = "S2" + s1.Substring(2);
                        if (mergeDT.Columns.Contains(key))
                            mergeDT.Columns.Remove(key);
                        mergeDT.Columns.Add(key);
                    }

                    dynamicCheck.Sort();
                    foreach (var df in dynamicCheck)
                    {
                        mergeDT.Columns.Remove(df);
                        mergeDT.Columns.Add(df);
                    }
                    dynamicField.Sort();
                    foreach (var df in dynamicField)
                    {
                        mergeDT.Columns.Remove(df);
                        mergeDT.Columns.Add(df);
                    }
                    dynamicAnsQ1.Sort();
                    for (int i = 1; i <= 4; i++)
                    {
                        foreach (var q1 in dynamicAnsQ1)
                        {
                            var key = "Dynamic/Q" + i + q1.Substring(10);
                            if (mergeDT.Columns.Contains(key))
                                mergeDT.Columns.Remove(key);
                            mergeDT.Columns.Add(key);
                        }
                    }
                    dynamicAnsS1.Sort();
                    for (int i = 1; i <= 2; i++)
                    {
                        foreach (var s1 in dynamicAnsS1)
                        {
                            var key = "Dynamic/S" + i + s1.Substring(10);
                            if (mergeDT.Columns.Contains(key))
                                mergeDT.Columns.Remove(key);
                            mergeDT.Columns.Add(key);
                        }
                    }


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
                        ReportConfiguration reportConfiguration = new Campus.Report.ReportConfiguration("ibshGradeYearReport.GradeYearReportCard.SY" + _SchoolYear + ".G" + grade);
                        if (reportConfiguration.Template != null)
                        {
                            var doc = new Document(new MemoryStream(reportConfiguration.Template.ToBinary()));

                            foreach (var field in doc.MailMerge.GetFieldNames())
                            {

                                if ((field.StartsWith("DynamicCheck/") || field.StartsWith("DynamicField/"))
                                    && !mergeDT.Columns.Contains(field))
                                {
                                    mergeDT.Columns.Add(field);
                                }
                            }

                            doc.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
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

        private void cboSY_SelectedIndexChanged(object sender, EventArgs e)
        {
            _SchoolYear = int.Parse(cboSY.Text);
        }
    }

    static class MergeFieldFix
    {
        public static string FixFieldName(this string item)
        {
            return item.Trim().Replace(' ', '_').Replace('"', '_').Replace('“', '_').Replace('”', '_');
        }
    }

    class MailMerge_MergeField : Aspose.Words.Reporting.IFieldMergingCallback
    {
        public void FieldMerging(Aspose.Words.Reporting.FieldMergingArgs args)
        {
            Node deleteNode = null;
            if (args.FieldName.StartsWith("DynamicField/"))
            {
                if (("" + args.FieldValue) == "")
                {
                    Node node = args.Field.Start;
                    while (node.ParentNode != null)
                    {
                        if (node is Aspose.Words.Tables.Row)
                        {
                            deleteNode = node;
                            break;
                        }
                        node = node.ParentNode;
                    }
                }
            }
            else if (args.FieldName.StartsWith("DynamicCheck/"))
            {
                //DocumentBuilder builder = new DocumentBuilder(args.Document);
                //builder.MoveToMergeField(args.DocumentFieldName, false, false);
                Node node = args.Field.Start;
                if (("" + args.FieldValue) == "")
                {
                    while (node.ParentNode != null)
                    {
                        if (node is Aspose.Words.Tables.Row)
                        {
                            deleteNode = node;
                            break;
                        }
                        node = node.ParentNode;
                    }
                }
                else
                {
                    args.Text = "";
                }
            }
            if (deleteNode != null)
                deleteNode.Remove();

            //if (args.FieldValue != null)
            //{
            //    DocumentBuilder builder = new DocumentBuilder(args.Document);
            //    builder.MoveToMergeField(args.DocumentFieldName, false, false);
            //    if (args.FieldValue is mailmergeSpecial)
            //    {
            //        builder.CellFormat.HorizontalMerge = (args.FieldValue as mailmergeSpecial).cellmerge;
            //    }
            //    builder.CellFormat.FitText = true;
            //    builder.CellFormat.WrapText = true;
            //    if (args.FieldValue is mailmergeSpecial)
            //    {
            //        args.Text = (args.FieldValue as mailmergeSpecial).value;
            //    }
            //}
        }
        public void ImageFieldMerging(Aspose.Words.Reporting.ImageFieldMergingArgs args)
        {
        }
    }
}
