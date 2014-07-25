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

        private const string config = "JH.IBSH.Report.PeriodicalExam.Config.1";
        private Dictionary<string, ReportConfiguration> custConfigs = new Dictionary<string, ReportConfiguration>();
        ReportConfiguration conf = new Campus.Report.ReportConfiguration(config);
        public string current = "";
        class filter {
            public SchoolYearSemester sys;
            public int exam;
        }
        public MainForm()
        {
            InitializeComponent();
            #region 設定comboBox選單
            foreach (string item in getCustConfig())
            {
                if (!string.IsNullOrWhiteSpace(item))
                {
                    custConfigs.Add(item, new ReportConfiguration(configNameRule(item)));
                    comboBoxEx1.Items.Add(item);
                }
            }
            foreach (string item in new string[] { "100", "101", "102" })
            {
                comboBoxEx2.Items.Add(item);
            }
            foreach (string item in new string[] { "1", "2" })
            {
                comboBoxEx3.Items.Add(item);
            }
            foreach (string item in new string[] { "1", "2" })
            {
                comboBoxEx4.Items.Add(item);
            }
            comboBoxEx1.Items.Add("新增");
            comboBoxEx1.SelectedIndex = 0;
            #endregion
            _bgw.DoWork += new DoWorkEventHandler(_bgw_DoWork);
            _bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgw_RunWorkerCompleted);
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            if (value == "新增") return;
            //畫面內容(範本內容,預設樣式
            Campus.Report.TemplateSettingForm TemplateForm;
            if (custConfigs[current].Template == null)
            {
                custConfigs[current].Template = new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(custConfigs[current].Template, new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = current + "樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                custConfigs[current].Template = TemplateForm.Template;
                custConfigs[current].Save();
            }
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            if (value == "新增") return;
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請先選擇學生");
                return;
            }
            btnPrint.Enabled = false;
            _bgw.RunWorkerAsync(new filter
                {
                    sys = new SchoolYearSemester(int.Parse(comboBoxEx2.Text), int.Parse(comboBoxEx3.Text)),
                    exam = int.Parse(comboBoxEx4.Text)
                }
            );
        }
        void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            filter f = (filter)e.Argument;
            Document document = new Document();
            Document template = (custConfigs[current].Template != null) //單頁範本
                 ? custConfigs[current].Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources.樣板, Campus.Report.TemplateType.Word).ToDocument();
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                return;
            List<string> sids = K12.Presentation.NLDPanels.Student.SelectedSource;

            int SchoolYear = f.sys.SchoolYear;
            int Semester = f.sys.Semester;
            int Exam = f.exam;

            DataTable dt = tool._Q.Select(@"select student.id,student.english_name,student.name,student.student_number,student.seat_no,class.class_name,teacher.teacher_name,class.grade_year,course.id as course_id,course.period as period,course.credit as credit,course.subject as subject,$ischool.subject.list.group as group,$ischool.subject.list.type as type,$ischool.subject.list.english_name as subject_english_name,sc_attend.ref_student_id as student_id,sce_take.ref_sc_attend_id as sc_attend_id,sce_take.ref_exam_id as exam_id,xpath_string(sce_take.extension,'//Score') as score 
from sc_attend
join sce_take on sce_take.ref_sc_attend_id=sc_attend.id
join course on course.id=sc_attend.ref_course_id
join $ischool.course.extend on $ischool.course.extend.ref_course_id=course.id
left join student on student.id=sc_attend.ref_student_id
left join class on student.ref_class_id=class.id
left join $ischool.subject.list on course.subject=$ischool.subject.list.name 
left join teacher on teacher.id=class.ref_teacher_id
where sc_attend.ref_student_id in (" +string.Join("," ,sids) + ") and course.school_year=" + SchoolYear + " and course.semester=" + Semester + " and sce_take.ref_exam_id = " + Exam + "");
            Dictionary<string, List<CustomSCETakeRecord>> dscetr = new Dictionary<string, List<CustomSCETakeRecord>>();
            foreach (DataRow row in dt.Rows)
            {
                string id = ""+row["id"] ;
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

            Dictionary<string, object> mailmerge = new Dictionary<string, object>();
            
            foreach (KeyValuePair<string,List<CustomSCETakeRecord>> row in dscetr )
            {
                mailmerge.Clear();
                int ClassGradeYear = int.Parse(row.Value[0].GradeYear); //should be the same
                mailmerge.Add("學年", SchoolYear);
                mailmerge.Add("學期", Semester);
                mailmerge.Add("學段", Exam);
                mailmerge.Add("班級", row.Value[0].ClassName);
                mailmerge.Add("座號", row.Value[0].SeatNo);
                mailmerge.Add("姓名", row.Value[0].Name);
                mailmerge.Add("英文名", row.Value[0].EnglishName);

                mailmerge.Add("學校名稱", School.ChineseName);
                mailmerge.Add("學校英文名稱", School.EnglishName);
                #region 學生成績
                int subjecti = 1 ;
                switch (ClassGradeYear)
                {
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        foreach (CustomSCETakeRecord item in row.Value)
                        {
                            mailmerge.Add(string.Format("科目{0}", subjecti), item.Name + " " + item.SubjectEnglishName);
                            mailmerge.Add(string.Format("成績{0}", subjecti), CourseGradeB.Tool.GPA.Eval(item.Score).Letter);
                            subjecti++;
                        }
                        
                        break;
                    case 9:
                    case 10:
                    case 11:
                    case 12:
                        int personalCreditCount = 0;
                        decimal personalGPACount = 0, personalAverageCount = 0;
                        foreach (CustomSCETakeRecord item in row.Value)
                        {
                            mailmerge.Add(string.Format("科目{0}", subjecti), item.Name + " " + item.SubjectEnglishName);
                            mailmerge.Add(string.Format("成績{0}", subjecti), item.Score);
                            personalCreditCount += item.CourseCredit;
                            CourseGradeB.Tool.GPA _gpa = CourseGradeB.Tool.GPA.Eval(item.Score) ;
                            if ( CourseGradeB.Tool.SubjectType.Honor == item.CourseType )
                                 personalGPACount +=  item.CourseCredit * _gpa.Honors ;
                            else if (CourseGradeB.Tool.SubjectType.Regular == item.CourseType)
                                personalGPACount += item.CourseCredit * _gpa.Regular;

                            personalAverageCount += item.CourseCredit * item.Score;
                            subjecti++;
                        }
                        if (personalCreditCount > 0)
                        {
                            mailmerge.Add("科目平均", personalAverageCount / personalCreditCount);
                            mailmerge.Add("GPA", personalGPACount / personalCreditCount);
                        }
                        GradePeriodicalExamGPARecord gpegpar = GradePeriodicalExamGPA.GetGradePeriodicalExamGPARecord(SchoolYear, Semester, ClassGradeYear, Exam);
                        List<GPADistributionPart> gpad = GradePeriodicalExamGPA.toGPADistribution(gpegpar);
                        int parti = 1;
                        foreach (GPADistributionPart gpadp in gpad)
                        {
                            mailmerge.Add(string.Format("GPA分段{0}", parti), gpadp.GPACeiling + "~" + gpadp.GPAFloor);
                            mailmerge.Add(string.Format("GPA計數{0}", parti), gpadp.Count);
                            parti++;
                        }
                        break;
                }
                mailmerge.Add(string.Format("科目{0}", subjecti), "(以下空白)");
                mailmerge.Add(string.Format("成績{0}", subjecti), "");
                for (; subjecti <= 20; subjecti++)
                {
                    mailmerge.Add(string.Format("科目{0}", subjecti), "");
                    mailmerge.Add(string.Format("成績{0}", subjecti), "");
                }
                 #endregion
                Document each = (Document)template.Clone(true);
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
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var value = (string)comboBoxEx1.SelectedItem;
            switch (value)
            {
                case "新增":
                    AddNew input = new AddNew();
                    if (input.ShowDialog() == DialogResult.OK)
                    {
                        input.name = System.Text.RegularExpressions.Regex.Replace(input.name, @"[\W_]+", "");
                        if (string.IsNullOrWhiteSpace(input.name))
                            FISCA.Presentation.Controls.MsgBox.Show("請輸入樣板名稱(中文或英文字母)");
                        else if (custConfigs.ContainsKey(input.name))
                            FISCA.Presentation.Controls.MsgBox.Show("樣板名稱已存在");
                        else
                        {
                            ReportConfiguration tmp_conf = new ReportConfiguration(configNameRule(input.name));
                            if (input.Template != null)
                                tmp_conf.Template = new ReportTemplate(input.Template);
                            tmp_conf.Save();
                            custConfigs.Add(input.name, tmp_conf);
                            addCustConfig(input.name);
                            comboBoxEx1.Items.Insert(0, input.name);
                            comboBoxEx1.SelectedIndex = 0;
                        }
                    }
                    break;
                default:
                    current = value;
                    break;
            }
        }
        private void delete_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string value = (string)comboBoxEx1.SelectedItem;
            switch (value)
            {
                case "新增":
                    break;
                default:
                    if (custConfigs.ContainsKey(value))
                    {
                        custConfigs[value].Template = null;
                        custConfigs[value].Save();
                        custConfigs.Remove(value);
                        comboBoxEx1.Items.Remove(value);
                        delCustConfig(value);
                    }
                    break;
            }
            comboBoxEx1.SelectedIndex = 0;
            current = (string)comboBoxEx1.SelectedItem;
        }
        private void addCustConfig(string custConfig)
        {
            List<string> tmp = conf.GetString("customs", "").Split(';').ToList<string>();
            tmp.Add(System.Text.RegularExpressions.Regex.Replace(custConfig, @"[\W_]+", ""));
            conf.SetString("customs", string.Join(";", tmp));
            conf.Save();
        }
        private void delCustConfig(string custConfig)
        {
            List<string> tmp = conf.GetString("customs", "").Split(';').ToList<string>();
            tmp.Remove(custConfig);
            conf.SetString("customs", string.Join(";", tmp));
            conf.Save();
        }
        private string[] getCustConfig()
        {
            return conf.GetString("customs", "").Split(';');
        }
        private static string configNameRule(string custConfigName)
        {
            return config + "." + custConfigName;
        }
        
    }
    
}
