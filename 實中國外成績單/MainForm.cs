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

namespace JH.IBSH.Report.Foreign
{
    public partial class MainForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bgw = new BackgroundWorker();

        private const string config3_6 = "JH.IBSH.Report.Foreign.Config.3_6";
        private const string config7_8 = "JH.IBSH.Report.Foreign.Config.7_8";
        private const string config9_12 = "JH.IBSH.Report.Foreign.Config.9_12";

        public static ReportConfiguration ReportConfiguration3_6 = new Campus.Report.ReportConfiguration(config3_6);
        public static ReportConfiguration ReportConfiguration7_8 = new Campus.Report.ReportConfiguration(config7_8);
        public static ReportConfiguration ReportConfiguration9_12 = new Campus.Report.ReportConfiguration(config9_12);

        class filter {
            public string GradeType;
        }
        class grade
        {
            public int grade_year;
            public int school_year;
            public SemesterScoreRecord sems1; // 1st semester 
            public SemesterScoreRecord sems2;// 2nd semester
        }
        class course
        {
            public string sems1_title;
            public decimal sems1_score;
            public string sems2_title;
            public decimal sems2_score;
        }
        string[] GradeTypes = new string[] { "3~6", "7~8", "9~12" };
        public MainForm()
        {
            InitializeComponent();
            #region 設定comboBox選單
            foreach (string item in GradeTypes)
            {
                comboBoxEx2.Items.Add(item);
            }
            comboBoxEx2.SelectedIndex = 0;
            #endregion
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
                    GradeType = comboBoxEx2.Text
                }
            );
        }
        void _bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            filter f = (filter)e.Argument;
            Document document = new Document();
            Document template ;
            if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                return;
            List<string> sids = K12.Presentation.NLDPanels.Student.SelectedSource;

            Dictionary<string, SemesterHistoryRecord> dshr = SemesterHistory.SelectByStudentIDs(sids).ToDictionary(x => x.RefStudentID, x => x);
            Dictionary<string, StudentRecord> dsr = Student.SelectByIDs(sids).ToDictionary(x => x.ID, x => x);
            Dictionary<string, SemesterScoreRecord> dssr = SemesterScore.SelectByStudentIDs(sids).ToDictionary(x => x.RefStudentID + "#" + x.SchoolYear + "#" + x.Semester, x => x);
            DataTable dt = tool._Q.Select("select id,enrollment_school_year from student where id in ('"+string.Join("','",sids)+"')");
            Dictionary<string, int> desy = new Dictionary<string, int>();
            int tmp;
            foreach (DataRow row in dt.Rows)
            {
                if ( !desy.ContainsKey(""+row["id"]))
                    desy.Add(""+row["id"],0);
                if ( int.TryParse(""+row["enrollment_school_year"],out tmp) )
                    desy[""+row["id"]] = tmp ;
            }
            List<string> gys;
            int domainDicKey;
            switch (f.GradeType)
            {
                case "3~6":
                case "6":
                    gys = new List<string> { "3", "4", "5", "6" };
                    domainDicKey = 6;
                    template = (ReportConfiguration3_6.Template != null) //單頁範本
                 ? ReportConfiguration3_6.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word).ToDocument();
                    break;
                case "7~8":
                case "8":
                    gys = new List<string> { "7", "8" };
                    domainDicKey = 8;
                    template = (ReportConfiguration7_8.Template != null) //單頁範本
                 ? ReportConfiguration7_8.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word).ToDocument();
                    break;
                case "9~12":
                case "12":
                    gys = new List<string> { "9", "10", "11", "12" };
                    domainDicKey = 12;
                    template = (ReportConfiguration9_12.Template != null) //單頁範本
                 ? ReportConfiguration9_12.Template.ToDocument()
                 : new Campus.Report.ReportTemplate(Properties.Resources._12樣版, Campus.Report.TemplateType.Word).ToDocument();
                    break;
                default:
                    return;
            }
            List<CourseGradeB.Tool.Domain> cgbdl = CourseGradeB.Tool.DomainDic[domainDicKey];
            cgbdl.Sort(delegate(CourseGradeB.Tool.Domain x, CourseGradeB.Tool.Domain y)
            {
                return x.DisplayOrder.CompareTo(y.DisplayOrder);
            });
            int domainCount;
            Dictionary<string, object> mailmerge = new Dictionary<string, object>();
            foreach (KeyValuePair<string,SemesterHistoryRecord> row in dshr )
            {//學生
                mailmerge.Clear();
                domainCount = 1;
                foreach (CourseGradeB.Tool.Domain domain in cgbdl)
                {
                    mailmerge.Add(string.Format("群{0}", domainCount), domain.Name);
                    mailmerge.Add(string.Format("群{0}_時數", domainCount), domain.Hours);
                    domainCount++;
                }
                #region 學生資料
                StudentRecord sr = dsr[row.Key];
                mailmerge.Add("姓名", sr.Name);
                mailmerge.Add("英文名", sr.EnglishName);
                string gender ;
                switch (sr.Gender)
                {
                    case "男":
                        gender = "Male";
                        break;
                    case "女":
                        gender = "Female";
                        break;
                    default:
                        gender = sr.Gender;
                        break;
                }
                mailmerge.Add("性別", gender);
                mailmerge.Add("國籍", sr.Nationality);
                mailmerge.Add("生日", sr.Birthday.HasValue?sr.Birthday.Value.ToString("d-MMMM-yyyy", new System.Globalization.CultureInfo("en-US")):"");
                string esy = "", edog = "";
                if (desy.ContainsKey(row.Key) && desy[row.Key] != 0)
                {
                    esy = "August-" + (desy[row.Key] + 1911);
                    edog = "June-" + (desy[row.Key] + 1911 + 12);
                }
                mailmerge.Add("入學日期", esy);
                mailmerge.Add("預計畢業日期", edog);

                //mailmerge.Add("Registrar", row.Value[0].SeatNo);
                //mailmerge.Add("Dean", row.Value[0].SeatNo);
                //mailmerge.Add("Principal", row.Value[0].SeatNo);
                #endregion

                #region 學生成績
                Dictionary<int, grade> dgrade = new Dictionary<int, grade>();
                //群 , 科目 , 分數
                Dictionary<string, Dictionary<string, course>> dcl = new Dictionary<string, Dictionary<string, course>>();
                foreach (SemesterHistoryItem shi in row.Value.SemesterHistoryItems)
                {
                    if (!gys.Contains(""+shi.GradeYear))
                        continue;

                    int _gradeYear = shi.GradeYear ;
                    string key = shi.RefStudentID + "#" + shi.SchoolYear + "#" + shi.Semester;
                    if (!dgrade.ContainsKey(_gradeYear))
                    {
                        dgrade.Add(_gradeYear, new grade()
                        {
                            grade_year = shi.GradeYear,
                            school_year = shi.SchoolYear
                        });
                    }
                    if (shi.Semester == 1 && dssr.ContainsKey(key))
                        dgrade[_gradeYear].sems1 = dssr[key];
                    else if (shi.Semester == 2 && dssr.ContainsKey(key))
                        dgrade[_gradeYear].sems2 = dssr[key];

                }
                int gradeCount = 1;
                foreach (string gy in gys)
                {//級別_
                    mailmerge.Add(string.Format("級別{0}", gradeCount), gy);
                    mailmerge.Add(string.Format("學年度{0}", gradeCount), "");
                    if (dgrade.ContainsKey(int.Parse(gy)))
                    {
                        grade g = dgrade[int.Parse(gy)];
                        mailmerge[string.Format("學年度{0}", gradeCount)]= (g.school_year + 1911) + "-" + (g.school_year + 1912);
                        if (g.sems1 != null)
                        {
                            foreach (KeyValuePair<string, SubjectScore> ss in g.sems1.Subjects)
                            {
                                if (!dcl.ContainsKey(ss.Value.Domain))
                                    dcl.Add(ss.Value.Domain, new Dictionary<string, course>());
                                if (!dcl[ss.Value.Domain].ContainsKey(ss.Value.Subject))
                                    dcl[ss.Value.Domain].Add(ss.Value.Subject, new course());
                                dcl[ss.Value.Domain][ss.Value.Subject].sems1_title = ss.Value.Subject;
                                dcl[ss.Value.Domain][ss.Value.Subject].sems1_score = ss.Value.Score.HasValue ? ss.Value.Score.Value : 0;
                            }
                        }
                        if (g.sems2 != null)
                        {
                            foreach (KeyValuePair<string, SubjectScore> ss in g.sems2.Subjects)
                            {
                                if (!dcl.ContainsKey(ss.Value.Domain))
                                    dcl.Add(ss.Value.Domain, new Dictionary<string, course>());
                                if (!dcl[ss.Value.Domain].ContainsKey(ss.Value.Subject))
                                    dcl[ss.Value.Domain].Add(ss.Value.Subject, new course());
                                dcl[ss.Value.Domain][ss.Value.Subject].sems2_title = ss.Value.Subject;
                                dcl[ss.Value.Domain][ss.Value.Subject].sems2_score = ss.Value.Score.HasValue ? ss.Value.Score.Value : 0;
                            }
                        }
                        #region 科目成績
                    }
                    #endregion
                    domainCount = 1;
                    foreach (CourseGradeB.Tool.Domain domain in cgbdl)
                    {//群
                        if (!dcl.ContainsKey(domain.Name))
                        {
                            mailmerge.Add(string.Format("群{0}_級{1}_學期1_科目{2}", domainCount, gradeCount, 1), "");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期2_科目{2}", domainCount, gradeCount, 1), "與前方合併");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期1_科目成績{2}", domainCount, gradeCount, 1), "N/A");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期2_科目成績{2}", domainCount, gradeCount, 1), "N/A");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期1_科目{2}", domainCount, gradeCount, 2), "");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期2_科目{2}", domainCount, gradeCount, 2), "與前方合併");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期1_科目成績{2}", domainCount, gradeCount, 2), "N/A");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期2_科目成績{2}", domainCount, gradeCount, 2), "N/A");
                            domainCount++;
                            continue;
                        }
                        int courseCount = 1;
                        foreach (course item in dcl[domain.Name].Values)
                        {
                            mailmerge.Add(string.Format("群{0}_級{1}_學期1_科目{2}", domainCount, gradeCount, courseCount), item.sems1_title);
                            mailmerge.Add(string.Format("群{0}_級{1}_學期2_科目{2}", domainCount, gradeCount, courseCount), item.sems2_title);
                            //都存在且相同才需要合併
                            if (item.sems1_title != null && item.sems2_title != null && item.sems1_title != item.sems2_title)
                                mailmerge[string.Format("群{0}_級{1}_學期2_科目{2}", domainCount, gradeCount, courseCount)] = "與前方合併";
                            mailmerge.Add(string.Format("群{0}_級{1}_學期1_科目成績{2}", domainCount, gradeCount, courseCount),
                                item.sems1_title != null ? "" + item.sems1_score : "N/A");
                            mailmerge.Add(string.Format("群{0}_級{1}_學期2_科目成績{2}", domainCount, gradeCount, courseCount),
                                item.sems2_title != null ? "" + item.sems2_score : "N/A");
                            courseCount++;
                        }
                        domainCount++;
                    }
                    gradeCount++;
                }
                string key2 = sr.ID+"#"+School.DefaultSchoolYear+"#"+School.DefaultSemester ;
                if (dssr.ContainsKey(key2))

                    mailmerge.Add("GPA", dssr[key2].CumulateGPA);
                GradeCumulateGPARecord gcgpar;
                if (sr.Class.GradeYear.HasValue)
                {
                    gcgpar = GradeCumulateGPA.GetGradeCumulateGPARecord(int.Parse(School.DefaultSchoolYear), int.Parse(School.DefaultSemester), sr.Class.GradeYear.Value);

                    mailmerge.Add("級最高GPA", gcgpar.MaxGPA);
                    mailmerge.Add("級平均GPA", gcgpar.AvgGPA);
                }
                 #endregion
                Document each = (Document)template.Clone(true);

                each.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
                each.MailMerge.Execute(mailmerge.Keys.ToArray(), mailmerge.Values.ToArray());
                document.Sections.Add(document.ImportNode(each.FirstSection, true));
            }
            
            document.Sections.RemoveAt(0);
            e.Result = document;
        }

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            if (e.FieldValue != null && e.FieldValue.ToString() == "與前方合併")
            {
                DocumentBuilder builder = new DocumentBuilder(e.Document);
                builder.MoveToField(e.Field, true);
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            }
        }
        void _bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document inResult = (Document)e.Result;
            btnPrint.Enabled = true;
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = "國外成績單";

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
    }
    
}
