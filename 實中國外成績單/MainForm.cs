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
using Aspose.Words.Tables;
//using Campus.ePaper;

namespace JH.IBSH.Report.Foreign
{
    public partial class MainForm : FISCA.Presentation.Controls.BaseForm
    {
        private const string config3_6 = "JH.IBSH.Report.Foreign.Config.3_6";
        private const string config7_8 = "JH.IBSH.Report.Foreign.Config.7_8";
        private const string config9_12 = "JH.IBSH.Report.Foreign.Config.9_12";

        public static ReportConfiguration ReportConfiguration3_6 = new Campus.Report.ReportConfiguration(config3_6);
        public static ReportConfiguration ReportConfiguration7_8 = new Campus.Report.ReportConfiguration(config7_8);
        public static ReportConfiguration ReportConfiguration9_12 = new Campus.Report.ReportConfiguration(config9_12);

        public static CourseGradeB.Tool.Domain English;
        public static CourseGradeB.Tool.Domain WesternSocialStudies;
        public static CourseGradeB.Tool.Domain Elective;

        public bool Choose3to6Grade = false;
        public bool Choose7to8Grade = false;

        class filter
        {
            public string GradeType;
        }
        class grade
        {
            public int grade_year;
            public int school_year;
            public int semester;
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
        class mailmergeSpecial
        {
            public Aspose.Words.Tables.CellMerge cellmerge;
            public string value;
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

            filter f = new filter
            {
                GradeType = comboBoxEx2.Text
            };
            Document document = new Document();
            BackgroundWorker bgw = new BackgroundWorker();
            Dictionary<StudentRecord, List<string>> errCheck = new Dictionary<StudentRecord, List<string>>();
            bgw.DoWork += delegate
            {
                #region DoWork
                Byte[] template;
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count <= 0)
                    return;
                List<string> sids = K12.Presentation.NLDPanels.Student.SelectedSource;

                Dictionary<string, SemesterHistoryRecord> dshr = SemesterHistory.SelectByStudentIDs(sids).ToDictionary(x => x.RefStudentID, x => x);
                Dictionary<string, StudentRecord> dsr = Student.SelectByIDs(sids).ToDictionary(x => x.ID, x => x);
                Dictionary<string, SemesterScoreRecord> dssr = SemesterScore.SelectByStudentIDs(sids).ToDictionary(x => x.RefStudentID + "#" + x.SchoolYear + "#" + x.Semester, x => x);
                DataTable dt = tool._Q.Select("select ref_student_id,entrance_date,leaving_date from $jhcore_bilingual.studentrecordext where ref_student_id in ('" + string.Join("','", sids) + "')");
                Dictionary<string, edld> dedld = new Dictionary<string, edld>();
                DateTime tmp;
                foreach (DataRow row in dt.Rows)
                {
                    if (!dedld.ContainsKey("" + row["ref_student_id"]))
                        dedld.Add("" + row["ref_student_id"], new edld() { });
                    if (DateTime.TryParse("" + row["entrance_date"], out tmp))
                        dedld["" + row["ref_student_id"]].entrance_date = tmp;
                    if (DateTime.TryParse("" + row["leaving_date"], out tmp))
                        dedld["" + row["ref_student_id"]].leaving_date = tmp;
                }
                List<string> gradeYearList;
                int domainDicKey;


                switch (f.GradeType)
                {
                    case "3~6":
                    case "6":
                        gradeYearList = new List<string> { "3", "4", "5", "6" };
                        domainDicKey = 6;
                        template = (ReportConfiguration3_6.Template != null) //單頁範本
                     ? ReportConfiguration3_6.Template.ToBinary()
                     : new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word).ToBinary();
                        Choose3to6Grade = true;
                        break;
                    case "7~8":
                    case "8":
                        gradeYearList = new List<string> { "7", "8" };
                        domainDicKey = 8;
                        template = (ReportConfiguration7_8.Template != null) //單頁範本
                     ? ReportConfiguration7_8.Template.ToBinary()
                     : new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word).ToBinary();
                        Choose7to8Grade = true;
                        break;
                    case "9~12":
                    case "12":
                        gradeYearList = new List<string> { "9", "10", "11", "12" };
                        domainDicKey = 12;
                        template = (ReportConfiguration9_12.Template != null) //單頁範本
                     ? ReportConfiguration9_12.Template.ToBinary()
                     : new Campus.Report.ReportTemplate(Properties.Resources._9_12_grade樣板, Campus.Report.TemplateType.Word).ToBinary();
                        break;
                    default:
                        return;
                }
                List<CourseGradeB.Tool.Domain> cgbdl = CourseGradeB.Tool.DomainDic[domainDicKey];

                // 2016/5/18 穎驊新增功能，因原本3~6年級其Domain 並無English、Western Social Studies ，會造成如果沒有成績，而不顯示N/A直接空白的問題
                if (Choose3to6Grade)
                {
                    English.Hours = 6;
                    English.Name = "English";
                    English.ShortName = "English";

                    WesternSocialStudies.Hours = 2;
                    WesternSocialStudies.Name = "Western Social Studies";
                    WesternSocialStudies.ShortName = "W.S.S";

                    cgbdl.Add(English);
                    cgbdl.Add(WesternSocialStudies);
                }
                // 2016/5/20(蔡英文上任)穎驊新增功能，因原本7~8年級其Domain 並無Elective ，會造成如果沒有成績，而不顯示N/A直接空白的問題
                if (Choose7to8Grade)
                {
                    Elective.Hours = 2;
                    Elective.Name = "Elective";
                    Elective.ShortName = "Elective";
                 
                    cgbdl.Add(Elective);
                }
              
                cgbdl.Sort(delegate(CourseGradeB.Tool.Domain x, CourseGradeB.Tool.Domain y)
                {
                    return x.DisplayOrder.CompareTo(y.DisplayOrder);
                });
                //int domainCount;
                Dictionary<string, string> NationalityMapping = K12.EduAdminDataMapping.Utility.GetNationalityMappingDict();
                Dictionary<string, object> mailmerge = new Dictionary<string, object>();

                GradeCumulateGPA gcgpa = new GradeCumulateGPA();
                foreach (var studentID in dshr.Keys)
                {//學生

                    System.IO.Stream docStream = new System.IO.MemoryStream(template);
                    Document each = new Document(docStream);
                    //DocumentBuilder db = new DocumentBuilder(each);
                    //Table table = (Table)each.GetChild(NodeType.Table, 1, true);
                    //table.AllowAutoFit = true;
                    //not work,why ?

                    // 2016/4/28 取得樣板上，所有的功能變數，以利以後核對使用。
                    string[] fieldNames = each.MailMerge.GetFieldNames();

                    grade lastGrade = null;
                    mailmerge.Clear();
                    mailmerge.Add("列印日期", DateTime.Today.ToString("MMMM d, yyyy", new System.Globalization.CultureInfo("en-US")));
                    #region 學生資料
                    StudentRecord sr = dsr[studentID];
                    mailmerge.Add("學生系統編號", sr.ID);
                    mailmerge.Add("學號", sr.StudentNumber);
                    mailmerge.Add("姓名", sr.Name);
                    mailmerge.Add("英文名", sr.EnglishName);
                    string gender;
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
                    if (NationalityMapping.ContainsKey(sr.Nationality))
                        mailmerge["國籍"] = NationalityMapping[sr.Nationality];

                    mailmerge.Add("生日", sr.Birthday.HasValue ? sr.Birthday.Value.ToString("d-MMMM-yyyy", new System.Globalization.CultureInfo("en-US")) : "");
                    string esy = "", edog = "";
                    if (dedld.ContainsKey(studentID))
                    {
                        if (dedld[studentID].entrance_date != null)
                            esy = dedld[studentID].entrance_date.Value.ToString("MMMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                        if (dedld[studentID].leaving_date != null)
                            edog = dedld[studentID].leaving_date.Value.ToString("MMMM-yyyy", new System.Globalization.CultureInfo("en-US"));
                    }
                    mailmerge.Add("入學日期", esy);
                    mailmerge.Add("預計畢業日期", edog);

                    //mailmerge.Add("Registrar", row.Value[0].SeatNo);
                    //mailmerge.Add("Dean", row.Value[0].SeatNo);
                    //mailmerge.Add("Principal", row.Value[0].SeatNo);
                    #endregion

                    #region 學生成績
                    Dictionary<int, grade> dgrade = new Dictionary<int, grade>();
                    #region 整理學生成績及年級
                    foreach (SemesterHistoryItem shi in dshr[studentID].SemesterHistoryItems)
                    {
                        if (!gradeYearList.Contains("" + shi.GradeYear))
                            continue;

                        int _gradeYear = shi.GradeYear;
                        string key = shi.RefStudentID + "#" + shi.SchoolYear + "#" + shi.Semester;
                        if (!dgrade.ContainsKey(_gradeYear))
                        {
                            dgrade.Add(_gradeYear, new grade()
                            {
                                grade_year = shi.GradeYear,
                                school_year = shi.SchoolYear
                            });
                        }
                        if (shi.Semester == 1)
                        {
                            dgrade[_gradeYear].semester = 1;
                            if (dssr.ContainsKey(key))
                                dgrade[_gradeYear].sems1 = dssr[key];
                        }
                        else if (shi.Semester == 2)
                        {
                            dgrade[_gradeYear].semester = 2;
                            if (dssr.ContainsKey(key))
                                dgrade[_gradeYear].sems2 = dssr[key];
                        }

                    }
                    #endregion

                    mailmerge.Add("GPA", "");
                    int gradeCount = 1;

                    foreach (string gy in gradeYearList)
                    {//級別_
                        //群 , 科目 , 分數
                        //Dictionary<string, Dictionary<string, course>> dcl = new Dictionary<string, Dictionary<string, course>>();
                        Dictionary<string, List<SubjectScore>> dcl = new Dictionary<string, List<SubjectScore>>();
                        mailmerge.Add(string.Format("級別{0}", gradeCount), gy);
                        mailmerge.Add(string.Format("學年度{0}", gradeCount), "");

                        if (dgrade.ContainsKey(int.Parse(gy)))
                        {
                            grade g = dgrade[int.Parse(gy)];
                            mailmerge[string.Format("學年度{0}", gradeCount)] = (g.school_year + 1911) + "-" + (g.school_year + 1912);
                            foreach (var semScore in new SemesterScoreRecord[] { g.sems1, g.sems2 })
                            {
                                if (semScore != null)
                                {
                                    foreach (var subjectScore in semScore.Subjects.Values)
                                    {
                                        //if (!dcl.ContainsKey(subjectScore.Domain))
                                        //    dcl.Add(subjectScore.Domain, new Dictionary<string, course>());
                                        //if (!dcl[subjectScore.Domain].ContainsKey(subjectScore.Subject))
                                        //    dcl[subjectScore.Domain].Add(subjectScore.Subject, new course());
                                        //switch (subjectScore.Semester)
                                        //{
                                        //    case 1:
                                        //        dcl[subjectScore.Domain][subjectScore.Subject].sems1_title = subjectScore.Subject;
                                        //        dcl[subjectScore.Domain][subjectScore.Subject].sems1_score = subjectScore.Score.HasValue ? Math.Round(subjectScore.Score.Value, 0, MidpointRounding.AwayFromZero) : 0;
                                        //        break;
                                        //    case 2:
                                        //        dcl[subjectScore.Domain][subjectScore.Subject].sems2_title = subjectScore.Subject;
                                        //        dcl[subjectScore.Domain][subjectScore.Subject].sems2_score = subjectScore.Score.HasValue ? Math.Round(subjectScore.Score.Value, 0, MidpointRounding.AwayFromZero) : 0;
                                        //        break;
                                        //}

                                        if (!dcl.ContainsKey(subjectScore.Domain))
                                            dcl.Add(subjectScore.Domain, new List<SubjectScore>());

                                        subjectScore.Score = subjectScore.Score.HasValue ? Math.Round(subjectScore.Score.Value, 0, MidpointRounding.AwayFromZero) : 0;

                                        dcl[subjectScore.Domain].Add(subjectScore);
                                    }
                                }
                            }
                            //使用學期歷程最後一筆的學年度學期
                            if (g.sems1 != null)
                                mailmerge["GPA"] = g.sems1.CumulateGPA;
                            if (g.sems2 != null)
                                mailmerge["GPA"] = g.sems2.CumulateGPA;
                            lastGrade = g;
                        }
                        //檢查預設清單，缺漏處補回空資料
                        foreach (CourseGradeB.Tool.Domain domain in cgbdl)
                        {
                            //if (!dcl.ContainsKey(domain.Name))
                            //    dcl.Add(domain.Name, new Dictionary<string, course>());
                            if (!dcl.ContainsKey(domain.Name))
                                dcl.Add(domain.Name, new List<SubjectScore>());
                        }
                        foreach (var domain in dcl.Keys)
                        {
                            foreach (var semester in new int[] { 1, 2 })
                            {
                                //群
                                int courseCount = 1;
                                foreach (var item in dcl[domain])
                                {
                                    if (item.Semester == semester)
                                    {
                                        mailmerge.Add(string.Format("{0}_級{1}_學期{2}_科目{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount), item.Subject);
                                        mailmerge.Add(string.Format("{0}_級{1}_學期{2}_科目級別{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount), "Level:" + item.Level);
                                        mailmerge.Add(string.Format("{0}_級{1}_學期{2}_科目成績{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount), "" + item.Score);


                                        // 2016/4/28 穎驊筆記，下面為檢察功能，fieldName為目前樣板的所有功能變數，假如樣版沒有完整的對應功能變數，會加入錯誤訊息提醒。
                                        if (!fieldNames.Contains(string.Format("{0}_級{1}_學期{2}_科目成績{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount)))
                                        {
                                            if (!errCheck.ContainsKey(sr))
                                                errCheck.Add(sr, new List<string>());
                                            errCheck[sr].Add("合併欄位「" + domain + " 學期" + semester + " 科目成績" + courseCount + "」在樣板中不存在(科目名稱 " + item.Subject + ")。");
                                        }

                                        courseCount++;
                                        //mailmerge.Add(string.Format("{0}_級{1}_學期1_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount), item.sems1_title);
                                        //mailmerge.Add(string.Format("{0}_級{1}_學期2_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount), item.sems2_title);


                                        ////都存在且相同才需要合併
                                        //if (item.sems1_title != null && item.sems2_title != null && item.sems1_title == item.sems2_title)
                                        //{
                                        //    mailmerge[string.Format("{0}_級{1}_學期1_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.First, value = item.sems1_title };
                                        //    mailmerge[string.Format("{0}_級{1}_學期2_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.Previous, value = item.sems2_title };
                                        //}
                                        //mailmerge.Add(string.Format("{0}_級{1}_學期1_科目成績{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount),
                                        //    item.sems1_title != null ? "" + item.sems1_score : "N/A");
                                        //mailmerge.Add(string.Format("{0}_級{1}_學期2_科目成績{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount),
                                        //    item.sems2_title != null ? "" + item.sems2_score : "N/A");

                                        //courseCount++;
                                    }
                                }

                                for (; courseCount <= 3; courseCount++)
                                {
                                    mailmerge.Add(string.Format("{0}_級{1}_學期{2}_科目{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount), "");
                                    mailmerge.Add(string.Format("{0}_級{1}_學期{2}_科目級別{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount), "");
                                    mailmerge.Add(string.Format("{0}_級{1}_學期{2}_科目成績{3}", domain.Trim().Replace(" ", "_"), gradeCount, semester, courseCount), "N/A");
                                    //mailmerge.Add(string.Format("{0}_級{1}_學期1_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount), new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.First, value = "" });
                                    //mailmerge.Add(string.Format("{0}_級{1}_學期2_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount), new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.Previous, value = "" });
                                    //mailmerge.Add(string.Format("{0}_級{1}_學期1_科目成績{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount), "N/A");
                                    //mailmerge.Add(string.Format("{0}_級{1}_學期2_科目成績{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount), "N/A");
                                }
                            }

                            for (var courseCount = 1; courseCount <= 3; courseCount++)
                            {
                                var subjName1 = "" + mailmerge[string.Format("{0}_級{1}_學期{2}_科目{3}", domain.Trim().Replace(" ", "_"), gradeCount, 1, courseCount)];
                                var subjName2 = "" + mailmerge[string.Format("{0}_級{1}_學期{2}_科目{3}", domain.Trim().Replace(" ", "_"), gradeCount, 2, courseCount)];
                                if (subjName1 == subjName2 || subjName1 == "" || subjName2 == "")
                                {
                                    mailmerge[string.Format("{0}_級{1}_學期1_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.First, value = subjName1 };
                                    mailmerge[string.Format("{0}_級{1}_學期2_科目{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.Previous, value = subjName1 };
                                }


                                var subjLevel1 = "" + mailmerge[string.Format("{0}_級{1}_學期{2}_科目級別{3}", domain.Trim().Replace(" ", "_"), gradeCount, 1, courseCount)];
                                var subjLevel2 = "" + mailmerge[string.Format("{0}_級{1}_學期{2}_科目級別{3}", domain.Trim().Replace(" ", "_"), gradeCount, 2, courseCount)];
                                if (subjLevel1 == subjLevel2)
                                {
                                    mailmerge[string.Format("{0}_級{1}_學期1_科目級別{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.First, value = subjLevel1 };
                                    mailmerge[string.Format("{0}_級{1}_學期2_科目級別{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.Previous, value = subjLevel1 };
                                }
                                if (subjLevel2 == ""&& subjLevel1 !="")
                                {
                                    mailmerge[string.Format("{0}_級{1}_學期1_科目級別{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.First, value = subjLevel1 };
                                    mailmerge[string.Format("{0}_級{1}_學期2_科目級別{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.Previous, value = "" };
                                }

                                if (subjLevel1 == "" && subjLevel2 !="")
                                {
                                    mailmerge[string.Format("{0}_級{1}_學期1_科目級別{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.First, value = subjLevel2 };
                                    mailmerge[string.Format("{0}_級{1}_學期2_科目級別{2}", domain.Trim().Replace(" ", "_"), gradeCount, courseCount)] = new mailmergeSpecial() { cellmerge = Aspose.Words.Tables.CellMerge.Previous, value ="" };
                                }
                            }
                        }
                        gradeCount++;
                    }


                    GradeCumulateGPARecord gcgpar;
                    mailmerge.Add("級最高GPA", "");
                    mailmerge.Add("級平均GPA", "");
                    if (lastGrade != null && sr.Class != null)
                    {
                        //if (lastGrade.semester != null)
                        //{
                        gcgpar = gcgpa.GetGradeCumulateGPARecord(lastGrade.school_year, lastGrade.semester, lastGrade.grade_year);
                        if (gcgpar != null)
                        {
                            mailmerge["級最高GPA"] = decimal.Round(gcgpar.MaxGPA, 2, MidpointRounding.AwayFromZero);
                            mailmerge["級平均GPA"] = decimal.Round(gcgpar.AvgGPA, 2, MidpointRounding.AwayFromZero);
                        }
                        //}
                    }
                    #endregion

                    // 正式把MailMerge資料 給填上去，2015/4/27 驊紀錄
                    each.MailMerge.FieldMergingCallback = new MailMerge_MergeField();
                    each.MailMerge.Execute(mailmerge.Keys.ToArray(), mailmerge.Values.ToArray());
                    each.MailMerge.DeleteFields();
                    document.Sections.Add(document.ImportNode(each.FirstSection, true));

                    //2016/4/28 以下是恩正給穎驊的程式碼，可以輸出本程式MergeField 所有功能變數的縮寫成一個獨立Doc檔
                    // 在未來如果要要大量更動新增表格很方便可以直接複製貼上使用，如要使用，將上方原本的輸出MergeField、下面的//document.Sections.RemoveAt(0);註解掉即可。

                    //{
                    //    Document doc = new Document();
                    //    DocumentBuilder bu = new DocumentBuilder(doc);
                    //    bu.MoveToDocumentStart();
                    //    bu.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //    bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    //    Table table = bu.StartTable();
                    //    foreach (String col in mailmerge.Keys)
                    //    {
                    //        bu.InsertCell();
                    //        bu.CellFormat.Width = 15;
                    //        bu.InsertField("MERGEFIELD " + col + @" \* MERGEFORMAT", "«.»");
                    //        bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    //        bu.InsertCell();
                    //        bu.CellFormat.Width = 125;
                    //        bu.Write(col);
                    //        bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    //        bu.EndRow();
                    //    }
                    //    table.AllowAutoFit = false;
                    //    bu.EndTable();
                    //    document = doc;
                    //    break;
                    //}

                }
                document.Sections.RemoveAt(0);
                #endregion
            };
            bgw.RunWorkerCompleted += delegate
            {
                if (errCheck.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var stuRec in errCheck.Keys)
                    {
                        foreach (var err in errCheck[stuRec])
                        {
                            sb.AppendLine(string.Format("{0} {1}({2}) {3}:{4}", stuRec.StudentNumber, stuRec.Class != null ? stuRec.Class.Name : "", stuRec.SeatNo, stuRec.Name, err));
                        }
                    }
                    MessageBox.Show(sb.ToString());
                }
                #region Completed
                btnPrint.Enabled = true;
                //if (e.Error != null)
                //{
                //    MessageBox.Show(e.Error.Message);
                //    return;
                //}
                Document inResult = document;
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
                #endregion
            };

            bgw.RunWorkerAsync();
        }
        class edld
        {
            public DateTime? entrance_date, leaving_date;
        }

        class MailMerge_MergeField : Aspose.Words.Reporting.IFieldMergingCallback
        {
            public void FieldMerging(Aspose.Words.Reporting.FieldMergingArgs args)
            {
                if (args.FieldValue != null)
                {
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName, false, false);
                    if (args.FieldValue is mailmergeSpecial)
                    {
                        builder.CellFormat.HorizontalMerge = (args.FieldValue as mailmergeSpecial).cellmerge;
                    }
                    builder.CellFormat.FitText = true;
                    builder.CellFormat.WrapText = true;
                    if (args.FieldValue is mailmergeSpecial)
                    {
                        args.Text = (args.FieldValue as mailmergeSpecial).value;
                    }
                }
            }
            public void ImageFieldMerging(Aspose.Words.Reporting.ImageFieldMergingArgs args)
            {
            }
        }

    }

}
