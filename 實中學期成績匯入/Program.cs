using Campus.DocumentValidator;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.UDT;
using K12.Data;
using K12.Data.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;

namespace JH.IBSH.Import.SemesterScore
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            FactoryProvider.FieldFactory.Add(new FieldValidatorFactory());

            MenuButton rbItemImport = MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"]["學務相關匯入"];
            rbItemImport["匯入學期成績"].Enable = Permissions.實中匯入學期成績權限;
            rbItemImport["匯入學期成績"].Click += delegate
            {
                
                ImportSemesterScore wizard = new ImportSemesterScore();
                wizard.Execute();
                if (wizard.dChangedStudentIDs.Count > 0)
                {
                    BackgroundWorker bgw = new BackgroundWorker();
                    bgw.WorkerReportsProgress = true;
                    bgw.DoWork += delegate
                    {
                        #region DoWork
                        //調整有上傳學期成績學生
                        bgw.ReportProgress(1);
                        List<string> tmp_sid = new List<string>();
                        foreach (KeyValuePair<SchoolYearSemester, List<string>> item in wizard.dChangedStudentIDs)
                        {
                            CourseGradeB.Tool.SetAverage(item.Value, item.Key.SchoolYear, item.Key.Semester);
                            tmp_sid.AddRange(item.Value);
                        }
                        bgw.ReportProgress(2);
                        List<SemesterHistoryRecord> dshr = K12.Data.SemesterHistory.SelectByStudentIDs(tmp_sid);
                        Dictionary<SchoolYearSemester, List<string>> dStudentSchoolYearSemester = new Dictionary<SchoolYearSemester, List<string>>();
                        foreach (SemesterHistoryRecord shr in dshr)
                        {
                            foreach (SemesterHistoryItem item in shr.SemesterHistoryItems)
                            {
                                SchoolYearSemester tmpsys = new SchoolYearSemester()
                                {
                                    SchoolYear = item.SchoolYear,
                                    Semester = item.Semester
                                };
                                if (!dStudentSchoolYearSemester.ContainsKey(tmpsys))
                                    dStudentSchoolYearSemester.Add(tmpsys, new List<string>());
                                if (!dStudentSchoolYearSemester[tmpsys].Contains(item.RefStudentID))
                                    dStudentSchoolYearSemester[tmpsys].Add(item.RefStudentID);
                            }
                        }
                        bgw.ReportProgress(3);
                        foreach (KeyValuePair<SchoolYearSemester, List<string>> item in dStudentSchoolYearSemester)
                        {
                            CourseGradeB.Tool.SetCumulateGPA(item.Value, item.Key.SchoolYear, item.Key.Semester);
                        }
                        bgw.ReportProgress(100);
                        #endregion
                    };
                    bgw.ProgressChanged += delegate(object sender, System.ComponentModel.ProgressChangedEventArgs e) {
                        string message = e.ProgressPercentage == 100 ? "計算完成" : "計算中...";
                        FISCA.Presentation.MotherForm.SetStatusBarMessage("匯入學期成績 平均與累計GPA " + message, e.ProgressPercentage);
                    };
                    bgw.RunWorkerAsync();
                }
            };
            Catalog detail1 = RoleAclSource.Instance["學生"]["匯出/匯入"];
            detail1.Add(new RibbonFeature(Permissions.實中匯入學期成績, "匯入學期成績"));
        }
        class mySchoolYearSemesterEqualityComparer : IEqualityComparer<SchoolYearSemester>
        {
            public bool Equals(SchoolYearSemester x, SchoolYearSemester y)
            {
                return x.SchoolYear == y.SchoolYear && x.Semester == y.Semester;
            }
            public int GetHashCode(SchoolYearSemester obj)
            {
                return obj.GetHashCode();
            }
        }
    }
}
