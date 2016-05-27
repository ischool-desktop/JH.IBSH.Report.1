using Campus.DocumentValidator;
using FISCA;
using FISCA.Presentation;
using K12.Data;
using K12.Data.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using FISCA.UDT;
namespace 實中學期成績匯入_含領域
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            FactoryProvider.FieldFactory.Add(new FieldValidatorFactory());
            
            MenuButton rbItemImport = MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"]["成績相關匯入(穎驊版)"];
            rbItemImport["匯入學期成績(穎驊版)"].Click += delegate
            {
                UpdateHelper uh = new UpdateHelper();
                //uh.Execute("TRUNCATE TABLE sems_subj_score");
                ImportSemesterScore2 wizard;
                try
                {
                    wizard = new ImportSemesterScore2();
                    wizard.IsSplit = false;
                    wizard.Execute();
                }
                catch (System.Exception)
                {
                    throw;
                }
                
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
