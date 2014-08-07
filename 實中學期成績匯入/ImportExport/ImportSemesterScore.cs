using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.Import;
using System.Xml;
using K12.Data;
using Campus.DocumentValidator;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using System.Data;
using System.Xml.Linq;

namespace JH.IBSH.Import.SemesterScore
{
    class ImportSemesterScore : ImportWizard
    {
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
        //設定檔
        private ImportOption myImportOption;
        private ImportKey useImportKey = ImportKey.ID;
        public Dictionary<SchoolYearSemester, List<string>> dChangedStudentIDs = new Dictionary<SchoolYearSemester, List<string>>();
        //student number,id
        Dictionary<string, string> dsn = new Dictionary<string, string>();
        enum ImportKey
        {
            ID, StudentNumber
        }
        /// <summary>
        /// 準備動作
        /// </summary>
        public override void Prepare(ImportOption Option)
        {
            myImportOption = Option;
        }
        /// <summary>
        /// 每1000筆資料,分批執行匯入
        /// Return是Log資訊
        /// </summary>
        public override string Import(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            //取得學生學號對比系統編號
            if (useImportKey == ImportKey.StudentNumber)
                dsn = GetStudent();

            Dictionary<string, List<SubjectScore>> dss = new Dictionary<string, List<SubjectScore>>();
            //group ref_student_id , school_year , semseter
            foreach (IRowStream each in Rows)
            {
                string id;
                switch (useImportKey)
                {
                    case ImportKey.ID:
                        id = each.GetValue("學生系統編號");
                        break;
                    case ImportKey.StudentNumber:
                        id = dsn[each.GetValue("學號")];
                        break;
                    default:
                        //沒有主key
                        return string.Empty;
                }
                string key = id + "#" + each.GetValue("學年度") + "#" + each.GetValue("學期");
                SchoolYearSemester tmpsys = new SchoolYearSemester()
                    {
                        SchoolYear = int.Parse(each.GetValue("學年度")),
                        Semester = int.Parse(each.GetValue("學期"))
                    };
                if (!dChangedStudentIDs.ContainsKey(tmpsys))
                    dChangedStudentIDs.Add(tmpsys, new List<string>());
                if (!dChangedStudentIDs[tmpsys].Contains(id))
                    dChangedStudentIDs[tmpsys].Add(id);

                if (!dss.ContainsKey(key))
                    dss.Add(key, new List<SubjectScore>());
                Dictionary<string, string> dsd = FieldValidatorFactory.GetSubjectDomain();
                dss[key].Add(new SubjectScore()
                {
                    Subject = each.GetValue("科目"),
                    Credit = decimal.Parse(each.GetValue("權數")),
                    Period = decimal.Parse(each.GetValue("節數")),
                    Score = decimal.Parse(each.GetValue("成績")),
                    GPA = decimal.Parse(each.GetValue("GPA")),
                    Level = int.Parse(each.GetValue("Level")),
                    Comment = "",
                    Domain = dsd[each.GetValue("科目")],
                    Effort = 0,
                    Text = ""
                });
            }
            try
            {
                foreach (KeyValuePair<SchoolYearSemester, List<string>> item in dChangedStudentIDs)
                {
                    Dictionary<string, SemesterScoreRecord> dssrUpdate = K12.Data.SemesterScore.SelectBySchoolYearAndSemester(item.Value, item.Key.SchoolYear, item.Key.Semester).ToDictionary(x => x.RefStudentID, x => x);
                    Dictionary<string, SemesterScoreRecord> dssrInsert = new Dictionary<string, SemesterScoreRecord>();
                    foreach (string id in item.Value.Distinct())
                    {
                        SemesterScoreRecord ssr;
                        if (!dssrUpdate.ContainsKey(id))
                        {
                            dssrInsert.Add(id, new SemesterScoreRecord(id, item.Key.SchoolYear, item.Key.Semester));
                            ssr = dssrInsert[id];
                        }
                        else
                            ssr = dssrUpdate[id];
                        foreach (SubjectScore ss in dss[id + "#" + item.Key.SchoolYear + "#" + item.Key.Semester])
                        {
                            if (ssr.Subjects.ContainsKey(ss.Subject))
                                ssr.Subjects[ss.Subject] = ss;
                            else
                                ssr.Subjects.Add(ss.Subject, ss);
                        }
                    }
                    //update semester score
                    if (dssrInsert.Count > 0)
                        K12.Data.SemesterScore.Insert(dssrInsert.Values);
                    if (dssrUpdate.Count > 0)
                        K12.Data.SemesterScore.Update(dssrUpdate.Values);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            //ApplicationLog.Log("匯入服務學習記錄(新增)", "匯入", "已匯入新增服務學習時數\n共" + InsertList.Count + "筆");
            return string.Empty;
        }

        /// <summary>
        /// 取得學生學號 vs 系統編號
        /// </summary>
        /// <param name="Rows"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetStudent()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            //取得比對序

            DataTable dt = tool._Q.Select("select id,student_number from student");
            foreach (DataRow row in dt.Rows)
            {
                string StudentID = "" + row[0];
                string Student_Number = "" + row[1];

                if (string.IsNullOrEmpty(Student_Number))
                    continue;

                if (!dic.ContainsKey(Student_Number))
                {
                    dic.Add(Student_Number, StudentID);
                }
            }
            return dic;
        }
        /// <summary>
        /// 取得驗證規則(動態建置XML內容)
        /// </summary>
        public override string GetValidateRule()
        {
            if (this.ImportOption.SelectedKeyFields.Count > 0)  //第二次會被呼叫
            {
                string Key = this.ImportOption.SelectedKeyFields[0];
                if (Key.Equals("學生系統編號"))
                {
                    this.useImportKey = ImportKey.ID;
                    return Properties.Resources.ImportSemesterScore_ID;
                }
                else if (Key.Equals("學號"))
                {
                    this.useImportKey = ImportKey.StudentNumber;
                    return Properties.Resources.ImportSemesterScore_StudentNumber;
                }
            }
            return Properties.Resources.ImportSemesterScore;
        }

        /// <summary>
        /// 設定匯入功能,所提供的匯入動作
        /// </summary>
        public override ImportAction GetSupportActions()
        {
            //新增或更新
            return ImportAction.InsertOrUpdate;
        }
    }
}
