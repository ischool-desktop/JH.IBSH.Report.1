using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using System.Data;

namespace JH.IBSH.Import.SemesterScore
{
    
    class FieldValidatorFactory : IFieldValidatorFactory
    {
        private Dictionary<string, List<string>> _studentdic;
        private Dictionary<string, List<string>> GetStudent()
        {
            if (_studentdic != null)
                return _studentdic;

            FISCA.Data.QueryHelper _queryHelper = new FISCA.Data.QueryHelper();
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            //取得比對序
            DataTable dt = _queryHelper.Select("select id,student_number from student");
            foreach (DataRow row in dt.Rows)
            {
                string StudentID = "" + row[0];
                string Student_Number = "" + row[1];

                if (string.IsNullOrEmpty(Student_Number))
                    continue;

                if (!dic.ContainsKey(Student_Number))
                {
                    dic.Add(Student_Number, new List<string>());
                }
                dic[Student_Number].Add(StudentID);
            }
            _studentdic = dic;
            return _studentdic;
        }
        private List<string> GetStudentID()
        {
            FISCA.Data.QueryHelper _queryHelper = new FISCA.Data.QueryHelper();
            List<string> l = new List<string>();
            //取得比對序
            DataTable dt = _queryHelper.Select("select id from student");
            foreach (DataRow row in dt.Rows)
            {
                if (string.IsNullOrEmpty("" + row[0]))
                    continue;
                l.Add("" + row[0]);
            }
            return l;
        }
        static public Dictionary<string,string> GetSubjectDomain()
        {
            FISCA.Data.QueryHelper _queryHelper = new FISCA.Data.QueryHelper();
            Dictionary<string, string > dic = new Dictionary<string, string>();
            //取得比對序
            DataTable dt = _queryHelper.Select("select name,$ischool.subject.list.group from $ischool.subject.list");
            foreach (DataRow row in dt.Rows)
            {
                string SubjectName = "" + row[0];
                string SubjectDomain = "" + row[1];

                if (string.IsNullOrEmpty(SubjectName))
                    continue;

                if (!dic.ContainsKey(SubjectName))
                {
                    dic.Add(SubjectName, SubjectDomain);
                }
            }
            return dic;
        }
        #region IFieldValidatorFactory 成員
        public IFieldValidator CreateFieldValidator(string typeName, System.Xml.XmlElement validatorDescription)
        {
            //Campus.DocumentValidator.ValidatorType.
            switch (typeName.ToUpper())
            {
                case "STUDENTNUMBEREXISTENCE":
                    return new StudentNumberExistenceValidator(this.GetStudent());
                case "STUDENTNUMBERREPEAT":
                    return new StudentNumberRepeatValidator(this.GetStudent());
                case "ISDECIMAL2":
                    return new IsDecimal2();
                case "ISDECIMAL5":
                    return new IsDecimal5();
                case "SUBJECTDOMAINEXISTENCE":
                    return new SubjectDomainExistence(GetSubjectDomain());
                case "IDEXISTENCE":
                    return new IdExistence(this.GetStudentID());
                default:
                    return null;
            }
        }
        #endregion
    }
    class SubjectDomainExistence : IFieldValidator
    {
        private Dictionary<string, string> ddomain;
        public SubjectDomainExistence(Dictionary<string,string> ddomain)
        {
            this.ddomain = ddomain;
        }
        public string Correct(string Value)
        {
            return Value;
        }
        public string ToString(string template)
        {
            return template;
        }
        public bool Validate(string Value)
        {
            return ddomain.ContainsKey(Value);
        }
    }
    class IdExistence : IFieldValidator
    {
        List<string> StudentIDs;
        public IdExistence(List<string> sidl)
        {
            this.StudentIDs = sidl;
        }
        public string Correct(string Value)
        {
            return Value;
        }
        public string ToString(string template)
        {
            return template;
        }
        public bool Validate(string Value)
        {
            return StudentIDs.Contains(Value);
        }
    }

    class IsDecimal2 : IFieldValidator
    {
        public string Correct(string Value)
        {
            return Value;
        }
        public string ToString(string template)
        {
            return template;
        }
        public bool Validate(string Value)
        {
            decimal tmp;
            if (decimal.TryParse(Value, out tmp))
            {
                decimal d = tmp * 100;
                return Math.Floor(d) == Math.Ceiling(d);
            }
            else
                return false;
        }
    }
    class IsDecimal5 : IFieldValidator
    {
        public string Correct(string Value)
        {
            return Value;
        }
        public string ToString(string template)
        {
            return template;
        }
        public bool Validate(string Value)
        {
            decimal tmp ;
            if ( decimal.TryParse(Value,out tmp) )
            {
                decimal d = tmp * 100000;
                return Math.Floor(d) == Math.Ceiling(d);
            }
            else 
                return false ;
        }
    }
    class StudentNumberExistenceValidator : IFieldValidator
    {
        Dictionary<string, List<string>> _StudentDic { get; set; }
        public StudentNumberExistenceValidator(Dictionary<string, List<string>> StudentDic)
        {
            _StudentDic = StudentDic;
        }
        #region IFieldValidator 成員

        //自動修正
        public string Correct(string Value)
        {
            return Value;
        }
        //回傳驗證訊息
        public string ToString(string template)
        {
            return template;
        }
        public bool Validate(string Value)
        {
            if (_StudentDic.ContainsKey(Value)) //包含此學號
            {
                return true;
            }
            else //不包含此學號
            {
                return false;
            }
        }
        /// <summary>
        /// 取得學生學號 vs 系統編號
        /// </summary>
        private Dictionary<string, List<string>> GetStudent()
        {
            FISCA.Data.QueryHelper _queryHelper = new FISCA.Data.QueryHelper();

            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            //取得比對序

            DataTable dt = _queryHelper.Select("select id,student_number from student");
            foreach (DataRow row in dt.Rows)
            {
                string StudentID = "" + row[0];
                string Student_Number = "" + row[1];

                if (string.IsNullOrEmpty(Student_Number))
                    continue;

                if (!dic.ContainsKey(Student_Number))
                {
                    dic.Add(Student_Number, new List<string>());
                }
                dic[Student_Number].Add(StudentID); //如果重覆也加進去
            }
            return dic;

        }

        #endregion
    }
    class StudentNumberRepeatValidator : IFieldValidator
    {
        Dictionary<string, List<string>> _StudentDic { get; set; }
        public StudentNumberRepeatValidator(Dictionary<string, List<string>> StudentDic)
        {
            _StudentDic = StudentDic;
        }

        #region IFieldValidator 成員

        public string Correct(string Value)
        {
            return Value;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(string Value)
        {
            if (_StudentDic.ContainsKey(Value)) //包含此學號
            {
                if (_StudentDic[Value].Count > 1) //多名學生為錯誤
                {
                    return false;
                }
            }
            return true;
        }
        #endregion
    }
}
