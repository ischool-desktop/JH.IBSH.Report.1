using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using K12.Data;
namespace JH.IBSH.Report.Foreign
{
    
    public class GradeCumulateGPARecord 
    {
        public GradeCumulateGPARecord(int SchoolYear, int Semester, int Grade, decimal MaxGPA, decimal AvgGPA)
        {
            this.SchoolYear = SchoolYear;
            this.Semester = Semester;
            this.Grade = Grade;
            this.MaxGPA = MaxGPA;
            this.AvgGPA = AvgGPA;
        }
        public int SchoolYear { get; private set; }
        public int Semester { get; private set; }
        public int Grade { get; private set; }
        public decimal MaxGPA { get; private set; }
        public decimal AvgGPA { get; private set; }
        
    }
    public class GradeCumulateGPA
    {
        private GradeCumulateGPA() { }
        private static Dictionary<string, GradeCumulateGPARecord> dgcgpar = new Dictionary<string, GradeCumulateGPARecord>();
        public static GradeCumulateGPARecord GetGradeCumulateGPARecord(int SchoolYear, int Semester, int Grade)
        {
            //logic : 同class.grade_year 且 學期成績為當學年度學期 之 學期成績資料
            string key = SchoolYear + "#" + Semester + "#" + Grade ;
            if (dgcgpar.ContainsKey(key))
            {
                return dgcgpar[key];
            }
            DataTable dt = tool._Q.Select(@"select xpath_string( '<root>'||score_info||'</root>','/root/CumulateGPA') as CumulateGPA,
student.id
from class
join student  on class.id=student.ref_class_id
join sems_subj_score on student.id=sems_subj_score.ref_student_id
where class.grade_year =" + Grade + " and sems_subj_score.school_year=" + SchoolYear + " and sems_subj_score.semester=" + Semester);
            decimal MaxGPA = 0,AvgGPA=0 ,TotalGPA =0;
            int validCount = 0 ;
            foreach (DataRow row in dt.Rows)
            {
                //row["id"]
                decimal CumulateGPA ;
                if ( decimal.TryParse(""+row["CumulateGPA"],out CumulateGPA))
                {
                    if ( MaxGPA < CumulateGPA )
                        MaxGPA = CumulateGPA ;
                    TotalGPA += CumulateGPA ;
                    validCount++;
                }
            }
            AvgGPA = TotalGPA / validCount;
            dgcgpar[key] = new GradeCumulateGPARecord(SchoolYear, Semester, Grade,MaxGPA,AvgGPA);
            return dgcgpar[key];
        }
    }
}
