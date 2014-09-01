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
            string key = SchoolYear + "#" + Semester + "#" + Grade;
            if (dgcgpar.ContainsKey(key))
            {
                return dgcgpar[key];
            }
            DataTable dt = tool._Q.Select(@"select * from $ischool.gparef where grade =" + Grade + " and school_year=" + SchoolYear + " and semester=" + Semester);
            decimal MaxGPA, AvgGPA;
            if (dt.Rows.Count > 0 && decimal.TryParse("" + dt.Rows[0]["max"], out MaxGPA) && decimal.TryParse("" + dt.Rows[0]["avg"], out AvgGPA))
            {
                dgcgpar[key] = new GradeCumulateGPARecord(SchoolYear, Semester, Grade, MaxGPA, AvgGPA);
                return dgcgpar[key];
            }
            else
                return null;
        }
    }
}
