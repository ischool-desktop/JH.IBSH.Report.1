using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JH.IBSH.Report.PeriodicalExam
{
    class Permissions
    {
        public static string 實中評量成績通知單 { get { return "JH.IBSH.Report.PeriodicalExam.cs"; } }
        public static bool 實中評量成績通知單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[實中評量成績通知單].Executable;
            }
        }
    }
}
