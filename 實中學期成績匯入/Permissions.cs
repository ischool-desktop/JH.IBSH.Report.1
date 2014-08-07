using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JH.IBSH.Import.SemesterScore
{
    class Permissions
    {
        public static string 實中匯入學期成績 { get { return "JH.IBSH.Import.SemesterScore.cs"; } }
        public static bool 實中匯入學期成績權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[實中匯入學期成績].Executable;
            }
        }
    }
}
