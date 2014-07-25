using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JH.IBSH.Report.Foreign
{
    class Permissions
    {
        public static string 實中國外成績單 { get { return "JH.IBSH.Report.Foreign.cs"; } }
        public static bool 實中國外成績單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[實中國外成績單].Executable;
            }
        }
    }
}
