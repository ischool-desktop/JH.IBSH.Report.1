using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;

namespace JH.IBSH.Report.Foreign
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            MenuButton item = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"];
            item["國外成績單"].Enable = Permissions.實中國外成績單權限;
            item["國外成績單"].Click += delegate
            {
                new MainForm().ShowDialog();
            };
            Catalog detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.實中國外成績單, "國外成績單"));
        }
    }
}
