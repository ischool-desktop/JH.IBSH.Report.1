using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;

namespace JH.IBSH.Report.PeriodicalExam
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            MenuButton item = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"];
            item["評量成績通知單"].Enable = Permissions.實中評量成績通知單權限;
            item["評量成績通知單"].Click += delegate
            {
                new MainForm().ShowDialog();
            };
            Catalog detail1 = RoleAclSource.Instance["學生"]["報表"];
            detail1.Add(new RibbonFeature(Permissions.實中評量成績通知單, "評量成績通知單"));
        }
    }
}
