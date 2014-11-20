using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JH.IBSH.Report.PeriodicalExam
{
    public partial class Config : BaseForm
    {
        public Config()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration3_6.Template == null)
            {
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration3_6.Template = new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration3_6.Template, new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "3~6 grade樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration3_6.Template = TemplateForm.Template;
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration3_6.Save();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration7_8.Template == null)
            {
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration7_8.Template = new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration7_8.Template, new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "7~8 grade樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration7_8.Template = TemplateForm.Template;
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration7_8.Save();
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration9_12.Template == null)
            {
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration9_12.Template = new Campus.Report.ReportTemplate(Properties.Resources._9樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration9_12.Template, new Campus.Report.ReportTemplate(Properties.Resources._9樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "9~12 grade樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration9_12.Template = TemplateForm.Template;
                JH.IBSH.Report.PeriodicalExam.MainForm.ReportConfiguration9_12.Save();
            }
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel ll = (LinkLabel)sender;
            try
            {
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = ll.Text.Replace("檢視", "");
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Aspose.Words.Document doc2 = new Aspose.Words.Document(new MemoryStream(Properties.Resources.合併欄位總表));
                    doc2.Save(SaveFileDialog1.FileName);
                    System.Diagnostics.Process.Start(SaveFileDialog1.FileName);
                }
                else
                    FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
            }
        }
    }
}
