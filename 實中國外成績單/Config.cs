using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JH.IBSH.Report.Foreign
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
            if (JH.IBSH.Report.Foreign.MainForm.ReportConfiguration3_6.Template == null)
            {
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration3_6.Template = new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.Report.Foreign.MainForm.ReportConfiguration3_6.Template, new Campus.Report.ReportTemplate(Properties.Resources._6樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "3~6 grade樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration3_6.Template = TemplateForm.Template;
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration3_6.Save();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.Report.Foreign.MainForm.ReportConfiguration7_8.Template == null)
            {
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration7_8.Template = new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.Report.Foreign.MainForm.ReportConfiguration7_8.Template, new Campus.Report.ReportTemplate(Properties.Resources._8樣版, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "7~8 grade樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration7_8.Template = TemplateForm.Template;
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration7_8.Save();
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Campus.Report.TemplateSettingForm TemplateForm;
            if (JH.IBSH.Report.Foreign.MainForm.ReportConfiguration9_12.Template == null)
            {
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration9_12.Template = new Campus.Report.ReportTemplate(Properties.Resources._9_12_grade樣板, Campus.Report.TemplateType.Word);
            }
            TemplateForm = new Campus.Report.TemplateSettingForm(JH.IBSH.Report.Foreign.MainForm.ReportConfiguration9_12.Template, new Campus.Report.ReportTemplate(Properties.Resources._9_12_grade樣板, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "9~12 grade樣板";
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration9_12.Template = TemplateForm.Template;
                JH.IBSH.Report.Foreign.MainForm.ReportConfiguration9_12.Save();
            }
        }
    }
}
