using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Word_Automatic_Generation
{
    public partial class frmGenPage9Page10 : Form
    {
        public frmGenPage9Page10()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            SampleData();
        }

        void SampleData()
        {
            lvData.Columns.Clear();
            lvData.Columns.Add("No", 50);
            lvData.Columns.Add("Class TH", 200);
            lvData.Columns.Add("Class EN", 100);
            lvData.Columns.Add("Student", 300);
            lvData.Columns.Add("Page 9", 300);
            lvData.Columns.Add("Page 10", 300);
            lvData.Columns.Add("Spend Time", 300);

            lvData.Items.Clear();

            ListViewItem lvi;
            for (int i = 1; i <= 10; i++)
            {
                lvi = new ListViewItem(new string[] { i.ToString(), "ชั้นประถมศึกษา 1A", "Year 1A", "Mr. Name " + i.ToString("00000"), "", "", "" });
                lvData.Items.Add(lvi);
            }

        }

        private void btnDirBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                using (FolderBrowserDialog dlg = new FolderBrowserDialog())
                {
                    dlg.ShowDialog();
                    txtDir.Text = dlg.SelectedPath;
                    btnGenerate.Enabled = true;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Browse directory"); }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                string Page9TemplateFile = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Template\Page9.doc");
                string Page10TemplateFile = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Template\Page10.doc");
                string Dir = txtDir.Text;

                int cnt = 0;
                foreach (ListViewItem lvi in lvData.Items)
                {
                    if (lvi.Checked)
                    {
                        string No = lvi.Text;
                        string ClassTh = lvi.SubItems[1].Text;
                        string ClassEn = lvi.SubItems[2].Text;
                        string Student = lvi.SubItems[3].Text;

                        string DestPage9FullFile = Dir + @"\" + Student + @" Page9.doc";
                        string DestPage9File = Student + @" Page9.doc";
                        string DestPage10FullFile = Dir + @"\" + Student + @" Page10.doc";
                        string DestPage10File = Student + @" Page10.doc";

                        Doc_Template_Management.DocTemplateManagement dtm = new Doc_Template_Management.DocTemplateManagement();

                        dtm.SetParameter("<CLASS_TH>", ClassTh);
                        dtm.SetParameter("<CLASS_EN>", ClassEn);
                        dtm.SetParameter("<STUDENT>", Student);
                        dtm.SetParameter("<NO>", No);

                        dtm.CreateWordDocument(Page9TemplateFile, DestPage9FullFile);
                        dtm.CreateWordDocument(Page10TemplateFile, DestPage10FullFile);

                        lvi.SubItems[4].Text = DestPage9File;
                        lvi.SubItems[5].Text = DestPage10File;
                        lvData.Invalidate();
                    }
                    else
                    {
                        cnt++;
                    }
                }
                if (cnt == lvData.Items.Count) MessageBox.Show("There is no selected student.", "Generate Word");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Generate Word"); }
        }
    }
}
