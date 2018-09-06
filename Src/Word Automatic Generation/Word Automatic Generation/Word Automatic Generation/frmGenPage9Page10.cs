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
            lvData.Columns.Add("Status", 100);

            lvData.Items.Clear();

            ListViewItem lvi;
            for (int i = 1; i <= 10; i++)
            {
                lvi = new ListViewItem(new string[] { i.ToString(), "ชั้นประถมศึกษา 1A", "Year 1A", "Mr. Name " + i.ToString("00000") });
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
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Browse directory"); }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                string Page9File = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Template\Page9.doc");
                string Page10File = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Template\Page10.doc");
                string Dir = txtDir.Text;                

                
                foreach (ListViewItem lvi in lvData.Items)
                {
                    
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Generate Word"); }
        }
    }
}
