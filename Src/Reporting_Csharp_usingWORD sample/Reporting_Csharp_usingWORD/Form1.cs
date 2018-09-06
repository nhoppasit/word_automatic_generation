using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

/////Add the references (new)
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Drawing.Drawing2D;
////

namespace Reporting_Csharp_usingWORD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Methode Find and Replace:
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText,
                        ref matchCase, ref matchWholeWord,
                        ref matchWildCards, ref matchSoundLike,
                        ref nmatchAllForms, ref forward,
                        ref wrap, ref format, ref replaceWithText,
                        ref replace, ref matchKashida,
                        ref matchDiactitics, ref matchAlefHamza,
                        ref matchControl);
        }



        string pathImage = null;
        //Methode Create the document :
        private void CreateWordDocument(object filename, object savaAs , object image)
        {
            List<int> processesbeforegen = getRunningProcesses();
            object missing = Missing.Value;
            string tempPath=null;

            Word.Application wordApp = new Word.Application();

            Word.Document aDoc = null;

            if (File.Exists((string)filename))
            {
                DateTime today = DateTime.Now;

                object readOnly = false; //default
                object isVisible = false;

                wordApp.Visible = false;

                aDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);

                aDoc.Activate();

                //Find and replace:
                this.FindAndReplace(wordApp, "$$name$$", tFirstname.Text);
                this.FindAndReplace(wordApp, "$$Lastname$$", tLastname.Text);
                this.FindAndReplace(wordApp, "$$tel$$", tPhone.Text);
                this.FindAndReplace(wordApp, "$$Company$$", tCompany.Text);
                this.FindAndReplace(wordApp, "$$Date$$", DateTime.Now.ToShortDateString());

                //insert the picture:
                Image img = resizeImage(pathImage, new Size(200, 90));
                tempPath = System.Windows.Forms.Application.StartupPath + "\\Images\\~Temp\\temp.jpg";
                img.Save(tempPath);

                Object oMissed = aDoc.Paragraphs[1].Range; //the position you want to insert
                Object oLinkToFile = false;  //default
                Object oSaveWithDocument = true;//default
                aDoc.InlineShapes.AddPicture(tempPath, ref  oLinkToFile, ref  oSaveWithDocument, ref oMissed);

                #region Print Document :
                /*object copies = "1";
                object pages = "1";
                object range = Word.WdPrintOutRange.wdPrintCurrentPage;
                object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                object oTrue = true;
                object oFalse = false;

                Word.Document document = aDoc;
                object nullobj = Missing.Value;
                int dialogResult = wordApp.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show(ref nullobj);
                wordApp.Visible = false;
                if (dialogResult == 1)
                {
                    document.PrintOut(
                    ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,
                    ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                    ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);
                }
                */
                #endregion

            }
            else
            {
                MessageBox.Show("file dose not exist.");
                return;
            }

            //Save as: filename
            aDoc.SaveAs2(ref savaAs, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            //Close Document:
            //aDoc.Close(ref missing, ref missing, ref missing);
            File.Delete(tempPath);
            MessageBox.Show("File created.");
            List<int> processesaftergen = getRunningProcesses();
            killProcesses(processesbeforegen, processesaftergen);
        }


        public List<int> getRunningProcesses()
        {
            List<int> ProcessIDs = new List<int>();
            //here we're going to get a list of all running processes on
            //the computer
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("WINWORD"))
                {
                    ProcessIDs.Add(clsProcess.Id);
                }
            }
            return ProcessIDs;
        }


        private void killProcesses(List<int> processesbeforegen, List<int> processesaftergen)
        {
            foreach (int pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (int pidbefore in processesbeforegen)
                {
                    if (pidafter == pidbefore)
                    {
                        processfound = true;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter);
                    clsProcess.Kill();
                }
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        //Méthode Enabled Controles:
        private void tEnabled(bool state)
        {
            tCompany.Enabled = state;
            tFirstname.Enabled = state;
            tPhone.Enabled = state;
            tLastname.Enabled = state;
            btnLogo.Enabled = state;
        }

        //Load the Template Document:
        private void button1_Click(object sender, EventArgs e)
        {
            if (LoadDoc.ShowDialog() == DialogResult.OK)
            {
                tFilename.Text = LoadDoc.FileName;
                tEnabled(true);
            }
        }

        //Load the logo picture :
        private void btnLogo_Click(object sender, EventArgs e)
        {
            if (LoadPic.ShowDialog() == DialogResult.OK)
            {
                pathImage = LoadPic.FileName;
                btnSubmit.Enabled = true;
            }
        }

        //Create your document:
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (SaveDoc.ShowDialog() == DialogResult.OK)
            {
                CreateWordDocument(tFilename.Text, SaveDoc.FileName, pathImage);
                tEnabled(false);
                //printDocument1.DocumentName = SaveDoc.FileName;
            }
        }

        //Change Picture Size :
        private static Image resizeImage(string filename, Size size)
        {
            Image imgToResize = Image.FromFile(filename);
            int sourceWidth = imgToResize.Width;
            int sourceHeight = imgToResize.Height;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)size.Width / (float)sourceWidth);
            nPercentH = ((float)size.Height / (float)sourceHeight);

            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();

            return (Image)b;
        }

        #region droped
        private void button2_Click(object sender, EventArgs e)
        {
           //pictureBox1 .Image = resizeImage(pathImage, new Size(200, 90));
           // pictureBox1.Image.Save(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\test.jpg");
        }
        #endregion
    }
}
