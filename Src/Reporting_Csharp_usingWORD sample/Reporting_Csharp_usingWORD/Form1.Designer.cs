namespace Reporting_Csharp_usingWORD
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.LoadDoc = new System.Windows.Forms.OpenFileDialog();
            this.SaveDoc = new System.Windows.Forms.SaveFileDialog();
            this.tFilename = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tFirstname = new System.Windows.Forms.TextBox();
            this.tLastname = new System.Windows.Forms.TextBox();
            this.tPhone = new System.Windows.Forms.TextBox();
            this.tCompany = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnLogo = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.LoadPic = new System.Windows.Forms.OpenFileDialog();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.SuspendLayout();
            // 
            // LoadDoc
            // 
            this.LoadDoc.FileName = "openFileDialog1";
            // 
            // tFilename
            // 
            this.tFilename.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tFilename.Location = new System.Drawing.Point(7, 47);
            this.tFilename.Name = "tFilename";
            this.tFilename.ReadOnly = true;
            this.tFilename.Size = new System.Drawing.Size(435, 27);
            this.tFilename.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(111, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(227, 28);
            this.button1.TabIndex = 1;
            this.button1.Text = "Load Word document template";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tFirstname
            // 
            this.tFirstname.Enabled = false;
            this.tFirstname.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tFirstname.Location = new System.Drawing.Point(129, 86);
            this.tFirstname.Name = "tFirstname";
            this.tFirstname.Size = new System.Drawing.Size(209, 27);
            this.tFirstname.TabIndex = 2;
            // 
            // tLastname
            // 
            this.tLastname.Enabled = false;
            this.tLastname.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tLastname.Location = new System.Drawing.Point(129, 119);
            this.tLastname.Name = "tLastname";
            this.tLastname.Size = new System.Drawing.Size(209, 27);
            this.tLastname.TabIndex = 3;
            // 
            // tPhone
            // 
            this.tPhone.Enabled = false;
            this.tPhone.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tPhone.Location = new System.Drawing.Point(129, 152);
            this.tPhone.Name = "tPhone";
            this.tPhone.Size = new System.Drawing.Size(145, 27);
            this.tPhone.TabIndex = 4;
            // 
            // tCompany
            // 
            this.tCompany.Enabled = false;
            this.tCompany.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tCompany.Location = new System.Drawing.Point(129, 185);
            this.tCompany.Name = "tCompany";
            this.tCompany.Size = new System.Drawing.Size(257, 27);
            this.tCompany.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(38, 89);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 19);
            this.label1.TabIndex = 7;
            this.label1.Text = "First name :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(39, 122);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 19);
            this.label2.TabIndex = 8;
            this.label2.Text = "Last name :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(13, 155);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 19);
            this.label3.TabIndex = 9;
            this.label3.Text = "Phone number :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(46, 188);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 19);
            this.label4.TabIndex = 10;
            this.label4.Text = "Company :";
            // 
            // btnLogo
            // 
            this.btnLogo.Enabled = false;
            this.btnLogo.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLogo.Location = new System.Drawing.Point(129, 218);
            this.btnLogo.Name = "btnLogo";
            this.btnLogo.Size = new System.Drawing.Size(75, 28);
            this.btnLogo.TabIndex = 11;
            this.btnLogo.Text = "upload...";
            this.btnLogo.UseVisualStyleBackColor = true;
            this.btnLogo.Click += new System.EventHandler(this.btnLogo_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(75, 222);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 19);
            this.label5.TabIndex = 12;
            this.label5.Text = "Logo :";
            // 
            // LoadPic
            // 
            this.LoadPic.FileName = "openFileDialog2";
            // 
            // btnSubmit
            // 
            this.btnSubmit.BackColor = System.Drawing.Color.Transparent;
            this.btnSubmit.Enabled = false;
            this.btnSubmit.FlatAppearance.BorderColor = System.Drawing.SystemColors.Highlight;
            this.btnSubmit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.AntiqueWhite;
            this.btnSubmit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.WhiteSmoke;
            this.btnSubmit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSubmit.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubmit.Location = new System.Drawing.Point(54, 254);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(340, 26);
            this.btnSubmit.TabIndex = 13;
            this.btnSubmit.Text = "Submit";
            this.btnSubmit.UseVisualStyleBackColor = false;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Document = this.printDocument1;
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.BackgroundImage = global::Reporting_Csharp_usingWORD.Properties.Resources.BKL;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(448, 286);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnLogo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tCompany);
            this.Controls.Add(this.tPhone);
            this.Controls.Add(this.tLastname);
            this.Controls.Add(this.tFirstname);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tFilename);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Reporting - using - Word Template";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog LoadDoc;
        private System.Windows.Forms.SaveFileDialog SaveDoc;
        private System.Windows.Forms.TextBox tFilename;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox tFirstname;
        private System.Windows.Forms.TextBox tLastname;
        private System.Windows.Forms.TextBox tPhone;
        private System.Windows.Forms.TextBox tCompany;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnLogo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.OpenFileDialog LoadPic;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Drawing.Printing.PrintDocument printDocument1;
    }
}

