namespace Muhasebe
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btMuhasebe2Banka = new System.Windows.Forms.Button();
            this.btBanka2Talimat = new System.Windows.Forms.Button();
            this.btExportExcel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.VeritabaninaYazOld = new System.Windows.Forms.Button();
            this.ExceleYaz = new System.Windows.Forms.Button();
            this.lbBilgi = new System.Windows.Forms.Label();
            this.btLink2Table = new System.Windows.Forms.Button();
            this.btKapat = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btMuhasebe2Banka
            // 
            this.btMuhasebe2Banka.Location = new System.Drawing.Point(46, 103);
            this.btMuhasebe2Banka.Name = "btMuhasebe2Banka";
            this.btMuhasebe2Banka.Size = new System.Drawing.Size(190, 50);
            this.btMuhasebe2Banka.TabIndex = 0;
            this.btMuhasebe2Banka.Text = "Muhasebe/Banka";
            this.btMuhasebe2Banka.UseVisualStyleBackColor = true;
            this.btMuhasebe2Banka.Click += new System.EventHandler(this.btMuhasebe2Banka_Click);
            // 
            // btBanka2Talimat
            // 
            this.btBanka2Talimat.Location = new System.Drawing.Point(46, 159);
            this.btBanka2Talimat.Name = "btBanka2Talimat";
            this.btBanka2Talimat.Size = new System.Drawing.Size(190, 50);
            this.btBanka2Talimat.TabIndex = 1;
            this.btBanka2Talimat.Text = "Banka/Talimat";
            this.btBanka2Talimat.UseVisualStyleBackColor = true;
            this.btBanka2Talimat.Click += new System.EventHandler(this.btBanka2Talimat_Click);
            // 
            // btExportExcel
            // 
            this.btExportExcel.Location = new System.Drawing.Point(46, 224);
            this.btExportExcel.Name = "btExportExcel";
            this.btExportExcel.Size = new System.Drawing.Size(190, 50);
            this.btExportExcel.TabIndex = 2;
            this.btExportExcel.Text = "Export > (Database2Excel)";
            this.btExportExcel.UseVisualStyleBackColor = true;
            this.btExportExcel.Click += new System.EventHandler(this.btExportExcel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 469);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(936, 23);
            this.progressBar1.TabIndex = 3;
            // 
            // VeritabaninaYazOld
            // 
            this.VeritabaninaYazOld.Location = new System.Drawing.Point(73, 345);
            this.VeritabaninaYazOld.Name = "VeritabaninaYazOld";
            this.VeritabaninaYazOld.Size = new System.Drawing.Size(190, 23);
            this.VeritabaninaYazOld.TabIndex = 10;
            this.VeritabaninaYazOld.Text = "Import < (Excel2Database) old";
            this.VeritabaninaYazOld.UseVisualStyleBackColor = true;
            this.VeritabaninaYazOld.Visible = false;
            this.VeritabaninaYazOld.Click += new System.EventHandler(this.VeritabaninaYaz_Click);
            // 
            // ExceleYaz
            // 
            this.ExceleYaz.Location = new System.Drawing.Point(73, 374);
            this.ExceleYaz.Name = "ExceleYaz";
            this.ExceleYaz.Size = new System.Drawing.Size(190, 23);
            this.ExceleYaz.TabIndex = 9;
            this.ExceleYaz.Text = "Excele Yaz old";
            this.ExceleYaz.UseVisualStyleBackColor = true;
            this.ExceleYaz.Visible = false;
            // 
            // lbBilgi
            // 
            this.lbBilgi.AutoSize = true;
            this.lbBilgi.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbBilgi.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lbBilgi.Font = new System.Drawing.Font("Segoe UI Semibold", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lbBilgi.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lbBilgi.Location = new System.Drawing.Point(0, 449);
            this.lbBilgi.Name = "lbBilgi";
            this.lbBilgi.Size = new System.Drawing.Size(0, 20);
            this.lbBilgi.TabIndex = 11;
            // 
            // btLink2Table
            // 
            this.btLink2Table.Location = new System.Drawing.Point(46, 47);
            this.btLink2Table.Name = "btLink2Table";
            this.btLink2Table.Size = new System.Drawing.Size(190, 50);
            this.btLink2Table.TabIndex = 12;
            this.btLink2Table.Text = "Import < (Excel2Database)";
            this.btLink2Table.UseVisualStyleBackColor = true;
            this.btLink2Table.Click += new System.EventHandler(this.Link2Table_Click);
            // 
            // btKapat
            // 
            this.btKapat.Location = new System.Drawing.Point(821, 391);
            this.btKapat.Name = "btKapat";
            this.btKapat.Size = new System.Drawing.Size(92, 50);
            this.btKapat.TabIndex = 13;
            this.btKapat.Text = "Kapat";
            this.btKapat.UseVisualStyleBackColor = true;
            this.btKapat.Click += new System.EventHandler(this.btKapat_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(936, 492);
            this.Controls.Add(this.btKapat);
            this.Controls.Add(this.btLink2Table);
            this.Controls.Add(this.lbBilgi);
            this.Controls.Add(this.VeritabaninaYazOld);
            this.Controls.Add(this.ExceleYaz);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btExportExcel);
            this.Controls.Add(this.btBanka2Talimat);
            this.Controls.Add(this.btMuhasebe2Banka);
            this.Name = "Form1";
            this.Text = "Muhasebe";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        #endregion

        private Button btMuhasebe2Banka;
        private Button btBanka2Talimat;
        private Button btExportExcel;
        private ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
       
        private Button VeritabaninaYazOld;
        private Button ExceleYaz;
        private Label lbBilgi;
        private Button btLink2Table;
        private Button btKapat;
    }
}