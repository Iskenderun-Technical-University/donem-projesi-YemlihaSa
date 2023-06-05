namespace Muhasebe
{
    partial class Form2
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
            this.ExceleYaz = new System.Windows.Forms.Button();
            this.VeritabaninaYaz = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ExceleYaz
            // 
            this.ExceleYaz.Location = new System.Drawing.Point(440, 293);
            this.ExceleYaz.Name = "ExceleYaz";
            this.ExceleYaz.Size = new System.Drawing.Size(75, 23);
            this.ExceleYaz.TabIndex = 0;
            this.ExceleYaz.Text = "Excele Yaz";
            this.ExceleYaz.UseVisualStyleBackColor = true;
            this.ExceleYaz.Click += new System.EventHandler(this.ExceleYaz_Click);
            // 
            // VeritabaninaYaz
            // 
            this.VeritabaninaYaz.Location = new System.Drawing.Point(12, 26);
            this.VeritabaninaYaz.Name = "VeritabaninaYaz";
            this.VeritabaninaYaz.Size = new System.Drawing.Size(224, 23);
            this.VeritabaninaYaz.TabIndex = 1;
            this.VeritabaninaYaz.Text = "Veritabanına Yaz";
            this.VeritabaninaYaz.UseVisualStyleBackColor = true;
            this.VeritabaninaYaz.Click += new System.EventHandler(this.VeritabaninaYaz_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.VeritabaninaYaz);
            this.Controls.Add(this.ExceleYaz);
            this.Name = "Form2";
            this.Text = "Form2";
            this.ResumeLayout(false);

        }

        #endregion

        private Button ExceleYaz;
        private Button VeritabaninaYaz;
    }
}