namespace WindowsFormsApp1
{
    partial class mainForm
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
            this.MLS_Input_File = new System.Windows.Forms.TextBox();
            this.AIM_Input_File = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialogMLS = new System.Windows.Forms.OpenFileDialog();
            this.OpenMLS = new System.Windows.Forms.Button();
            this.OpenAIM = new System.Windows.Forms.Button();
            this.openFileDialogAIM = new System.Windows.Forms.OpenFileDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // MLS_Input_File
            // 
            this.MLS_Input_File.Location = new System.Drawing.Point(41, 64);
            this.MLS_Input_File.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MLS_Input_File.Name = "MLS_Input_File";
            this.MLS_Input_File.Size = new System.Drawing.Size(701, 22);
            this.MLS_Input_File.TabIndex = 0;
            this.MLS_Input_File.Text = "C:\\Users\\Origami1105\\source\\repos\\BuschEric97\\HatcoFilesClosedDeterminer\\TestFile" +
    "s\\MLSData.xlsx";
            // 
            // AIM_Input_File
            // 
            this.AIM_Input_File.Location = new System.Drawing.Point(41, 150);
            this.AIM_Input_File.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.AIM_Input_File.Name = "AIM_Input_File";
            this.AIM_Input_File.Size = new System.Drawing.Size(701, 22);
            this.AIM_Input_File.TabIndex = 1;
            this.AIM_Input_File.Text = "C:\\Users\\Origami1105\\source\\repos\\BuschEric97\\HatcoFilesClosedDeterminer\\TestFile" +
    "s\\AIMData.xlsx";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 41);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(181, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "MLS Input/Output FIle Path:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(41, 127);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "AIM+ Input File Path:";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(41, 225);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(811, 108);
            this.button1.TabIndex = 6;
            this.button1.Text = "Run";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialogMLS
            // 
            this.openFileDialogMLS.FileName = "MLS.xlsx";
            // 
            // OpenMLS
            // 
            this.OpenMLS.Location = new System.Drawing.Point(752, 62);
            this.OpenMLS.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OpenMLS.Name = "OpenMLS";
            this.OpenMLS.Size = new System.Drawing.Size(100, 28);
            this.OpenMLS.TabIndex = 7;
            this.OpenMLS.Text = "Open";
            this.OpenMLS.UseVisualStyleBackColor = true;
            this.OpenMLS.Click += new System.EventHandler(this.OpenMLS_Click);
            // 
            // OpenAIM
            // 
            this.OpenAIM.Location = new System.Drawing.Point(752, 148);
            this.OpenAIM.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OpenAIM.Name = "OpenAIM";
            this.OpenAIM.Size = new System.Drawing.Size(100, 28);
            this.OpenAIM.TabIndex = 8;
            this.OpenAIM.Text = "Open";
            this.OpenAIM.UseVisualStyleBackColor = true;
            this.OpenAIM.Click += new System.EventHandler(this.OpenAIM_Click);
            // 
            // openFileDialogAIM
            // 
            this.openFileDialogAIM.FileName = "AIM.xlsx";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(41, 389);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(811, 28);
            this.progressBar1.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(41, 366);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 17);
            this.label3.TabIndex = 10;
            this.label3.Text = "Progress";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(893, 450);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.OpenAIM);
            this.Controls.Add(this.OpenMLS);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AIM_Input_File);
            this.Controls.Add(this.MLS_Input_File);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.Name = "mainForm";
            this.Text = "Hatco Files Closed Determiner";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox MLS_Input_File;
        private System.Windows.Forms.TextBox AIM_Input_File;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialogMLS;
        private System.Windows.Forms.Button OpenMLS;
        private System.Windows.Forms.Button OpenAIM;
        private System.Windows.Forms.OpenFileDialog openFileDialogAIM;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label3;
    }
}

