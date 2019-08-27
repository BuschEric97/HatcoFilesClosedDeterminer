namespace WindowsFormsApp1
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
            this.MLS_Input_File.Location = new System.Drawing.Point(31, 52);
            this.MLS_Input_File.Name = "MLS_Input_File";
            this.MLS_Input_File.Size = new System.Drawing.Size(527, 20);
            this.MLS_Input_File.TabIndex = 0;
            this.MLS_Input_File.Text = "C:\\Users\\Origami1105\\source\\repos\\BuschEric97\\HatcoFilesClosedDeterminer\\TestFile" +
    "s\\MLSData.xlsx";
            // 
            // AIM_Input_File
            // 
            this.AIM_Input_File.Location = new System.Drawing.Point(31, 122);
            this.AIM_Input_File.Name = "AIM_Input_File";
            this.AIM_Input_File.Size = new System.Drawing.Size(527, 20);
            this.AIM_Input_File.TabIndex = 1;
            this.AIM_Input_File.Text = "C:\\Users\\Origami1105\\source\\repos\\BuschEric97\\HatcoFilesClosedDeterminer\\TestFile" +
    "s\\AIMData.xlsx";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(141, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "MLS Input/Output FIle Path:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 103);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "AIM+ Input File Path:";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(31, 183);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(608, 88);
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
            this.OpenMLS.Location = new System.Drawing.Point(564, 50);
            this.OpenMLS.Name = "OpenMLS";
            this.OpenMLS.Size = new System.Drawing.Size(75, 23);
            this.OpenMLS.TabIndex = 7;
            this.OpenMLS.Text = "Open";
            this.OpenMLS.UseVisualStyleBackColor = true;
            this.OpenMLS.Click += new System.EventHandler(this.OpenMLS_Click);
            // 
            // OpenAIM
            // 
            this.OpenAIM.Location = new System.Drawing.Point(564, 120);
            this.OpenAIM.Name = "OpenAIM";
            this.OpenAIM.Size = new System.Drawing.Size(75, 23);
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
            this.progressBar1.Location = new System.Drawing.Point(31, 316);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(608, 23);
            this.progressBar1.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 297);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(48, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Progress";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(670, 366);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.OpenAIM);
            this.Controls.Add(this.OpenMLS);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AIM_Input_File);
            this.Controls.Add(this.MLS_Input_File);
            this.MaximizeBox = false;
            this.Name = "Form1";
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

