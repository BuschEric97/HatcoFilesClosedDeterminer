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
            this.SuspendLayout();
            // 
            // MLS_Input_File
            // 
            this.MLS_Input_File.Location = new System.Drawing.Point(41, 64);
            this.MLS_Input_File.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MLS_Input_File.Name = "MLS_Input_File";
            this.MLS_Input_File.Size = new System.Drawing.Size(439, 22);
            this.MLS_Input_File.TabIndex = 0;
            this.MLS_Input_File.Text = "C:\\Users\\Origami1105\\Desktop\\MLSData.xlsx";
            // 
            // AIM_Input_File
            // 
            this.AIM_Input_File.Location = new System.Drawing.Point(41, 150);
            this.AIM_Input_File.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.AIM_Input_File.Name = "AIM_Input_File";
            this.AIM_Input_File.Size = new System.Drawing.Size(439, 22);
            this.AIM_Input_File.TabIndex = 1;
            this.AIM_Input_File.Text = "C:\\Users\\Origami1105\\Desktop\\AIMData.xlsx";
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
            this.button1.Location = new System.Drawing.Point(644, 64);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(296, 108);
            this.button1.TabIndex = 6;
            this.button1.Text = "Run";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1067, 228);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AIM_Input_File);
            this.Controls.Add(this.MLS_Input_File);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
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
    }
}

