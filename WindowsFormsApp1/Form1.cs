using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;


namespace WindowsFormsApp1
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            Determiner det = new Determiner();
            progressBar1.Maximum = 100;
            progressBar1.Minimum = 0;
            var progress = new Progress<int>(v =>
            {
               progressBar1.Increment(v);
            });

            try
            {
                // run main determiner function to perform the main function of the program
                progressBar1.Value = progressBar1.Minimum;
                await Task.Run(() => det.mainDeterminer(MLS_Input_File.Text, AIM_Input_File.Text,
                    0.75, 0.75, 0.75, 0.75, progress));
                progressBar1.Value = progressBar1.Maximum;
                MessageBox.Show("Complete!");
            }
            catch (Exception ex)
            {
                // display any exceptions that are thrown as a popup message box
                MessageBox.Show(ex.ToString());
            }
        }

        private void OpenMLS_Click(object sender, EventArgs e)
        {
            openFileDialogMLS.ShowHelp = true;
            openFileDialogMLS.ShowDialog();
            MLS_Input_File.Text = openFileDialogMLS.FileName;
        }

        private void OpenAIM_Click(object sender, EventArgs e)
        {
            openFileDialogAIM.ShowHelp = true;
            openFileDialogAIM.ShowDialog();
            AIM_Input_File.Text = openFileDialogAIM.FileName;
        }

        /// <summary>
        /// Event Handler for when the whole application closes. It is in Form1.cs
        /// so that it has access to the file names to make sure they are closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static public void OnApplicationExit(object sender, EventArgs e)
        {
            Console.WriteLine("Exiting Application");
        }
    }
}
