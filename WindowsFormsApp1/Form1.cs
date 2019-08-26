﻿using System;
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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Determiner det = new Determiner();
            try
            {
                // run main determiner function to perform the main function of the program
                det.mainDeterminer(MLS_Input_File.Text, AIM_Input_File.Text, 0.2, 0.2, 0.5, 0.5);
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
            Console.WriteLine("entered OpenMLS_Click()");
            openFileDialogMLS.ShowHelp = true;
            Console.WriteLine("set openFileDialogMLS.ShowHelp = true");
            openFileDialogMLS.ShowDialog();
            Console.WriteLine("exited openFileDialogMLS.ShowDialog()");
            MLS_Input_File.Text = openFileDialogMLS.FileName;
            Console.WriteLine("exited OpenMLS_Click()");
        }

        private void OpenAIM_Click(object sender, EventArgs e)
        {
            openFileDialogAIM.ShowHelp = true;
            openFileDialogAIM.ShowDialog();
            AIM_Input_File.Text = openFileDialogAIM.FileName;
        }
    }
}
