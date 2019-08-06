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
            Console.WriteLine("MLS Input File Path: " + MLS_Input_File.Text);
            Console.WriteLine("AIM Input File Path: " + AIM_Input_File.Text);
            Console.WriteLine("Output File Path: " + Output_File.Text);
        }
    }
}
