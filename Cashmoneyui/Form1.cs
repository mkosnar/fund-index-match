using System;
using System.IO;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Cashmoneyui
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string sFileName;
        ExcelData ed;

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog1.FileName;
            }

            if (sFileName is not null && sFileName != "")
            {
                ed = new ExcelData(sFileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sDate = textBox2.Text;
            
            try
            {
                var day = DateTime.Parse(sDate);
                textBox1.Text = ed.GetRate(day).ToString();
            }
            catch(Exception exc)
            {
                textBox1.Text = exc.ToString();
            }
        }
    }
}
