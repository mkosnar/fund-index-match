using System;
using System.Windows.Forms;

namespace Cashmoneyui
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        ExcelData ed;

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Excel File to Edit";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";

            string sFileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog1.FileName;
            }

            if (sFileName is not null && sFileName != "")
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                ed = new ExcelData(sFileName, 3);
                watch.Stop();
                textBoxLoadTime.Text = watch.ElapsedMilliseconds.ToString();
            }
        }

        private void buttonWrite_Click(object sender, EventArgs e)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            ed.WriteMatches();
            watch.Stop();
            textBoxWriteTime.Text = watch.ElapsedMilliseconds.ToString();
        }
    }
}
