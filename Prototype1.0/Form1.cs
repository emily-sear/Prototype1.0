using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
//using Excel = Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Prototype1._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dataToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void resultsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void spreadSheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream myStream = null;

            //Creates an open file dialog box to pull the .xml file from 
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File with Data";

            //We don't want anything expect .xml files 
            theDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            theDialog.InitialDirectory = @"C:\";

            if(theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if((myStream = theDialog.OpenFile()) != null)
                    {
                        //generates some info about the file location in order to read the file 
                        System.IO.FileInfo fInfo = new System.IO.FileInfo(theDialog.FileName);
                        string strFileLocation = fInfo.FullName;

                        string pathName = theDialog.FileName;
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(theDialog.FileName);
                        DataTable tbContainer = new DataTable();
                        string strConn = string.Empty;
                        string sheetName = fileName;

                        FileInfo file = new FileInfo(pathName);
                        if (!file.Exists) { throw new Exception("Error, file doesn't exists!"); }
                        string extension = file.Extension;
                        switch (extension)
                       {
                           case ".xls":
                                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                               break;
                            case ".xlsx":
                                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                               break;
                           default:
                                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                                break;
                       }
                        OleDbConnection cnnxls = new OleDbConnection(strConn);
                        OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), cnnxls);
                        oda.Fill(tbContainer);

                        //dtGrid.DataSource = tbContainer;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error" + ex.Message);
                }
            }

        }

    }
}
