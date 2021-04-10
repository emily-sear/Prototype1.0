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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Prototype1._0
{
    /**
     * Emily Sear (Programmer), Trey Karsten (Design Architecture), Andrew Cizek, and Turner Frazier
     * 10 April 2021 
     * Sprint 1 
     * **/
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

            //We don't want anything expect excel files 
            theDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            theDialog.InitialDirectory = @"C:\";

            DataTable theDataContainer = new DataTable();
            DataRow row;

            if (theDialog.ShowDialog() == DialogResult.OK)
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

                        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                        Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(pathName);

                        Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                        Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

                        int rowCount = excelRange.Rows.Count; //get row count of excel sheet
                        int colCount = excelRange.Columns.Count; //get column cout of excel data 

                        //get row data of Excel sheet 
                        for(int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                theDataContainer.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                            }
                            break;
                        }

                        
                        int rowCounter; //used for row index number 
                        for(int i = 2; i < rowCount; i++)
                        {
                            row = theDataContainer.NewRow(); //assign new row to DataTable
                            rowCounter = 0;
                            for(int j = 1; j <= colCount; j++) // loop for available column of excel data 
                            {
                                //check to see if the cell is empty 

                                if(excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                                {
                                    row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                                }
                                else
                                {
                                    row[i] = "";
                                }
                                rowCounter++;
                            }
                            theDataContainer.Rows.Add(row); //add the row to the DataTable
                        }
                       

                        dataGridView1.DataSource = theDataContainer; //assign DataTable as Datasource for DataGridview

                        //close and clean excel process
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(excelRange);
                        Marshal.ReleaseComObject(excelWorksheet);

                        //quit apps
                        excelWorkbook.Close();
                        Marshal.ReleaseComObject(excelWorkbook);
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. \nOriginal error " + ex.Message);
                }
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
