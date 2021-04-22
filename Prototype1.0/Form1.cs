using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Prototype1._0
{
    /**
     * Emily Sear (Programmer), Trey Karsten (Design Architecture), Andrew Cizek, and Turner Frazier
     * 17 April 2021 
     * Sprint 2
     * **/

    /**
     * Functionality of the sprint: 
     * Focus on the "View" portion of the application 
     * Make a working graph section 
     * Allow user to view the student names
     * allow user to enter in instructor values (not allowed to grade quite yet
     **/


    public partial class Form1 : Form
    {
        //generates table for imported info
        public static DataTable theDataContainer = new DataTable();
        public static Double[] theData;
        public static int rowCount;
        public static int colCount;
        int theDataCount = 0;
        double instructorsValue;
        int nameSpot = 0;
        int graduatedCylinderSpot = 1;
        int buretteSpot = 2;
        int hyrdometerSpot = 3;
        int thermometerSpot = 4;
        int balanceSpot = 5;


       public static double lowValue = 36.4;

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

            //Creates an open file dialog box to pull the excel file from
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File with Data";

            //Only allows excel files to be uploaded
            theDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            theDialog.InitialDirectory = @"C:\";

            DataRow row;

            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if((myStream = theDialog.OpenFile()) != null)
                    { 
                        
                        System.IO.FileInfo fInfo = new System.IO.FileInfo(theDialog.FileName);
                        string strFileLocation = fInfo.FullName;
                        

                        //used for passing on the file name for future functions
                        string pathName = theDialog.FileName;
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(theDialog.FileName);


                        //Opens the info from the first page of the excel sheet
                        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(pathName);
                        Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                        Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                      

                        rowCount = excelRange.Rows.Count; //get row count of excel sheet
                        colCount = excelRange.Columns.Count; //get column cout of excel data 

                        theData = new Double[rowCount];
                        //get the labels of Excel sheet
                            for (int i = 1; i <= rowCount; i++)
                            {
                                for (int j = 1; j <= colCount; j++)
                                {
                                    
                                    string columnName = excelRange.Cells[i, j].Value2.ToString();
                                    theDataContainer.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                                    if (columnName.Contains("Graduated"))
                                    {
                                        graduatedCylinderSpot = j - 1;
                                    Console.WriteLine(j);
                                }
                                    else if(columnName.Contains("Names"))
                                    {
                                        nameSpot = j -1 ;
                                        Console.WriteLine(j);
                                    }
                                    else if(columnName.Contains("Bur"))
                                    {
                                        buretteSpot = j-1;
                                        Console.WriteLine(j);
                                    }
                                    else if(columnName.Contains("Hyrdo"))
                                    {
                                        hyrdometerSpot = j-1;
                                    Console.WriteLine(j);
                                }
                                    else if(columnName.Contains("Thermometer"))
                                    {
                                        thermometerSpot = j-1;
                                    Console.WriteLine(j);
                                }
                                    else if(columnName.Contains("Balance"))
                                    {
                                        balanceSpot = j-1;
                                    Console.WriteLine(j);
                                }
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
                                if(excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                                {
                                    row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                                    if (j == 2)
                                    {
                                        theData[theDataCount] = excelRange.Cells[i, j].Value2;
                                        theDataCount++;
                                    }
                                    
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
                        dataGridView1.Columns[nameSpot].Visible = false; //makes the names invisible originally
                        dataGridView1.Columns[graduatedCylinderSpot].Visible = true;
                        dataGridView1.Columns[hyrdometerSpot].Visible = false;
                        dataGridView1.Columns[buretteSpot].Visible = false;
                        dataGridView1.Columns[thermometerSpot].Visible = false;
                        dataGridView1.Columns[balanceSpot].Visible = false;

                        //clean up the excel sheet and close it 
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(excelRange);
                        Marshal.ReleaseComObject(excelWorksheet);
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

        private void graphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Creates and displays table when the "graph" button is clicked
            Form2 graphForm = new Form2();
            graphForm.Show();
           

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void studentNamesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[nameSpot].Visible = true; //changes the names to visible
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            instructorsValue = Convert.ToDouble(textBox4.Text); //need to put this in a grade method 
        }

        private void graduatedCylinderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[graduatedCylinderSpot].Visible = true;
            dataGridView1.Columns[hyrdometerSpot].Visible = false;
            dataGridView1.Columns[buretteSpot].Visible = false;
            dataGridView1.Columns[thermometerSpot].Visible = false;
            dataGridView1.Columns[balanceSpot].Visible = false;

        }

        private void hydrometerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[graduatedCylinderSpot].Visible = false;
            dataGridView1.Columns[hyrdometerSpot].Visible = true;
            dataGridView1.Columns[buretteSpot].Visible = false;
            dataGridView1.Columns[thermometerSpot].Visible = false;
            dataGridView1.Columns[balanceSpot].Visible = false;

            //changes the image to the new image
            pictureBox1.Image = Properties.Resources.glass_cylinder_new;
            pictureBox1.Refresh();
            pictureBox1.Visible = true;


        }

        private void buretteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[graduatedCylinderSpot].Visible = false;
            dataGridView1.Columns[hyrdometerSpot].Visible = false;
            dataGridView1.Columns[buretteSpot].Visible = true;
            dataGridView1.Columns[thermometerSpot].Visible = false;
            dataGridView1.Columns[balanceSpot].Visible = false;
        }

        private void thermometerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[graduatedCylinderSpot].Visible = false;
            dataGridView1.Columns[hyrdometerSpot].Visible = false;
            dataGridView1.Columns[buretteSpot].Visible = false;
            dataGridView1.Columns[thermometerSpot].Visible = true;
            dataGridView1.Columns[balanceSpot].Visible = false;
        }

        private void analyticalBalanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[graduatedCylinderSpot].Visible = false;
            dataGridView1.Columns[hyrdometerSpot].Visible = false;
            dataGridView1.Columns[buretteSpot].Visible = false;
            dataGridView1.Columns[thermometerSpot].Visible = false;
            dataGridView1.Columns[balanceSpot].Visible = true;
        }
    }
}
