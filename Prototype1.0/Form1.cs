using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Prototype1._0
{
    /**
     * Emily Sear (Programmer), Trey Karsten (Design Architecture), Andrew Cizek, and Turner Frazier
     * 24 April 2021 
     * Sprint 3
     * **/

    /**
     * Functionality of the sprint: 
     * Focus on Results/Data tab: 
     *      Allow users to flip through specified tool that the user wants & all corresponding data change
     *      Allow user to get rid of outliers 
     *      Allow for screen size adjustment and for the application to not look as weird 
     **/


    public partial class Form1 : Form
    {
        //generates tables for each corresponding instrument 
        public static DataTable theDataContainerGraduatedCylinder = new DataTable();
        public static DataTable theDataContainerHydrometer = new DataTable();
        public static DataTable theDataContainerBurette = new DataTable();
        public static DataTable theDataContainerThermometer = new DataTable();
        public static DataTable theDataContainerBalance = new DataTable();

        //Next sprint will go through these and see what is actually still being used 
        public static Double[] theData;
        public static int rowCount;
        public static int colCount;
        int theDataCount = 0;
        double instructorsValue;

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
            //Put the spreadsheet into the corresponding datatables and the datagridview
            Stream myStream = null;

            //Creates an open file dialog box to pull the excel file from
            OpenFileDialog theDialog = new OpenFileDialog();
            theDialog.Title = "Open Excel File with Data";

            //Only allows excel files to be uploaded
            theDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            theDialog.InitialDirectory = @"C:\";


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

                        Console.WriteLine(colCount);

                        theData = new Double[rowCount];

                        //Add the labels to each dataContainer
                        theDataContainerGraduatedCylinder.Columns.Add("Names");
                        theDataContainerHydrometer.Columns.Add("Names");
                        theDataContainerBurette.Columns.Add("Names");
                        theDataContainerThermometer.Columns.Add("Names");
                        theDataContainerBalance.Columns.Add("Names");

                        theDataContainerGraduatedCylinder.Columns.Add("Values");
                        theDataContainerHydrometer.Columns.Add("Values");
                        theDataContainerBurette.Columns.Add("Values");
                        theDataContainerThermometer.Columns.Add("Values");
                        theDataContainerBalance.Columns.Add("Values");

                        theDataContainerGraduatedCylinder.Columns.Add("Sig Figs");
                        theDataContainerHydrometer.Columns.Add("Sig Figs");
                        theDataContainerBurette.Columns.Add("Sig Figs");
                        theDataContainerThermometer.Columns.Add("Sig Figs");
                        theDataContainerBalance.Columns.Add("Sig Figs");

                        theDataContainerGraduatedCylinder.Columns.Add("Units");
                        theDataContainerHydrometer.Columns.Add("Units");
                        theDataContainerBurette.Columns.Add("Units");
                        theDataContainerThermometer.Columns.Add("Units");
                        theDataContainerBalance.Columns.Add("Units");

                        theDataContainerGraduatedCylinder.Columns.Add("Counted Values");
                        theDataContainerHydrometer.Columns.Add("Counted Values");
                        theDataContainerBurette.Columns.Add("Counted Values");
                        theDataContainerThermometer.Columns.Add("Counted Values");
                        theDataContainerBalance.Columns.Add("Counted Values");

                        for (int i = 2; i < rowCount; i++)
                        {
                            //add data to each specified data Container based on which part of the excel sheet it is at
                           DataRow row = theDataContainerGraduatedCylinder.NewRow(); //assign new row to DataTable
                            row[0] = excelRange.Cells[i, 1].Value2.ToString();
                            string[] currentRow = excelRange.Cells[i, 2].Value2.ToString().Split();
                            row[1] = currentRow[0];
                            row[2] = currentRow[0];
                            row[3] = currentRow[1];
                            row[4] = currentRow[0];
                            theDataContainerGraduatedCylinder.Rows.Add(row);

                            Console.WriteLine("Made it");

                            DataRow row2  = theDataContainerHydrometer.NewRow(); //assign new row to DataTable
                            row2[0] = excelRange.Cells[i, 1].Value2.ToString();
                            string[] currentRow2 = excelRange.Cells[i, 3].Value2.ToString().Split();
                            row2[1] = currentRow2[0];
                            row2[2] = currentRow2[0];
                            row2[3] = currentRow2[1];
                            row2[4] = currentRow2[0];
                            theDataContainerHydrometer.Rows.Add(row2);

                            Console.WriteLine("Made it");

                            DataRow row3 = theDataContainerBurette.NewRow(); //assign new row to DataTable
                            row3[0] = excelRange.Cells[i, 1].Value2.ToString();
                            string[] currentRow3 = excelRange.Cells[i, 4].Value2.ToString().Split();
                            row3[1] = currentRow3[0];
                            row3[2] = currentRow3[0];
                            //row3[3] = currentRow3[1];
                            row3[4] = currentRow3[0];
                            theDataContainerBurette.Rows.Add(row3);

                            Console.WriteLine("Made it");

                            DataRow row4 = theDataContainerThermometer.NewRow(); //assign new row to DataTable
                            row4[0] = excelRange.Cells[i, 1].Value2.ToString();
                            string[] currentRow4 = excelRange.Cells[i, 5].Value2.ToString().Split();
                            row4[1] = currentRow4[0];
                            row4[2] = currentRow4[0];
                            row4[3] = currentRow4[1];
                            row4[4] = currentRow4[0];
                            theDataContainerThermometer.Rows.Add(row4);

                            Console.WriteLine("Made it");

                            DataRow row5 = theDataContainerBalance.NewRow(); //assign new row to DataTable
                            row5[0] = excelRange.Cells[i, 1].Value2.ToString();
                            string[] currentRow5 = excelRange.Cells[i, 6].Value2.ToString().Split();
                            row5[1] = currentRow5[0];
                            row5[2] = currentRow5[0];
                            row5[3] = currentRow5[1];
                            row5[4] = currentRow5[0];
                            theDataContainerBalance.Rows.Add(row5);

                            Console.WriteLine("Made it");
                        
                        }
                        
                        dataGridView1.DataSource = theDataContainerGraduatedCylinder; //assign the graduated cylinder dataTable as Datasource for DataGridview
                        dataGridView1.Columns[0].Visible = false; //makes the names invisible originally

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
                    MessageBox.Show("Error: Could not read file from disk. \nOriginal error: " + ex.Message);
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
            dataGridView1.Columns[0].Visible = true; //changes the names to visible
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            instructorsValue = Convert.ToDouble(textBox4.Text); //need to put this in a grade method 
        }

        private void graduatedCylinderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = theDataContainerGraduatedCylinder; //changes the dataSource to the Graduated Cylinder

            //changes the image to the new image
            pictureBox1.Image = Properties.Resources._25014;
            pictureBox1.Refresh();
            pictureBox1.Visible = true;

        }

        private void hydrometerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = theDataContainerHydrometer;

            //changes the image to the new image
            pictureBox1.Image = Properties.Resources._25014;
            pictureBox1.Refresh();
            pictureBox1.Visible = true;


        }

        private void buretteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = theDataContainerBurette;

            pictureBox1.Image = Properties.Resources.Capture22;
            pictureBox1.Refresh();
            pictureBox1.Visible = true;

        }

        private void thermometerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = theDataContainerThermometer;

            pictureBox1.Image = Properties.Resources._71Bt_oSijtL__SL1500_;
            pictureBox1.Refresh();
            pictureBox1.Visible = true;

        }

        private void analyticalBalanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = theDataContainerBalance;

            pictureBox1.Image = Properties.Resources._118976197_183548856633170_5248159684347062085_o;
            pictureBox1.Refresh();
            pictureBox1.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //gets rid of outliers specified by the instructor
            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                if(dataGridView1.Rows[i].Cells[1].Value != null)
                {
                    if (Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()) < Convert.ToDouble(textBox1.Text) || Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()) > Convert.ToDouble(textBox2.Text))
                    {
                        dataGridView1.Rows.RemoveAt(i);
                        i--;
                    }
                }

            }
        }
    }
}
