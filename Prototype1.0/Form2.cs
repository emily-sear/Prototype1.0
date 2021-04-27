using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
namespace Prototype1._0
{
    public partial class Form2 : Form
    {
        public static double[] theData;
        public Form2()
        {
            InitializeComponent();
            var theDataSeries = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                //Initialize the table (name, and some general parameters).
                //
                Name = "TheDataSeries",
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Column,
                IsXValueIndexed = true
                //
            };

            //Pre-defining veriables.
            double num1 = 0;
            double num2 = 0;
            double num3 = 0;
            double num4 = 0;
            double num5 = 0; //mid point variable
            double num6 = 0;
            double num7 = 0;
            double num8 = 0;
            double num9 = 0;
            
            //This entire nested for-if is used to sort the data gotten in Form1 from the spreadsheet into different columns.
            //
            for (int i = 0; i < theData.Length; i++)
            {
                Console.WriteLine(theData[i]);
                if (theData[i] <= Form1.lowValue)
                {
                    num1++;
                }
                else if (theData[i] == Form1.lowValue + 1)
                {
                    num2++;
                }
                else if (theData[i] == Form1.lowValue + 2)
                {
                    num3++;
                }
                else if (theData[i] == Form1.lowValue + 3)
                {
                    num4++;
                }
                else if (theData[i] == Form1.lowValue + 4)
                {
                    num5++;
                }
                else if (theData[i] == Form1.lowValue + 5)
                {
                    num6++;
                }
                else if (theData[i] == Form1.lowValue + 6)
                {
                    num7++;
                }
                else if (theData[i] == Form1.lowValue + 7)
                {
                    num8++;
                }
                else
                {
                    num9++;
                }

            }

            double midPoint = (Form1.highValue + Form1.lowValue) / 2;

            //This points data to the correct X,Y coordinate on the table.
            theDataSeries.Points.AddXY(Form1.lowValue, num1);
            theDataSeries.Points.AddXY(2, num2);
            theDataSeries.Points.AddXY(3, num3);
            theDataSeries.Points.AddXY(4, num4);
            theDataSeries.Points.AddXY(midPoint, num5);
            theDataSeries.Points.AddXY(6, num6);
            theDataSeries.Points.AddXY(7, num7);
            theDataSeries.Points.AddXY(8, num8);
            theDataSeries.Points.AddXY(Form1.highValue, num9);
            

            this.chart1.Series.Add(theDataSeries); //This stores the data into the table.
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
