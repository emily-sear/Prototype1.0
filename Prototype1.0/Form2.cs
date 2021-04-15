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

        public Form2()
        {
            InitializeComponent();
            var theDataSeries = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "TheDataSeries",
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Column,
                IsXValueIndexed = true
            };

            
            double num1 = 0;
            double num2 = 0;
            double num3 = 0;
            double num4 = 0;
            double num5 = 0;
            double num6 = 0;
            double num7 = 0;
            double num8 = 0;
            double num9 = 0;
            double num10 = 0;
            
            for (int i = 0; i < Form1.theData.Length; i++)
            {
                Console.WriteLine(Form1.theData[i]);
                if (Form1.theData[i] <= Form1.lowValue)
                {
                    num1++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 1)
                {
                    num2++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 2)
                {
                    num3++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 3)
                {
                    num4++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 4)
                {
                    num5++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 5)
                {
                    num6++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 6)
                {
                    num7++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 7)
                {
                    num8++;
                }
                else if (Form1.theData[i] == Form1.lowValue + 8)
                {
                    num9++;
                }
                else
                {
                    num10++;
                }

            }

            theDataSeries.Points.AddXY(1, num1);
            theDataSeries.Points.AddXY(2, num2);
            theDataSeries.Points.AddXY(3, num3);
            theDataSeries.Points.AddXY(4, num4);
            theDataSeries.Points.AddXY(5, num5);
            theDataSeries.Points.AddXY(6, num6);
            theDataSeries.Points.AddXY(7, num7);
            theDataSeries.Points.AddXY(8, num8);
            theDataSeries.Points.AddXY(9, num9);
            theDataSeries.Points.AddXY(10, num10);

            this.chart1.Series.Add(theDataSeries);
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
