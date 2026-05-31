using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboTrendlineDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Column B – values for the column series
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Column C – values for the line series (to which we will add a trendline)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);
            sheet.Cells["C5"].PutValue(70);

            // Add a Combo chart (initially a Column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // First series – column type (Sales)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Sales";

            // Second series – will be displayed as a line (Profit)
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Profit";
            chart.NSeries[1].Type = ChartType.Line; // Convert this series to a line

            // Add a linear trendline to the line series and display its equation
            int trendlineIdx = chart.NSeries[1].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[1].TrendLines[trendlineIdx];
            trendline.DisplayEquation = true;      // Show equation on the chart
            trendline.DisplayRSquared = false;     // Optional: hide R‑squared
            trendline.Color = Color.Blue;          // Optional: set trendline color

            // Save the workbook
            workbook.Save("ComboChartWithTrendline.xlsx");
        }
    }
}