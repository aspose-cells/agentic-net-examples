using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendControl
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with six series (more than five)
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Series 2
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Series 3
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);
            sheet.Cells["D5"].PutValue(42);

            // Series 4
            sheet.Cells["E1"].PutValue("Series4");
            sheet.Cells["E2"].PutValue(18);
            sheet.Cells["E3"].PutValue(28);
            sheet.Cells["E4"].PutValue(38);
            sheet.Cells["E5"].PutValue(48);

            // Series 5
            sheet.Cells["F1"].PutValue("Series5");
            sheet.Cells["F2"].PutValue(14);
            sheet.Cells["F3"].PutValue(24);
            sheet.Cells["F4"].PutValue(34);
            sheet.Cells["F5"].PutValue(44);

            // Series 6 (extra series to trigger legend hiding)
            sheet.Cells["G1"].PutValue("Series6");
            sheet.Cells["G2"].PutValue(16);
            sheet.Cells["G3"].PutValue(26);
            sheet.Cells["G4"].PutValue(36);
            sheet.Cells["G5"].PutValue(46);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including all series)
            chart.SetChartDataRange("A1:G5", true);

            // Determine if the number of series exceeds five
            if (chart.NSeries.Count > 5)
            {
                // Hide the legend to reduce visual clutter
                chart.ShowLegend = false;
            }

            // Save the workbook
            workbook.Save("ChartWithConditionalLegend.xlsx");
        }
    }
}