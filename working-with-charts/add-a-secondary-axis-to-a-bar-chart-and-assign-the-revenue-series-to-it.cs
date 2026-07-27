using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    class AddSecondaryAxisDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B1"].PutValue("Units Sold");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C1"].PutValue("Revenue");
            sheet.Cells["C2"].PutValue(3000);
            sheet.Cells["C3"].PutValue(4500);
            sheet.Cells["C4"].PutValue(5400);

            // Add a bar chart
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add series: Units Sold and Revenue
            chart.NSeries.Add("B2:B4", true); // Units Sold series
            chart.NSeries.Add("C2:C4", true); // Revenue series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the Revenue series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Customize the secondary value axis (optional)
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Revenue (USD)";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // Save the workbook
            workbook.Save("BarChartWithSecondaryAxis.xlsx");
        }
    }
}