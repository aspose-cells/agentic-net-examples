using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace VerifyLeaderLines
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Configure data labels and offset them slightly
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd;
            // Offset the data label horizontally by 10 pixels (example)
            series.DataLabels.XPixel = series.DataLabels.XPixel + 10;

            // Verify leader lines visibility by reading the property
            bool leaderLinesVisible = series.HasLeaderLines;
            Console.WriteLine("Leader lines enabled: " + leaderLinesVisible);

            // Save the workbook
            workbook.Save("VerifyLeaderLines.xlsx");
        }
    }
}