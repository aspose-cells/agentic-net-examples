using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLeaderLinesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart (leader lines are not supported for this chart type)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Get the first series of the chart
            Series series = chart.NSeries[0];

            // Attempt to configure leader lines with proper error handling
            try
            {
                // Enable leader lines
                series.HasLeaderLines = true;

                // Customize leader line appearance
                series.LeaderLines.IsAuto = false;
                series.LeaderLines.Style = LineType.Dot;
                series.LeaderLines.WeightPt = 1.5;
                series.LeaderLines.Color = Color.Blue;

                Console.WriteLine("Leader lines configured successfully.");
            }
            catch (Exception ex)
            {
                // Handle cases where the chart type does not support leader lines
                Console.WriteLine($"Error configuring leader lines: {ex.Message}");
                // Optionally, you can disable leader lines if they were partially set
                series.HasLeaderLines = false;
            }

            // Save the workbook
            workbook.Save("LeaderLinesWithErrorHandling.xlsx");
        }
    }
}