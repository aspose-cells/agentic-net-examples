using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class OffsetDataLabelsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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

            // Add a pie chart (leader lines are most visible with pie charts)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and place them outside the slices
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd; // moves labels outward

            // Enable leader lines so the labels are connected to the data points
            series.HasLeaderLines = true;

            // Optional: customize leader line appearance
            series.LeaderLines.IsAuto = false;
            series.LeaderLines.WeightPt = 1.0;
            series.LeaderLines.Color = Color.DarkGray;

            // Save the workbook
            workbook.Save("OffsetDataLabelsDemo.xlsx");
        }
    }
}