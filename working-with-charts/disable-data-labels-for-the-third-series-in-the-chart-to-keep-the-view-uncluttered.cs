using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class DisableThirdSeriesDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for three series
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

                // First series values
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Second series values
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Third series values
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(12);
                sheet.Cells["D3"].PutValue(22);
                sheet.Cells["D4"].PutValue(32);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add three series to the chart
                chart.NSeries.Add("=Sheet1!$B$2:$B$4", true); // Series 1
                chart.NSeries.Add("=Sheet1!$C$2:$C$4", true); // Series 2
                chart.NSeries.Add("=Sheet1!$D$2:$D$4", true); // Series 3

                // Set category (X) data
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

                // Enable data labels for the first two series (optional)
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[1].DataLabels.ShowValue = true;

                // ----- Disable data labels for the third series -----
                Series thirdSeries = chart.NSeries[2];

                // Hide all possible label components
                thirdSeries.DataLabels.ShowValue = false;
                thirdSeries.DataLabels.ShowCategoryName = false;
                thirdSeries.DataLabels.ShowPercentage = false;
                thirdSeries.DataLabels.ShowSeriesName = false;
                thirdSeries.DataLabels.ShowLegendKey = false;
                thirdSeries.DataLabels.ShowBubbleSize = false;
                thirdSeries.DataLabels.ShowCellRange = false;

                // Save the workbook
                string outputPath = "DisableThirdSeriesDataLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableThirdSeriesDataLabels.Run();
        }
    }
}