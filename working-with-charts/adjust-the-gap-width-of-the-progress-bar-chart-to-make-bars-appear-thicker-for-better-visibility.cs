using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsProgressBarDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for a progress‑bar style column chart
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["A4"].PutValue("Testing");

                sheet.Cells["B1"].PutValue("Progress");
                sheet.Cells["B2"].PutValue(70);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(30);

                // Add a 2‑D column chart (used as a progress bar)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Reduce the gap width to make the bars appear thicker.
                // GapWidth is a percentage of the bar width (0‑500). Smaller values = thicker bars.
                chart.GapWidth = 30; // e.g., 30% of the default gap

                // Optional: make the chart look like a progress bar
                chart.NSeries[0].Overlap = -100; // bars will touch each other
                // Enable data labels and show the value
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Ensure the output directory exists
                string outputPath = "ProgressBarThickBars.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}