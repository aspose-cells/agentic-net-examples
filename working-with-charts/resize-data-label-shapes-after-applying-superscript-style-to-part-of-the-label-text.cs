using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    public class ResizeDataLabelAfterSuperscript
    {
        public static void Run()
        {
            try
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

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Iterate through each point to customize its data label
                foreach (ChartPoint point in series.Points)
                {
                    // Access the data label of the current point
                    DataLabels dl = point.DataLabels;

                    // Set the full text first (e.g., "10 (x10ⁿ)")
                    dl.Text = $"{point.YValue} (x10)";

                    // Apply superscript to the last character "n"
                    int start = dl.Text.Length - 1;
                    dl.Characters(start, 1).Font.IsSuperscript = true;

                    // Ensure the shape resizes to fit the modified text
                    dl.IsResizeShapeToFitText = true;

                    // Apply the font changes to all child nodes of the data label
                    dl.ApplyFont();
                }

                // Define output file path
                string outputPath = "ResizeDataLabelAfterSuperscript.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ResizeDataLabelAfterSuperscript.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}