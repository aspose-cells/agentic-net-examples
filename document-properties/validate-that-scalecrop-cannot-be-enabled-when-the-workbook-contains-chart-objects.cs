using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ScaleCropValidationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Determine if any worksheet contains chart objects
                bool containsChart = false;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Charts.Count > 0)
                    {
                        containsChart = true;
                        break;
                    }
                }

                // Access built‑in document properties (type inferred)
                var properties = workbook.BuiltInDocumentProperties;

                // Validate ScaleCrop: it must not be enabled when charts are present
                if (containsChart)
                {
                    // Disable ScaleCrop when charts exist
                    if (properties["ScaleCrop"] != null)
                        properties["ScaleCrop"].Value = false;
                    Console.WriteLine("ScaleCrop remains disabled because the workbook contains chart objects.");
                }
                else
                {
                    // Enable ScaleCrop when no charts exist
                    if (properties["ScaleCrop"] != null)
                        properties["ScaleCrop"].Value = true;
                    Console.WriteLine("ScaleCrop enabled: no chart objects detected.");
                }

                // Define output file path
                string outputPath = "ScaleCropValidationResult.xlsx";

                // Save the workbook (save rule)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}