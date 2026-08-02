// Title: Copy a Chart Between Worksheets with Shapes.AddCopy in Aspose.Cells for .NET
// Description: Demonstrates how to duplicate a column chart from Sheet1 to Sheet3 using Shapes.AddCopy, then re‑assign the series and category ranges so the copied chart continues to reference the original data on Sheet1.
// Keywords: Aspose.Cells | C# | .NET | Shapes.AddCopy | copy chart between worksheets | chart shape | preserve data source | ChartShape | Excel chart copy | Aspose.Cells example
// Common Searches: Aspose.Cells copy chart to another sheet | Shapes.AddCopy chart example C# | keep chart data source after copying worksheet | duplicate chart with original references Aspose.Cells | how to use ChartShape in Aspose.Cells
// Developer Intent: Duplicate an existing chart on a different worksheet while retaining its original data links.
// Use Cases: Build a summary dashboard that shows live‑updating charts from source sheets without copying the data. | Generate a reporting workbook where charts are placed on a presentation sheet but still pull values from the original data tables. | Automate the replication of multiple charts onto a single sheet for a consolidated view, preserving each chart’s source range.
// AI Prompts: Show how to copy several charts from different worksheets to one dashboard sheet using Shapes.AddCopy, keeping each chart linked to its source data. | Provide a snippet that updates the copied chart’s data range dynamically based on the source worksheet name after the copy operation. | Explain best‑practice error handling for a null return from Shapes.AddCopy when copying a chart in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTransfer
{
    // Demonstrates how to duplicate a column chart from Sheet1 to Sheet3 using Shapes.AddCopy, then re‑assign the series and category ranges so the copied chart continues to reference the original data on Sheet1.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Add a third worksheet where the chart will be copied to
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Populate sample data in Sheet1 (the chart source)
                sheet1.Cells["A1"].PutValue("Category");
                sheet1.Cells["A2"].PutValue("A");
                sheet1.Cells["A3"].PutValue("B");
                sheet1.Cells["A4"].PutValue("C");
                sheet1.Cells["B1"].PutValue("Value");
                sheet1.Cells["B2"].PutValue(10);
                sheet1.Cells["B3"].PutValue(20);
                sheet1.Cells["B4"].PutValue(30);

                // Add a column chart to Sheet1
                int chartIndex = sheet1.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart sourceChart = sheet1.Charts[chartIndex];
                sourceChart.NSeries.Add("B2:B4", true);
                sourceChart.NSeries.CategoryData = "A2:A4";

                // Get the underlying shape of the chart
                ChartShape chartShape = sourceChart.ChartObject;

                // Copy the chart shape to Sheet3 using Shapes.AddCopy
                Shape copiedShape = sheet3.Shapes.AddCopy(chartShape, 5, 0, 15, 5);
                if (copiedShape == null)
                {
                    throw new InvalidOperationException("Failed to copy the chart shape.");
                }

                // Retrieve the Chart object from the copied shape
                Chart copiedChart = ((ChartShape)copiedShape).Chart;

                // Ensure the copied chart still references the original data range on Sheet1
                copiedChart.NSeries[0].Values = "Sheet1!B2:B4";
                copiedChart.NSeries.CategoryData = "Sheet1!A2:A4";

                // Define output file path
                string outputPath = "ChartTransferred.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
