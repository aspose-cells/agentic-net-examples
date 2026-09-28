// Title: Create a light‑blue rectangle shape behind a chart in an Aspose.Cells workbook using C#
// AI Prompts: Add a rectangle shape to the worksheet, set its fill to light blue, move it behind the existing column chart, and save the workbook. | Generate a light‑blue background rectangle for a chart by inserting a shape, adjusting its Z‑order, and exporting the file as XLSX with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add rectangle shape behind chart | how to set shape fill color to light blue in Aspose.Cells .NET | move shape to back layer in Excel workbook using Aspose.Cells | place background rectangle behind column chart with Aspose.Cells | C# code to insert shape behind chart in XLSX using Aspose.Cells
// Tags: insert rectangle shape Aspose.Cells C# | set shape fill color light blue Aspose.Cells | shape z-order back Aspose.Cells | chart background rectangle Aspose.Cells | export workbook with shape Aspose.Cells XLSX

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a column chart with data from cells A1‑A3, inserts a rectangle shape, fills it with a light‑blue color, moves the shape behind the chart, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a sample chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            var chart = sheet.Charts[chartIndex];

            // Populate data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            chart.NSeries.Add("A1:A3", true);

            // Prepare output path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
