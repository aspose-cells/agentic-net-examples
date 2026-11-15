// Title: Set custom minimum and maximum values for a chart Y‑axis with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Y axis minimum value for column chart | How to define maximum value for chart value axis using Aspose.Cells .NET | C# code to fix Y axis range of an Excel chart with Aspose.Cells | Adjust chart Y axis scaling in an Aspose.Cells workbook | Set major tick unit on chart value axis with Aspose.Cells C#
// Tags: set y‑axis min max Aspose.Cells | chart value axis scaling .NET | custom y‑axis range Excel chart Aspose.Cells | define major unit chart axis Aspose.Cells | column chart y‑axis limits Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAxisScaling
{
    // The example creates a workbook, adds sample data, inserts a column chart, and configures the chart's Y‑axis by setting MinValue to 0, MaxValue to 100, and MajorUnit to 20 before saving the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // Column A: Categories, Column B: Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(55);
                sheet.Cells["B4"].PutValue(35);
                sheet.Cells["B5"].PutValue(80);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart series
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // ----- Axis scaling -----
                // Define custom minimum and maximum values for the Y‑axis
                chart.ValueAxis.MinValue = 0;    // Minimum value
                chart.ValueAxis.MaxValue = 100; // Maximum value

                // Optional: set major unit for better tick spacing
                chart.ValueAxis.MajorUnit = 20;

                // Determine output file path
                string outputPath = "ChartWithCustomYAxis.xlsx";

                // Ensure the directory exists before saving
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
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
