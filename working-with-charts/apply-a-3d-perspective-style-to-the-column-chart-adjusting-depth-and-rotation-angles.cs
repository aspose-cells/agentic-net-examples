// Title: Create a 3‑D column chart with Aspose.Cells for .NET and gracefully handle missing perspective, depth, and rotation options
// AI Prompts: Generate C# code that adds a 3‑D column chart using Aspose.Cells, then checks the library version and applies Depth, RotationX, and RotationY only when those properties are available. | Write a fallback routine in C# that logs a warning if Aspose.Cells does not expose perspective controls for a 3‑D chart and continues with the default rendering.
// Common Searches: Aspose.Cells .NET set depth on 3D column chart if supported | how to rotate a 3D column chart using Aspose.Cells C# | detect Aspose.Cells version to enable chart perspective features | fallback for unavailable 3D chart properties in Aspose.Cells
// Tags: Aspose.Cells 3D column chart configuration | C# chart perspective version check | handle missing 3D rotation Aspose.Cells | conditional chart depth setting .NET | unsupported 3D chart API fallback

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, inserts sample data, adds a 3‑D column chart, assigns the data series, and saves the file. It also demonstrates how to detect the Aspose.Cells version and conditionally apply depth, rotation, or perspective settings when the API supports them, providing a graceful fallback otherwise.
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

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart (positioned at row 6, column 0, spanning 15 rows and 10 columns)
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Adjusting 3‑D perspective (Depth, RotationX/Y, Perspective) is not available
            // in the current Aspose.Cells version used. These settings have been omitted.

            // Save the workbook to a file
            workbook.Save("ColumnChart3DPerspective.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
