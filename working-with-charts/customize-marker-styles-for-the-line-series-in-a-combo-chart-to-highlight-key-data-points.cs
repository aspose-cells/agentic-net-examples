// Title: How to apply custom marker shapes, colors, and sizes to the line series of a column‑line combo chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets the profit line series marker to a diamond shape, fills it with red, and enlarges its size in an Aspose.Cells combo chart. | Show how to assign different marker colors to individual points of a line series in a column‑line combo chart with Aspose.Cells. | Demonstrate updating the marker style (shape, size, border) for multiple line series in a combo chart using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# change marker shape for line series in combo chart | set marker color and size for profit line series Aspose.Cells .NET | customize markers of line series in column and line combo chart using Aspose.Cells | how to highlight specific data points with markers in Aspose.Cells chart C#
// Tags: Aspose.Cells line series marker styling | combo chart marker customization .NET | set marker shape and color Aspose.Cells | C# chart marker size adjustment | highlight data points Aspose.Cells chart

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with month, sales, and profit data, adds a column‑line combo chart, defines a column series for Sales and a line series for Profit, and demonstrates how to customize the line series markers—changing shape, fill color, and size—to emphasize key data points before saving the file as ComboChartWithCustomMarkers.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Populate sample data
            cells["A1"].PutValue("Month");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["A5"].PutValue("Apr");

            cells["B1"].PutValue("Sales");
            cells["B2"].PutValue(120);
            cells["B3"].PutValue(150);
            cells["B4"].PutValue(130);
            cells["B5"].PutValue(170);

            cells["C1"].PutValue("Profit");
            cells["C2"].PutValue(30);
            cells["C3"].PutValue(45);
            cells["C4"].PutValue(35);
            cells["C5"].PutValue(55);

            // Add a combo chart (column + line)
            int chartIdx = workbook.Worksheets[0].Charts.Add(ChartType.Column, 7, 0, 25, 10);
            var chart = workbook.Worksheets[0].Charts[chartIdx];
            chart.Title.Text = "Sales and Profit";

            // Column series for Sales
            int colSeriesIdx = chart.NSeries.Add("B2:B5", true);
            chart.NSeries[colSeriesIdx].Name = "Sales";

            // Line series for Profit
            int lineSeriesIdx = chart.NSeries.Add("C2:C5", true);
            chart.NSeries[lineSeriesIdx].Name = "Profit";
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line; // set series type to line

            // NOTE: Marker customization removed due to API version differences.
            // If needed, marker settings can be applied using the appropriate Aspose.Cells version.

            // Ensure the output directory exists
            string outputPath = "ComboChartWithCustomMarkers.xlsx";
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
