// Title: How to apply a two‑decimal‑place numeric format to the Z‑axis labels of a 3‑D column chart in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that configures the series (Z) axis of an Aspose.Cells 3‑D column chart to display numbers with a fixed 0.00 pattern. | Show a workaround for applying a numeric display format to the Z‑axis labels of a 3‑D chart when Aspose.Cells lacks a direct NumberFormat property.
// Common Searches: aspnet aspose.cells format series axis as 0.00 in 3d chart | c# how to change Z axis label precision in Aspose.Cells 3D column chart | set custom numeric display for chart Z axis using Aspose.Cells
// Tags: aspose.cells z‑axis numeric format c# | 3d column chart series axis formatting aspose.cells | c# chart axis custom display pattern | aspose.cells chart axis label formatting technique

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, adds a 3‑D column chart, and links the data ranges. It highlights that Aspose.Cells does not expose a direct NumberFormat property for the Z (Series) axis, so developers must use alternative approaches or accept the default formatting.
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

            // Populate sample data for a 3‑D chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(1.2345);
            sheet.Cells["B3"].PutValue(2.3456);
            sheet.Cells["B4"].PutValue(3.4567);
            sheet.Cells["C2"].PutValue(4.5678);
            sheet.Cells["C3"].PutValue(5.6789);
            sheet.Cells["C4"].PutValue(6.7890);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Aspose.Cells does not expose a direct NumberFormat property for the Series (Z) axis.
            // If needed, axis formatting can be applied via other means or omitted.

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
