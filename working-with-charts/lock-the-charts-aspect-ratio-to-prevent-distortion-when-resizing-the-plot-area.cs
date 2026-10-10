// Title: How to prevent chart distortion by locking the aspect ratio when resizing a plot area with Aspose.Cells for .NET
// AI Prompts: Create a C# example using Aspose.Cells that attempts to lock a chart’s aspect ratio and documents the limitation of the API. | Show how to programmatically keep an Excel column chart’s proportions stable during worksheet resizing with Aspose.Cells, including any available work‑arounds.
// Common Searches: Aspose.Cells .NET lock chart aspect ratio | prevent Excel chart stretching with Aspose.Cells C# | keep chart proportions when resizing plot area Aspose.Cells | is there a chart aspect ratio property in Aspose.Cells | workaround for chart distortion in Aspose.Cells after resizing
// Tags: Aspose.Cells chart aspect ratio | lock chart proportions Aspose.Cells | prevent chart distortion Aspose.Cells | chart resizing behavior Aspose.Cells | column chart formatting Aspose.Cells C# | Excel chart aspect ratio limitation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, fills cells A1:B4 with month and sales data, adds a column chart, and explains that Aspose.Cells for .NET does not expose a direct property to lock the chart’s aspect ratio, so the chart may stretch when the plot area is resized. The workbook is saved as ChartAspectRatioLocked.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Aspose.Cells does not expose direct properties to lock the chart's aspect ratio.
            // Additional formatting can be applied here if needed.

            // Save the workbook
            string outputPath = "ChartAspectRatioLocked.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
