// Title: How to position a column chart at the top‑left corner of a PDF page with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets a chart’s Placement to FreeFloating and moves it to the origin (row 0, column 0) before exporting the workbook to PDF using Aspose.Cells. | Show the steps to create a worksheet, add sample data, insert a column chart, and align the chart to the page’s top‑left edge for PDF output in a .NET application.
// Common Searches: Aspose.Cells C# set chart coordinates (0,0) for PDF generation | How to move a worksheet chart to the page origin before saving as PDF with Aspose.Cells | Example of floating chart placement in Aspose.Cells .NET | Adjust chart location to top corner when converting workbook to PDF using Aspose.Cells
// Tags: chart absolute placement Aspose.Cells | top-left chart PDF layout .NET | column chart PDF positioning C# | Aspose.Cells chart origin alignment | chart coordinate configuration for PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // Required for PlacementType

// The sample creates a workbook, fills cells A1:B4 with data, adds a column chart, sets its Placement to FreeFloating, moves the chart to the page origin (top‑left corner), and saves the workbook as ChartTopLeft.pdf, ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the chart at the top‑left corner of the page
            chart.Placement = PlacementType.FreeFloating; // Allows absolute positioning

            // Define output file path
            string outputPath = "ChartTopLeft.pdf";

            // Ensure the directory for the output file exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook (with the chart) to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
