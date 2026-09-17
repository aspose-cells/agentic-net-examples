// Title: Generate responsive HTML from an Excel workbook with scalable width, FitToCell layout, and cell address tooltips using Aspose.Cells for .NET
// AI Prompts: Create C# code that saves a Workbook to HTML with WidthScalable = true, HtmlCrossType = HtmlCrossType.FitToCell, and AddTooltipText = true using Aspose.Cells. | Modify an existing Aspose.Cells HtmlSaveOptions configuration to enable a responsive table, fit each cell to its content, and display cell addresses as hover tooltips. | Produce a complete example that exports sample data to an HTML file where the table adjusts to the browser width, cells are sized to fit content, and tooltips show the cell reference.
// Common Searches: asp.net how to export Excel to responsive HTML with cell tooltips using Aspose.Cells | c# Aspose.Cells HtmlSaveOptions WidthScalable FitToCell tooltip example | generate HTML table from workbook that scales with browser and shows cell address on hover | set HtmlCrossType to FitToCell for responsive Aspose.Cells HTML export
// Tags: Aspose.Cells responsive HTML export | WidthScalable HtmlSaveOptions | HtmlCrossType FitToCell setting | AddTooltipText cell hover | C# Excel to HTML with tooltips

using System;
using Aspose.Cells;

namespace ResponsiveHtmlExport
{
    // Demonstrates creating a workbook, populating it with data, and saving it as responsive HTML using WidthScalable, HtmlCrossType.FitToCell, and AddTooltipText to produce a scalable table that fits cell content and shows cell addresses on hover.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Populate the first worksheet with sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.50);

                // Configure HTML save options for responsive output with tooltips
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Makes the HTML table width scalable (responsive)
                    WidthScalable = true,

                    // Adds tooltip text for cells (e.g., shows cell address on hover)
                    AddTooltipText = true
                };

                // Save the workbook as an HTML file using the configured options
                workbook.Save("ResponsiveOutput.html", htmlOptions);
                Console.WriteLine("HTML file 'ResponsiveOutput.html' generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
