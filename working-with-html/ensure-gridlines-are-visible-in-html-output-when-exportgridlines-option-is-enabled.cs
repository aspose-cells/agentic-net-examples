// Title: Export Excel to HTML with Gridlines Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable worksheet gridlines, populate sample data, configure HtmlSaveOptions with ExportGridLines = true (and optionally ExportActiveWorksheetOnly), and save the workbook as an HTML file that displays the gridlines.
// Keywords: aspnet aspose.cells html export gridlines | htmlsaveoptions exportgridlines example | excel gridlines in html c# | aspose.cells export active worksheet only | c# convert excel to html with borders
// Common Searches: aspose.cells export gridlines to html | htmlsaveoptions exportgridlines not working | show excel gridlines in html output | c# save worksheet as html with borders | export only active sheet to html aspose
// Developer Intent: Generate an HTML file from an Excel workbook that preserves the worksheet’s gridlines.
// Use Cases: Publish a spreadsheet‑based report on a website while keeping the familiar grid layout. | Create a single‑sheet HTML snapshot for embedding in documentation or intranet portals. | Produce printable HTML versions of Excel data that retain cell borders for visual consistency.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, ensuring gridlines are visible and adding custom CSS for styling. | Explain why both Worksheet.IsGridlinesVisible and HtmlSaveOptions.ExportGridLines must be set for gridlines to appear in the HTML output. | Provide a C# snippet that iterates through all worksheets and saves each as a separate HTML file with gridlines enabled.

using System;
using Aspose.Cells;

namespace AsposeCellsGridlinesHtmlDemo
{
    // Demonstrates how to enable worksheet gridlines, populate sample data, configure HtmlSaveOptions with ExportGridLines = true (and optionally ExportActiveWorksheetOnly), and save the workbook as an HTML file that displays the gridlines.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Enable gridlines visibility in the worksheet
            sheet.IsGridlinesVisible = true;

            // Add some sample data so the gridlines can be observed
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Enable gridline export
                ExportActiveWorksheetOnly = true      // Export only the active sheet (optional)
            };

            // Save the workbook as HTML with gridlines visible
            workbook.Save("GridlinesOutput.html", htmlOptions);

            Console.WriteLine("HTML file saved with gridlines exported.");
        }
    }
}
