// Title: Export only a defined print area of an Excel worksheet to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets a worksheet's print area to A1:C4 and saves only that area as an HTML file using Aspose.Cells. | Show how to configure HtmlSaveOptions.ExportPrintAreaOnly to export a specific cell range to HTML in Aspose.Cells for .NET. | Create a sample workbook, define a print range, and produce an HTML output that contains only the defined cells with Aspose.Cells.
// Common Searches: Aspose.Cells C# export only selected range to HTML | How to save a print area as HTML with Aspose.Cells .NET | HtmlSaveOptions ExportPrintAreaOnly usage example | Set print area before HTML conversion in Aspose.Cells | Export Excel print area to HTML file using Aspose.Cells for .NET
// Tags: html export print area Aspose.Cells | HtmlSaveOptions ExportPrintAreaOnly C# | define worksheet print area Aspose.Cells | save specific cell range as HTML .NET | Aspose.Cells limited range HTML output

using System;
using Aspose.Cells;

// Creates a workbook, fills cells A1:C4 with data, sets the worksheet's print area to that range, configures HtmlSaveOptions to export only the print area, and saves the result as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(0.5);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue(0.3);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(15);
        sheet.Cells["C4"].PutValue(0.8);

        // Define the print area (only cells A1:C4 will be exported)
        sheet.PageSetup.PrintArea = "A1:C4";

        // Configure HTML save options to export only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportPrintAreaOnly = true;

        // Export the workbook to HTML
        workbook.Save("ExportedPrintArea.html", htmlOptions);
    }
}
