// Title: Export a specific print area from an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that defines a print area on a worksheet and saves only that area as an HTML file with Aspose.Cells. | Show how to enable HtmlSaveOptions.ExportPrintAreaOnly when converting a workbook to HTML in C#. | Create a sample workbook, set the print range A1:D10, and produce an HTML output that contains just that range.
// Common Searches: Aspose.Cells C# export only the defined print area to HTML | How to use HtmlSaveOptions to limit HTML output to a print range in .NET | Save Excel worksheet as HTML with ExportPrintAreaOnly flag using Aspose.Cells | Set print area before converting workbook to HTML in C# Aspose.Cells example
// Tags: Aspose.Cells HtmlSaveOptions print area | C# export worksheet range to HTML | Aspose.Cells limit HTML output to selected cells | define print area Aspose.Cells .NET | HTML conversion of Excel print area

using System;
using Aspose.Cells;

// The example creates a new workbook, fills cells A1:D10 with data, sets the worksheet's print area to that range, configures HtmlSaveOptions with ExportPrintAreaOnly enabled, and saves the workbook as an HTML file. The resulting HTML contains only the defined print area.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data (A1:D10)
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define the print area (only A1:D10 will be considered for export)
        sheet.PageSetup.PrintArea = "A1:D10";

        // Set HTML save options to export only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportPrintAreaOnly = true   // ensures only the print area appears in the HTML
        };

        // Export the workbook to HTML
        workbook.Save("ExportedPrintArea.html", htmlOptions);
    }
}
