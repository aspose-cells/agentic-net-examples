// Title: Export Excel to HTML with Gridlines as CSS Borders using Aspose.Cells for .NET
// Description: Shows how to save a workbook as HTML with worksheet gridlines rendered as CSS borders by enabling IsGridlinesVisible and setting HtmlSaveOptions.ExportGridLines = true.
// Keywords: Aspose.Cells HTML export gridlines | ExportGridLines true | Excel to HTML CSS borders | Preserve Excel gridlines in HTML | HtmlSaveOptions Aspose.Cells C#
// Common Searches: Aspose.Cells export Excel to HTML with gridlines | HtmlSaveOptions ExportGridLines example C# | How to keep Excel gridlines when converting to HTML | Render Excel cell borders as CSS in HTML output
// Developer Intent: Generate an HTML file from an Excel workbook that keeps the worksheet’s gridlines as CSS border styles.
// Use Cases: Display a spreadsheet in a web portal with the exact grid layout from Excel. | Create email‑ready HTML reports that preserve the original cell borders. | Provide an on‑the‑fly preview of uploaded Excel files without losing visual formatting.
// AI Prompts: Give a C# example that loads an existing .xlsx, ensures gridlines are visible, and saves it to HTML with borders using Aspose.Cells. | Explain the relationship between IsGridlinesVisible and HtmlSaveOptions.ExportGridLines and how to customize border colors. | Show how to export multiple worksheets to a single HTML page while preserving each sheet's gridlines as CSS borders.

using System;
using Aspose.Cells;

// Shows how to save a workbook as HTML with worksheet gridlines rendered as CSS borders by enabling IsGridlinesVisible and setting HtmlSaveOptions.ExportGridLines = true.
class ExportExcelToHtmlWithGridlines
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Make sure gridlines are visible in the worksheet
        sheet.IsGridlinesVisible = true;

        // Add some sample data
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue("Data1");
        sheet.Cells["B2"].PutValue("Data2");

        // Configure HTML save options to export gridlines as CSS borders
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportGridLines = true
        };

        // Save the workbook as an HTML file
        workbook.Save("output.html", options);
    }
}
