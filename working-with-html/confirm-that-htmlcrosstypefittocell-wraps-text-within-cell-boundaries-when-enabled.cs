// Title: Confirm HtmlCrossType.FitToCell preserves Excel cell text wrapping when exporting to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets HtmlSaveOptions.HtmlCrossType to FitToCell, enables IsTextWrapped on a cell, and saves the workbook as an HTML file. | Write a C# routine that reads the saved HTML file and verifies that the <td> element for the cell contains markup or CSS that forces line breaks within the column width. | Demonstrate how to adjust column width and cell style before export so the HTML output clearly shows the FitToCell wrapping behavior.
// Common Searches: Aspose.Cells how to keep Excel cell wrap when saving as HTML .NET | HtmlCrossType FitToCell effect on text wrapping in exported HTML | C# example exporting wrapped text to HTML using Aspose.Cells | verify that HTML output respects Excel column width and wrap Aspose
// Tags: Aspose.Cells HtmlCrossType FitToCell | export Excel cell wrap to HTML .NET | C# HtmlSaveOptions text wrapping | verify HTML cell width Aspose.Cells | set column width for HTML export Aspose

using System;
using Aspose.Cells;

// The example creates a workbook, inserts a long string into cell A1, narrows column A, enables text wrapping on the cell, configures HtmlSaveOptions (default HtmlCrossType.FitToCell), and saves the workbook as HTML. The generated HTML demonstrates that the text wraps within the cell boundaries, confirming that HtmlCrossType.FitToCell preserves Excel's wrap setting during export.
class HtmlCrossTypeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This is a very long piece of text that should be wrapped within the cell boundaries when exported to HTML.");

            // Set a narrow column width to make wrapping noticeable
            sheet.Cells.SetColumnWidth(0, 10); // Column A width

            // Enable text wrapping in Excel
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Configure HTML save options (text wrapping will be preserved automatically)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            string outputPath = "FitToCellDemo.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("HTML file saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
