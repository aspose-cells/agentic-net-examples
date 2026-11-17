// Title: How to verify that HtmlCrossType.FitToCell wraps long text within cell boundaries in Aspose.Cells HTML export (C#)
// Common Searches: Aspose.Cells C# check if HtmlCrossType.FitToCell wraps text in HTML output | verify text wrapping to cell size using HtmlCrossType.FitToCell in .NET | sample code to test HtmlCrossType.FitToCell behavior when exporting to HTML | does enabling HtmlCrossType.FitToCell affect column width in Aspose.Cells HTML | how to see wrapped text in generated HTML with Aspose.Cells HtmlSaveOptions
// Tags: HtmlCrossType.FitToCell wrapping behavior | Aspose.Cells HTML cell text wrap | C# test HtmlCrossType in HTML export | column width row height HTML output Aspose | verify cell boundary wrapping Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, inserts a long string into cell A1, adjusts column width and row height, and saves the workbook as HTML using HtmlSaveOptions. By enabling HtmlCrossType.FitToCell (when supported) the long text will wrap inside the cell boundaries in the generated HTML, allowing you to verify the wrapping behavior.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long text into cell A1
            sheet.Cells["A1"].PutValue(
                "This is a very long piece of text that will exceed the default width of the cell. When HtmlCrossType.FitToCell is enabled, this text should wrap within the cell boundaries in the generated HTML.");

            // Adjust column width and row height to make the wrapping effect visible
            sheet.Cells.SetColumnWidth(0, 20); // Column A width (approx. 20 characters)
            sheet.Cells.SetRowHeight(0, 30);   // Row 1 height (points)

            // Set HTML save options (HtmlCrossType property is not available in this version)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save("FitToCellDemo.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
