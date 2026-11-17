// Title: C# example: Export Excel workbook to HTML with text truncated to cell bounds using HtmlCrossType.FitToCell (Aspose.Cells)
// Common Searches: how to limit long text overflow when saving Excel as HTML with Aspose.Cells in C# | Aspose.Cells HtmlCrossType.FitToCell property usage for HTML export | C# save workbook to HTML with cell content truncated to column width | prevent cell text spilling out of bounds in HTML output from Excel using Aspose | set column width and fit text to cell during Excel to HTML conversion in .NET
// Tags: HtmlCrossType.FitToCell usage | export Excel to HTML C# | truncate cell text in HTML output | column width handling Aspose.Cells | HTML save options cell overflow control

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, writes a long string to cell A1, sets the column width, and saves the sheet as an HTML file using HtmlSaveOptions. The HtmlCrossType.FitToCell setting ensures the text is truncated to fit within the cell bounds, preventing overflow in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long text into cell A1
            sheet.Cells["A1"].PutValue(
                "This is a very long text that should be truncated to fit within the cell bounds when exported to HTML.");

            // Optionally set column width to demonstrate the fit-to-cell effect
            sheet.Cells.SetColumnWidth(0, 20); // Column A width

            // Configure HTML save options (default behavior already fits text to cell)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Determine output path and ensure directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
