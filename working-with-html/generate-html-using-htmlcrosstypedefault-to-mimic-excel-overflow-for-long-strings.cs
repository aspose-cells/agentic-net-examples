// Title: Generate HTML with Aspose.Cells (C#) that mimics Excel long‑text overflow across adjacent cells
// AI Prompts: Create a C# program that builds a workbook, writes a lengthy string to a cell, sets column widths, and saves it as HTML using Aspose.Cells HtmlSaveOptions to retain Excel‑style overflow. | Show how to configure HtmlSaveOptions (or rely on its defaults) so the generated HTML displays the text spilling into neighboring empty cells like Excel does. | Provide code that exports a worksheet to HTML while preserving column width settings that make the overflow visible in the output.
// Common Searches: how to keep text overflow when converting Excel to HTML with Aspose.Cells C# | Aspose.Cells HtmlSaveOptions long string spills into next cells | C# export worksheet to HTML preserving Excel cell overflow behavior | set column width before HTML export Aspose.Cells to show overflow | HTML output from Aspose.Cells that mimics Excel overflow
// Tags: Aspose.Cells HtmlSaveOptions overflow | C# export Excel to HTML with text spill | preserve Excel cell overflow in HTML | set column widths for HTML export Aspose.Cells | HTML output mimicking Excel overflow behavior

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // This example creates a workbook, inserts a long string into cell A1, defines column widths for columns A‑C, and saves the workbook as an HTML file using Aspose.Cells HtmlSaveOptions, which by default allows the text to overflow into adjacent empty cells, reproducing Excel's visual behavior.
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

            // Put a long string into cell A1
            sheet.Cells["A1"].PutValue(
                "This is a very long string that should overflow into adjacent empty cells when viewed in Excel.");

            // Set column widths for visibility (columns A, B, C)
            sheet.Cells.SetColumnWidth(0, 20); // Column A
            sheet.Cells.SetColumnWidth(1, 20); // Column B
            sheet.Cells.SetColumnWidth(2, 20); // Column C

            // Configure HTML save options (default behavior already allows overflow)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Determine output file path
            string outputPath = "OverflowExample.html";

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
