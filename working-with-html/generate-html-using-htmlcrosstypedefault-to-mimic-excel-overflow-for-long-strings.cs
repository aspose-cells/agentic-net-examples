// Title: How to export Excel cell overflow to HTML using Aspose.Cells HtmlCrossType.Default in C#
// Common Searches: C# Aspose.Cells export cell text overflow to HTML with HtmlCrossType.Default | preserve Excel overflow when saving worksheet as HTML using Aspose.Cells | generate HTML that shows long string overflow like Excel using Aspose.Cells | how to keep Excel cell overflow in HTML output with Aspose.Cells .NET | set narrow column width and export overflowed text to HTML in C#
// Tags: HtmlCrossType.Default overflow handling | Aspose.Cells HTML export with overflow support | C# HtmlSaveOptions active worksheet only | column width overflow simulation Aspose.Cells | Excel cell overflow HTML output .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, narrows column A to force a long string in cell A1 to overflow, configures HtmlSaveOptions to export only the active worksheet, and saves the file as HTML. Using HtmlCrossType.Default, the generated HTML mimics Excel's cell overflow behavior.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "OverflowDemo";

            // Set a narrow column width to force overflow of long text
            sheet.Cells.SetColumnWidth(0, 10); // Column A width

            // Insert a long string into cell A1
            sheet.Cells["A1"].PutValue("This is a very long string that should overflow into adjacent empty cells when viewed in Excel.");

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportActiveWorksheetOnly = true;

            // Save the workbook as HTML
            string outputPath = "OverflowDemo.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
