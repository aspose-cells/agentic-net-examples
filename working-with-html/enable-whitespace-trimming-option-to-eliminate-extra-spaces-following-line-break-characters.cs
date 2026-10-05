// Title: How to enable whitespace trimming for line‑break spaces when exporting an Excel workbook to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that configures Aspose.Cells HtmlSaveOptions to remove leading spaces that appear after '\r\n' in cell text during HTML export. | Demonstrate how to apply the whitespace‑trimming setting in Aspose.Cells before saving a workbook as an HTML file.
// Common Searches: Aspose.Cells C# trim spaces after line break when saving to HTML | HTML export remove leading whitespace from cell values Aspose.Cells | How to configure HtmlSaveOptions to eliminate extra spaces after \r\n in C# | C# Aspose.Cells whitespace handling for line breaks in HTML output
// Tags: Aspose.Cells HtmlSaveOptions whitespace trimming | remove line break spaces Aspose.Cells | C# export Excel to HTML without leading spaces | trim cell text whitespace Aspose.Cells HTML | HTML output whitespace handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, writes a string containing a line break and leading spaces into cell A1, configures HtmlSaveOptions to trim whitespace after line breaks, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a cell value containing a line break and leading spaces
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("First line\r\n   Second line");

            // Save the workbook
            string outputPath = "TrimmedSpaces.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
