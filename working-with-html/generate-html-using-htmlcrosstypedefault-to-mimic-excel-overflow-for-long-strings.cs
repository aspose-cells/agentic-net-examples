// Title: Create HTML from an Aspose.Cells workbook that mimics Excel cell overflow for long text using HtmlCrossType.Default (C#)
// AI Prompts: Write C# code that narrows a column, places a lengthy string in a cell, and saves the worksheet as HTML with HtmlCrossType.Default so the text spills into neighboring cells. | Show how to configure HtmlSaveOptions in Aspose.Cells to export a sheet to HTML while preserving Excel‑style overflow for long strings.
// Common Searches: asp.net aspose.cells export worksheet to html with text overflow | how to keep Excel cell overflow when converting to html using aspose.cells c# | HtmlCrossType.Default example for preserving overflow in html output | set column width and overflow long string in aspose.cells html export
// Tags: Aspose.Cells HTML export overflow | HtmlCrossType.Default C# | set column width Aspose.Cells | long string cell overflow Aspose.Cells | export worksheet to HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, narrows column A, inserts a long string into A1, and saves the active worksheet as HTML using HtmlSaveOptions with the default HtmlCrossType, demonstrating Excel‑style text overflow in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and set its name
            var sheet = workbook.Worksheets[0];
            sheet.Name = "OverflowDemo";

            // Set a narrow column width to force overflow (width in characters)
            sheet.Cells.SetColumnWidth(0, 10);

            // Insert a long string that exceeds the column width
            sheet.Cells["A1"].PutValue("This is a very long string that should overflow into adjacent empty cells when rendered as HTML.");

            // Configure HTML save options (default overflow behavior is sufficient)
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true
            };

            // Determine output file path and ensure the directory exists
            string outputPath = "OverflowDemo.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"HTML file saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
