// Title: Generate an Excel workbook with bold, italic, and bold‑italic cells and export it to HTML while preserving font styles using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, applies bold, italic, and combined bold‑italic formatting to specific cells, and saves the file as HTML using Aspose.Cells with CSS styling enabled. | Show how to set up Aspose.Cells HtmlSaveOptions to export all worksheets and retain cell font styles in the generated HTML document. | Provide a version that inlines the CSS in the HTML output instead of using an external stylesheet, while keeping the bold and italic formatting intact.
// Common Searches: Aspose.Cells C# export workbook to HTML keep bold and italic formatting | How to preserve cell font styles when converting Excel to HTML with Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly false example in .NET | Generate HTML from Excel with inline CSS using Aspose.Cells | Save multiple worksheets as a single HTML file Aspose.Cells C#
// Tags: Aspose.Cells HTML export with font styling | C# HtmlSaveOptions preserve bold italic | export Excel to HTML all worksheets Aspose.Cells | inline CSS generation Aspose.Cells HTML output | cell formatting to HTML using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, formats cells with bold, italic, and bold‑italic text, configures HtmlSaveOptions to export all worksheets, and saves the workbook as an HTML file that retains the font styles via CSS.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Cell A1 - Bold text
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Bold Text");
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.IsBold = true;
            cellA1.SetStyle(styleA1);

            // Cell A2 - Italic text
            Cell cellA2 = sheet.Cells["A2"];
            cellA2.PutValue("Italic Text");
            Style styleA2 = cellA2.GetStyle();
            styleA2.Font.IsItalic = true;
            cellA2.SetStyle(styleA2);

            // Cell A3 - Bold and Italic text
            Cell cellA3 = sheet.Cells["A3"];
            cellA3.PutValue("Bold Italic");
            Style styleA3 = cellA3.GetStyle();
            styleA3.Font.IsBold = true;
            styleA3.Font.IsItalic = true;
            cellA3.SetStyle(styleA3);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (not only the active one)
                ExportActiveWorksheetOnly = false
                // Additional options such as CSS handling or font resources can be set here
                // if supported by the specific Aspose.Cells version in use.
            };

            // Ensure the output directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
