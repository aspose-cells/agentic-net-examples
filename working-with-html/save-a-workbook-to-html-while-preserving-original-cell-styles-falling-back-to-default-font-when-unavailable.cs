// Title: Save Aspose.Cells Workbook to HTML with Original Styles and Fallback Font (C#)
// Description: Demonstrates how to export an Aspose.Cells workbook to HTML while keeping all cell formatting and specifying a default font (e.g., Arial) that Aspose.Cells uses when a cell's font is unavailable on the target system.
// Keywords: Aspose.Cells | HtmlSaveOptions | DefaultFontName | C# | HTML export | preserve cell styles | fallback font | missing font handling | Excel to HTML | style preservation .NET
// Common Searches: Aspose.Cells set default font for HTML export | preserve Excel cell formatting when saving as HTML C# | fallback font for unavailable fonts Aspose.Cells | HtmlSaveOptions.ExcludeUnusedStyles effect | export styled workbook to HTML using Aspose.Cells
// Developer Intent: Export a workbook to HTML, retain every cell's visual style, and automatically replace fonts that are not installed with a defined default font.
// Use Cases: Create web‑ready reports from Excel files that look identical to the original spreadsheet. | Generate HTML emails or dashboards where the recipient may not have custom fonts installed. | Produce HTML versions of spreadsheets that can be re‑imported without losing style information.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML, preserving all cell styles and using Arial as a fallback font for missing fonts. | Explain how HtmlSaveOptions.DefaultFontName works and how it affects font substitution during HTML export. | Describe the impact of HtmlSaveOptions.ExcludeUnusedStyles on round‑trip conversion between Excel and HTML.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to export an Aspose.Cells workbook to HTML while keeping all cell formatting and specifying a default font (e.g., Arial) that Aspose.Cells uses when a cell's font is unavailable on the target system.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with various styles
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Styled Text");
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.Name = "Calibri";
            styleA1.Font.Size = 12;
            styleA1.Font.Color = Color.Blue;
            styleA1.Font.IsBold = true;
            cellA1.SetStyle(styleA1);

            Cell cellB2 = sheet.Cells["B2"];
            cellB2.PutValue("Another Style");
            Style styleB2 = cellB2.GetStyle();
            styleB2.Font.Name = "NonExistentFont"; // This font may not be available on the system
            styleB2.Font.Size = 14;
            styleB2.Font.Color = Color.Green;
            cellB2.SetStyle(styleB2);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Preserve original cell styles (default behavior) and specify a fallback font
            // When a font is not found, Aspose.Cells will use this default font
            htmlOptions.DefaultFontName = "Arial";

            // Optional: keep all styles even if they are not used (helps when later importing back)
            // htmlOptions.ExcludeUnusedStyles = false;

            // Save the workbook as HTML
            string outputPath = "StyledWorkbook.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML at: {outputPath}");
        }
    }
}
