// Title: Export Excel to HTML with full style preservation and fallback font using Aspose.Cells for .NET
// Description: Shows how to create a workbook, apply custom styles (including a missing font), and save it as HTML with HtmlSaveOptions.ExcludeUnusedStyles=false and DefaultFontName='Arial' so every cell's formatting is retained and unavailable fonts fall back to Arial.
// Keywords: Aspose.Cells | C# | .NET | HTML export | preserve cell styles | fallback font | DefaultFontName | ExcludeUnusedStyles | Excel to HTML | Aspose.Cells HtmlSaveOptions | web report generation | US | Europe
// Common Searches: Aspose.Cells export Excel to HTML preserving styles | How to keep cell formatting when saving workbook as HTML with Aspose.Cells | Set default font for missing fonts in Aspose.Cells HTML export | HtmlSaveOptions ExcludeUnusedStyles example C# | C# Aspose.Cells HTML fallback font
// Developer Intent: Generate an HTML file from an Excel workbook that keeps every cell’s formatting and substitutes a default font when the original font isn’t installed.
// Use Cases: Web dashboards that display Excel data with exact styling. | Emailing spreadsheet content as HTML where custom fonts may be unavailable. | Automated documentation pipelines converting Excel reports to HTML. | Cross‑region reporting where font availability varies.
// AI Prompts: Provide C# code using Aspose.Cells to save a workbook as HTML with all styles preserved and Arial as fallback font. | Explain how HtmlSaveOptions.ExcludeUnusedStyles and DefaultFontName affect the generated HTML and CSS. | Suggest ways to embed the exported HTML into a web page while ensuring consistent appearance across browsers.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to create a workbook, apply custom styles (including a missing font), and save it as HTML with HtmlSaveOptions.ExcludeUnusedStyles=false and DefaultFontName='Arial' so every cell's formatting is retained and unavailable fonts fall back to Arial.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with different styles
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("Styled Text");
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.Color = System.Drawing.Color.Red;
            styleA1.Font.IsBold = true;
            cellA1.SetStyle(styleA1);

            Cell cellB2 = sheet.Cells["B2"];
            cellB2.PutValue("Another Style");
            Style styleB2 = cellB2.GetStyle();
            styleB2.Font.Name = "NonExistentFont"; // This font may not be installed
            styleB2.Font.Size = 14;
            cellB2.SetStyle(styleB2);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Preserve all styles (do not exclude unused ones)
            saveOptions.ExcludeUnusedStyles = false;

            // Specify a fallback font when the original font is unavailable
            saveOptions.DefaultFontName = "Arial";

            // Save the workbook as HTML
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to HTML at: {outputPath}");
        }
    }
}
