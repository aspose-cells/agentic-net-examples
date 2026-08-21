// Title: Aspose.Cells for .NET – Set Times New Roman as Default Font in HTML Export
// Description: Shows how to assign "Times New Roman" to HtmlSaveOptions.DefaultFontName and save a Workbook as HTML, ensuring cells without explicit font settings use the specified typeface.
// Keywords: Aspose.Cells | HtmlSaveOptions | DefaultFontName | Times New Roman | HTML export | C# | Excel to HTML | set default font | workbook.Save | .NET
// Common Searches: Aspose.Cells set default HTML font | HtmlSaveOptions DefaultFontName C# example | Export Excel to HTML Times New Roman | Change default font in Aspose.Cells HTML output | C# Aspose.Cells HTML export font setting
// Developer Intent: Configure Aspose.Cells to use Times New Roman as the fallback font when converting a workbook to HTML.
// Use Cases: Create web‑ready reports from Excel files that match corporate typography. | Publish spreadsheets on intranet portals where a uniform font simplifies CSS maintenance. | Generate HTML snapshots of data sheets for email distribution with a consistent appearance.
// AI Prompts: Write C# code that reads a font name from appsettings.json and applies it as the default font for HTML export using Aspose.Cells. | Explain the impact of HtmlSaveOptions.DefaultFontName on generated CSS and how to override it for individual cells. | Provide a step‑by‑step guide to change the default HTML font to Arial in an Aspose.Cells .NET project.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlDefaultFontDemo
{
    // Shows how to assign "Times New Roman" to HtmlSaveOptions.DefaultFontName and save a Workbook as HTML, ensuring cells without explicit font settings use the specified typeface.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to demonstrate the font rendering
            worksheet.Cells["A1"].PutValue("Sample text with default font");
            worksheet.Cells["A2"].PutValue("Another line of text");

            // Configure HTML save options and set the default font to Times New Roman
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DefaultFontName = "Times New Roman";

            // Save the workbook as HTML using the configured options (lifecycle save)
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with default font '{htmlOptions.DefaultFontName}'.");
        }
    }
}
