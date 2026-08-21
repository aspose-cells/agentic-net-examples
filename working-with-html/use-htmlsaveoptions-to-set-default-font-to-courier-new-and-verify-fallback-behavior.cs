// Title: Set a Default Fallback Font with Aspose.Cells HtmlSaveOptions (C#) and Verify HTML Output
// Description: Demonstrates how to assign "Courier New" as the DefaultFontName in HtmlSaveOptions, export a workbook containing a missing font to HTML, and programmatically confirm that the fallback font appears in the generated markup.
// Keywords: Aspose.Cells HtmlSaveOptions | DefaultFontName | fallback font HTML export | C# .NET | Courier New | missing font handling | font substitution Aspose.Cells | verify HTML output | unit test HTML conversion | Aspose.Cells HTML conversion
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions DefaultFontName example C# | how to use fallback font with Aspose.Cells HTML | verify font substitution in Aspose.Cells HTML output | C# code to check default font in saved HTML
// Developer Intent: Define a fallback font for HTML conversion and confirm its presence in the saved file.
// Use Cases: Ensure consistent typography in HTML reports when source workbooks reference unavailable fonts. | Create automated tests that validate font fallback behavior after HTML export. | Generate web‑ready spreadsheets on servers lacking the original fonts.
// AI Prompts: Write C# code that sets HtmlSaveOptions.DefaultFontName to "Courier New" and checks the resulting HTML for that font name. | Explain Aspose.Cells font resolution during HTML export and how DefaultFontName influences fallback selection. | Provide a C# unit‑test that saves a workbook with a non‑existent font to HTML and asserts the fallback font is used.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlDefaultFontDemo
{
    // Demonstrates how to assign "Courier New" as the DefaultFontName in HtmlSaveOptions, export a workbook containing a missing font to HTML, and programmatically confirm that the fallback font appears in the generated markup.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put sample text into a cell
            worksheet.Cells["A1"].PutValue("Text with missing font");

            // Apply a style that uses a font which does not exist on the system
            Style missingFontStyle = workbook.CreateStyle();
            missingFontStyle.Font.Name = "NonExistentFont";
            worksheet.Cells["A1"].SetStyle(missingFontStyle);

            // Configure HTML save options with a default fallback font
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DefaultFontName = "Courier New";

            // Save the workbook to HTML using the configured options
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the fallback font appears in the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);
            bool fallbackUsed = htmlContent.Contains("Courier New", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"HTML saved to: {htmlPath}");
            Console.WriteLine($"Fallback font \"Courier New\" applied: {fallbackUsed}");
        }
    }
}
