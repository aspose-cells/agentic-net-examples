// Title: Aspose.Cells .NET: Set HtmlSaveOptions.DefaultFontName to "Courier New" for HTML export and verify fallback
// Description: Demonstrates how to configure HtmlSaveOptions.DefaultFontName to "Courier New" so that Aspose.Cells substitutes a missing font with a reliable fallback when saving a workbook as HTML. The example creates a workbook, applies a non‑existent font to a cell, saves to HTML, and checks the generated markup for the fallback font name.
// Keywords: Aspose.Cells HtmlSaveOptions | DefaultFontName | Courier New fallback | HTML export missing font | C# Aspose.Cells example | verify fallback font | .NET spreadsheet to HTML | font substitution Aspose
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions.DefaultFontName usage | fallback font when saving workbook to HTML | how to handle missing fonts in Aspose.Cells HTML output | verify default font in generated HTML Aspose
// Developer Intent: Set a guaranteed fallback font for HTML conversion and confirm that Aspose.Cells applies it when the original cell font is unavailable.
// Use Cases: Ensure consistent HTML rendering across browsers by defining a fallback font. | Replace unsupported or missing fonts in a spreadsheet with a known web‑safe font during export. | Programmatically validate that the fallback font appears in the saved HTML markup.
// AI Prompts: Show C# code that sets HtmlSaveOptions.DefaultFontName to "Courier New" and checks the output HTML for the fallback font. | Generate a unit test in .NET that asserts the saved HTML contains "Courier New" when a cell uses a non‑existent font. | Explain how Aspose.Cells selects a fallback font during HTML conversion and how to customize it with HtmlSaveOptions.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to configure HtmlSaveOptions.DefaultFontName to "Courier New" so that Aspose.Cells substitutes a missing font with a reliable fallback when saving a workbook as HTML. The example creates a workbook, applies a non‑existent font to a cell, saves to HTML, and checks the generated markup for the fallback font name.
class HtmlDefaultFontDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample text to a cell
        worksheet.Cells["A1"].PutValue("Text with a missing font");

        // Apply a style that uses a font that does not exist on the system
        Style missingFontStyle = workbook.CreateStyle();
        missingFontStyle.Font.Name = "NonExistentFont";
        worksheet.Cells["A1"].SetStyle(missingFontStyle);

        // Configure HTML save options to use "Courier New" as the default fallback font
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DefaultFontName = "Courier New";

        // Save the workbook to HTML
        string htmlFile = "output.html";
        workbook.Save(htmlFile, saveOptions);

        // Verify that the fallback font appears in the generated HTML
        string htmlContent = File.ReadAllText(htmlFile);
        bool fallbackUsed = htmlContent.Contains("Courier New");
        Console.WriteLine("Fallback to default font used: " + fallbackUsed);
    }
}
