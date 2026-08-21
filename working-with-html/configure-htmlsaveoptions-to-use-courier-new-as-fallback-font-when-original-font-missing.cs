// Title: Aspose.Cells .NET – Use HtmlSaveOptions.DefaultFontName "Courier New" as fallback for missing fonts
// Description: Demonstrates how to configure HtmlSaveOptions.DefaultFontName to "Courier New" so that any unavailable fonts in a workbook are automatically replaced when exporting to HTML with Aspose.Cells for .NET.
// Keywords: Aspose.Cells HtmlSaveOptions default font | fallback font HTML export .NET | Courier New Aspose.Cells | HtmlSaveOptions.DefaultFontName example | export Excel to HTML missing font
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions.DefaultFontName usage | fallback font when saving workbook as HTML | how to replace missing fonts in Aspose.Cells HTML output | use Courier New as default font in Aspose.Cells
// Developer Intent: Configure HtmlSaveOptions so that HTML output uses "Courier New" when the original cell font cannot be found.
// Use Cases: Create HTML reports from Excel files that contain custom fonts not installed on the server. | Build an automated conversion pipeline that guarantees consistent text rendering across browsers by applying a web‑safe fallback font. | Support multi‑user environments where client machines have different font libraries, ensuring uniform appearance of exported HTML.
// AI Prompts: Show me C# code to set HtmlSaveOptions.DefaultFontName to "Courier New" in Aspose.Cells. | How can I export a workbook to HTML with a fallback font for missing typefaces using Aspose.Cells .NET? | Explain the impact of HtmlSaveOptions.DefaultFontName on font handling during HTML conversion.

using System;
using Aspose.Cells;

// Demonstrates how to configure HtmlSaveOptions.DefaultFontName to "Courier New" so that any unavailable fonts in a workbook are automatically replaced when exporting to HTML with Aspose.Cells for .NET.
class HtmlFallbackFontDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample text and assign a font that is unlikely to be installed
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Text with a missing font");
        Style style = workbook.CreateStyle();
        style.Font.Name = "NonExistentFont";
        cell.SetStyle(style);

        // Configure HtmlSaveOptions to use "Courier New" when the original font is not found
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DefaultFontName = "Courier New";

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
