// Title: Configure HtmlSaveOptions.DefaultFontName to "Courier New" for missing fonts in AspNet Aspose.Cells HTML export (C#)
// Description: This C# example creates a workbook, applies a non‑existent font to a cell, and sets HtmlSaveOptions.DefaultFontName to "Courier New" so the generated HTML uses this font whenever the original font cannot be found, then saves the file as output.html.
// Keywords: Aspose.Cells HtmlSaveOptions | DefaultFontName | fallback font HTML export | C# Aspose.Cells example | missing font handling | Courier New fallback | Excel to HTML conversion
// Common Searches: Aspose.Cells set default font for HTML export | HtmlSaveOptions fallback font C# | use Courier New when font missing Aspose.Cells | how to handle missing fonts in Aspose.Cells HTML output | default font name Aspose.Cells HtmlSaveOptions
// Developer Intent: Define a default font that Aspose.Cells will use in the HTML output when the workbook references fonts that are not installed on the server.
// Use Cases: Ensure consistent appearance of exported HTML reports when source spreadsheets contain custom or unavailable fonts. | Prevent browser warnings about missing fonts by providing a reliable fallback such as Courier New. | Simplify styling of generated HTML by applying a single fallback font across all cells with unknown fonts.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML using "Courier New" as the fallback font for any missing fonts. | Explain the role of HtmlSaveOptions.DefaultFontName and how to configure it for HTML export in Aspose.Cells for .NET. | Provide a complete Aspose.Cells example that sets a default fallback font and also enables CSS embedding and cell formatting preservation.

using System;
using Aspose.Cells;

// This C# example creates a workbook, applies a non‑existent font to a cell, and sets HtmlSaveOptions.DefaultFontName to "Courier New" so the generated HTML uses this font whenever the original font cannot be found, then saves the file as output.html.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample text and assign a font that likely does not exist on the system
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Text with a missing font");
        Style style = workbook.CreateStyle();
        style.Font.Name = "NonExistentFont";
        cell.SetStyle(style);

        // Configure HTML save options to use "Courier New" when the original font is unavailable
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DefaultFontName = "Courier New";

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
