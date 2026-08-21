// Title: Set Arial as Default Font for HTML Export with Aspose.Cells (C#)
// Description: Learn how to use Aspose.Cells HtmlSaveOptions.DefaultFontName to set Arial as the default font when converting an Excel workbook to HTML in .NET. The sample creates a workbook, adds text, configures the font, and saves the HTML file.
// Keywords: Aspose.Cells HTML export | HtmlSaveOptions DefaultFontName | set default font Arial C# | export Excel to HTML .NET | Aspose.Cells font settings | C# Excel to HTML conversion | Aspose.Cells sample code
// Common Searches: Aspose.Cells set default HTML font | HtmlSaveOptions.DefaultFontName example | C# export workbook to HTML with Arial | how to change default font in Aspose.Cells HTML export | Aspose.Cells HTML export font configuration
// Developer Intent: Configure Aspose.Cells to use Arial as the default font for all text in the generated HTML file.
// Use Cases: Create branded HTML reports where Arial is the corporate typeface. | Generate web‑ready spreadsheets without applying cell‑level styles. | Automate email‑template creation from Excel data with a consistent font.
// AI Prompts: Show a C# snippet that sets HtmlSaveOptions.DefaultFontName to "Arial" and saves a workbook as HTML using Aspose.Cells. | Explain the impact of DefaultFontName on the resulting HTML and how to verify the font in a browser. | Combine DefaultFontName with other HtmlSaveOptions (e.g., ExportImagesAsBase64, PreserveCellBorder) for a full‑featured HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Learn how to use Aspose.Cells HtmlSaveOptions.DefaultFontName to set Arial as the default font when converting an Excel workbook to HTML in .NET. The sample creates a workbook, adds text, configures the font, and saves the HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to demonstrate the font effect
            worksheet.Cells["A1"].PutValue("Hello, World!");
            worksheet.Cells["A2"].PutValue("This text will use Arial as the default HTML font.");

            // Create HTML save options and set the default font to Arial
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.DefaultFontName = "Arial";

            // Save the workbook as HTML using the specified options (lifecycle: save)
            workbook.Save("ExportWithArial.html", saveOptions);

            Console.WriteLine("HTML file saved with default font set to Arial.");
        }
    }
}
