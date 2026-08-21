// Title: Export Excel to UTF‑8 HTML using Aspose.Cells for .NET
// Description: Demonstrates how to save a Workbook as an HTML file with UTF‑8 encoding via HtmlSaveOptions, ensuring Japanese and other Unicode characters render correctly.
// Keywords: Aspose.Cells HTML export | UTF-8 encoding C# | HtmlSaveOptions Encoding | multilingual Excel to HTML | .NET spreadsheet to web | Unicode Excel export
// Common Searches: Aspose.Cells export HTML UTF-8 | C# save workbook as HTML with Unicode | HtmlSaveOptions set encoding Aspose | convert Excel to HTML multilingual | UTF-8 HTML output Aspose.Cells
// Developer Intent: Generate an HTML version of an Excel workbook that preserves international characters by applying UTF‑8 encoding.
// Use Cases: Publish Excel‑based reports on multilingual websites without garbled text. | Create email‑ready HTML previews of spreadsheets containing Japanese, Arabic, or other non‑Latin scripts. | Automate batch conversion of Excel templates to web‑friendly HTML for global audiences.
// AI Prompts: Provide C# code that loads an existing .xlsx, sets HtmlSaveOptions.Encoding to UTF‑8, and saves it as HTML with Aspose.Cells. | Explain how to test that the generated HTML correctly displays Japanese characters after applying UTF‑8 encoding. | Show how to add custom CSS to the HTML export while keeping UTF‑8 encoding for Unicode text.

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates how to save a Workbook as an HTML file with UTF‑8 encoding via HtmlSaveOptions, ensuring Japanese and other Unicode characters render correctly.
class ExportExcelToHtmlUtf8
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample data containing international characters
        workbook.Worksheets[0].Cells["A1"].PutValue("こんにちは世界"); // Japanese greeting

        // Initialize HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set the output encoding to UTF‑8 for proper character support
        saveOptions.Encoding = Encoding.UTF8;

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
