// Title: Convert Excel to HTML with a custom TableCssId using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, sets HtmlSaveOptions.TableCssId to a custom identifier, and saves the file as HTML, enabling the generated <table> element to be styled via CSS.
// Keywords: Aspose.Cells | C# | Excel to HTML | HtmlSaveOptions | TableCssId | custom CSS id | HTML export | save workbook as HTML
// Common Searches: Aspose.Cells set TableCssId | C# export Excel to HTML with custom table id | How to add CSS id to HTML table when converting Excel | HtmlSaveOptions TableCssId example | Aspose.Cells HTML export custom CSS identifier
// Developer Intent: Create an HTML version of an Excel workbook and assign a specific CSS id to the output table for targeted styling.
// Use Cases: Generate web reports that rely on a predefined CSS table style by applying a custom id during conversion. | Embed Excel data into existing web pages where CSS rules target a particular table identifier. | Batch‑convert multiple workbooks to HTML while preserving a consistent table id for centralized styling.
// AI Prompts: Write C# code with Aspose.Cells to convert input.xlsx to output.html and set TableCssId='report-table'. | Explain how the TableCssId property changes the HTML markup and how to reference it in a stylesheet. | Provide a CSS snippet that styles the table with id 'report-table' after Aspose.Cells HTML export. | Show how to process a folder of Excel files, exporting each to HTML with the same TableCssId using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, sets HtmlSaveOptions.TableCssId to a custom identifier, and saves the file as HTML, enabling the generated <table> element to be styled via CSS.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and set a custom TableCssId
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "custom-table-style";

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
