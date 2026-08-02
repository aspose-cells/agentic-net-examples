// Title: Export Excel to HTML with Row/Column Headers and a Custom Table ID – Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds sample data with column headings, and uses Aspose.Cells HtmlSaveOptions to generate an HTML file that shows Excel row and column labels. The table element receives a custom CSS identifier via the TableCssId property, enabling scoped styling.
// Keywords: Aspose.Cells HTML export C# | Export row and column headings | HtmlSaveOptions TableCssId | Excel to HTML with custom CSS id | C# Aspose.Cells sample
// Common Searches: Aspose.Cells include row column headings in HTML | Set custom table id when saving Excel as HTML | HtmlSaveOptions ExportRowColumnHeadings C# example | How to add CSS identifier to Aspose.Cells HTML table | C# export worksheet to HTML with styled table
// Developer Intent: Create an HTML version of a worksheet that displays Excel’s row/column labels and assigns a developer‑defined ID to the generated <table> for targeted styling.
// Use Cases: Web dashboards that need visible spreadsheet coordinates for user reference. | Multi‑tenant portals where each exported table must have a unique CSS scope to prevent style clashes. | Automated UI tests that verify table structure using a known DOM ID.
// AI Prompts: Show how to keep column widths in the HTML output while preserving the custom TableCssId. | Generate a C# snippet that reads the saved HTML file and injects additional CSS rules targeting #custom-table. | Explain how to hide gridlines in the exported HTML but still include row and column headings.

using System;
using Aspose.Cells;

// This C# example creates a workbook, adds sample data with column headings, and uses Aspose.Cells HtmlSaveOptions to generate an HTML file that shows Excel row and column labels. The table element receives a custom CSS identifier via the TableCssId property, enabling scoped styling.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data with column headings
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1.20);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(0.80);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Include row and column headings in the HTML output
        saveOptions.ExportRowColumnHeadings = true;
        // Prefix table related CSS classes/ids with a custom identifier
        saveOptions.TableCssId = "custom-table";

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
