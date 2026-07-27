// Title: Export Excel to HTML with Column Letters and Row Numbers using Aspose.Cells for .NET
// Description: Shows how to build a workbook, populate cells, enable HtmlSaveOptions.ExportRowColumnHeadings, and save the sheet as an HTML file that displays Excel‑style column letters (A, B, …) and row indices as table headers.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportRowColumnHeadings | C# | .NET | Excel to HTML conversion | column letters in HTML | row numbers in HTML table | workbook export | HTML table from spreadsheet
// Common Searches: Aspose.Cells export HTML column letters C# | HtmlSaveOptions ExportRowColumnHeadings example | how to include row numbers when saving Excel as HTML | save worksheet as HTML with column headers .NET | generate HTML table with A B C headings from workbook
// Developer Intent: Create an HTML file from a workbook that retains Excel‑style column labels and row numbers for clearer web‑based reporting.
// Use Cases: Web dashboards that need familiar spreadsheet column identifiers for end‑users. | Printable HTML snapshots of spreadsheets that preserve the original A‑B‑C column scheme. | Embedding spreadsheet previews in web applications where navigation relies on row/column headings.
// AI Prompts: Provide C# code to export a worksheet to HTML with column letters and customize the table’s CSS. | Show how to export multiple worksheets to separate HTML files while keeping row and column headings. | Explain the effect of ExportRowColumnHeadings and how to turn it off for a minimalist table layout.

using System;
using Aspose.Cells;

// Shows how to build a workbook, populate cells, enable HtmlSaveOptions.ExportRowColumnHeadings, and save the sheet as an HTML file that displays Excel‑style column letters (A, B, …) and row indices as table headers.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to the worksheet
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(15);

        // Configure HTML save options to include row and column headings (e.g., A, B, C...)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportRowColumnHeadings = true; // Enables column letters and row numbers in the HTML table

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output_with_column_headers.html", saveOptions);
    }
}
