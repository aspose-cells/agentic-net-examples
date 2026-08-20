// Title: Aspose.Cells for .NET – Export Excel to HTML with Row Numbers (C#)
// Description: Learn how to use Aspose.Cells in C# to save an Excel workbook as HTML while displaying the original row numbers. The example configures HtmlSaveOptions.ExportRowColumnHeadings, adds sample data, and generates an HTML file with row headers for easy reference.
// Keywords: Aspose.Cells HTML export C# | ExportRowColumnHeadings | row numbers in HTML | Excel to HTML with row headers | HtmlSaveOptions Aspose.Cells | .NET spreadsheet to web | display Excel row indices in HTML
// Common Searches: Aspose.Cells export HTML with row numbers C# | HtmlSaveOptions ExportRowColumnHeadings example | how to show row headers when saving Excel as HTML | C# code to generate HTML from Excel with row indices | Aspose.Cells HTML output row and column headings
// Developer Intent: Add Excel row numbers to the HTML file produced by Aspose.Cells.
// Use Cases: Web dashboards that need to reference original Excel row positions. | Technical documentation where row indices help readers locate data quickly. | Audit‑friendly HTML reports that retain Excel row numbering for traceability.
// AI Prompts: Generate C# code using Aspose.Cells to export a workbook to HTML with row numbers and custom CSS classes. | Show how to disable row numbers but keep column headings when saving Excel as HTML with Aspose.Cells. | Provide a sample that iterates through all worksheets and saves each as a separate HTML file with both row and column headings enabled.

using System;
using Aspose.Cells;

// Learn how to use Aspose.Cells in C# to save an Excel workbook as HTML while displaying the original row numbers. The example configures HtmlSaveOptions.ExportRowColumnHeadings, adds sample data, and generates an HTML file with row headers for easy reference.
class HtmlExportWithRowHeaders
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to cells
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(15);

        // Configure HTML save options to export row and column headings (row numbers)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportRowColumnHeadings = true; // Enables row numbers in the HTML output

        // Save the workbook as an HTML file
        workbook.Save("output_with_row_numbers.html", saveOptions);

        Console.WriteLine("HTML file saved with row numbers displayed.");
    }
}
