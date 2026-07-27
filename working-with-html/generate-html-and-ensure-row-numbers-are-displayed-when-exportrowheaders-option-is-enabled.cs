// Title: C# – Export Excel to HTML with Row Numbers Using Aspose.Cells
// Description: Shows how to build a workbook, add sample data, enable ExportRowColumnHeadings in HtmlSaveOptions, and save the worksheet as an HTML file that displays row numbers next to each record.
// Keywords: Aspose.Cells C# HTML export | ExportRowColumnHeadings | row numbers in HTML output | save workbook as HTML | Excel to web page conversion | HTML table with row headers
// Common Searches: Aspose.Cells export HTML with row numbers C# | HtmlSaveOptions ExportRowColumnHeadings example | how to show row headers when saving Excel as HTML | C# generate HTML from Excel workbook Aspose | display row indices in HTML table using Aspose.Cells
// Developer Intent: Create an HTML file from an Excel worksheet that includes visible row indices.
// Use Cases: Generate a product catalog web page where each entry is numbered for easy reference. | Publish Excel data to a website while preserving row headings for accessibility. | Automate HTML report creation from spreadsheets with row numbers to aid navigation and debugging.
// AI Prompts: Provide code to also export column letters as headings together with row numbers using Aspose.Cells. | Show how to write the HTML to a MemoryStream and return it from an ASP.NET Core controller while keeping row numbers. | Explain how to style the automatically generated row‑number column with custom CSS in the HTML output.

using System;
using Aspose.Cells;

// Shows how to build a workbook, add sample data, enable ExportRowColumnHeadings in HtmlSaveOptions, and save the worksheet as an HTML file that displays row numbers next to each record.
class HtmlExportWithRowNumbers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to the worksheet
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(1.2);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(0.8);

        // Configure HTML save options to export row and column headings (row numbers)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportRowColumnHeadings = true; // Enable row numbers in the HTML output

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ProductList.html", saveOptions);

        Console.WriteLine("HTML file saved with row numbers displayed.");
    }
}
