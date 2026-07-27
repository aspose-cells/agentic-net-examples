// Title: Set UTF-8 Encoding and Custom TableCssId for HTML Export in Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, fill cells, and save it as HTML with Aspose.Cells using HtmlSaveOptions configured for UTF-8 encoding and a custom TableCssId (e.g., "custom-table") so the output tables receive predictable CSS class names.
// Keywords: Aspose.Cells HTML export | HtmlSaveOptions Encoding UTF-8 | TableCssId | C# Aspose.Cells | custom CSS class Aspose.Cells | export workbook to HTML .NET | UTF-8 HTML Aspose.Cells
// Common Searches: Aspose.Cells set HTML encoding UTF-8 | Aspose.Cells HtmlSaveOptions TableCssId example | C# export workbook to HTML with custom table CSS | How to change table CSS ID in Aspose.Cells HTML output | UTF-8 HTML export Aspose.Cells .NET
// Developer Intent: Configure HtmlSaveOptions to use UTF-8 encoding and assign a custom TableCssId, producing HTML tables with a consistent CSS selector.
// Use Cases: Create multilingual HTML reports where UTF-8 guarantees correct character rendering and a custom TableCssId enables uniform styling across all tables. | Generate a batch of HTML files that share a single stylesheet by prefixing table class names with a known ID, simplifying front‑end CSS maintenance. | Export Excel data to HTML for web applications while ensuring both proper encoding and predictable CSS selectors for JavaScript or CSS manipulation.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML with Encoding = Encoding.UTF8 and TableCssId = "my-table". | Explain how HtmlSaveOptions.Encoding and HtmlSaveOptions.TableCssId affect the structure and styling of the HTML produced by Aspose.Cells. | Provide step‑by‑step instructions to configure UTF‑8 encoding and a custom TableCssId for HTML export using Aspose.Cells for .NET.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a workbook, fill cells, and save it as HTML with Aspose.Cells using HtmlSaveOptions configured for UTF-8 encoding and a custom TableCssId (e.g., "custom-table") so the output tables receive predictable CSS class names.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to demonstrate the HTML output
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Create HTML save options for HTML format
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Set the encoding to UTF-8 (default is UTF-8, but we set it explicitly)
            htmlOptions.Encoding = Encoding.UTF8;

            // Apply a custom TableCssId to prefix CSS class names for table elements
            htmlOptions.TableCssId = "custom-table";

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("HTML file saved with UTF-8 encoding and TableCssId = \"custom-table\".");
        }
    }
}
