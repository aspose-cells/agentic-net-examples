// Title: Export Excel to HTML with Column Letter Headers using AspNet.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate cells, enable ExportRowColumnHeadings in HtmlSaveOptions, and save the sheet as an HTML file where the table shows column letters (A, B, C…) as headers.
// Keywords: Aspose.Cells | C# | .NET | HTML export | ExportRowColumnHeadings | column letter headers | Excel to HTML | workbook to HTML | table column headings | sample code
// Common Searches: Aspose.Cells export HTML column letters | HtmlSaveOptions ExportRowColumnHeadings example C# | save Excel as HTML with column headers .NET | how to show A B C headers in HTML export Aspose | C# export worksheet to HTML with column headings
// Developer Intent: Generate an HTML representation of an Excel worksheet that includes column letters as table headers.
// Use Cases: Display spreadsheet data on a web page with familiar column identifiers. | Create printable reports where column labels must be visible. | Integrate Excel‑to‑HTML conversion into a .NET web application for documentation purposes.
// AI Prompts: Provide C# code to export a workbook to HTML with column letters and custom header styling using Aspose.Cells. | Show how to export multiple worksheets to separate HTML files while keeping column headings enabled. | Explain how to disable row headings but retain column headings in HtmlSaveOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, populate cells, enable ExportRowColumnHeadings in HtmlSaveOptions, and save the sheet as an HTML file where the table shows column letters (A, B, C…) as headers.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(0.80);

            // Configure HTML save options to include column letters as headers
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportRowColumnHeadings = true   // Enables A, B, C... column headers in the HTML table
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save("ExportWithColumnHeaders.html", htmlOptions);

            Console.WriteLine("HTML file saved with column letter headers.");
        }
    }
}
