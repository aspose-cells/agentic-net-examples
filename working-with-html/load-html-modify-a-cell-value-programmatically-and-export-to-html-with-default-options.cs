// Title: Load an HTML workbook, edit a cell, and export to HTML with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an existing HTML spreadsheet using Aspose.Cells, modify a specific cell (e.g., A1), apply the default HtmlSaveOptions, and save the workbook back to HTML in a C# console application.
// Keywords: Aspose.Cells HTML load | C# modify cell | HtmlSaveOptions default | export workbook to HTML | .NET spreadsheet editing | programmatic HTML spreadsheet update
// Common Searches: Aspose.Cells edit cell in HTML file C# | Save modified HTML workbook with default options | Load HTML into Aspose.Cells and re‑export | C# change value of A1 in HTML spreadsheet
// Developer Intent: Open an HTML file as a workbook, change a cell value via code, and write the result back to HTML using the library’s default save settings.
// Use Cases: Automatically update header rows in HTML reports before distribution. | Batch‑process multiple HTML spreadsheets to inject a company logo or name into a designated cell. | Expose a web service that receives an HTML spreadsheet, applies data corrections, and returns the revised HTML file.
// AI Prompts: Generate C# code that loads an HTML workbook with Aspose.Cells, sets cell B2 to today's date, and saves the file as HTML using default options. | Explain the purpose of HtmlSaveOptions in Aspose.Cells and how to keep default behavior while customizing specific settings. | Provide a step‑by‑step tutorial for loading an HTML workbook, updating several cells, and exporting the modified workbook to HTML in a .NET console app.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExample
{
    // Demonstrates how to open an existing HTML spreadsheet using Aspose.Cells, modify a specific cell (e.g., A1), apply the default HtmlSaveOptions, and save the workbook back to HTML in a C# console application.
    class Program
    {
        static void Main()
        {
            // Load an existing HTML file into a workbook
            // (Workbook constructor loads the file based on its format)
            Workbook workbook = new Workbook("input.html");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify a cell value programmatically (e.g., cell A1)
            worksheet.Cells["A1"].PutValue("Modified Value");

            // Create default HTML save options (uses the provided constructor rule)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook back to HTML with default options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("HTML file has been saved with the modified cell.");
        }
    }
}
