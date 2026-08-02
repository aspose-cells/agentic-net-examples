// Title: Fast HTML Export of Large Workbooks with Aspose.Cells Using HtmlCrossType.Cross (C#)
// Description: This example builds a 5,000‑row by 20‑column workbook, sets HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.Cross, optionally adds a page title, and saves the file as a single HTML document. It demonstrates how to achieve high‑speed, low‑memory HTML conversion for massive spreadsheets in C#.
// Keywords: Aspose.Cells HTML export | HtmlCrossType.Cross | HtmlCrossStringType performance | large workbook to HTML C# | high‑speed Excel to HTML conversion | low memory HTML export Aspose | generate HTML from Excel | Aspose.Cells C# example
// Common Searches: Aspose.Cells HtmlCrossType.Cross example | export large Excel file to HTML quickly | HtmlSaveOptions HtmlCrossStringType usage | C# high performance HTML export for big spreadsheets | reduce memory usage when converting Excel to HTML
// Developer Intent: Export a massive workbook to HTML quickly by enabling HtmlCrossType.Cross in Aspose.Cells.
// Use Cases: Create a single‑page HTML report from a 5,000‑row worksheet with minimal memory overhead. | Integrate fast HTML conversion into a web service that processes large Excel uploads. | Generate HTML previews for massive data sets where cell values span multiple columns.
// AI Prompts: Write C# code that streams a 10,000‑row workbook to an HTTP response as HTML using HtmlCrossType.Cross. | Explain the performance benefits of HtmlCrossStringType.Cross and when to prefer other HtmlCrossType values. | Show how to add custom CSS and a header to the HTML output while keeping the export fast for large workbooks.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossDemo
{
    // This example builds a 5,000‑row by 20‑column workbook, sets HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.Cross, optionally adds a page title, and saves the file as a single HTML document. It demonstrates how to achieve high‑speed, low‑memory HTML conversion for massive spreadsheets in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the default constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large amount of data to demonstrate high‑performance HTML export
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    // Put a sample string that will span across cells
                    sheet.Cells[row, col].PutValue($"Row{row}_Col{col}");
                }
            }

            // Create HTML save options (uses the HtmlSaveOptions() constructor rule)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the cross‑cell string type to Cross for faster large‑file generation
            // (uses the HtmlCrossStringType property rule)
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            // Optional: set a page title for the generated HTML
            htmlOptions.PageTitle = "Large Workbook Export with HtmlCrossType.Cross";

            // Save the workbook as HTML using the configured options
            // (uses the Workbook.Save method with options)
            workbook.Save("LargeWorkbook.html", htmlOptions);

            Console.WriteLine("HTML file generated with HtmlCrossType.Cross for high‑performance export.");
        }
    }
}
