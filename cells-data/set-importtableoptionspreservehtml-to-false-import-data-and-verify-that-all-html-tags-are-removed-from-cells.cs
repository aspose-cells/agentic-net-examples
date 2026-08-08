// Title: C# – Import a DataTable into Aspose.Cells with PreserveHtml disabled to remove HTML tags
// Description: Demonstrates how to set ImportTableOptions.PreserveHtml to false, import a DataTable that contains HTML markup, verify that the tags are stripped from the resulting cells, and optionally save the workbook.
// Keywords: Aspose.Cells PreserveHtml false | ImportTableOptions C# | strip HTML tags Aspose.Cells | ImportData DataTable without HTML | Aspose.Cells HTML removal example | C# Excel import plain text | GitHub Aspose.Cells ImportTableOptions
// Common Searches: How to disable HTML preservation when importing a DataTable with Aspose.Cells | Aspose.Cells ImportTableOptions PreserveHtml false C# example | Remove HTML tags from cells during ImportData in Aspose.Cells | Verify HTML removal after importing data into an Excel workbook using Aspose.Cells
// Developer Intent: Import a DataTable into an Excel workbook while ensuring any HTML markup in the source strings is discarded.
// Use Cases: Cleaning user‑generated content that may contain HTML before exporting to Excel. | Generating plain‑text reports from databases where description fields store markup. | Automated validation that imported cells no longer contain '<' or '>' characters.
// AI Prompts: Provide C# code that sets ImportTableOptions.PreserveHtml = false, imports a DataTable with HTML strings into an Aspose.Cells workbook, checks a cell for remaining tags, and saves the file. | Explain the difference between ImportTableOptions.IsHtmlString and PreserveHtml, and show which property to use to strip HTML during import. | Write a unit‑test in C# that verifies HTML tags are removed after calling Cells.ImportData with PreserveHtml disabled.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    // Demonstrates how to set ImportTableOptions.PreserveHtml to false, import a DataTable that contains HTML markup, verify that the tags are stripped from the resulting cells, and optionally save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a DataTable containing HTML tags in one of its cells
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Content", typeof(string));
            table.Rows.Add(1, "<a href='https://www.example.com'>Example Link</a>");
            table.Rows.Add(2, "<b>Bold Text</b> and <i>Italic Text</i>");

            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure import options: set IsHtmlString to false so HTML tags are ignored/removed
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // import column headers
                IsHtmlString = false       // do NOT treat cell values as HTML
            };

            // Import the DataTable starting at row 0, column 0
            cells.ImportData(table, 0, 0, importOptions);

            // Verify that HTML tags have been stripped from the imported cells
            // Cell A2 corresponds to the first data row (ID = 1)
            string contentCell = cells["B2"].StringValue; // column B contains the HTML content
            Console.WriteLine("Imported cell value: " + contentCell);

            // Simple check: the value should not contain '<' or '>'
            bool containsHtmlTags = contentCell.Contains("<") || contentCell.Contains(">");
            Console.WriteLine("HTML tags removed: " + (!containsHtmlTags));

            // Save the workbook (optional, just to visualize the result)
            workbook.Save("HtmlImportResult.xlsx");
        }
    }
}
