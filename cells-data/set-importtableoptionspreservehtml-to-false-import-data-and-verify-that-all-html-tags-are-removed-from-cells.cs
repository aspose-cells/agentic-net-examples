// Title: How to import a DataTable into an Aspose.Cells worksheet with HTML tags removed by setting PreserveHtml to false (IsHtmlString = true)
// AI Prompts: Generate C# code that uses ImportTableOptions with IsHtmlString = true to import a DataTable into a workbook so that cells contain only plain text. | Provide a loop that checks each imported cell in Aspose.Cells for leftover HTML tags and prints a verification message. | Show how to save the workbook after importing HTML‑containing data as plain text and output the verification results.
// Common Searches: Aspose.Cells C# import DataTable without preserving HTML tags | Set PreserveHtml false in ImportTableOptions example | IsHtmlString true usage to strip HTML during ImportData | Verify that HTML tags are removed from worksheet cells after import | ImportData from DataTable as plain text Aspose.Cells
// Tags: ImportTableOptions IsHtmlString true C# | Aspose.Cells strip HTML on import | ImportData DataTable plain text Aspose.Cells | verify HTML removal worksheet cells | PreserveHtml false Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    // The example creates a DataTable with HTML strings, configures ImportTableOptions with IsHtmlString = true (equivalent to PreserveHtml = false) to strip HTML during import, imports the table into the first worksheet, iterates over the cells to confirm that no HTML tags remain, prints verification messages, and saves the workbook as HtmlImportResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a DataTable with HTML content
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Content", typeof(string));
            table.Columns.Add("Time", typeof(string));

            // Add rows containing HTML tags
            table.Rows.Add(1, "<a href='https://www.example.com'>Example Link</a>", "2:30 PM");
            table.Rows.Add(2, "<b>Bold Text</b> and <i>Italic Text</i>", "3:45 PM");

            // Create a new workbook and get the first worksheet's cells
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure import options:
            // IsHtmlString = true tells Aspose.Cells that the source values contain HTML.
            // The library will parse the HTML and store only the plain text (HTML tags are removed).
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                IsHtmlString = true   // Equivalent to PreserveHtml = false
            };

            // Import the DataTable starting at cell A1 (row 0, column 0)
            cells.ImportData(table, 0, 0, importOptions);

            // Verify that HTML tags have been stripped from the imported cells
            for (int row = 0; row < table.Rows.Count + 1; row++) // +1 for header row
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    string cellValue = cells[row, col].StringValue;
                    bool containsHtmlTag = cellValue.Contains("<") && cellValue.Contains(">");
                    Console.WriteLine($"Cell {cells[row, col].Name}: \"{cellValue}\" " +
                                      (containsHtmlTag ? "(HTML tags still present)" : "(HTML tags removed)"));
                }
            }

            // Save the workbook to verify the result manually if needed
            workbook.Save("HtmlImportResult.xlsx");
        }
    }
}
