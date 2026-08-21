// Title: C# – Count non‑empty HTML cells after importing a DataTable with Aspose.Cells
// Description: The sample builds a DataTable with plain text and HTML strings, imports it into an Aspose.Cells workbook using ImportData with IsHtmlString enabled, iterates through every cell, detects strings that contain HTML tags, tallies the matches, and saves the workbook.
// Keywords: Aspose.Cells | C# | HTML cell count | ImportData IsHtmlString | worksheet cell iteration | detect HTML tags | non‑empty HTML | DataTable import | cell value check
// Common Searches: Aspose.Cells count HTML cells C# | How to detect HTML strings in worksheet cells after ImportData | C# iterate cells to find HTML tags Aspose | ImportData IsHtmlString count cells | Aspose.Cells example counting HTML content
// Developer Intent: Find out how many cells contain non‑empty HTML after a DataTable import.
// Use Cases: Verify that HTML strings were imported correctly before rendering. | Create a summary of cells that require HTML rendering or special formatting. | Highlight or filter rows based on the presence of HTML tags in their cells.
// AI Prompts: Generate C# code with Aspose.Cells that counts cells containing HTML tags after importing a DataTable. | Show how to modify the loop to collect cell addresses of HTML‑containing cells instead of only counting them. | Explain how to use Worksheet.Cells.Find or a regular expression to locate HTML content more efficiently.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCountDemo
{
    // The sample builds a DataTable with plain text and HTML strings, imports it into an Aspose.Cells workbook using ImportData with IsHtmlString enabled, iterates through every cell, detects strings that contain HTML tags, tallies the matches, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a sample DataTable with HTML content
                DataTable table = new DataTable();
                table.Columns.Add("ID");
                table.Columns.Add("Content");
                table.Columns.Add("Time");

                table.Rows.Add("1", "<a href='https://www.example.com'>Example Link</a>", "2:30 PM");
                table.Rows.Add("2", "Plain text", "3:45 PM");
                table.Rows.Add("3", "<b>Bold Text</b>", "4:00 PM");

                // Create a new workbook and import the DataTable with HTML flag set
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells.ImportData(table, 0, 0, new ImportTableOptions
                {
                    IsFieldNameShown = true,
                    IsHtmlString = true,
                    // Use empty strings instead of null for non‑nullable entries
                    NumberFormats = new string[] { "", "", "h:mm AM/PM" }
                });

                // Count cells that contain non‑empty HTML
                long htmlCellCount = 0;
                foreach (Cell cell in worksheet.Cells)
                {
                    if (cell.Value != null && cell.Value is string str && !string.IsNullOrWhiteSpace(str))
                    {
                        if (str.Contains("<") && str.Contains(">"))
                        {
                            htmlCellCount++;
                        }
                    }
                }

                Console.WriteLine($"Number of cells containing non‑empty HTML content: {htmlCellCount}");

                // Save the workbook safely
                string outputPath = "HtmlContentCountDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
