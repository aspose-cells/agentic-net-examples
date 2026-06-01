using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCountDemo
{
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
                table.Columns.Add("Notes");

                table.Rows.Add("1", "<a href='https://example.com'>Link</a>", "Sample note");
                table.Rows.Add("2", "Plain text", "Another note");
                table.Rows.Add("3", "<b>Bold</b> and <i>Italic</i>", "More notes");

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Import the DataTable into the worksheet, treating values as HTML strings
                worksheet.Cells.ImportData(table, 0, 0, new ImportTableOptions
                {
                    IsFieldNameShown = true,
                    IsHtmlString = true
                });

                // Count cells that contain non‑empty HTML content
                long htmlCellCount = 0;
                foreach (Cell cell in worksheet.Cells)
                {
                    if (cell.Value is string text && !string.IsNullOrWhiteSpace(text))
                    {
                        // Simple HTML detection: presence of '<' and '>'
                        if (text.Contains("<") && text.Contains(">"))
                        {
                            htmlCellCount++;
                        }
                    }
                }

                Console.WriteLine($"Number of cells containing non‑empty HTML content: {htmlCellCount}");

                // Save the workbook (ensure the directory exists)
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
                // General runtime safety
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}