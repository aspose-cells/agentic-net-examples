// Title: Count non‑empty HTML cells after importing a DataTable into an Aspose.Cells worksheet (C#)
// AI Prompts: Write C# code that imports a DataTable with HTML strings into an Aspose.Cells worksheet and returns the number of cells containing HTML tags. | Modify the cell enumeration loop to increment a counter only when the cell's string value includes both '<' and '>' characters. | Add code to log the address (e.g., A1, B2) of each cell that contains HTML content while counting them.
// Common Searches: how to detect HTML strings in worksheet cells using Aspose.Cells for .NET | C# Aspose.Cells import DataTable with HTML and count cells containing tags | enumerate cells in Aspose.Cells workbook and check for non‑empty HTML content | Aspose.Cells count cells that have HTML markup after ImportData | sample code for counting HTML cells in a worksheet with Aspose.Cells C#
// Tags: import DataTable as HTML strings Aspose.Cells | enumerate worksheet cells Aspose.Cells C# | detect HTML markup in cell values Aspose.Cells | count cells containing HTML tags C# | Aspose.Cells cell iteration HTML detection

using System;
using System.Data;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsHtmlCountDemo
{
    // Demonstrates importing a DataTable with HTML strings into an Aspose.Cells worksheet, iterating all cells, counting those whose values contain both '<' and '>' characters, optionally logging their addresses, and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Prepare a DataTable with sample HTML content
            DataTable table = new DataTable();
            table.Columns.Add("ID");
            table.Columns.Add("Content");
            table.Columns.Add("Time");

            table.Rows.Add("1", "<a href='https://www.example.com'>Example Link</a>", "2:30 PM");
            table.Rows.Add("2", "Plain text", "3:45 PM");
            table.Rows.Add("3", "<b>Bold Text</b>", "4:00 PM");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Import the DataTable with HTML strings recognized
            worksheet.Cells.ImportData(table, 0, 0, new ImportTableOptions
            {
                IsFieldNameShown = true,
                IsHtmlString = true,
                NumberFormats = new string[] { null, null, "h:mm AM/PM" }
            });

            // Iterate through all cells and count those containing non‑empty HTML
            int htmlCellCount = 0;
            IEnumerator enumerator = worksheet.Cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Get the cell value as string (if any)
                string value = cell.Value?.ToString();
                // Consider it HTML if it contains at least one opening and closing tag
                if (!string.IsNullOrEmpty(value) && value.Contains("<") && value.Contains(">"))
                {
                    htmlCellCount++;
                }
            }

            Console.WriteLine($"Number of cells containing non‑empty HTML content: {htmlCellCount}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("HtmlContentCountDemo.xlsx");
        }
    }
}
