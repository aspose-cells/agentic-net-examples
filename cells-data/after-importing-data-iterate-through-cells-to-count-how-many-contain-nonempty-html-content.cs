using System;
using System.Data;
using Aspose.Cells;
using System.Collections;

namespace AsposeCellsHtmlCountDemo
{
    class Program
    {
        static void Main()
        {
            // Prepare a DataTable with some HTML content
            DataTable table = new DataTable();
            table.Columns.Add("ID");
            table.Columns.Add("Content");
            table.Columns.Add("Notes");

            table.Rows.Add("1", "<a href='https://example.com'>Link</a>", "Sample note");
            table.Rows.Add("2", "Plain text", "Another note");
            table.Rows.Add("3", "<b>Bold</b> and <i>Italic</i>", "More notes");

            // Create a new workbook and import the DataTable.
            // Set IsHtmlString = true so that HTML strings are kept as‑is.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.ImportData(table, 0, 0, new ImportTableOptions
            {
                IsFieldNameShown = true,
                IsHtmlString = true
            });

            // Iterate through all cells and count those that contain non‑empty HTML.
            int htmlCellCount = 0;
            IEnumerator enumerator = sheet.Cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                if (cell.Value is string text && !string.IsNullOrWhiteSpace(text))
                {
                    // Simple heuristic: presence of '<' and '>' indicates HTML tags.
                    if (text.Contains("<") && text.Contains(">"))
                    {
                        htmlCellCount++;
                    }
                }
            }

            Console.WriteLine($"Number of cells containing HTML content: {htmlCellCount}");

            // Save the workbook (optional, demonstrates standard save usage)
            workbook.Save("HtmlContentDemo.xlsx");
        }
    }
}