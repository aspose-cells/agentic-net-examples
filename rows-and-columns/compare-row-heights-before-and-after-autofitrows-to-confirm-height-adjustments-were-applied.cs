using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a cell with long wrapped text to cause row height change
            cells["A1"].PutValue("This is a very long piece of text that will require the row to expand when AutoFitRows is applied.");
            Style style = cells["A1"].GetStyle();
            style.IsTextWrapped = true;               // Enable text wrapping
            cells["A1"].SetStyle(style);

            // Record the original row height (in points) before auto-fitting
            double originalHeight = cells.GetRowHeight(0);
            Console.WriteLine($"Original row height (point): {originalHeight}");

            // Perform auto‑fit on all rows in the worksheet
            worksheet.AutoFitRows();

            // Record the new row height after auto‑fit
            double newHeight = cells.GetRowHeight(0);
            Console.WriteLine($"Row height after AutoFitRows (point): {newHeight}");

            // Verify that the height has been adjusted
            bool heightChanged = Math.Abs(newHeight - originalHeight) > 0.01;
            Console.WriteLine($"Height adjusted: {heightChanged}");

            // Optionally, check the IsHeightMatched flag which should be true after auto‑fit
            bool isMatched = worksheet.Cells.Rows[0].IsHeightMatched;
            Console.WriteLine($"IsHeightMatched after auto‑fit: {isMatched}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("RowHeightComparison.xlsx");
        }
    }
}