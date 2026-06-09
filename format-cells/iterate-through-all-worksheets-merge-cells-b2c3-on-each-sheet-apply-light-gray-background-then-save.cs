using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration (optional)
            workbook.Worksheets.Add(); // Sheet2
            workbook.Worksheets.Add(); // Sheet3

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Merge cells B2:C3 (zero‑based indices: row 1, column 1, 2 rows, 2 columns)
                sheet.Cells.Merge(1, 1, 2, 2);

                // Apply a light gray background to the merged cell (upper‑left cell of the range)
                Style style = sheet.Cells[1, 1].GetStyle();
                style.ForegroundColor = Color.LightGray;
                style.Pattern = BackgroundType.Solid;
                sheet.Cells[1, 1].SetStyle(style);
            }

            // Save the workbook to a file
            workbook.Save("MergedCellsOutput.xlsx");
        }
    }
}