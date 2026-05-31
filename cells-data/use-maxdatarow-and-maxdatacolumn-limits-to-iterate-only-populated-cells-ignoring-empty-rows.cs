using System;
using Aspose.Cells;

namespace AsposeCellsMaxDataIteration
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (including some empty cells)
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);
            // Row 4 is intentionally left empty to demonstrate skipping

            // Retrieve the maximum data row and column indices once
            int maxRow = sheet.Cells.MaxDataRow;       // zero‑based index of last row containing data
            int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of last column containing data

            Console.WriteLine($"Iterating rows 0..{maxRow}, columns 0..{maxCol}");

            // Iterate only over the populated area
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Process only cells that actually have a value
                    if (cell.Value != null)
                    {
                        Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                    }
                }
            }

            // Example: apply a style to the whole data range using MaxDataRow/MaxDataColumn
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            StyleFlag flag = new StyleFlag { FontBold = true };

            // CreateRange expects the number of rows and columns, so add 1 to include the last index
            sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1).ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("IteratedData.xlsx");
        }
    }
}