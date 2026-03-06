using System;
using Aspose.Cells;

namespace AsposeCellsRangeMoveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (create rule is used internally by the Load constructor)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the source range to move (e.g., A1:B2)
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,      // Row index for A (zero‑based)
                StartColumn = 0,   // Column index for 1 (zero‑based)
                EndRow = 1,        // Row index for B (zero‑based)
                EndColumn = 1      // Column index for 2 (zero‑based)
            };

            // Destination start row and column (e.g., move down 2 rows and right 3 columns)
            int destRow = sourceArea.StartRow + 2;    // Row 2 (C‑based index 3)
            int destColumn = sourceArea.StartColumn + 3; // Column 3 (C‑based index D)

            // Move the defined range to the new location using the Cells.MoveRange method
            cells.MoveRange(sourceArea, destRow, destColumn);

            // Save the modified workbook (save rule)
            workbook.Save("output.xlsx");
        }
    }
}