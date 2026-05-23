using System;
using Aspose.Cells;

namespace AsposeCellsMoveRangeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example data and formulas in the source range (E5:G10)
            // This block is optional; remove if the workbook already contains data
            for (int row = 4; row <= 9; row++)          // Rows 5 to 10 (zero‑based index)
            {
                for (int col = 4; col <= 6; col++)      // Columns E to G (zero‑based index)
                {
                    // Put a sample value
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");

                    // Add a simple formula referencing the cell itself (to demonstrate formula preservation)
                    cells[row, col].Formula = $"=SUM({cells[row, col].Name})";
                }
            }

            // Define the source range E5:G10 using zero‑based indices
            CellArea sourceArea = new CellArea
            {
                StartRow = 4,    // Row 5
                StartColumn = 4, // Column E
                EndRow = 9,      // Row 10
                EndColumn = 6    // Column G
            };

            // Destination start cell J5 (row 4, column 9)
            int destRow = 4;    // Row 5
            int destColumn = 9; // Column J

            // Move the range while preserving formulas
            cells.MoveRange(sourceArea, destRow, destColumn);

            // Save the workbook
            workbook.Save("MovedRangeOutput.xlsx");
        }
    }
}