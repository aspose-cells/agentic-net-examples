using System;
using Aspose.Cells;

namespace AsposeCellsRangeReplaceDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the range where the search and replace will be performed (A1:C10)
            int startRow = 0;      // Row index for A1 (zero‑based)
            int startColumn = 0;   // Column index for A1 (zero‑based)
            int endRow = 9;        // Row index for row 10
            int endColumn = 2;     // Column index for column C

            // The text to find and its replacement
            string oldText = "OldValue";
            string newText = "NewValue";

            // Iterate through each cell in the defined range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains the target string (case‑sensitive)
                    if (cell.Type == CellValueType.IsString && cell.StringValue == oldText)
                    {
                        // Replace the cell's value with the new string
                        cell.PutValue(newText);
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}