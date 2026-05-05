using System;
using Aspose.Cells;

namespace AsposeCellsGroupBeforeRange
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the range block where new cells will be inserted
            // For example, insert a block covering rows 5-7 and columns B-D (zero‑based indices)
            CellArea rangeBlock = new CellArea
            {
                StartRow = 4,   // Row 5 (zero‑based)
                EndRow = 6,     // Row 7
                StartColumn = 1, // Column B
                EndColumn = 3   // Column D
            };

            // Insert a GROUP marker (row grouping) immediately before the range block
            // Group the rows that are just above the range block (rows 2‑4 in this example)
            int groupFirstRow = rangeBlock.StartRow - 3; // adjust as needed
            int groupLastRow = rangeBlock.StartRow - 1;
            if (groupFirstRow >= 0 && groupLastRow >= groupFirstRow)
            {
                // Group rows without hiding them
                cells.GroupRows(groupFirstRow, groupLastRow, false);
            }

            // Insert the range block, shifting existing cells down
            cells.InsertRange(rangeBlock, ShiftType.Down);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}