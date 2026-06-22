using System;
using Aspose.Cells;

namespace FreezePaneDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Desired freeze pane parameters
            int desiredRow = 3;          // zero‑based row index where the pane is split
            int desiredColumn = 3;       // zero‑based column index where the pane is split
            int desiredFrozenRows = 3;   // number of rows to freeze (must be <= desiredRow)
            int desiredFrozenColumns = 3; // number of columns to freeze (must be <= desiredColumn)

            // Retrieve current freeze pane state
            int currentRow, currentColumn, currentFrozenRows, currentFrozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out currentRow, out currentColumn,
                                                       out currentFrozenRows, out currentFrozenColumns);

            // Determine whether the worksheet already has the desired frozen state
            bool alreadyDesired = hasFreeze &&
                                  currentRow == desiredRow &&
                                  currentColumn == desiredColumn &&
                                  currentFrozenRows == desiredFrozenRows &&
                                  currentFrozenColumns == desiredFrozenColumns;

            // Apply FreezePanes only if the current state differs from the desired one
            if (!alreadyDesired)
            {
                worksheet.FreezePanes(desiredRow, desiredColumn, desiredFrozenRows, desiredFrozenColumns);
            }

            // Save the workbook
            workbook.Save("FreezePaneResult.xlsx");
        }
    }
}