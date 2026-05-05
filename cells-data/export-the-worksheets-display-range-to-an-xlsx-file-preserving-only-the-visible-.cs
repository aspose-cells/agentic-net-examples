using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExportVisibleCells
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Access the first worksheet (or any worksheet you need)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Determine the used range dimensions
            int maxRow = sourceSheet.Cells.MaxDataRow;      // zero‑based index of last used row
            int maxCol = sourceSheet.Cells.MaxDataColumn;   // zero‑based index of last used column

            // Create a new workbook to hold only the visible data
            Workbook resultWorkbook = new Workbook();
            Worksheet resultSheet = resultWorkbook.Worksheets[0];

            int resultRow = 0;
            // Iterate through rows, copying only visible ones
            for (int r = 0; r <= maxRow; r++)
            {
                if (sourceSheet.Cells.Rows[r].IsHidden) continue;

                int resultCol = 0;
                // Iterate through columns, copying only visible ones
                for (int c = 0; c <= maxCol; c++)
                {
                    if (sourceSheet.Cells.Columns[c].IsHidden) continue;

                    object value = sourceSheet.Cells[r, c].Value;
                    resultSheet.Cells[resultRow, resultCol].PutValue(value);
                    resultCol++;
                }
                resultRow++;
            }

            // Save the result workbook to XLSX format, preserving only visible cells
            resultWorkbook.Save("visible_cells_output.xlsx", SaveFormat.Xlsx);
        }
    }
}