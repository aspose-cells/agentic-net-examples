// Title: Backup an Excel worksheet by enumerating cells with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, adds a "BackupCopy" sheet, determines the used range of the first worksheet, iterates through each populated cell, copies values or formulas to the new sheet, and saves the file. Empty cells are skipped and formulas are retained.
// Keywords: Aspose.Cells backup worksheet C# | enumerate cells Aspose.Cells | copy Excel sheet programmatically | preserve formulas Aspose.Cells | iterate used range .NET
// Common Searches: How to copy a worksheet to a new sheet with Aspose.Cells C# | Backup Excel data using Aspose.Cells .NET | Copy cells with formulas to another sheet Aspose.Cells | Programmatic Excel worksheet snapshot C#
// Developer Intent: Create a duplicate of a worksheet’s populated cells in a new sheet while keeping formulas intact.
// Use Cases: Generate versioned backups before applying bulk data transformations. | Create read‑only snapshots for audit trails or reporting. | Clone a sheet for temporary calculations without modifying the original.
// AI Prompts: Write C# code using Aspose.Cells that adds a backup worksheet, loops through the used range of the first sheet, copies each non‑empty cell’s value or formula to the backup sheet, and saves the workbook. | Provide an Aspose.Cells example that uses MaxDataRow and MaxDataColumn, skips empty cells, and preserves formulas when duplicating worksheet data.

using System;
using Aspose.Cells;

namespace AsposeCellsBackupExample
{
    // Loads a workbook, adds a "BackupCopy" sheet, determines the used range of the first worksheet, iterates through each populated cell, copies values or formulas to the new sheet, and saves the file. Empty cells are skipped and formulas are retained.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or any worksheet you want to back up)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet that will hold the backup copy
            Worksheet backupSheet = workbook.Worksheets.Add("BackupCopy");

            // Get the Cells collections for source and backup sheets
            Cells sourceCells = sourceSheet.Cells;
            Cells backupCells = backupSheet.Cells;

            // Determine the used range limits
            int maxRow = sourceCells.MaxDataRow;      // zero‑based index of the last row with data
            int maxCol = sourceCells.MaxDataColumn;   // zero‑based index of the last column with data

            // Enumerate each cell in the used range and copy its content
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell srcCell = sourceCells[row, col];

                    // Skip empty cells
                    if (srcCell == null || srcCell.Type == CellValueType.IsNull)
                        continue;

                    Cell destCell = backupCells[row, col];

                    // Preserve formulas if present; otherwise copy the evaluated value
                    if (!string.IsNullOrEmpty(srcCell.Formula))
                    {
                        destCell.Formula = srcCell.Formula;
                    }
                    else
                    {
                        destCell.PutValue(srcCell.Value);
                    }
                }
            }

            // Save the workbook with the backup sheet
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
