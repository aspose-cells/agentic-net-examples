// Title: Backup a worksheet by enumerating cells and copying values with Aspose.Cells for .NET
// Description: Demonstrates how to create a backup sheet in the same workbook by determining the used range of the source worksheet, iterating through each cell, copying non‑null values, and saving the file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | copy worksheet data | backup Excel sheet | enumerate cells | used range | cell iteration | programmatic Excel backup | Excel data duplication
// Common Searches: Aspose.Cells copy worksheet to new sheet | How to backup Excel data with Aspose.Cells .NET | Iterate over used range Aspose.Cells C# | Copy cell values between worksheets Aspose.Cells | Create backup sheet in same workbook Aspose.Cells
// Developer Intent: Create a backup copy of a worksheet’s data by looping through each cell and writing the values to a newly added sheet.
// Use Cases: Preserve original data before applying transformations or calculations. | Generate an internal version‑controlled snapshot for audit trails. | Provide a quick undo mechanism by keeping a duplicate of the source sheet within the same workbook.
// AI Prompts: Write C# code using Aspose.Cells that copies all non‑null cell values from a source worksheet to a newly added backup worksheet. | Show an Aspose.Cells .NET example that iterates over the used range of a sheet and duplicates the data into another sheet while preserving cell formatting. | Create a reusable method that accepts a Workbook and a source sheet name, adds a backup sheet, and copies the data cell‑by‑cell with Aspose.Cells.

using Aspose.Cells;
using System;

// Demonstrates how to create a backup sheet in the same workbook by determining the used range of the source worksheet, iterating through each cell, copying non‑null values, and saving the file using Aspose.Cells for C#.
class BackupWorksheet
{
    static void Main()
    {
        try
        {
            // Create a source workbook and add some sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Original";

            sourceSheet.Cells["A1"].PutValue("Name");
            sourceSheet.Cells["B1"].PutValue("Score");
            sourceSheet.Cells["A2"].PutValue("Alice");
            sourceSheet.Cells["B2"].PutValue(85);
            sourceSheet.Cells["A3"].PutValue("Bob");
            sourceSheet.Cells["B3"].PutValue(92);

            // Add a new worksheet that will hold the backup copy
            Worksheet backupSheet = sourceWorkbook.Worksheets.Add("Backup");

            // Get the Cells collections for source and backup sheets
            Cells srcCells = sourceSheet.Cells;
            Cells backupCells = backupSheet.Cells;

            // Determine the used range in the source sheet
            int maxRow = srcCells.MaxDataRow;      // zero‑based index of the last row with data
            int maxCol = srcCells.MaxDataColumn;   // zero‑based index of the last column with data

            // Copy each cell value from source to backup
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var srcCell = srcCells[row, col];
                    if (srcCell != null && srcCell.Type != CellValueType.IsNull)
                    {
                        backupCells[row, col].PutValue(srcCell.Value);
                    }
                }
            }

            // Save the workbook containing both the original and the backup sheet
            sourceWorkbook.Save("BackupDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
