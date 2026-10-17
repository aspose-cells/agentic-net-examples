// Title: Create a backup worksheet by enumerating cells and copying values and styles with Aspose.Cells for .NET
// AI Prompts: Enumerate the used range of the first worksheet, copy each non‑empty cell's value and formatting to a newly added "Backup" sheet, and save the workbook. | Programmatically add a backup worksheet and duplicate all populated cells with their styles from the source sheet using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# copy non‑empty cells with formatting to a new worksheet | How to create a backup sheet in an Excel file using Aspose.Cells .NET | Iterate over used range and duplicate cell styles Aspose.Cells example | Programmatically backup first worksheet to another sheet with Aspose.Cells
// Tags: enumerate cells Aspose.Cells C# | copy cell values and styles Aspose.Cells | add backup worksheet Aspose.Cells .NET | iterate used range Aspose.Cells | duplicate worksheet data Aspose.Cells

using System;
using Aspose.Cells;

// The program loads an existing workbook, adds a "Backup" worksheet, iterates over the used range of the first sheet, copies non‑null cell values and their styles to the backup sheet, and saves the result as a new file.
class BackupWorksheet
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the source worksheet (first worksheet in the workbook)
        Worksheet sourceSheet = workbook.Worksheets[0];

        // Add a new worksheet that will hold the backup copy
        Worksheet backupSheet = workbook.Worksheets.Add("Backup");

        // References to the cells collections of source and backup sheets
        Cells sourceCells = sourceSheet.Cells;
        Cells backupCells = backupSheet.Cells;

        // Determine the used range to limit enumeration
        int maxRow = sourceCells.MaxDataRow;
        int maxCol = sourceCells.MaxDataColumn;

        // Enumerate each cell in the used range and copy its value and style
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell srcCell = sourceCells[row, col];

                // Skip empty cells to avoid unnecessary objects
                if (srcCell.Type != CellValueType.IsNull)
                {
                    Cell destCell = backupCells[row, col];
                    destCell.PutValue(srcCell.Value);
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }
        }

        // Save the workbook with the new backup sheet
        string outputPath = "output_with_backup.xlsx";
        workbook.Save(outputPath);
    }
}
