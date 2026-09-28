// Title: How to sync only changed cells, formulas, and formatting from one Excel worksheet to another using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates two worksheets, detects cells with different values or formulas, and copies the updated content and style to the target sheet. | Generate a method that compares each cell in a source workbook to a target workbook and updates only the mismatched cells while preserving formulas and formatting using Aspose.Cells. | Create a C# routine that synchronizes an Excel sheet by copying modified cell values, formulas, and styles from source.xlsx to target.xlsx with Aspose.Cells.
// Common Searches: Aspose.Cells compare two worksheets and update only changed cells in C# | C# copy cell formulas and styles from one Excel file to another using Aspose.Cells | How to synchronize Excel sheets while preserving formatting with Aspose.Cells .NET | Iterate over cells range to detect differences between workbooks Aspose.Cells
// Tags: cell synchronization Aspose.Cells | copy modified cell values Aspose.Cells C# | preserve formulas during worksheet sync Aspose.Cells | transfer cell styles between Excel workbooks Aspose.Cells | compare worksheets max row column Aspose.Cells

using System;
using Aspose.Cells;

// The program loads source.xlsx and target.xlsx, determines the maximum used rows and columns, iterates each cell, compares values (including formulas), and copies only differing values, formulas, and styles from the source worksheet to the target worksheet, then saves the synchronized workbook as target_synced.xlsx.
class SyncWorksheets
{
    static void Main()
    {
        // Load the source and target workbooks
        Workbook sourceWb = new Workbook("source.xlsx");
        Workbook targetWb = new Workbook("target.xlsx");

        // Access the first worksheets (adjust index if needed)
        Worksheet sourceWs = sourceWb.Worksheets[0];
        Worksheet targetWs = targetWb.Worksheets[0];

        // Get the cells collections for both worksheets
        Cells sourceCells = sourceWs.Cells;
        Cells targetCells = targetWs.Cells;

        // Determine the range to iterate (maximum rows and columns present in either sheet)
        int maxRow = Math.Max(sourceCells.MaxRow, targetCells.MaxRow);
        int maxColumn = Math.Max(sourceCells.MaxColumn, targetCells.MaxColumn);

        // Iterate through each cell in the determined range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxColumn; col++)
            {
                // Retrieve the source cell; if it doesn't exist, skip to next cell
                Cell sourceCell = sourceCells[row, col];
                if (sourceCell == null) continue;

                // Retrieve the corresponding target cell (may be null)
                Cell targetCell = targetCells[row, col];

                // Compare the underlying values (including formulas)
                object sourceValue = sourceCell.Value;
                object targetValue = targetCell?.Value;

                // If values differ, copy the source cell's content to the target cell
                if (!object.Equals(sourceValue, targetValue))
                {
                    // Ensure the target cell object exists
                    if (targetCell == null)
                    {
                        targetCell = targetCells[row, col];
                    }

                    // Copy the value
                    targetCell.PutValue(sourceValue);

                    // Copy the formula if present
                    if (!string.IsNullOrEmpty(sourceCell.Formula))
                    {
                        targetCell.Formula = sourceCell.Formula;
                    }

                    // Copy the style to keep formatting consistent
                    targetCell.SetStyle(sourceCell.GetStyle());
                }
            }
        }

        // Save the synchronized target workbook
        targetWb.Save("target_synced.xlsx");
    }
}
