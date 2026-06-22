using System;
using Aspose.Cells;

class UnmergeAndCopy
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Retrieve the original value from the merged cell (top‑left cell of the range)
        Cell mergedCell = cells["B2"];               // B2 is row 1, column 1 (zero‑based)
        object originalValue = mergedCell.Value;

        // Unmerge the range B2:C2
        // Option 1: using Cells.UnMerge
        cells.UnMerge(1, 1, 1, 2); // firstRow, firstColumn, totalRows, totalColumns

        // After unmerge, copy the original value to the adjacent cell (C2)
        Cell targetCell = cells["C2"];               // C2 is row 1, column 2
        targetCell.PutValue(originalValue);

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}