using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Create a new (empty) destination workbook
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Define the zero‑based indices of the rows you want to copy.
        // Example: copy rows 2, 4 and 5 from the source sheet.
        int[] rowsToCopy = new int[] { 1, 3, 4 };

        // Destination row index starts at 0 and increments after each copy.
        int destRowIndex = 0;

        // Copy each selected row individually.
        // Cells.CopyRow copies data and formats of a whole row and is memory‑efficient.
        foreach (int srcRowIndex in rowsToCopy)
        {
            destinationSheet.Cells.CopyRow(sourceSheet.Cells, srcRowIndex, destRowIndex);
            destRowIndex++;
        }

        // Save the new workbook containing only the selected rows.
        destinationWorkbook.Save("selectedRowsCopy.xlsx", SaveFormat.Xlsx);
    }
}

// Author: Example showing how to copy selected rows with minimal memory consumption using Aspose.Cells.