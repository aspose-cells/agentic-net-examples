using System;
using Aspose.Cells;

class SplitWorkbookToCsv
{
    static void Main()
    {
        // Path to the source workbook (can be .xlsx, .xls, etc.)
        string sourcePath = "input.xlsx";

        // Load the source workbook
        Workbook source = new Workbook(sourcePath);
        Worksheet srcSheet = source.Worksheets[0];
        Cells srcCells = srcSheet.Cells;

        // Get the total number of rows that contain data in the first column (zero‑based index)
        int totalRows = srcCells.GetLastDataRow(0) + 1;

        const int batchSize = 10000; // rows per CSV file
        int partIndex = 0;

        while (partIndex * batchSize < totalRows)
        {
            // Clone the whole workbook so we can trim it without affecting the original
            Workbook part = new Workbook();
            part.Copy(source);

            Worksheet partSheet = part.Worksheets[0];
            Cells partCells = partSheet.Cells;

            int startRow = partIndex * batchSize;

            // Remove rows that appear before the current batch
            if (startRow > 0)
            {
                partCells.DeleteRows(0, startRow);
            }

            // Determine how many rows to keep in this batch
            int rowsToKeep = Math.Min(batchSize, totalRows - startRow);

            // After deleting preceding rows, find the last row that still has data
            int lastRowIndex = partCells.GetLastDataRow(0);
            int rowsAfter = (lastRowIndex + 1) - rowsToKeep;

            // Remove rows that appear after the current batch
            if (rowsAfter > 0)
            {
                partCells.DeleteRows(rowsToKeep, rowsAfter);
            }

            // Save the trimmed workbook as a CSV file (only the active sheet is exported)
            string outFile = $"output_part_{partIndex + 1}.csv";
            part.Save(outFile, SaveFormat.Csv);

            partIndex++;
        }
    }
}