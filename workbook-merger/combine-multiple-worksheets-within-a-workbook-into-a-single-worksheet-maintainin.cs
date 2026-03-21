using System;
using Aspose.Cells;

namespace CombineWorksheetsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains multiple worksheets
            Workbook sourceWorkbook = new Workbook("SourceWorkbook.xlsx");

            // Add a new worksheet that will hold the combined data
            int combinedIndex = sourceWorkbook.Worksheets.Add();
            Worksheet combinedSheet = sourceWorkbook.Worksheets[combinedIndex];
            combinedSheet.Name = "Combined";

            // Keep track of the next free row in the combined sheet
            int currentRow = 0;

            // Iterate through all worksheets in the source workbook
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                Worksheet ws = sourceWorkbook.Worksheets[i];

                // Skip the sheet we are using to collect the data
                if (ws.Index == combinedSheet.Index)
                    continue;

                // Determine the last used row in the current worksheet.
                // Here we use column 0 (A) as a reference; adjust if needed.
                Cell lastCellInColumnA = ws.Cells.EndCellInColumn(0);
                int rowsToCopy = lastCellInColumnA.Row + 1; // rows are zero‑based

                // Copy all rows from the source worksheet to the combined worksheet,
                // preserving both values and formatting.
                ws.Cells.CopyRows(ws.Cells, 0, currentRow, rowsToCopy);

                // Optionally add a blank row as a separator between sheets
                currentRow += rowsToCopy + 1;
            }

            // Remove all original worksheets, leaving only the combined one
            for (int i = sourceWorkbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (sourceWorkbook.Worksheets[i].Index != combinedSheet.Index)
                {
                    sourceWorkbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook
            sourceWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}