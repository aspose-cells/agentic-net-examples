using System;
using Aspose.Cells;

class CopyRowHeightsDemo
{
    static void Main()
    {
        // ---------- Create source workbook and set custom row heights ----------
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Set distinct heights for three rows (0‑based indices)
        sourceSheet.Cells.SetRowHeight(0, 25); // Row 1
        sourceSheet.Cells.SetRowHeight(1, 30); // Row 2
        sourceSheet.Cells.SetRowHeight(2, 35); // Row 3

        // Optional: add some data so rows are visible
        sourceSheet.Cells["A1"].PutValue("Source Row 1");
        sourceSheet.Cells["A2"].PutValue("Source Row 2");
        sourceSheet.Cells["A3"].PutValue("Source Row 3");

        // ---------- Create destination workbook ----------
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Optional: add placeholder data in destination rows
        destinationSheet.Cells["A6"].PutValue("Dest Row 1");
        destinationSheet.Cells["A7"].PutValue("Dest Row 2");
        destinationSheet.Cells["A8"].PutValue("Dest Row 3");

        // ---------- Define source and destination ranges ----------
        int sourceStartRow = 0;   // first row of source range (Row 1)
        int destinationStartRow = 5; // first row of destination range (Row 6)
        int rowsToCopy = 3;       // number of rows whose heights will be copied

        // ---------- Copy row heights ----------
        for (int i = 0; i < rowsToCopy; i++)
        {
            // Get height from source row
            double height = sourceSheet.Cells.GetRowHeight(sourceStartRow + i);
            // Apply the same height to the corresponding destination row
            destinationSheet.Cells.SetRowHeight(destinationStartRow + i, height);
        }

        // ---------- Save the result ----------
        destinationWorkbook.Save("RowHeightsCopied.xlsx");
    }
}