using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create source workbook and populate it with sample data.
        Workbook srcWb = new Workbook();
        Worksheet srcSheet = srcWb.Worksheets[0];
        srcSheet.Name = "Source";

        srcSheet.Cells["A1"].PutValue("Header");
        srcSheet.Cells["A2"].PutValue("Row 1");
        srcSheet.Cells["A3"].PutValue("Row 2");
        srcSheet.Cells["A4"].PutValue("Row 3");

        // Set custom heights for the rows to be copied.
        srcSheet.Cells.Rows[1].Height = 30; // Row 2
        srcSheet.Cells.Rows[2].Height = 40; // Row 3
        srcSheet.Cells.Rows[3].Height = 50; // Row 4

        // Create destination workbook.
        Workbook destWb = new Workbook();
        Worksheet destSheet = destWb.Worksheets[0];
        destSheet.Name = "Destination";

        // Define the range of rows to copy.
        int sourceStartRow = 1; // zero‑based index (Row 2)
        int rowCount = 3;       // rows 2,3,4
        int destStartRow = 0;   // paste starting at Row 1 in destination

        // Copy rows (data, formats, hyperlinks, etc.).
        destSheet.Cells.CopyRows(srcSheet.Cells, sourceStartRow, destStartRow, rowCount);

        // Preserve exact row heights and other settings.
        for (int i = 0; i < rowCount; i++)
        {
            Row srcRow = srcSheet.Cells.Rows[sourceStartRow + i];
            Row destRow = destSheet.Cells.Rows[destStartRow + i];
            destRow.CopySettings(srcRow, false);
        }

        // Auto‑fit rows that do not have custom heights.
        destSheet.AutoFitRows(true);

        // Save both workbooks.
        srcWb.Save("Source.xlsx");
        destWb.Save("Destination.xlsx");
    }
}

// Author: Example demonstrating row copy with height preservation using Aspose.Cells for .NET.