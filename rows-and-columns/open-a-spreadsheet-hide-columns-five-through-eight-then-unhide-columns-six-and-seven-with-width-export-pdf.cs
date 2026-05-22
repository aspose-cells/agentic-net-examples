using System;
using Aspose.Cells;

class HideUnhideColumnsAndExportPdf
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns 5 through 8 (zero‑based indices 4,5,6,7)
        int startHideColumn = 4;      // column E
        int hideColumnCount = 4;      // columns E, F, G, H
        cells.HideColumns(startHideColumn, hideColumnCount);

        // Unhide columns 6 and 7 (zero‑based indices 5 and 6) with a specific width
        int startUnhideColumn = 5;    // column F
        int unhideColumnCount = 2;    // columns F and G
        double columnWidth = 15.0;    // desired width in characters
        cells.UnhideColumns(startUnhideColumn, unhideColumnCount, columnWidth);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}