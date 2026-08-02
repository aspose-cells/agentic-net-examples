using System;
using Aspose.Cells;

namespace AsposeCellsUnhideRowsToPdf
{
    class Program
    {
        static void Main()
        {
            // Load the existing Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Unhide rows 40 to 45 (zero‑based index) and set a custom height (e.g., 20 points)
            // totalRows = 6 because rows 40,41,42,43,44,45 are to be processed
            cells.UnhideRows(40, 6, 20);

            // Export the workbook to PDF format
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}