using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Get the second worksheet (zero‑based index)
        Worksheet sheet = workbook.Worksheets[1];

        // Unhide all columns in this sheet using the default column width.
        // Width = -1 tells Aspose.Cells to apply the standard width.
        sheet.Cells.UnhideColumns(0, 16384, -1);

        // Save the modified workbook as a PDF document.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}