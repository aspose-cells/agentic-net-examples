using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the active worksheet (first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide row 10 (zero‑based index 9) and let Aspose.Cells auto‑fit the height
        worksheet.Cells.UnhideRow(9, -1);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}