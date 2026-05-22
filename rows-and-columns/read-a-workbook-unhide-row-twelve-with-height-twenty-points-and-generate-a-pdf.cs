using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Unhide row 12 (zero‑based index 11) and set its height to 20 points
        workbook.Worksheets[0].Cells.UnhideRow(11, 20);

        // Save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}