using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet's cells collection
        Cells cells = workbook.Worksheets[0].Cells;

        // Hide rows 10 through 20 (1‑based). Zero‑based start index is 9, total rows = 11.
        cells.HideRows(9, 11);

        // Unhide rows 15 through 18 (1‑based). Zero‑based start index is 14, total rows = 4.
        // Set a positive height (e.g., 15 points) to adjust the row height when unhiding.
        cells.UnhideRows(14, 4, 15.0);

        // Save the modified workbook as PDF.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}