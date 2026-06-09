using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells D5:F5 (zero‑based row 4, column 3, 1 row, 3 columns)
        cells.Merge(4, 3, 1, 3);

        // Set currency number format for the merged cell
        Style style = cells[4, 3].GetStyle();
        style.Number = 164; // Built‑in currency format
        cells[4, 3].SetStyle(style);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}