using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths to the input Excel file and the output PDF file
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Work with the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Delete column Z (zero‑based index 25)
        cells.DeleteColumn(25);

        // Hide rows 50 to 55 (zero‑based start index 49, total 6 rows)
        cells.HideRows(49, 6);

        // Save the modified workbook as PDF
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}