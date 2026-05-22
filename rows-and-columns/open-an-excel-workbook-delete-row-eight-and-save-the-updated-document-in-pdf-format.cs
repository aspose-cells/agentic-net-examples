using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string sourcePath = "input.xlsx";

        // Path where the PDF will be saved
        string outputPath = "output.pdf";

        // Load the workbook from the file (lifecycle: load)
        Workbook workbook = new Workbook(sourcePath);

        // Delete row 8 (zero‑based index 7) from the first worksheet
        workbook.Worksheets[0].Cells.DeleteRow(7);

        // Save the modified workbook as PDF (lifecycle: save)
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}