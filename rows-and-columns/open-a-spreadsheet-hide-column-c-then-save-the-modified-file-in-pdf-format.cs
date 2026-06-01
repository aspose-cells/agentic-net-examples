using System;
using Aspose.Cells;

class HideColumnAndSavePdf
{
    static void Main()
    {
        // Paths for the source Excel file and the resulting PDF file
        string sourceFile = "input.xlsx";
        string pdfFile = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(sourceFile);

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index 2)
        worksheet.Cells.HideColumn(2);

        // Save the modified workbook as PDF
        workbook.Save(pdfFile, SaveFormat.Pdf);
    }
}