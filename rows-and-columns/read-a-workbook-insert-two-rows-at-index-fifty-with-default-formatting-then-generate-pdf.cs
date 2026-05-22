using System;
using Aspose.Cells;

class InsertRowsAndConvertToPdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Load the workbook (create + load)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert two rows at index 50 (zero‑based). This uses the InsertRows method.
        // The inserted rows will inherit the default formatting of the worksheet.
        worksheet.Cells.InsertRows(50, 2);

        // Save the modified workbook as a PDF (save rule)
        string outputPdf = "output.pdf";
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}