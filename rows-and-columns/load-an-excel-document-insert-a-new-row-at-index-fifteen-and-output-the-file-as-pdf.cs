using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class InsertRowAndConvertToPdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (default sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new row at index 15 (zero‑based). Existing rows 15 and below are shifted down.
        sheet.Cells.InsertRow(15);

        // Save the modified workbook to a temporary file (save rule)
        string tempFile = "modified.xlsx";
        workbook.Save(tempFile);

        // Convert the modified Excel file to PDF (conversion rule)
        string outputPdf = "output.pdf";
        ConversionUtility.Convert(tempFile, outputPdf);

        Console.WriteLine("Row inserted and PDF generated successfully.");
    }
}