using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class DeleteRowsAndConvertToPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Delete rows 60 through 65 (zero‑based index: 59, total 6 rows)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteRows(59, 6);

        // Save the modified workbook to a temporary file
        string tempPath = "temp_modified.xlsx";
        workbook.Save(tempPath, SaveFormat.Xlsx);

        // Convert the temporary Excel file to PDF
        string pdfPath = "output.pdf";
        ConversionUtility.Convert(tempPath, pdfPath);

        // Optional: clean up the temporary file
        if (System.IO.File.Exists(tempPath))
        {
            System.IO.File.Delete(tempPath);
        }

        Console.WriteLine("Rows deleted and PDF saved to: " + pdfPath);
    }
}