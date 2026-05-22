using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class InsertColumnAndSavePdf
{
    static void Main()
    {
        // Paths for the source workbook, a temporary modified workbook, and the final PDF
        string sourceFile = "input.xlsx";
        string tempFile = "modified.xlsx";
        string pdfFile = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(sourceFile);

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a new column at the second position (index 1) and update references
        cells.InsertColumn(1, true);

        // Copy styles (and data) from the original first column (index 0) to the new column (index 1)
        cells.CopyColumn(cells, 0, 1);

        // Save the modified workbook to a temporary Excel file (required before conversion)
        workbook.Save(tempFile);

        // Convert the temporary Excel file to PDF using the provided ConversionUtility rule
        ConversionUtility.Convert(tempFile, pdfFile);

        // Optional: clean up the temporary file
        try { System.IO.File.Delete(tempFile); } catch { }

        Console.WriteLine("Column inserted, styles copied, and PDF saved successfully.");
    }
}