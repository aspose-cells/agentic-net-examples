using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureToLastWorksheet
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Optionally add some data or additional worksheets here
        // For demonstration, add two more worksheets so the last one is not the default
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Determine the index of the last worksheet using the worksheet count
        int lastWorksheetIndex = workbook.Worksheets.Count - 1;

        // Get the last worksheet
        Worksheet lastWorksheet = workbook.Worksheets[lastWorksheetIndex];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            AllowComments = true,
            IsLine = true,
            ShowSignedDate = true,
            Instructions = "Please sign to approve."
        };

        // Add the signature line to the worksheet at a specific cell position (row 5, column 2)
        // Shapes.AddSignatureLine expects row and column indices (zero‑based)
        lastWorksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("WorkbookWithSignature.xlsx");
    }
}