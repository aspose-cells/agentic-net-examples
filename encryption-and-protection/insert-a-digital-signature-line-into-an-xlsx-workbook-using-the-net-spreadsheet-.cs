using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertSignatureLineDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";
        signatureLine.Email = "john.doe@example.com";
        signatureLine.Instructions = "Please sign to approve the document.";
        signatureLine.AllowComments = true;
        signatureLine.ShowSignedDate = true;
        signatureLine.IsLine = true; // display as a line

        // Add the signature line shape to the worksheet at row 5, column 2 (zero‑based indices)
        // This returns a Picture object that represents the inserted signature line
        Picture picture = worksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Optionally, you can adjust the size of the picture if needed
        picture.Width = 200;
        picture.Height = 50;

        // Save the workbook to an XLSX file
        workbook.Save("WorkbookWithSignatureLine.xlsx");
    }
}