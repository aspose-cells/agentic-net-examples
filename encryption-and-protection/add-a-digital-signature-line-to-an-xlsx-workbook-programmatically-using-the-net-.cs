using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the signature line
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";
        signatureLine.Email = "john.doe@example.com";
        signatureLine.Instructions = "Please sign the document.";
        signatureLine.AllowComments = true;
        signatureLine.ShowSignedDate = true;
        signatureLine.IsLine = true;

        // Add the signature line shape to the worksheet (row 5, column 2)
        Picture signaturePicture = worksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Save the workbook with the signature line
        workbook.Save("WorkbookWithSignatureLine.xlsx");
    }
}