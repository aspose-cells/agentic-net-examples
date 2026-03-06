using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureLineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Configure the signature line properties
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            Instructions = "Please sign to approve.",
            AllowComments = true,
            ShowSignedDate = true,
            IsLine = true
        };

        // Add the signature line at row 5, column 2 (zero‑based indices)
        Picture addedSignature = shapes.AddSignatureLine(5, 2, signatureLine);

        // Save the workbook as an XLSX file
        workbook.Save("SignatureLineDemo.xlsx");
    }
}