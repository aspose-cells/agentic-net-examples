using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and configure its properties
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

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        Picture addedSignature = worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Save the workbook to an XLSX file
        workbook.Save("SignatureLineDemo.xlsx");
    }
}