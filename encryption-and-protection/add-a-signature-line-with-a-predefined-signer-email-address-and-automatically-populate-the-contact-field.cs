using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureLineDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create and configure the signature line
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";                     // predefined signer name
        signatureLine.Email = "john.doe@example.com";         // predefined signer email (contact field)
        signatureLine.Title = "Approver";
        signatureLine.IsLine = true;                          // display as a line
        signatureLine.ShowSignedDate = true;                  // show signed date when signed

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Save the workbook
        workbook.Save("SignatureLineWithEmail.xlsx");
    }
}