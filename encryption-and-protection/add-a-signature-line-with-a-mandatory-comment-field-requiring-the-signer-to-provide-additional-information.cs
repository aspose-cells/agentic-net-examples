using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureLineWithComment
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the signature line
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";                     // Signer's name
        signatureLine.Title = "Approver";                     // Signer's title
        signatureLine.Email = "john.doe@example.com";         // Signer's email
        signatureLine.IsLine = true;                          // Mark as a signature line
        signatureLine.AllowComments = true;                   // Require comments
        signatureLine.Instructions = "Please add your comments before signing."; // Prompt for comments
        signatureLine.ShowSignedDate = true;                  // Show signed date

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Save the workbook
        workbook.Save("SignatureLineWithComments.xlsx");
    }
}