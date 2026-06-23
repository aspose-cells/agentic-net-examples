using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureAndProtect
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Create and configure a signature line
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.AllowComments = true;
        signatureLine.Email = "example@example.com";
        signatureLine.Instructions = "Sign to confirm the excel content.";
        signatureLine.IsLine = true;
        signatureLine.ShowSignedDate = true;
        signatureLine.Signer = "User";
        signatureLine.Title = "tester";

        // Add the signature line at row 5, column 5 (zero‑based indices)
        shapes.AddSignatureLine(5, 5, signatureLine);

        // Protect the worksheet with a password, locking all protection types
        worksheet.Protect(ProtectionType.All, "MyPassword123", null);

        // Save the workbook
        workbook.Save("SignatureLineProtected.xlsx");
    }
}