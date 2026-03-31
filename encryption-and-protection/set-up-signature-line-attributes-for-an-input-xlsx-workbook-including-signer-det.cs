using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SignatureLineSetup
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and configure its visual and signer attributes
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";                     // Name of the signer
        signatureLine.Title = "Project Manager";              // Signer's title
        signatureLine.Email = "john.doe@example.com";         // Signer's email
        signatureLine.Instructions = "Please sign to approve the document."; // Prompt shown to user
        signatureLine.AllowComments = true;                   // Allow comments with the signature
        signatureLine.ShowSignedDate = true;                  // Display the signed date
        signatureLine.IsLine = true;                          // Render as a line (not a picture)
        signatureLine.SignatureLineType = SignatureType.Default; // Use default signature type

        // Add the signature line to the worksheet at row 5, column 2 (zero‑based indices)
        Picture signaturePicture = worksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Adjust the visual size of the signature placeholder (optional)
        signaturePicture.Width = 150;   // Width in points
        signaturePicture.Height = 50;   // Height in points

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}