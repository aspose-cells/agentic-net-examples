using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureLineExample
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create and configure a SignatureLine object
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";
        signatureLine.Email = "john.doe@example.com";
        signatureLine.Instructions = "Please sign to approve the document.";
        signatureLine.IsLine = true;               // Display as a line
        signatureLine.AllowComments = true;       // Allow comments
        signatureLine.ShowSignedDate = true;      // Show the signed date

        // Add the signature line to the worksheet.
        // topRow and leftColumn are zero‑based indices (e.g., 2,2 corresponds to cell C3).
        Picture addedSignature = worksheet.Shapes.AddSignatureLine(2, 2, signatureLine);

        // Save the workbook with the new signature line
        workbook.Save("output.xlsx");
    }
}