using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ConfigureSignatureLine
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";
        signatureLine.Email = "john.doe@example.com";
        signatureLine.Instructions = "Please sign to approve.";
        signatureLine.IsLine = true;
        signatureLine.AllowComments = true;
        signatureLine.ShowSignedDate = true;
        signatureLine.SignatureLineType = SignatureType.Default; // can be Stamp or Custom as needed

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}