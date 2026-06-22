using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureLineExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (or specify by index/name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and set signer name and title
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";      // signer name
        signatureLine.Title = "Software Engineer"; // signer title

        // Add the signature line to the worksheet at the desired cell position
        // Parameters: topRow (zero‑based), leftColumn (zero‑based), signatureLine object
        worksheet.Shapes.AddSignatureLine(2, 2, signatureLine); // places at cell C3

        // Save the workbook to a file
        workbook.Save("SignatureLineDemo.xlsx");
    }
}