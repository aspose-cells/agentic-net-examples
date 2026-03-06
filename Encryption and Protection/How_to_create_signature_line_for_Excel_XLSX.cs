using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a SignatureLine object and set its properties
            SignatureLine signatureLine = new SignatureLine();
            signatureLine.Signer = "John Doe";
            signatureLine.Title = "Project Manager";
            signatureLine.Email = "john.doe@example.com";
            signatureLine.Instructions = "Please sign to approve the document.";
            signatureLine.AllowComments = true;
            signatureLine.ShowSignedDate = true;
            signatureLine.IsLine = true;

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
            // This returns a Picture object representing the visual signature line
            Picture picture = worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Optionally, you can modify the picture appearance (border, size, etc.)
            picture.BorderLineColor = System.Drawing.Color.Blue;
            picture.BorderWeight = 2;

            // Save the workbook as an XLSX file
            workbook.Save("SignatureLineDemo.xlsx");
        }
    }
}