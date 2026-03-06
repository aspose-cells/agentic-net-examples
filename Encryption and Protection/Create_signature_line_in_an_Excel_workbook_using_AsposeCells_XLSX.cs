using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a SignatureLine object and set its properties
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",
                Title = "Project Manager",
                Email = "john.doe@example.com",
                AllowComments = true,
                IsLine = true,
                ShowSignedDate = true,
                Instructions = "Please sign to approve the document."
            };

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
            // ShapeCollection.AddSignatureLine returns a Picture object representing the line
            Picture addedSignature = worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Optionally, you can adjust the size or position of the picture if needed
            // addedSignature.Width = 200;
            // addedSignature.Height = 30;

            // Save the workbook to an XLSX file (lifecycle save)
            workbook.Save("SignatureLineDemo.xlsx");
        }
    }
}