using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a signature line object and set its properties
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
            // This uses the ShapeCollection.AddSignatureLine method as defined in the API
            worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Save the workbook to an XLSX file
            workbook.Save("SignatureLineDemo.xlsx");
        }
    }
}