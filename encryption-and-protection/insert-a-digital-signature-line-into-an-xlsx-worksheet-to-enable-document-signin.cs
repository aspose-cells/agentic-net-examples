using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a SignatureLine object and set its properties
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",               // Name of the signer
                Title = "Manager",                 // Title of the signer
                Email = "john.doe@example.com",    // Email of the signer
                Instructions = "Please sign to approve.", // Prompt shown to the user
                AllowComments = true,              // Allow comments when signing
                ShowSignedDate = true,             // Show the date after signing
                IsLine = true                      // Render as a line (not a picture)
            };

            // Add the signature line to the worksheet at row 5, column 2 (zero‑based indices)
            // ShapeCollection.AddSignatureLine returns a Picture object representing the line
            Picture picture = worksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

            // Optionally, adjust the size of the signature line picture
            picture.Width = 200;
            picture.Height = 50;

            // Save the workbook to an XLSX file (lifecycle: save)
            workbook.Save("SignatureLineDemo.xlsx");
        }
    }
}