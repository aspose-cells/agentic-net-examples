using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook (XLSX) that will receive the signature line
            string inputPath = "input.xlsx";

            // Path where the modified workbook will be saved
            string outputPath = "output.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Create and configure a SignatureLine object
            SignatureLine signatureLine = new SignatureLine
            {
                AllowComments = true,
                Email = "example@example.com",
                Instructions = "Please sign to confirm the content.",
                IsLine = true,
                ShowSignedDate = true,
                Signer = "John Doe",
                Title = "Manager"
            };

            // Add the signature line to the worksheet at the desired cell position
            // topRow and leftColumn are zero‑based indices (e.g., row 5, column 2 => cell C6)
            int topRow = 5;      // Row index where the signature line starts
            int leftColumn = 2;  // Column index where the signature line starts
            Picture addedSignature = shapes.AddSignatureLine(topRow, leftColumn, signatureLine);

            // Optionally, you can adjust the size of the signature line picture
            addedSignature.Width = 150;
            addedSignature.Height = 50;

            // Save the workbook with the new signature line
            workbook.Save(outputPath);
        }
    }
}