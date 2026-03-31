using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertSignatureLine
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or select the desired one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Approver",
            Email = "john.doe@example.com",
            Instructions = "Please sign to confirm.",
            AllowComments = true,
            ShowSignedDate = true,
            IsLine = true
        };

        // Add the signature line to the worksheet at the desired cell position
        // topRow and leftColumn are zero‑based indices (e.g., row 5, column 2 -> indices 4, 1)
        int topRow = 4;      // Row 5
        int leftColumn = 1;  // Column B
        Picture picture = worksheet.Shapes.AddSignatureLine(topRow, leftColumn, signatureLine);

        // Optionally, adjust the size of the signature line picture
        picture.Width = 150;
        picture.Height = 50;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}