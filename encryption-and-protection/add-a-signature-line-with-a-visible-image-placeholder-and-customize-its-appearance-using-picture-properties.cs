using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddSignatureLineWithPlaceholder
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a picture placeholder (no image) at cell B2 (row index 1, column index 1)
        int pictureIndex = sheet.Pictures.Add(1, 1, (string)null);
        Picture picture = sheet.Pictures[pictureIndex];

        // Customize the picture appearance
        picture.BorderLineColor = Color.Blue;   // Set border color
        picture.BorderWeight = 2;               // Border width in points
        picture.Width = 150;                    // Width in pixels
        picture.Height = 50;                    // Height in pixels
        picture.IsPlacedInCell = true;          // Place the picture inside the cell

        // Create and configure a signature line
        SignatureLine signature = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            IsLine = true,
            AllowComments = true,
            ShowSignedDate = true,
            Instructions = "Please sign here"
        };

        // Assign the signature line to the picture
        picture.SignatureLine = signature;

        // Save the workbook
        workbook.Save("SignatureLineWithPlaceholder.xlsx");
    }
}