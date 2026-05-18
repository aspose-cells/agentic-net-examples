using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class LinkedPictureExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define picture position and size (in pixels)
        int topRow = 1;      // Row index (0‑based)
        int leftColumn = 1;  // Column index (0‑based)
        int height = 150;    // Height in pixels
        int width = 150;     // Width in pixels

        // URL of the image to link
        string imageUrl = "https://example.com/sample.jpg";

        // Add a linked picture to the worksheet
        Picture linkedPicture = worksheet.Shapes.AddLinkedPicture(topRow, leftColumn, height, width, imageUrl);

        // Ensure the picture is marked as linked (IsLink = true) and data is not embedded
        linkedPicture.IsLink = true;          // Explicitly set, though AddLinkedPicture does this automatically
        if (linkedPicture.Data != null)
        {
            // If data was somehow embedded, clear it to keep the picture linked only
            linkedPicture.Data = null;
        }

        // Optional: verify linking status
        Console.WriteLine("Is linked picture: " + linkedPicture.IsLink);
        Console.WriteLine("Picture data is null (not embedded): " + (linkedPicture.Data == null));

        // Save the workbook
        workbook.Save("LinkedPictureOutput.xlsx");
    }
}