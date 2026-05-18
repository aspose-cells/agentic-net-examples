using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertPictureWithHyperlink
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Path to the image file (can be a local file or a URL)
        string imagePath = @"C:\Images\sample.jpg";

        // Add a linked picture to the worksheet at row 2, column 2 (zero‑based indexes)
        // Height and width are specified in pixels
        Picture picture = sheet.Shapes.AddLinkedPicture(1, 1, 150, 150, imagePath);

        // Attach a hyperlink to the picture.
        // When the picture is clicked, the linked document will be opened.
        // Use a file URL or a web URL as needed.
        string documentLink = @"file:///C:/Documents/TargetDocument.pdf";
        picture.AddHyperlink(documentLink);

        // Save the workbook
        workbook.Save("Workbook_With_Picture_Hyperlink.xlsx");
    }
}