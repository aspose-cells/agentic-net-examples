using System;
using Aspose.Cells;

class InsertPictureFromUrl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // URL of the image to insert
        string imageUrl = "https://example.com/sample.jpg";

        // Add a linked picture at cell R2.
        // Row index for R2 is 1 (zero‑based), column index for R is 17 (A=0).
        // Height and width are set to 100 pixels each (adjust as needed).
        worksheet.Shapes.AddLinkedPicture(1, 17, 100, 100, imageUrl);

        // Save the workbook
        workbook.Save("OutputWithLinkedPicture.xlsx");
    }
}