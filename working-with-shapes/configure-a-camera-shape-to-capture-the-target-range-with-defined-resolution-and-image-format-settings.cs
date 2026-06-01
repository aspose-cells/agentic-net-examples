using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class CameraShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data that will be captured by the camera
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apples");
        worksheet.Cells["B2"].PutValue(150);
        worksheet.Cells["A3"].PutValue("Oranges");
        worksheet.Cells["B3"].PutValue(200);

        // Use the Camera method to take a picture of the range A1:B3.
        // The picture will be placed with its top‑left corner at row 5, column 1.
        int pictureIndex = worksheet.Pictures.Camera(5, 1, "A1:B3");

        // Retrieve the picture (which is a Shape) from the collection
        Picture cameraPicture = worksheet.Pictures[pictureIndex];

        // Configure image options: 300 DPI resolution and PNG format
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            HorizontalResolution = 300,
            VerticalResolution = 300,
            ImageType = ImageType.Png
        };

        // Export the captured picture to an image file using the defined options
        cameraPicture.ToImage("CapturedRange.png", imgOptions);

        // Save the workbook (optional, shows the picture inside the sheet as well)
        workbook.Save("CameraShapeDemo.xlsx");
    }
}