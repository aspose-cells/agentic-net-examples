using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsScreenshotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that we want to capture
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apples";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Bananas";
            sheet.Cells["B3"].Value = 85;
            sheet.Cells["A4"].Value = "Cherries";
            sheet.Cells["B4"].Value = 60;

            // Get the Pictures collection from the worksheet
            PictureCollection pictures = sheet.Pictures;

            // Use the Camera method to take a picture of the range A1:B4.
            // The picture will be placed with its top‑left corner at row 6, column 1 (zero‑based indexing).
            int pictureIndex = pictures.Camera(5, 0, "A1:B4");

            // Optionally, you can access the created picture to modify its properties
            Picture capturedPicture = pictures[pictureIndex];
            capturedPicture.BorderWeight = 2; // set border thickness
            capturedPicture.BorderLineColor = System.Drawing.Color.Blue;

            // Save the workbook (lifecycle: save)
            workbook.Save("ScreenshotDemo.xlsx");

            Console.WriteLine($"Picture captured at index {pictureIndex} and workbook saved as 'ScreenshotDemo.xlsx'.");
        }
    }
}