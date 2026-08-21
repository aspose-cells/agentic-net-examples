// Title: Refresh a Camera Shape in Aspose.Cells for .NET – Re‑create Picture to Display Updated Range
// Description: Demonstrates how to programmatically refresh a camera shape by deleting the existing picture from a worksheet's PictureCollection and inserting a new camera picture for the same range, ensuring the latest cell values are shown before saving the workbook.
// Keywords: Aspose.Cells .NET camera shape refresh | update camera picture programmatically | PictureCollection.Camera API | refresh Excel camera image | C# Aspose.Cells example
// Common Searches: how to refresh a camera shape in Aspose.Cells | Aspose.Cells update camera picture after cell change | C# remove and recreate camera picture Excel | Aspose.Cells refresh range image without reopening workbook
// Developer Intent: Update a camera shape so it reflects the current data in its source range.
// Use Cases: After modifying worksheet data, regenerate the camera picture to keep dashboards accurate. | Automate report generation where embedded camera images must show the latest calculations. | Refresh multiple camera shapes in a sheet after a batch data import.
// AI Prompts: Generate C# code using Aspose.Cells to refresh a camera shape after changing its source cells. | Show how to delete an existing camera picture and add a new one for the same range in Aspose.Cells for .NET. | Explain a method to programmatically update several camera shapes in a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to programmatically refresh a camera shape by deleting the existing picture from a worksheet's PictureCollection and inserting a new camera picture for the same range, ensuring the latest cell values are shown before saving the workbook.
    public class RefreshCameraShapeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a range with initial data
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);

                // Add a camera picture that captures the range A1:B3
                PictureCollection pictures = sheet.Pictures;
                int pictureIndex = pictures.Camera(5, 1, "A1:B3"); // row=5, column=1 (zero‑based)

                // Modify the source range data after the picture has been created
                sheet.Cells["B2"].PutValue(15); // change quantity for Apple
                sheet.Cells["B3"].PutValue(25); // change quantity for Banana

                // Refresh the camera picture:
                // 1. Remove the old picture.
                pictures.RemoveAt(pictureIndex);
                // 2. Re‑add the camera picture for the same range.
                pictureIndex = pictures.Camera(5, 1, "A1:B3");

                // Save the workbook to verify that the picture reflects the updated data
                workbook.Save("RefreshCameraShapeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshCameraShapeDemo.Run();
        }
    }
}
