// Title: C# – Refresh an Aspose.Cells Camera Shape to Show Updated Range Data
// Description: Demonstrates how to add a camera picture to a worksheet, modify the source cells, and programmatically refresh the camera shape using the Aspose.Cells Camera API so the latest values are displayed. The example saves the workbook before and after the refresh.
// Keywords: Aspose.Cells | C# | .NET | camera picture | refresh camera shape | range snapshot | PictureCollection.Camera | update cell values | dynamic workbook | API example
// Common Searches: Aspose.Cells refresh camera picture after data change | C# create and update camera shape in Excel | How to programmatically refresh a camera shape in Aspose.Cells | Aspose.Cells .NET camera method example | Refresh snapshot of a range using Aspose.Cells
// Developer Intent: Update an existing camera shape so it reflects the current content of its source range without recreating the workbook.
// Use Cases: Generate a sales report, capture a table with a camera shape, change quantities, and refresh the image to show new totals. | Create a dashboard where a pivot table snapshot must stay in sync after data recalculation. | Automate workbook generation that imports data, then refreshes multiple camera pictures to keep all visual snapshots current.
// AI Prompts: Write C# code that refreshes an Aspose.Cells camera picture after modifying cell values without adding a duplicate image. | Explain how to locate an existing camera picture in a worksheet and update its image to reflect changed data. | Show how to replace an old camera picture with a refreshed version while preserving its original position and size.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCameraRefreshDemo
{
    // Demonstrates how to add a camera picture to a worksheet, modify the source cells, and programmatically refresh the camera shape using the Aspose.Cells Camera API so the latest values are displayed. The example saves the workbook before and after the refresh.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate initial data in the range that will be photographed
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);

            // Get the Pictures collection from the worksheet
            PictureCollection pictures = sheet.Pictures;

            // Create a camera picture that captures the range A1:B3.
            // The picture will be placed with its top‑left corner at row 5, column 1.
            int pictureIndex = pictures.Camera(5, 1, "A1:B3");
            Console.WriteLine($"Initial camera picture added at index {pictureIndex}.");

            // Save the workbook showing the initial camera picture
            workbook.Save("CameraInitial.xlsx");

            // -----------------------------------------------------------------
            // Update the source range data – the camera picture should reflect this.
            // -----------------------------------------------------------------
            sheet.Cells["B2"].PutValue(15); // Change quantity for Apple
            sheet.Cells["B3"].PutValue(25); // Change quantity for Banana

            // Refresh the camera picture by creating a new one for the same range.
            // This demonstrates the API method to capture the current content.
            int refreshedPictureIndex = pictures.Camera(5, 1, "A1:B3");
            Console.WriteLine($"Refreshed camera picture added at index {refreshedPictureIndex}.");

            // Save the workbook after refreshing the camera picture
            workbook.Save("CameraRefreshed.xlsx");
        }
    }
}
