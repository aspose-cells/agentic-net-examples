// Title: Aspose.Cells for .NET – Create a Camera Shape for a Range and Save as PNG
// Description: This C# example shows how to add a camera shape that captures a specific cell range (A1:B2), convert the camera picture to a PNG image, write the image to disk, and save the workbook with the camera shape using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | camera shape | Worksheet.Pictures.Camera | export range as image | save PNG from workbook | convert range to picture
// Common Searches: Aspose.Cells camera shape C# example | How to capture a cell range as an image with Aspose.Cells | Save camera picture to PNG using Aspose.Cells | Export worksheet range to image .NET | Create picture from range Aspose.Cells
// Developer Intent: Generate a camera picture for a defined cell range and export it as an image file.
// Use Cases: Snapshot a table section for inclusion in reports or presentations. | Create thumbnail previews of worksheet areas for dashboards. | Export selected data as PNG for embedding in emails or external documents.
// AI Prompts: Show C# code to create a camera shape for range C3:D10 and export it as a JPEG with Aspose.Cells. | Provide a snippet that adds a border to a camera picture before saving it as an image. | Explain how to reposition and resize a camera shape after it is created using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCameraDemo
{
    // This C# example shows how to add a camera shape that captures a specific cell range (A1:B2), convert the camera picture to a PNG image, write the image to disk, and save the workbook with the camera shape using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data that will be captured by the camera
            worksheet.Cells["A1"].Value = "Header 1";
            worksheet.Cells["B1"].Value = "Header 2";
            worksheet.Cells["A2"].Value = "Row 1 Col 1";
            worksheet.Cells["B2"].Value = "Row 1 Col 2";

            // Obtain the Pictures collection from the worksheet
            PictureCollection pictures = worksheet.Pictures;

            // Create a camera picture that captures the range A1:B2.
            // Parameters: top‑left row index, top‑left column index, source range.
            int pictureIndex = pictures.Camera(5, 1, "A1:B2");

            // Retrieve the created picture object
            Picture cameraPicture = pictures[pictureIndex];

            // Convert the camera picture to an image and save it to a file
            using (MemoryStream imageStream = new MemoryStream())
            {
                cameraPicture.ToImage(imageStream, ImageType.Png);
                File.WriteAllBytes("CameraCapture.png", imageStream.ToArray());
            }

            // Save the workbook containing the camera picture
            workbook.Save("CameraDemo.xlsx");
        }
    }
}
