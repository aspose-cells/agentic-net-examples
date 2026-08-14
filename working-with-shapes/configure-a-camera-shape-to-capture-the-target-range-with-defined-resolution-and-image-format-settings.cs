// Title: Aspose.Cells .NET – Add a Camera Shape and Export as 300 DPI PNG (C#)
// Description: Demonstrates how to insert a camera shape that captures a specific cell range, configure ImageOrPrintOptions for PNG format and 300 DPI resolution, render the shape to a memory stream, save the image file, and persist the workbook.
// Keywords: Aspose.Cells camera shape C# | export camera picture PNG | set DPI Aspose.Cells | ImageOrPrintOptions resolution | capture cell range as image | Aspose.Cells .NET image export | high‑resolution workbook snapshot
// Common Searches: Aspose.Cells set camera shape DPI | save camera picture as PNG in C# | render cell range to image Aspose.Cells | configure ImageOrPrintOptions for camera shape | how to export workbook snapshot high resolution
// Developer Intent: Create a camera shape that captures a defined range and export it as a high‑resolution PNG image.
// Use Cases: Embed a crisp snapshot of a report section into a PDF or PowerPoint slide. | Display a table as a PNG on a web page while preserving exact layout and DPI. | Generate high‑resolution images of charts or data blocks for technical documentation.
// AI Prompts: Show me how to change the camera picture export to JPEG with 150 DPI using Aspose.Cells .NET. | Provide C# code to capture a non‑contiguous range with a camera shape and save it as BMP. | Explain how to reuse a single ImageOrPrintOptions instance for multiple camera shapes in the same workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsCameraExample
{
    // Demonstrates how to insert a camera shape that captures a specific cell range, configure ImageOrPrintOptions for PNG format and 300 DPI resolution, render the shape to a memory stream, save the image file, and persist the workbook.
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
            worksheet.Cells["A2"].Value = 123;
            worksheet.Cells["B2"].Value = 456;
            worksheet.Cells["A3"].Value = "Row 3";
            worksheet.Cells["B3"].Value = "Data";

            // Add a camera picture that captures the range A1:B3.
            // Parameters: top‑left row index, column index, and the source range.
            int pictureIndex = worksheet.Pictures.Camera(5, 1, "A1:B3");

            // Retrieve the picture (inherits from Shape) that was just added
            Picture cameraPicture = worksheet.Pictures[pictureIndex];

            // Configure image options: format (PNG) and resolution (300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the camera picture to a memory stream using the configured options
            using (MemoryStream imageStream = new MemoryStream())
            {
                cameraPicture.ToImage(imageStream, imgOptions);

                // Save the image to a file for verification
                File.WriteAllBytes("CameraCapture.png", imageStream.ToArray());
            }

            // Save the workbook that contains the camera shape
            workbook.Save("CameraDemo.xlsx");
        }
    }
}
