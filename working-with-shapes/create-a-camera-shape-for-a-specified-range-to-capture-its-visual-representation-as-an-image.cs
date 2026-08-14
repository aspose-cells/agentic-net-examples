// Title: Aspose.Cells for .NET – Add a Camera Shape to Capture a Range and Export as PNG
// Description: Shows how to create a workbook, populate cells, insert a camera shape that captures a specific range (e.g., A1:B2), position the picture, convert the shape to a PNG image using a MemoryStream, write the image to disk, and save the workbook.
// Keywords: Aspose.Cells | C# camera shape | Excel range to image | Picture.Camera method | Export PNG | ToImage | .NET workbook | MemoryStream image | Excel snapshot | Aspose.Cells .NET
// Common Searches: Aspose.Cells add camera shape C# | Export Excel range as PNG Aspose.Cells | Capture worksheet range as picture .NET | Picture.Camera example Aspose | Convert camera shape to image Aspose.Cells
// Developer Intent: Insert a camera shape that captures a defined cell range and generate an image file from that shape.
// Use Cases: Create a visual snapshot of a report section for embedding in PDFs or presentations. | Generate thumbnail images of data tables for web dashboards or portals. | Automate the export of selected Excel ranges as PNG files for email or API delivery.
// AI Prompts: Provide C# code to add a camera shape for range C3:D10 and export it as a JPEG using Aspose.Cells. | Show how to place a camera picture at row 10, column 5 and programmatically set its width and height. | Explain how to obtain the picture index after adding a camera shape and handle potential exceptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, populate cells, insert a camera shape that captures a specific range (e.g., A1:B2), position the picture, convert the shape to a PNG image using a MemoryStream, write the image to disk, and save the workbook.
class CameraShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in the range to be photographed
        worksheet.Cells["A1"].Value = "Header1";
        worksheet.Cells["B1"].Value = "Header2";
        worksheet.Cells["A2"].Value = 123;
        worksheet.Cells["B2"].Value = 456;

        // Add a camera picture that captures the range A1:B2.
        // The picture will be placed with its top‑left corner at row 5, column 1.
        PictureCollection pictures = worksheet.Pictures;
        int pictureIndex = pictures.Camera(5, 1, "A1:B2");

        // Retrieve the created picture
        Picture cameraPicture = pictures[pictureIndex];

        // Export the camera picture to an image file (PNG)
        using (MemoryStream imageStream = new MemoryStream())
        {
            cameraPicture.ToImage(imageStream, ImageType.Png);
            File.WriteAllBytes("CameraPicture.png", imageStream.ToArray());
        }

        // Save the workbook containing the camera shape
        workbook.Save("CameraDemo.xlsx");
    }
}
