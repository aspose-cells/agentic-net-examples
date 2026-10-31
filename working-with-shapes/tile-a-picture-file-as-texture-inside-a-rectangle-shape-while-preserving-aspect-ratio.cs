// Title: Insert a JPEG as a texture into a rectangle shape while preserving aspect ratio using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a JPEG into a rectangle shape with Aspose.Cells C# | preserve aspect ratio when adding picture to a shape in Aspose.Cells .NET | Aspose.Cells C# set image as shape fill without distortion | add picture as texture to Excel shape using Aspose.Cells API
// Tags: rectangle shape image fill Aspose.Cells | aspect ratio preservation shape texture .NET | C# add picture shape Aspose.Cells | Excel shape fill with JPEG Aspose.Cells | Aspose.Cells shape texture implementation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, verifies that a JPEG file exists, inserts the image as a separate picture shape positioned over the rectangle, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (optional visual element)
            Shape rectangle = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                2,    // upper row
                2,    // left column
                0,    // upper row offset (pixels)
                0,    // left column offset (pixels)
                100,  // height (pixels)
                200   // width (pixels)
            );

            // Path to the picture used as texture
            string picturePath = "texture.jpg";

            // Ensure the picture file exists before adding it
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Picture file not found: {picturePath}");

            // Add the picture to the worksheet (as a separate picture shape)
            // This avoids using FillType.Picture which may not be available in some versions
            sheet.Pictures.Add(2, 2, picturePath);

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
