// Title: Aspose.Cells for .NET: Move a Picture Shape to the Front (Z‑Order) in an Excel Worksheet
// Description: C# example that creates a workbook, inserts a PNG picture, adds a rectangle shape, then uses the ToFrontOrBack method to bring the picture to the front of the drawing order so it overlays the rectangle. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# | picture front | Z‑order | ToFrontOrBack | overlay shapes | Excel drawing order | move image forward | shape layering | worksheet picture
// Common Searches: Aspose.Cells bring picture to front | How to set Z order of an image in Excel using Aspose.Cells | Move picture above other shapes .NET | Overlay PNG over rectangle Aspose.Cells | ToFrontOrBack method example C#
// Developer Intent: Place a picture on top of all other drawing objects in the generated worksheet.
// Use Cases: Display a company logo over a chart so it stays visible. | Add a watermark image that covers cell borders and other graphics. | Show a photo annotation above highlighted rectangles. | Overlay a badge icon on top of data‑driven shapes. | Ensure a dynamically generated image appears above any existing drawings.
// AI Prompts: Generate C# code that adds several pictures to a worksheet and moves a chosen picture to the front using Aspose.Cells. | Show how to check the Z‑order of shapes after calling ToFrontOrBack in an Aspose.Cells workbook. | Provide an example that places a PNG logo over a rectangle shape and saves the file with Aspose.Cells for .NET. | Write a script that toggles a picture between front and back positions in an Excel sheet using Aspose.Cells. | Explain how to combine ToFrontOrBack with grouping shapes for complex layering scenarios.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExample
{
    // C# example that creates a workbook, inserts a PNG picture, adds a rectangle shape, then uses the ToFrontOrBack method to bring the picture to the front of the drawing order so it overlays the rectangle. The workbook is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the picture file
                string picturePath = "sample.png";

                // Verify that the picture file exists to avoid FileNotFoundException
                if (!File.Exists(picturePath))
                {
                    Console.WriteLine($"Image file '{picturePath}' not found. Ensure the file exists in the execution directory.");
                    return;
                }

                // Add the picture to the worksheet; Pictures.Add returns the index of the picture
                int pictureIndex = worksheet.Pictures.Add(5, 5, picturePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Add a rectangle shape that will be underneath the picture initially
                Shape rectangle = worksheet.Shapes.AddRectangle(6, 6, 100, 100, 0, 0);
                // Optional: set fill color if needed (commented out to avoid API compatibility issues)
                // rectangle.FillColor = Color.LightBlue;

                // Bring the picture to the front so it overlays other shapes
                picture.ToFrontOrBack(10);

                // Save the workbook
                workbook.Save("PictureOnTop.xlsx");
                Console.WriteLine("Workbook saved as 'PictureOnTop.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
