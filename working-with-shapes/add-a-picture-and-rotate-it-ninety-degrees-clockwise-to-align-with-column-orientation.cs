// Title: C# – Add a picture to a worksheet cell and rotate it 90° clockwise using Aspose.Cells
// Description: The sample creates a new Workbook, verifies an image file, inserts the picture into cell B2, sets the picture's RotationAngle to 90 degrees, and saves the workbook as PictureRotated.xlsx.
// Keywords: Aspose.Cells | C# | .NET | insert picture | Excel picture rotation | RotationAngle | add image to cell | rotate image 90 degrees | save workbook | Worksheet.Pictures.Add
// Common Searches: Aspose.Cells rotate picture 90 degrees | how to add image to specific cell in Aspose.Cells C# | set picture rotation angle Aspose.Cells | insert and rotate picture in Excel using .NET | Worksheet.Pictures.Add example
// Developer Intent: Place an image in a target cell and apply a 90‑degree clockwise rotation.
// Use Cases: Show a company logo in the report header, rotated to align with column headings. | Fit a scanned diagram into a narrow column by rotating it for better visual fit. | Add a portrait photo to a worksheet and rotate it to match the page orientation before saving.
// AI Prompts: Write C# code that adds a picture to cell D4 in an Aspose.Cells workbook and rotates it 180°. | Demonstrate how to check for an image file, insert it, and rotate it based on column width using Aspose.Cells for .NET. | Provide a snippet that inserts multiple pictures into different cells, each with its own rotation angle.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureRotationDemo
{
    // The sample creates a new Workbook, verifies an image file, inserts the picture into cell B2, sets the picture's RotationAngle to 90 degrees, and saves the workbook as PictureRotated.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                const string imagePath = "sample.jpg";

                // Verify that the image file exists before adding it to the worksheet
                if (File.Exists(imagePath))
                {
                    // Add a picture to the worksheet at cell B2 (row 1, column 1)
                    int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);

                    // Retrieve the added picture object
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Rotate the picture 90 degrees clockwise
                    picture.RotationAngle = 90;
                }
                else
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Skipping picture insertion.");
                }

                // Save the workbook (lifecycle: save)
                workbook.Save("PictureRotated.xlsx");
                Console.WriteLine("Workbook saved successfully as PictureRotated.xlsx.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
