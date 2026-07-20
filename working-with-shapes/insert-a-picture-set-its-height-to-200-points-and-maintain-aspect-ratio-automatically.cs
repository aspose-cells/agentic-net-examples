// Title: Aspose.Cells .NET – Insert a picture with 200 pt height and locked aspect ratio
// Description: Demonstrates how to create a new Workbook, verify an image file, add the picture to cell A1, set HeightPt to 200 points, enable IsAspectRatioLocked to keep proportions, and save the result as output.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | .NET | C# | insert image Excel | picture height points | lock aspect ratio | HeightPt | IsAspectRatioLocked | Excel automation | worksheet picture | image resizing
// Common Searches: Aspose.Cells add picture to Excel worksheet | set picture height in points Aspose.Cells | lock aspect ratio when inserting image Excel .NET | C# resize picture to fixed height without distortion | how to use HeightPt and IsAspectRatioLocked
// Developer Intent: Add an image to a worksheet, force its height to 200 pt, and preserve the original width‑to‑height ratio.
// Use Cases: Embedding a company logo at a uniform height across generated reports. | Building a template that accepts user photos and displays them with consistent sizing. | Batch‑processing a collection of pictures into spreadsheets while ensuring each retains its proportions.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a picture into cell B2, sets WidthPt to 150, and keeps the aspect ratio locked. | Explain the interaction between HeightPt, WidthPt, and IsAspectRatioLocked in Aspose.Cells picture objects. | Provide robust error‑handling patterns for image insertion in an Aspose.Cells console application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertPictureExample
{
    // Demonstrates how to create a new Workbook, verify an image file, add the picture to cell A1, set HeightPt to 200 points, enable IsAspectRatioLocked to keep proportions, and save the result as output.xlsx using Aspose.Cells for C#.
    class InsertPictureWithAspectRatio
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string imagePath = "image.jpg";

                // Verify that the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add picture at top‑left cell (row 0, column 0)
                    int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Set picture height to 200 points and lock aspect ratio
                    picture.HeightPt = 200;
                    picture.IsAspectRatioLocked = true;
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
