// Title: Insert and proportionally scale a picture within a cell range using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, defines a target cell block, validates an image file, adds the picture to that rectangle, locks its aspect ratio, sets PlacementType.MoveAndSize so the image moves and resizes with the cells, and saves the workbook.
// Keywords: Aspose.Cells C# picture insert | scale image to fit cells | lock aspect ratio Aspose.Cells | PlacementType.MoveAndSize | add picture to worksheet .NET | fit image within cell range | Aspose.Cells image handling
// Common Searches: how to add an image to a specific cell range in Aspose.Cells | Aspose.Cells keep picture aspect ratio when resizing | C# insert picture and fit it inside merged cells | Aspose.Cells picture placement MoveAndSize example | scale picture to cell block without distortion
// Developer Intent: Place an image inside a defined cell range and have it automatically resize proportionally as the cells change size.
// Use Cases: Insert a company logo across cells C3:G7 while preserving its proportions. | Add product thumbnails to a table so they expand or shrink with row and column adjustments. | Generate a report where each record includes a picture that fits within a merged cell area.
// AI Prompts: Show C# code to insert a picture into a worksheet range with Aspose.Cells and lock its aspect ratio. | Explain how to set PlacementType.MoveAndSize so an image resizes with its surrounding cells. | Provide an example of retrieving the picture index after adding an image with worksheet.Pictures.Add.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertPictureFitInRange
{
    // Creates a new workbook, defines a target cell block, validates an image file, adds the picture to that rectangle, locks its aspect ratio, sets PlacementType.MoveAndSize so the image moves and resizes with the cells, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cell range (rows 2‑6, columns 2‑6).
                int topRow = 2;      // zero‑based index (row 3 in Excel)
                int leftColumn = 2;  // zero‑based index (column C)
                int bottomRow = 6;   // row 7
                int rightColumn = 6; // column G

                // Verify that the image file exists before adding it.
                string imagePath = "image.jpg";
                if (!File.Exists(imagePath))
                {
                    throw new FileNotFoundException($"Image file not found: {imagePath}");
                }

                // Add the picture to the worksheet within the specified range.
                // This overload stretches the picture to the rectangle defined by the cells.
                int pictureIndex = worksheet.Pictures.Add(topRow, leftColumn, bottomRow, rightColumn, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Lock the aspect ratio so the picture scales proportionally.
                picture.IsAspectRatioLocked = true;

                // Ensure the picture moves and resizes with the cells.
                picture.Placement = PlacementType.MoveAndSize;

                // Save the workbook.
                string outputPath = "PictureFitInRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
