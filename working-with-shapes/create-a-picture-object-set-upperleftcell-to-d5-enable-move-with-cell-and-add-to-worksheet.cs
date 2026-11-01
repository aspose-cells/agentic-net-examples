// Title: Add a picture to cell D5 and set its placement to Move using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert image at D5 cell and keep it anchored to the cell | how to set picture placement to Move in Aspose.Cells .NET | C# Aspose.Cells add picture with upper‑left corner at a specific cell | Aspose.Cells picture placement options Move versus FreeFloating | insert PNG into worksheet cell using Aspose.Cells C#
// Tags: insert image into worksheet cell Aspose.Cells | configure picture placement Move Aspose.Cells | upper‑left cell anchor for picture C# | add PNG picture to Excel using Aspose.Cells | picture moves with cell Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new Workbook, accesses the first Worksheet, checks for an existing PNG file, adds the picture so its upper‑left corner aligns with cell D5, retrieves the Picture object, sets its Placement property to Move so the image moves with the cell, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "image.png";

            // Verify that the image file exists before loading
            if (File.Exists(imagePath))
            {
                // Define the cell where the picture's upper-left corner will be placed (D5)
                int upperLeftRow = 4;      // Row index for row 5 (zero‑based)
                int upperLeftColumn = 3;   // Column index for column D (zero‑based)

                // Add the picture to the worksheet using the file path
                int pictureIndex = sheet.Pictures.Add(upperLeftRow, upperLeftColumn, imagePath);

                // Retrieve the picture object
                Picture picture = sheet.Pictures[pictureIndex];

                // Set the picture to move with the cell (but not resize)
                picture.Placement = PlacementType.Move;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
