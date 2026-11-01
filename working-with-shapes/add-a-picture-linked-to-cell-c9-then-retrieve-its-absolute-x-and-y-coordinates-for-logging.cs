// Title: Insert an image into cell C9, link it with MoveAndSize, and log its absolute row and column coordinates using Aspose.Cells for .NET
// AI Prompts: Add a picture from a file path to cell C9 and set its Placement property to MoveAndSize with Aspose.Cells. | After linking the picture, read the UpperLeftRow and UpperLeftColumn properties, convert them to 1‑based values, and write the coordinates to the console. | Check whether the image file exists, handle the error gracefully, then save the workbook as Output.xlsx.
// Common Searches: Aspose.Cells C# add image to a specific cell and obtain its position | how to get picture upper left row and column after linking to a cell in Aspose.Cells | C# Aspose.Cells picture placement MoveAndSize example with absolute coordinates | retrieve absolute cell coordinates of a picture inserted in an Aspose.Cells workbook
// Tags: add picture to cell Aspose.Cells | picture placement MoveAndSize Aspose.Cells | retrieve picture absolute coordinates Aspose.Cells | handle missing image file Aspose.Cells | save workbook Output.xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an image file, inserts it into cell C9 of a new worksheet, links the picture to the cell using MoveAndSize placement, logs the picture's upper‑left row and column as 1‑based coordinates, and saves the workbook as Output.xlsx while handling a missing image file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "image.png";

            // Ensure the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add a picture to cell C9 (row index 8, column index 2)
            int targetRow = 8;      // C9 -> row 9 (zero‑based index 8)
            int targetColumn = 2;   // C9 -> column C (zero‑based index 2)

            // Add returns the picture index; retrieve the Picture object
            int pictureIndex = sheet.Pictures.Add(targetRow, targetColumn, imagePath);
            Picture pic = sheet.Pictures[pictureIndex];

            // Link the picture to the cell so it moves/resizes with the cell
            pic.Placement = PlacementType.MoveAndSize;

            // Retrieve cell coordinates (row and column are zero‑based)
            int upperLeftRow = pic.UpperLeftRow;
            int upperLeftColumn = pic.UpperLeftColumn;

            // Log the coordinates
            Console.WriteLine("Picture linked to C9:");
            Console.WriteLine($"  Upper‑left cell: Row {upperLeftRow + 1}, Column {upperLeftColumn + 1}");

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
