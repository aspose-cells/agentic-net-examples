// Title: Insert a PNG picture at cell H12, move it to I13, and verify UpperLeftCell with Aspose.Cells for .NET
// AI Prompts: Add a PNG image to a worksheet at cell H12, set its placement to MoveAndSize, then change its UpperLeftRow and UpperLeftColumn to point to cell I13 and print the new cell address. | Create a new workbook, insert a picture linked to a specific cell, reposition the picture to another cell, and output the updated UpperLeftCell coordinates using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# insert image at specific cell H12 | how to change picture UpperLeftRow and UpperLeftColumn in Aspose.Cells | move picture from H12 to I13 programmatically Aspose.Cells .NET | retrieve cell name from picture coordinates Aspose.Cells example | set picture placement MoveAndSize Aspose.Cells C#
// Tags: insert picture linked to cell Aspose.Cells | picture UpperLeftCell reposition Aspose.Cells | picture placement MoveAndSize .NET | retrieve cell address from picture coordinates Aspose.Cells | save workbook after image reposition Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Demonstrates inserting a PNG picture anchored to cell H12, setting its placement to MoveAndSize, moving it to cell I13 by updating UpperLeftRow/UpperLeftColumn, printing both original and new cell addresses, and saving the workbook as PictureDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file (ensure the file exists at this location)
            string imagePath = "sample.png";

            // Prevent FileNotFoundException for the image
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert picture linked to cell H12 (row index 11, column index 7)
            int row = 11;      // H12 -> row 12 (zero‑based index)
            int column = 7;    // H12 -> column H (zero‑based index)
            int pictureIndex = sheet.Pictures.Add(row, column, imagePath);
            Picture pic = sheet.Pictures[pictureIndex];

            // Set placement so the picture moves and resizes with the cell
            pic.Placement = PlacementType.MoveAndSize;

            // Verify initial UpperLeftCell (should be H12)
            string initialCell = CellsHelper.CellIndexToName(pic.UpperLeftRow, pic.UpperLeftColumn);
            Console.WriteLine("Initial UpperLeftCell: " + initialCell);

            // Change position to I13 (row index 12, column index 8)
            pic.UpperLeftRow = 12;      // I13 -> row 13 (zero‑based)
            pic.UpperLeftColumn = 8;    // I13 -> column I (zero‑based)

            // Verify the new position
            string newCell = CellsHelper.CellIndexToName(pic.UpperLeftRow, pic.UpperLeftColumn);
            Console.WriteLine("New UpperLeftCell: " + newCell);
            Console.WriteLine("UpperLeftRow (zero‑based): " + pic.UpperLeftRow);
            Console.WriteLine("UpperLeftColumn (zero‑based): " + pic.UpperLeftColumn);

            // Save the workbook
            string outputPath = "PictureDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
