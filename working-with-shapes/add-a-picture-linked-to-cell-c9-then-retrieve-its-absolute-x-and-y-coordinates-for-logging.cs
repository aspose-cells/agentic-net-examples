// Title: How to insert a PNG picture into cell C9 and get its absolute X/Y coordinates using Aspose.Cells for .NET
// AI Prompts: Insert a PNG image into worksheet cell C9 and set its Placement to MoveAndSize. | Obtain the picture's absolute X and Y positions (in points) after insertion. | Print the coordinates to the console and save the workbook as output.xlsx.
// Common Searches: Aspose.Cells add image to specific cell and obtain its position | C# get picture X and Y coordinates after inserting into worksheet with Aspose.Cells | link picture to a cell so it moves and resizes in Aspose.Cells .NET | retrieve absolute coordinates of a shape in an Aspose.Cells workbook
// Tags: insert picture into worksheet cell Aspose.Cells | picture placement MoveAndSize Aspose.Cells | picture absolute X Y coordinates .NET | link image to cell C9 Aspose.Cells | log shape position Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Inserts a PNG image into cell C9 of a new workbook, links it to the cell with MoveAndSize placement, logs its absolute X and Y coordinates, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the image file to be inserted
            string imagePath = "sample.png"; // replace with your image file path

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the picture to cell C9 (row index 8, column index 2)
            int pictureIndex = sheet.Pictures.Add(8, 2, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Link the picture to the cell so it moves and sizes with the cell
            picture.Placement = PlacementType.MoveAndSize;

            // Retrieve the absolute X and Y coordinates (in points) of the picture
            double absoluteX = picture.X;
            double absoluteY = picture.Y;

            // Log the coordinates
            Console.WriteLine($"Picture absolute X: {absoluteX}");
            Console.WriteLine($"Picture absolute Y: {absoluteY}");

            // Save the workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
