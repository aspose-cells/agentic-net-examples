// Title: Add an image to cell C9 and obtain its absolute X‑Y pixel coordinates with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add image to cell C9 and get pixel coordinates | retrieve absolute X and Y position of a picture anchored to a worksheet cell using Aspose.Cells | C# Aspose.Cells picture placement MoveAndSize read upper left row and column | how to log picture location after inserting into Excel with Aspose.Cells .NET | Aspose.Cells get picture absolute position in points or pixels
// Tags: add picture to worksheet cell Aspose.Cells | retrieve picture absolute coordinates Aspose.Cells | picture placement MoveAndSize Aspose.Cells | Aspose.Cells image pixel position | C# insert image anchored to cell

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts an image from a file into cell C9, sets the picture’s placement to MoveAndSize, reads the picture’s upper‑left row and column (absolute cell coordinates) which can be used to calculate its X‑Y pixel position, logs this information to the console, and saves the workbook as an XLSX file.
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

            // Ensure the image file exists before attempting to load it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
            }
            else
            {
                try
                {
                    // Add the picture anchored to cell C9 (row 8, column 2)
                    int pictureIndex;
                    using (FileStream imgStream = File.OpenRead(imagePath))
                    {
                        pictureIndex = sheet.Pictures.Add(8, 2, imgStream);
                    }

                    // Retrieve the added picture
                    Picture addedPic = sheet.Pictures[pictureIndex];

                    // Set the picture to move and size with the cell
                    addedPic.Placement = PlacementType.MoveAndSize;

                    // Output basic location information
                    Console.WriteLine($"Picture added at row {addedPic.UpperLeftRow}, column {addedPic.UpperLeftColumn}");
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Error adding picture: {picEx.Message}");
                }
            }

            // Save the workbook (optional)
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
