// Title: Aspose.Cells C# – Insert picture at D5 and set Placement to Move
// Description: Creates a new workbook, adds an image as a Picture object to cell D5 (row 4, column 3), aligns its UpperLeftCell, sets Placement = Move so the picture follows cell moves without resizing, and saves the file.
// Keywords: Aspose.Cells insert picture | C# picture UpperLeftCell | PlacementType.Move | add image to Excel worksheet | .NET Aspose.Cells picture example | Excel picture move with cell | Aspose.Cells picture positioning
// Common Searches: Aspose.Cells add image to specific cell | set picture placement move with cells Aspose | C# UpperLeftCell picture Aspose.Cells | prevent picture resizing Aspose.Cells | insert logo at D5 using Aspose.Cells
// Developer Intent: Add an image to cell D5 and configure it to move with the cell while keeping its original size.
// Use Cases: Anchor a company logo to a fixed cell so it stays aligned when rows/columns are inserted or deleted. | Attach product thumbnails to rows in a report, ensuring they shift with the data layout. | Place a watermark at a designated cell that follows sheet modifications without scaling.
// AI Prompts: Show C# code to insert a picture at D5 with Aspose.Cells and set its placement to Move. | How do I align a Picture object's UpperLeftCell to D5 and prevent resizing in Aspose.Cells? | Example of checking an image file exists before adding it to an Aspose.Cells worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds an image as a Picture object to cell D5 (row 4, column 3), aligns its UpperLeftCell, sets Placement = Move so the picture follows cell moves without resizing, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "image.jpg";

            // Verify that the image file exists before attempting to add it
            if (File.Exists(imagePath))
            {
                // Add a picture to the worksheet at cell D5 (row 4, column 3)
                int pictureIndex = worksheet.Pictures.Add(4, 3, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Ensure the picture is placed at cell D5 (upper‑left corner)
                picture.Move(4, 3);

                // Set the placement type so the picture moves with the cell but does not resize
                picture.Placement = PlacementType.Move;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
