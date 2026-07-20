// Title: Aspose.Cells .NET – Insert a picture, anchor its upper‑left corner to cell D5, and set Placement to Move
// Description: Shows how to create a workbook, check for an image file, add the picture at A1, reposition it so its upper‑left corner lines up with cell D5, configure Placement = Move so the image follows cell moves without resizing, and save the workbook.
// Keywords: Aspose.Cells | C# | add picture to worksheet | anchor picture to cell D5 | Placement Move | picture.Move | PlacementType.Move | image insertion Aspose.Cells | worksheet picture positioning | move with cells
// Common Searches: Aspose.Cells insert image at specific cell | Set picture placement to Move in Aspose.Cells C# | Align picture upper left corner to D5 Aspose.Cells | How to make a picture follow cell movements in .NET | Aspose.Cells picture.Move example
// Developer Intent: Add an image, align it to cell D5, and enable it to move with the cell.
// Use Cases: Place a company logo at D5 so it stays with the header when rows are added or removed. | Insert a diagram next to a data table at D5 that automatically shifts as the table expands. | Add a watermark anchored to D5 that follows sheet modifications without changing size.
// AI Prompts: Generate C# code using Aspose.Cells to insert a PNG image at cell D5 and set its Placement to Move without resizing. | Explain the difference between PlacementType.Move and PlacementType.MoveAndSize in Aspose.Cells and when to use each. | Provide a step‑by‑step guide for checking an image file's existence before adding it as a picture in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, check for an image file, add the picture at A1, reposition it so its upper‑left corner lines up with cell D5, configure Placement = Move so the image follows cell moves without resizing, and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "image.jpg";

            // Add picture only if the file exists
            if (File.Exists(imagePath))
            {
                // Add picture at cell A1 (row 0, column 0)
                int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Move picture so its upper‑left corner aligns with cell D5 (row 4, column 3)
                picture.Move(4, 3);

                // Enable the picture to move with the cell (but not resize)
                picture.Placement = PlacementType.Move;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
