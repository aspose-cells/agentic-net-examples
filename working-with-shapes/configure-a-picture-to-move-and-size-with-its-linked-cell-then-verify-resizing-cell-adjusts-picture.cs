// Title: Set Picture Placement to MoveAndSize and Verify Automatic Resizing with Linked Cell in Aspose.Cells for .NET
// Description: Shows how to insert an image into a worksheet, apply PlacementType.MoveAndSize, change the column width and row height of the linked cell, and programmatically confirm that the picture scales accordingly using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# picture placement | PlacementType.MoveAndSize | image resize with cell | Excel picture scaling | Aspose.Cells .NET example | cell linked picture | dynamic image sizing | Excel automation C#
// Common Searches: Aspose.Cells move and size picture C# | Set picture placement to MoveAndSize in .NET | Resize cell and picture automatically Aspose.Cells | Link picture size to cell dimensions Excel C# | Verify picture dimensions after cell resize Aspose
// Developer Intent: Configure a picture to move and resize with its underlying cell and validate the size adjustment through code.
// Use Cases: Insert a logo in a header cell that automatically adapts when the row height changes. | Create reports where product images expand or shrink with column width adjustments. | Build Excel templates with embedded graphics that stay proportionate after end‑user edits.
// AI Prompts: Generate C# code that adds a picture to a worksheet, sets PlacementType.MoveAndSize, resizes the target cell, and prints the picture's width and height before and after the change. | Explain how PlacementType.MoveAndSize differs from Move and FreeFloating in Aspose.Cells. | Provide troubleshooting steps when a picture does not resize after changing the linked cell dimensions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPicturePlacementDemo
{
    // Shows how to insert an image into a worksheet, apply PlacementType.MoveAndSize, change the column width and row height of the linked cell, and programmatically confirm that the picture scales accordingly using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the picture file
                string picturePath = "example.jpg";

                // Verify that the picture file exists before adding it
                if (!File.Exists(picturePath))
                {
                    Console.WriteLine($"Error: The picture file \"{picturePath}\" was not found.");
                    return;
                }

                // Add a picture to cell B2 (row index 1, column index 1)
                int pictureIndex = worksheet.Pictures.Add(1, 1, picturePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Configure the picture to move and size with the cells beneath it
                picture.Placement = PlacementType.MoveAndSize;

                // Output original picture dimensions
                Console.WriteLine($"Original picture size: Width={picture.Width} px, Height={picture.Height} px");

                // Resize the linked cell (B2) by changing its column width and row height
                worksheet.Cells.SetColumnWidth(1, 30); // Column B width = 30 characters
                worksheet.Cells.SetRowHeight(1, 40);   // Row 2 height = 40 points

                // After resizing, the picture should have adjusted its size automatically
                Console.WriteLine($"After cell resize picture size: Width={picture.Width} px, Height={picture.Height} px");

                // Save the workbook to verify the result visually if needed
                string outputPath = "PictureMoveAndSizeWithCell.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
