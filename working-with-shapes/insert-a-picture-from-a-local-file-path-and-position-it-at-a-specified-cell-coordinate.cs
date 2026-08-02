// Title: Insert a Local Image into a Specific Cell with MoveAndSize using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, validates a local PNG/JPEG file, adds the picture to a target cell (e.g., C5) using zero‑based row/column indices, sets the picture's Placement to MoveAndSize so it follows cell resizing, and saves the file as an XLSX document.
// Keywords: Aspose.Cells insert picture C# | add image to cell Aspose.Cells | PlacementType.MoveAndSize | load local PNG Aspose.Cells | embed picture in worksheet .NET | Aspose.Cells picture placement | C# Excel image insertion | Aspose.Cells picture cell coordinates
// Common Searches: how to add an image to a specific cell with Aspose.Cells | Aspose.Cells C# picture placement MoveAndSize example | insert PNG into Excel cell using Aspose.Cells .NET | validate image file before adding to worksheet Aspose | Aspose.Cells picture at C5 row 5 column 3
// Developer Intent: Place a local image into a designated worksheet cell and bind it to the cell so it moves and resizes with the cell.
// Use Cases: Attach product thumbnails to catalog rows for automated report generation. | Anchor a company logo in a header cell that scales with column width. | Insert employee photos next to data entries in HR spreadsheets.
// AI Prompts: Generate C# code that inserts a JPEG into cell B2 with Aspose.Cells and sets Placement to MoveAndSize. | Provide robust error handling for missing image files when adding pictures to an Aspose.Cells worksheet. | Show how to programmatically adjust picture height and width after placing it in a specific cell using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertPictureInCellApp
{
    // C# example that creates a workbook, validates a local PNG/JPEG file, adds the picture to a target cell (e.g., C5) using zero‑based row/column indices, sets the picture's Placement to MoveAndSize so it follows cell resizing, and saves the file as an XLSX document.
    class InsertPictureInCell
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cell (e.g., C5). Row and column indices are zero‑based.
                int targetRow = 4;   // Row 5
                int targetColumn = 2; // Column C

                // Local image file path
                string imagePath = "example.png";

                // Verify that the image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Add the picture to the worksheet at the specified cell position
                int pictureIndex = worksheet.Pictures.Add(targetRow, targetColumn, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Place the picture inside the cell (move and size with cells)
                picture.Placement = PlacementType.MoveAndSize;

                // Save the workbook
                string outputPath = "PictureInCell.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
