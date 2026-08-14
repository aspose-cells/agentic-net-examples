// Title: C# – Insert a picture at cell H12, move it to I13, and verify its position using Aspose.Cells
// Description: Demonstrates how to add a JPEG image to cell H12 in a new workbook, programmatically relocate the picture so its upper‑left corner aligns with cell I13, confirm the new UpperLeftRow and UpperLeftColumn values, and save the file as an XLSX document.
// Keywords: Aspose.Cells picture insert C# | move picture to another cell Aspose.Cells | verify picture position Aspose.Cells | UpperLeftRow UpperLeftColumn | C# spreadsheet image placement | Aspose.Cells picture anchoring
// Common Searches: How to anchor an image to H12 and move it to I13 with Aspose.Cells | Aspose.Cells C# change picture UpperLeftCell | Check picture coordinates after moving in Aspose.Cells | Insert and reposition picture in Excel using Aspose.Cells .NET
// Developer Intent: Add an image to a worksheet at H12, shift it to I13 programmatically, and validate the new cell coordinates.
// Use Cases: Place a logo in a placeholder cell and later move it to the final header location in automated reports. | Ensure dynamically added product photos line up with calculated cells in a catalog worksheet. | Validate layout adjustments after inserting images into generated spreadsheets.
// AI Prompts: Show C# code that inserts a picture at cell H12 with Aspose.Cells, moves it to I13, and verifies UpperLeftRow and UpperLeftColumn. | Provide error handling for missing image files and confirm workbook saving after repositioning the picture.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a JPEG image to cell H12 in a new workbook, programmatically relocate the picture so its upper‑left corner aligns with cell I13, confirm the new UpperLeftRow and UpperLeftColumn values, and save the file as an XLSX document.
class PicturePositionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "sample.jpg";

            // Verify that the image file exists before attempting to add it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert a picture anchored to cell H12 (row index 11, column index 7)
            int pictureIndex = worksheet.Pictures.Add(11, 7, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Move the picture so its upper‑left cell becomes I13 (row index 12, column index 8)
            picture.Move(12, 8);

            // Verify the new position
            int upperRow = picture.UpperLeftRow;       // zero‑based row index
            int upperColumn = picture.UpperLeftColumn; // zero‑based column index

            if (upperRow == 12 && upperColumn == 8)
            {
                Console.WriteLine("Picture successfully moved to cell I13.");
            }
            else
            {
                Console.WriteLine($"Picture position mismatch. Current position: Row {upperRow + 1}, Column {upperColumn + 1}");
            }

            // Save the workbook
            workbook.Save("PicturePositionResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
