// Title: C# – Insert a Local Image into a Specific Excel Cell with Aspose.Cells
// Description: This example creates a new workbook, verifies that a PNG file exists at a given path, adds the picture to the top‑left corner of a target cell (e.g., C3) using Worksheet.Pictures.Add, sets IsPlacedInCell to true so the image is anchored inside the cell, and saves the file as Output.xlsx.
// Keywords: Aspose.Cells add picture to cell | C# insert image into Excel cell | .NET embed PNG in worksheet | IsPlacedInCell property | save workbook with image
// Common Searches: Aspose.Cells insert image into specific cell C# | how to anchor picture inside Excel cell using Aspose | check image file exists before adding to worksheet Aspose.Cells | place logo in Excel cell programmatically .NET | exception handling when adding picture Aspose.Cells
// Developer Intent: Embed a local image file into a designated Excel cell and persist the workbook.
// Use Cases: Add a company logo to the title cell of an automated report. | Insert product thumbnail pictures into a catalog sheet at exact cell positions. | Place a digital signature image into an approval cell for workflow documents.
// AI Prompts: Generate C# code that loads a JPEG from a file path, inserts it into cell B5 of an Excel worksheet with Aspose.Cells, and ensures the picture is anchored inside the cell. | Show how to verify an image file exists, add it to a worksheet, handle possible exceptions, and save the workbook using Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a new workbook, verifies that a PNG file exists at a given path, adds the picture to the top‑left corner of a target cell (e.g., C3) using Worksheet.Pictures.Add, sets IsPlacedInCell to true so the image is anchored inside the cell, and saves the file as Output.xlsx.
class InsertPictureInCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the target cell (zero‑based indices). Example: cell C3
        int targetRow = 2;    // Row index for C3
        int targetColumn = 2; // Column index for C3

        // Local image file path
        string imagePath = @"C:\Images\sample.png";

        // Verify that the image file exists before attempting to add it
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        try
        {
            // Add the picture to the worksheet at the specified cell's top‑left corner
            int pictureIndex = worksheet.Pictures.Add(targetRow, targetColumn, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Place the picture inside the cell (instead of floating over cells)
            picture.IsPlacedInCell = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error adding picture: {ex.Message}");
            return;
        }

        try
        {
            // Save the workbook
            workbook.Save("Output.xlsx");
            Console.WriteLine("Workbook saved successfully as Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
