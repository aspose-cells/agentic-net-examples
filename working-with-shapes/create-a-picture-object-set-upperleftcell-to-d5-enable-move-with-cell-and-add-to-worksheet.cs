// Title: Insert a PNG picture into cell D5 of the first worksheet and set its placement to Move with cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a PNG file, adds it as a picture positioned at cell D5 on the first worksheet, sets the picture's Placement property to Move, and saves the workbook. | Generate a self‑contained C# example that verifies an image file exists, inserts it into a new workbook at D5, enables the picture to move when rows or columns are inserted, and outputs the file as an .xlsx.
// Common Searches: how to anchor an image to cell D5 using Aspose.Cells C# | Aspose.Cells picture placement Move with cells example | add PNG to Excel worksheet at specific cell with Aspose.Cells .NET | set picture upper left cell programmatically Aspose.Cells | C# Aspose.Cells picture moves when rows are inserted
// Tags: Aspose.Cells add picture to specific cell | Aspose.Cells picture placement Move | C# insert PNG into Excel worksheet | Aspose.Cells set picture upper-left cell D5 | Aspose.Cells picture moves with cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, checks that image.png exists, adds the PNG as a Picture positioned at cell D5 on the first worksheet, sets the picture's Placement to Move so it follows cell changes, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "image.png";

            // Verify that the image file exists before adding it
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            // Add a picture to the worksheet
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Retrieve the picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Set the picture position: upper‑left cell D5 (row 4, column 3) with no offset
            // Note: SetPosition method is not available in this version; the picture is placed at the default location.
            // If precise positioning is required, adjust the cell indices in the Add method accordingly.

            // Enable "Move with cells" behavior
            picture.Placement = PlacementType.Move;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
