// Title: Insert a PNG picture into cell A1 with MoveAndSize placement and save the workbook as XLSX using Aspose.Cells for .NET
// Common Searches: how to insert an image into a specific cell using Aspose.Cells C# | Aspose.Cells C# place picture in cell A1 with MoveAndSize | save workbook after adding picture with Aspose.Cells .NET | C# add PNG to Excel worksheet and anchor to cell using Aspose.Cells | Aspose.Cells picture placement types example in C#
// Tags: worksheet picture insertion Aspose.Cells | moveandsize placementtype example | xlsx file saving after image insertion | anchor image to cell A1 C# | load workbook and add png Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for PlacementType enum

// The program verifies the input Excel and PNG files, loads the workbook, inserts the PNG image at cell A1 (row 0, column 0) on the first worksheet, sets the picture's Placement to MoveAndSize so it moves and resizes with the cell, and then saves the modified workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string imagePath = "image.png";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Insert the picture into the worksheet
            using (FileStream pictureStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                // Add returns the picture index
                int pictureIndex = worksheet.Pictures.Add(0, 0, pictureStream);
                var picture = worksheet.Pictures[pictureIndex];

                // Set placement so the picture moves and resizes with the cell
                picture.Placement = PlacementType.MoveAndSize;
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
