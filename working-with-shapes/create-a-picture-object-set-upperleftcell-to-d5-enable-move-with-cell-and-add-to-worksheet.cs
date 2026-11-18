// Title: Add a picture to an Excel worksheet at cell D5 and set it to move with cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert image at specific cell D5 | how to set picture placement to Move in Aspose.Cells workbook | position picture using UpperLeftColumn and UpperLeftRow Aspose.Cells | C# Aspose.Cells picture move with cells property | add PNG image to Excel sheet and anchor to cell D5 using Aspose.Cells
// Tags: insert picture into worksheet Aspose.Cells | anchor picture to cell D5 Aspose.Cells | picture placement Move Aspose.Cells | Aspose.Cells picture UpperLeftColumn UpperLeftRow | C# add image to Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, loads a PNG file, adds it as a Picture to the first worksheet, anchors the picture's upper‑left corner to cell D5 by setting UpperLeftColumn and UpperLeftRow, configures the picture to move with cells using PlacementType.Move, and saves the workbook as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the image file to be inserted
                const string imagePath = "image.png";

                // Verify that the image file exists to avoid FileNotFoundException
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Error: Image file \"{imagePath}\" not found.");
                    return;
                }

                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a picture to the worksheet; the Add method returns the picture index
                int pictureIndex = worksheet.Pictures.Add(0, 0, imagePath);

                // Retrieve the picture object
                Picture picture = worksheet.Pictures[pictureIndex];

                // Set the upper‑left cell of the picture to D5 (zero‑based indices: column 3, row 4)
                picture.UpperLeftColumn = 3; // Column D
                picture.UpperLeftRow = 4;    // Row 5

                // Enable "move with cells" behavior
                picture.Placement = PlacementType.Move;

                // Save the workbook (lifecycle rule: save)
                const string resultPath = "Result.xlsx";
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
