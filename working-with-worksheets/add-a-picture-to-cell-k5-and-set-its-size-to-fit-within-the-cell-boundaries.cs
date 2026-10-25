// Title: Insert a PNG image into cell K5 and automatically resize it to fit the cell using Aspose.Cells for .NET
// Common Searches: c# insert PNG image into Excel cell K5 with Aspose.Cells | resize picture to match Excel cell dimensions using Aspose.Cells .NET | set picture placement to MoveAndSize so it moves with the cell in Aspose.Cells | get pixel width and height of a worksheet cell for image scaling in C# | anchor an image to a specific cell and fit it inside the cell boundaries Aspose.Cells
// Tags: add picture to specific worksheet cell Aspose.Cells | scale image to cell pixel size .NET | picture placement MoveAndSize Aspose.Cells example | retrieve column width pixel Aspose.Cells API | anchor PNG to cell K5 C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, calculates the pixel width and height of cell K5, adds a PNG picture anchored to that cell, resizes the picture to exactly fill the cell, sets its placement to MoveAndSize so it moves and scales with the cell, and saves the workbook as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the target cell K5 (zero‑based indices)
                int targetRow = 4;      // Row 5
                int targetColumn = 10; // Column K

                // Get the pixel dimensions of the target cell
                double cellWidthPx = sheet.Cells.GetColumnWidthPixel(targetColumn);
                double cellHeightPx = sheet.Cells.GetRowHeightPixel(targetRow);

                // Path to the image file
                string imagePath = @"C:\Images\MyPicture.png";

                if (File.Exists(imagePath))
                {
                    // Add the picture anchored to K5
                    int pictureIdx = sheet.Pictures.Add(targetRow, targetColumn, imagePath);
                    Picture picture = sheet.Pictures[pictureIdx];

                    // Resize the picture to fit exactly within the cell boundaries
                    picture.Width = (int)cellWidthPx;
                    picture.Height = (int)cellHeightPx;

                    // Ensure the picture moves and resizes with the cell
                    picture.Placement = PlacementType.MoveAndSize;
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                }

                // Ensure the output directory exists
                string outputPath = @"C:\Output\Result.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
