// Title: Aspose.Cells .NET: Insert an Image into Cell C9 and Retrieve Its Absolute X/Y Coordinates
// Description: Demonstrates how to add a PNG picture to cell C9 (row 8, column 2) using Aspose.Cells for .NET, set the picture's Placement to MoveAndSize, log its absolute X and Y pixel positions, save the workbook, reload it, and verify that the coordinates persist.
// Keywords: Aspose.Cells add picture to cell | C# insert image Aspose.Cells | picture absolute coordinates | PlacementType.MoveAndSize example | retrieve picture X Y properties | save workbook with image | reload workbook picture position
// Common Searches: how to insert an image into a specific cell with Aspose.Cells .NET | get pixel position of a picture after adding it to a worksheet | persist picture coordinates after saving and loading a workbook | set picture placement to move and size with cell in Aspose.Cells
// Developer Intent: Add a PNG to cell C9, configure it to move and size with the cell, and obtain its absolute X and Y pixel coordinates before and after saving the workbook.
// Use Cases: Log picture coordinates to validate layout alignment in automated report generation. | Calculate relative positions for additional shapes or annotations based on the image's absolute location. | Ensure consistent picture placement when a workbook is transferred between systems or edited later.
// AI Prompts: Generate C# code that inserts a PNG into cell C9 with Aspose.Cells, sets PlacementType.MoveAndSize, and prints the picture's X and Y coordinates. | Show how to retrieve and log picture coordinates after reloading a workbook saved with Aspose.Cells. | Explain how PlacementType.MoveAndSize influences a picture's absolute coordinates in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Picture class and PlacementType enum

namespace AsposeCellsExample
{
    // Demonstrates how to add a PNG picture to cell C9 (row 8, column 2) using Aspose.Cells for .NET, set the picture's Placement to MoveAndSize, log its absolute X and Y pixel positions, save the workbook, reload it, and verify that the coordinates persist.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file (ensure it exists)
                string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Add the picture to cell C9 (row index 8, column index 2)
                int pictureIndex = worksheet.Pictures.Add(8, 2, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Embed the picture inside the cell (move and size with the cell)
                picture.Placement = PlacementType.MoveAndSize;

                // Log absolute X/Y coordinates (pixels from worksheet origin)
                Console.WriteLine($"Picture absolute X: {picture.X} pixels");
                Console.WriteLine($"Picture absolute Y: {picture.Y} pixels");

                // Save the workbook
                string fileName = "output.xlsx";
                workbook.Save(fileName);

                // Verify the file was saved
                if (!File.Exists(fileName))
                {
                    Console.WriteLine($"Failed to save workbook: {fileName}");
                    return;
                }

                // Reload the workbook to demonstrate persistence
                Workbook loadedWorkbook;
                try
                {
                    loadedWorkbook = new Workbook(fileName);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Error loading workbook: {loadEx.Message}");
                    return;
                }

                Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

                // Ensure a picture exists after reload
                if (loadedWorksheet.Pictures.Count > 0)
                {
                    Picture loadedPicture = loadedWorksheet.Pictures[0];
                    Console.WriteLine($"After reload - Picture absolute X: {loadedPicture.X} pixels");
                    Console.WriteLine($"After reload - Picture absolute Y: {loadedPicture.Y} pixels");
                }
                else
                {
                    Console.WriteLine("No pictures found after reload.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }
}
