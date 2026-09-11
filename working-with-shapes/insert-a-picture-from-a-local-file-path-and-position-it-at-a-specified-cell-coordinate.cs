// Title: Insert a local PNG picture into cell B2 of a new Excel workbook with FreeFloating placement using Aspose.Cells for .NET
// AI Prompts: Add a picture from a file path to a specific cell (e.g., B2) in a new workbook with Aspose.Cells C# and retrieve the picture object. | Set the inserted picture's Placement property to FreeFloating to allow independent movement. | Create the output directory if missing, then save the workbook to a given path after adding the picture.
// Common Searches: Aspose.Cells C# insert image at cell address B2 | How to add a PNG picture to an Excel worksheet using Aspose.Cells | Change picture placement mode in Aspose.Cells .NET | Save workbook after adding picture with Aspose.Cells example | Convert cell name to row and column indexes Aspose.Cells
// Tags: insert image into worksheet cell Aspose.Cells | PNG picture insertion Aspose.Cells .NET | FreeFloating picture placement Aspose.Cells | create output folder before saving workbook C# | cell name to row column conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example checks that the PNG file exists, creates a new workbook, converts cell B2 to row/column indexes, inserts the picture at that location, sets its placement to FreeFloating, ensures the output directory is present, and saves the workbook as Result.xlsx.
class InsertPictureExample
{
    static void Main()
    {
        // Path to the picture file to be inserted
        string picturePath = @"C:\Images\MyPicture.png";

        // Path where the resulting workbook will be saved
        string outputPath = @"C:\Output\Result.xlsx";

        try
        {
            // Verify picture file exists
            if (!File.Exists(picturePath))
            {
                Console.WriteLine($"Picture file not found: {picturePath}");
                return;
            }

            // Ensure output directory exists (if any)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell for the picture's top‑left corner
            string targetCell = "B2";

            // Convert cell name to row and column indexes
            int row = worksheet.Cells[targetCell].Row;
            int column = worksheet.Cells[targetCell].Column;

            // Add the picture; Add returns the picture index
            int pictureIndex = worksheet.Pictures.Add(row, column, picturePath);

            // Retrieve the Picture object using the index
            Picture picture = worksheet.Pictures[pictureIndex];

            // Optional: set placement type (FreeFloating allows independent movement)
            picture.Placement = PlacementType.FreeFloating;

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
