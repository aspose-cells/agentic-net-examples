// Title: How to anchor a picture to a merged cell range and make it move and resize with the cells using Aspose.Cells for .NET
// AI Prompts: Insert a PNG file into a worksheet, merge cells B2:D4, and bind the image to the merged area so it moves and resizes together with the cells using Aspose.Cells for .NET. | Configure a picture's Placement to MoveAndSize and call SetPosition with the start and end rows/columns to lock the picture to a merged cell block.
// Common Searches: Aspose.Cells C# anchor image to merged cells B2:D4 | C# set picture placement move and size merged range Aspose.Cells | how to bind a picture to a merged cell area in Aspose.Cells for .NET | SetPosition picture merged cells Aspose.Cells example
// Tags: picture anchoring to merged cells Aspose.Cells | PlacementType.MoveAndSize worksheet image | SetPosition method picture anchoring Aspose.Cells | add picture to merged cells C# | merge cells and attach image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, merges cells B2:D4, adds a PNG picture at the start of the merged range, sets its Placement to MoveAndSize so it moves and resizes with the cells, optionally uses SetPosition to bind the picture to the full merged area, and saves the file as MergedPicture.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged cell range (B2:D4)
            int startRow = 1;      // B2 (zero‑based)
            int startColumn = 1;   // B2
            int endRow = 3;        // D4
            int endColumn = 3;     // D4

            // Merge the cells
            sheet.Cells.Merge(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);

            // Path to the picture file
            string picturePath = @"C:\Images\logo.png";

            // Ensure the picture file exists to avoid FileNotFoundException
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Image file not found: {picturePath}");

            // Add the picture to the worksheet
            int pictureIndex = sheet.Pictures.Add(startRow, startColumn, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Make the picture move and resize with the cells (including merged cells)
            picture.Placement = PlacementType.MoveAndSize;

            // Optional: anchor the picture to the merged range using SetPosition if available
            // Uncomment the following line if your Aspose.Cells version supports SetPosition:
            // picture.SetPosition(startRow, startColumn, endRow, endColumn);

            // Save the workbook
            string outputPath = "MergedPicture.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
