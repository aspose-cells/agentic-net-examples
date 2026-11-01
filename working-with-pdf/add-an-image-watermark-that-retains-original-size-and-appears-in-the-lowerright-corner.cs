// Title: Add a fixed‑size image watermark to the lower‑right corner of every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# place image watermark in bottom‑right corner of each sheet without scaling | how to insert a free‑floating picture as a watermark in Excel with Aspose.Cells .NET | preserve original image dimensions when adding a PNG watermark to an Excel workbook using Aspose.Cells | iterate all worksheets and add a watermark at the last used cell with Aspose.Cells | set picture placement to FreeFloating in Aspose.Cells to keep its size
// Tags: add image watermark Aspose.Cells | bottom right picture placement Excel | picture placement FreeFloating Aspose.Cells | keep image size unchanged Aspose.Cells | iterate worksheets add watermark .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, reads a PNG watermark into memory, loops through each worksheet, finds the last used row and column, inserts the image at that cell, sets the picture's Placement to FreeFloating so the original dimensions are retained, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string imagePath = "watermark.png";
            const string outputPath = "output.xlsx";

            // Verify required files exist to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Watermark image not found: {imagePath}");

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Load the watermark image bytes once for reuse.
            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // Iterate through all worksheets and place the watermark.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Identify the lower‑right cell of the used range.
                int lastRow = sheet.Cells.MaxDataRow;
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Add the picture to the worksheet at the identified cell.
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    int pictureIndex = sheet.Pictures.Add(lastRow, lastColumn, ms);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Ensure the picture does not move or resize with cells.
                    picture.Placement = PlacementType.FreeFloating;
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
