// Title: Batch insert a picture into cell A1 of each worksheet and save the workbook with Aspose.Cells for .NET
// Common Searches: aspnet insert same image into every sheet at A1 using Aspose.Cells | c# loop through worksheets and add picture anchored to cell A1 with Aspose.Cells | set picture placement to MoveAndSize for all sheets in Aspose.Cells workbook | save modified Excel file after adding images to each worksheet with Aspose.Cells .NET | how to verify input workbook and image file before adding pictures in Aspose.Cells
// Tags: batch add picture to worksheets Aspose.Cells | anchor image to cell A1 C# Aspose.Cells | picture placement MoveAndSize Aspose.Cells | save workbook after shape modifications .NET | verify input files before processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example validates the presence of an input Excel file and a PNG image, loads the workbook, iterates over every worksheet, inserts the image anchored at cell A1, sets the picture's placement to MoveAndSize so it moves and resizes with the cells, ensures the output directory exists, and saves the updated workbook. Exception handling is included to report any errors.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input workbook exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Verify that the picture file exists
            const string picturePath = "image.png";
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Picture file not found: {picturePath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Process every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add the picture to the sheet (initial position is cell A1)
                int pictureIndex = sheet.Pictures.Add(0, 0, picturePath);

                // Retrieve the inserted picture object
                Picture picture = sheet.Pictures[pictureIndex];

                // Make the picture move and size with the cells
                picture.Placement = PlacementType.MoveAndSize;

                // The picture is already anchored to cell A1 by Add(0,0,…)
                // If further positioning is required, set the row/column offsets directly:
                // picture.UpperLeftRow = 0;
                // picture.UpperLeftColumn = 0;
                // picture.UpperLeftRowOffset = 0;
                // picture.UpperLeftColumnOffset = 0;
            }

            // Ensure the output directory exists
            const string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
