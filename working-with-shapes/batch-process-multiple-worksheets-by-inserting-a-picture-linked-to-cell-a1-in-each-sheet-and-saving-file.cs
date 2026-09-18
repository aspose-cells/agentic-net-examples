// Title: Batch insert a PNG picture into cell A1 of every worksheet and save the workbook using Aspose.Cells for .NET
// AI Prompts: Insert a PNG image from a file into cell A1 of each worksheet in an existing workbook, set the picture placement to MoveAndSize, and save the modified workbook. | Iterate over all worksheets in a Workbook, add a picture from a byte array at the top‑left cell, configure it to move and resize with the cell, then write the result to a new file.
// Common Searches: how to add the same image to cell A1 of all sheets in an Excel file using Aspose.Cells C# | Aspose.Cells batch picture insertion across multiple worksheets with MoveAndSize placement | insert picture from memory stream into each worksheet cell A1 in .NET
// Tags: batch picture insertion Aspose.Cells C# | add image to cell A1 multiple worksheets | picture placement MoveAndSize Aspose.Cells | load picture from byte array Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing workbook, reads a PNG file into a byte array, loops through every worksheet, inserts the picture at cell A1 with MoveAndSize placement so it moves and resizes with the cell, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputFile = "input.xlsx";
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Picture file path
            string pictureFile = "picture.png";
            if (!File.Exists(pictureFile))
            {
                Console.WriteLine($"Picture file not found: {pictureFile}");
                return;
            }

            // Read picture data into a byte array
            byte[] pictureBytes = File.ReadAllBytes(pictureFile);

            // Insert picture into each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a memory stream for the picture (reset for each sheet)
                using (MemoryStream picStream = new MemoryStream(pictureBytes))
                {
                    // Insert picture at cell A1 (row 0, column 0)
                    int picIndex = sheet.Pictures.Add(0, 0, picStream);
                    Picture picture = sheet.Pictures[picIndex];

                    // Ensure the picture moves and resizes with the cell
                    picture.Placement = PlacementType.MoveAndSize;
                }
            }

            // Save the modified workbook
            string outputFile = "output.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
