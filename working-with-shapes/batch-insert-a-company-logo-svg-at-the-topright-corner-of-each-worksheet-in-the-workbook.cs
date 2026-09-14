// Title: Batch insert a company logo SVG into the top‑right corner of every worksheet using Aspose.Cells for .NET
// AI Prompts: Add a free‑floating SVG logo to row 1 and the last used column of each worksheet in a workbook with Aspose.Cells for .NET. | Loop through all worksheets, compute the maximum column index, insert the SVG image, set PlacementType.FreeFloating, and save the updated workbook.
// Common Searches: how to add the same SVG image to the top right cell of every sheet in an Excel file using Aspose.Cells C# | Aspose.Cells C# place company logo in the first row and last column of each worksheet | determine last used column in Aspose.Cells and insert picture programmatically | batch add free floating SVG picture to all worksheets with Aspose.Cells .NET
// Tags: Aspose.Cells add SVG picture to worksheets | free floating image placement Aspose.Cells | determine max column Aspose.Cells | batch insert logo Excel workbook C# | picture insertion top right corner Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, checks for a logo.svg file, iterates over every worksheet, finds the last used column, inserts the SVG logo as a free‑floating picture at the first row of that column, and saves the modified workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Path to the SVG logo to be inserted
            string logoPath = "logo.svg";
            bool logoExists = File.Exists(logoPath);

            // Loop through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the index of the last used column (top‑right corner)
                int lastColumn = sheet.Cells.MaxColumn;

                if (logoExists)
                {
                    // Insert the SVG picture at the first row (0) and the last column
                    int pictureIndex = sheet.Pictures.Add(0, lastColumn, logoPath);

                    // Adjust picture placement if needed
                    Picture picture = sheet.Pictures[pictureIndex];
                    picture.Placement = PlacementType.FreeFloating;
                    picture.Top = 0;
                    picture.Left = 0;
                }
                else
                {
                    Console.WriteLine($"Logo file not found: {logoPath}. Skipping picture insertion for sheet '{sheet.Name}'.");
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
