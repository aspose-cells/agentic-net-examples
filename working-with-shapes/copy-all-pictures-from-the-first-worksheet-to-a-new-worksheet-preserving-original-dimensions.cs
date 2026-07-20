// Title: Copy pictures from first worksheet to a new sheet while keeping size – Aspose.Cells C#
// Description: Loads a workbook, adds a worksheet named "CopiedPictures", iterates over the Picture collection of the first sheet and uses ShapeCollection.AddCopy to duplicate each image to the new sheet with the same row, column, height and width, then saves the file.
// Keywords: Aspose.Cells | C# | copy pictures | ShapeCollection.AddCopy | duplicate images worksheet | preserve picture dimensions | Excel image copy .NET | worksheet shapes
// Common Searches: Aspose.Cells copy images between worksheets | C# copy picture preserving size | AddCopy shape Aspose.Cells example | duplicate pictures from first sheet | how to move pictures in Excel using Aspose
// Developer Intent: Duplicate every picture from the first worksheet onto a newly created worksheet, retaining its exact position and dimensions.
// Use Cases: Generate a visual‑only summary sheet from a data‑rich template | Apply corporate branding by copying logos to multiple report tabs | Create a printable snapshot of a dashboard containing only graphics
// AI Prompts: Write C# code with Aspose.Cells to copy all pictures from the first worksheet to a new worksheet, preserving row, column, height, and width. | Explain the parameters of ShapeCollection.AddCopy and how they affect picture placement in Aspose.Cells. | Suggest robust error‑handling patterns for file I/O and picture copying in Aspose.Cells .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureCopyDemo
{
    // Loads a workbook, adds a worksheet named "CopiedPictures", iterates over the Picture collection of the first sheet and uses ShapeCollection.AddCopy to duplicate each image to the new sheet with the same row, column, height and width, then saves the file.
    class Program
    {
        static void Main()
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(sourcePath);

                // Get the first worksheet (index 0)
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Add a new worksheet to hold the copied pictures
                Worksheet destinationSheet = workbook.Worksheets.Add("CopiedPictures");

                // Iterate through all pictures in the source worksheet
                foreach (Picture sourcePicture in sourceSheet.Pictures)
                {
                    // Copy each picture to the destination worksheet preserving its original position and size.
                    // ShapeCollection.AddCopy copies the shape using its upper‑left row/column and size.
                    destinationSheet.Shapes.AddCopy(
                        sourcePicture,
                        sourcePicture.UpperLeftRow,
                        sourcePicture.UpperLeftColumn,
                        sourcePicture.Height,
                        sourcePicture.Width);
                }

                // Save the workbook with the copied pictures
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
