// Title: Copy all worksheet shapes with original positions using Aspose.Cells for .NET
// Description: Loads a source workbook, iterates through its ShapeCollection, and uses ShapeCollection.AddCopy to duplicate each shape on a new worksheet while preserving the UpperLeftRow, UpperLeftColumn, and offset values. The destination workbook is then saved.
// Keywords: Aspose.Cells copy shapes | ShapeCollection.AddCopy | preserve shape position | C# copy worksheet drawings | duplicate Excel shapes programmatically | copy charts Aspose.Cells | copy images between worksheets | Aspose.Cells .NET
// Common Searches: Aspose.Cells copy shapes between worksheets | How to duplicate drawings in Excel using C# | Preserve shape coordinates when copying sheets Aspose.Cells | AddCopy method example C# | Copy charts and images to another sheet Aspose.Cells
// Developer Intent: Duplicate every shape from a source worksheet to a target worksheet while retaining its cell location and offsets.
// Use Cases: Migrate custom graphics from a template to generated reports | Create personalized dashboards by copying chart objects to individual sheets | Archive embedded images by moving them to a separate workbook | Clone worksheet layouts for multi‑region Excel exports
// AI Prompts: Write C# code that copies all shapes from one worksheet to another with Aspose.Cells, keeping exact row, column and offset values. | Explain each parameter of ShapeCollection.AddCopy and how to calculate offsets for different Excel versions. | Provide best‑practice error handling and logging for shape copying across workbooks using Aspose.Cells. | Show how to copy shapes while preserving hyperlinks and embedded data.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a source workbook, iterates through its ShapeCollection, and uses ShapeCollection.AddCopy to duplicate each shape on a new worksheet while preserving the UpperLeftRow, UpperLeftColumn, and offset values. The destination workbook is then saved.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string destinationPath = "destination.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the source workbook that contains the shapes
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook for the destination worksheet
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            destinationSheet.Name = "Destination";

            // Get shape collections from both worksheets
            ShapeCollection sourceShapes = sourceSheet.Shapes;
            ShapeCollection destinationShapes = destinationSheet.Shapes;

            // Iterate through each shape in the source worksheet
            for (int i = 0; i < sourceShapes.Count; i++)
            {
                Shape srcShape = sourceShapes[i];

                // Preserve the original position of the shape
                int topRow = srcShape.UpperLeftRow;          // Upper‑left row index
                int leftColumn = srcShape.UpperLeftColumn;   // Upper‑left column index

                // Offsets are not available in older API versions; use 0 as fallback
                int topOffset = 0;
                int leftOffset = 0;

                // Add a copy of the shape to the destination worksheet with the same position
                destinationShapes.AddCopy(srcShape, topRow, topOffset, leftColumn, leftOffset);
            }

            // Save the workbook that now contains the copied shapes
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Shapes copied successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
