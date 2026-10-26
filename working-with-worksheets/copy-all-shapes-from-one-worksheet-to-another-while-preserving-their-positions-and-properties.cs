// Title: Copy all shapes from one worksheet to another while preserving positions with Aspose.Cells for .NET
// Common Searches: how to copy all shapes from one Excel sheet to another using Aspose.Cells C# | preserve shape coordinates when moving drawings between worksheets in Aspose.Cells | Aspose.Cells AddCopy shape method example for copying shapes between sheets | duplicate Excel worksheet drawings programmatically with Aspose.Cells .NET | copy shapes with exact size and location from source sheet to target sheet Aspose
// Tags: Aspose.Cells copy shapes between worksheets | preserve shape coordinates Excel .NET | AddCopy shape method Aspose.Cells | duplicate worksheet drawings programmatically | copy Excel drawing objects with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, obtains a source and destination worksheet (creating the destination if needed), iterates through each Shape in the source sheet, captures its UpperLeftRow/Column and LowerRightRow/Column values, and uses Shapes.AddCopy to place an identical shape at the same coordinates on the destination sheet before saving the workbook.
class ShapeCopyExample
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the workbook containing the source and destination worksheets
            Workbook workbook = new Workbook(sourcePath);

            // Get the source worksheet (replace with your actual sheet name or index)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
                throw new ArgumentException("Source worksheet 'Sheet1' does not exist.");

            // Get the destination worksheet (replace with your actual sheet name or index)
            Worksheet destinationSheet = workbook.Worksheets["Sheet2"];
            // If the destination worksheet does not exist, create it
            if (destinationSheet == null)
                destinationSheet = workbook.Worksheets.Add("Sheet2");

            // Iterate through all shapes in the source worksheet and copy them
            foreach (Shape shape in sourceSheet.Shapes)
            {
                // Preserve the original position of the shape using properties
                int upperLeftRow = shape.UpperLeftRow;
                int upperLeftColumn = shape.UpperLeftColumn;
                int lowerRightRow = shape.LowerRightRow;
                int lowerRightColumn = shape.LowerRightColumn;

                // Add a copy of the shape to the destination worksheet at the same position
                destinationSheet.Shapes.AddCopy(shape, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            }

            // Save the workbook with the copied shapes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
