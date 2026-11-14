// Title: Duplicate a rectangle shape, link the copy to another cell, and compare shape texts with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# duplicate a shape and attach it to a different worksheet cell | compare text of two shapes in an Excel workbook using Aspose.Cells .NET | add a copy of a rectangle shape at a specific cell with Aspose.Cells API | set shape placement to MoveAndSize for cell‑linked shapes in Aspose.Cells | use AddCopy method to duplicate a shape and verify its Text property in C#
// Tags: AddCopy shape method Aspose.Cells | shape placement MoveAndSize linking | compare shape Text property C# | duplicate rectangle shape Excel .NET | link shape to cell by row column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape at cell C3, assigns text, and links it to the cell with MoveAndSize placement. It then duplicates the shape to cell E5 using the AddCopy method, updates the duplicate's text, links it similarly, compares the Text properties of both shapes, and saves the workbook as DuplicatedShapes.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape at cell C3 (row 2, column 2) with size 100x30 points
            Shape originalShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // top offset in points
                0,   // left offset in points
                100, // width in points
                30   // height in points
            );

            originalShape.Text = "Original Shape";

            // Link the original shape to cell C3 (so it moves/resizes with the cell)
            originalShape.Placement = PlacementType.MoveAndSize;
            originalShape.UpperLeftRow = 2;
            originalShape.UpperLeftColumn = 2;

            // Duplicate the shape and place the copy at cell E5 (row 4, column 4)
            // AddCopy overload requires row/column offsets; use 0 for both
            Shape duplicateShape = sheet.Shapes.AddCopy(originalShape, 4, 4, 0, 0);
            duplicateShape.Text = "Duplicate Shape";

            // Link the duplicate shape to cell E5
            duplicateShape.Placement = PlacementType.MoveAndSize;
            duplicateShape.UpperLeftRow = 4;
            duplicateShape.UpperLeftColumn = 4;

            // Compare the displayed text of both shapes
            bool textsAreEqual = originalShape.Text == duplicateShape.Text;
            Console.WriteLine("Are the shape texts equal? " + textsAreEqual);

            // Save the workbook
            string outputPath = "DuplicatedShapes.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
