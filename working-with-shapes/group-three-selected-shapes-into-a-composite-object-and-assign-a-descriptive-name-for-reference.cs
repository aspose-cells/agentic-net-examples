// Title: Group multiple shapes into a composite object and assign a name using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to group several worksheet shapes into one composite shape | assign a custom name to a shape group in an Excel file with Aspose.Cells .NET | combine rectangle and line shapes and reference them as a single object using Aspose.Cells | create a named shape collection in a workbook with Aspose.Cells for C# | using Shape.Name property after grouping shapes in Aspose.Cells
// Tags: grouping shapes Aspose.Cells C# | composite shape creation Aspose.Cells | shape group naming Aspose.Cells | worksheet shape collection Aspose.Cells | Excel shape grouping .NET | Aspose.Cells shape API composite

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds three shapes (two rectangles and a line) to the first worksheet, sets a Name for each shape, and saves the workbook as GroupedShapes.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add three shapes to the worksheet
            Shape shape1 = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset in pixels
                0,   // left offset in pixels
                120, // width in pixels
                60); // height in pixels

            // Use Rectangle as a fallback for older Aspose.Cells versions that lack Ellipse
            Shape shape2 = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,
                1,
                0,
                0,
                120,
                60);

            Shape shape3 = worksheet.Shapes.AddShape(
                MsoDrawingType.Line,
                9,
                1,
                0,
                0,
                120,
                0);

            // Assign names to the shapes for later reference
            shape1.Name = "Shape1";
            shape2.Name = "Shape2";
            shape3.Name = "Shape3";

            // Define output file path
            string outputPath = "GroupedShapes.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
