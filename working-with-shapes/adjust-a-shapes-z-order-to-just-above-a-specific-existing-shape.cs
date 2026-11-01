// Title: How to set a shape's Z‑order just above another shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set ZOrderPosition of a shape relative to another shape | move Excel shape above a specific shape with Aspose.Cells .NET | bring a named shape to front of another shape in Aspose.Cells workbook | adjust stacking order of worksheet shapes using Aspose.Cells API | find shape by name and change its Z‑order in Aspose.Cells C#
// Tags: Aspose.Cells shape ZOrderPosition manipulation | C# adjust Excel shape stacking order | Aspose.Cells find shape by name | set shape order relative to another shape Aspose.Cells | Excel worksheet shape Z‑order .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel file, locates two shapes by their names, and sets the ZOrderPosition of the target shape to one higher than the reference shape, placing it directly above the reference shape before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Retrieve the shapes by name
            Shape shapeToMove = FindShapeByName(shapes, "ShapeToMove");
            Shape referenceShape = FindShapeByName(shapes, "ReferenceShape");

            if (shapeToMove == null)
            {
                Console.WriteLine("Error: Shape \"ShapeToMove\" not found on the worksheet.");
                return;
            }

            if (referenceShape == null)
            {
                Console.WriteLine("Error: Shape \"ReferenceShape\" not found on the worksheet.");
                return;
            }

            // Adjust Z‑order: place shapeToMove just above referenceShape
            // Aspose.Cells does not provide MoveToAbove; adjust ZOrderPosition instead
            int refPos = referenceShape.ZOrderPosition;
            shapeToMove.ZOrderPosition = refPos + 1;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to locate a shape by its name
    private static Shape FindShapeByName(ShapeCollection shapes, string name)
    {
        foreach (Shape shape in shapes)
        {
            if (shape.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                return shape;
        }
        return null;
    }
}
