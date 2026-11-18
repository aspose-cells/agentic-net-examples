// Title: How to set a shape's Z-order just above another shape in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape Z-order above another shape | move Excel drawing object to top of specific shape with Aspose.Cells | change ZOrderPosition of a shape relative to another shape in .NET | place one worksheet shape over another using Aspose.Cells API | adjust stacking order of Excel shapes programmatically in C#
// Tags: Aspose.Cells shape ZOrderPosition manipulation | C# reorder Excel drawing objects | set shape stacking order in workbook | move shape above reference shape .NET | Excel shape layering with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel file, retrieves two named shapes, sets the ZOrderPosition of the first shape to one greater than the second shape's ZOrderPosition (placing it directly above), and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape to move
            Shape shapeToMove = sheet.Shapes["ShapeToMove"];
            if (shapeToMove == null)
            {
                Console.WriteLine("Error: Shape \"ShapeToMove\" not found.");
                return;
            }

            // Retrieve the reference shape that should be directly below the moved shape
            Shape referenceShape = sheet.Shapes["ReferenceShape"];
            if (referenceShape == null)
            {
                Console.WriteLine("Error: Shape \"ReferenceShape\" not found.");
                return;
            }

            // Adjust Z-order: place shapeToMove just above referenceShape
            shapeToMove.ZOrderPosition = referenceShape.ZOrderPosition + 1;

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
