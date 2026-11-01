// Title: Programmatically increase a shape's Z-order by five positions in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to a higher value | how to move an Excel shape forward by multiple layers with Aspose.Cells | increase Z-order of a shape in a .xlsx file using Aspose.Cells API | C# code to adjust shape layering in an Excel workbook with Aspose.Cells | retrieve and modify shape ZOrderPosition in Aspose.Cells for .NET
// Tags: shape ZOrderPosition adjustment Aspose.Cells | increase shape layer order .xlsx C# | modify Excel shape stacking Aspose.Cells | programmatic shape ordering in workbook | Aspose.Cells shape layering manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx workbook, accesses the first worksheet, selects the first shape, reads its current ZOrderPosition, adds five to bring the shape forward in the layer stack, prints the before/after values, and saves the workbook. It includes file existence verification and basic exception handling.
class AdjustShapeZOrder
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Ensure there is at least one shape to work with
            if (shapes.Count > 0)
            {
                // Select the first shape (or any shape by index)
                Shape shape = shapes[0];

                // Retrieve the current Z-order (layer position) of the shape
                int currentZOrder = shape.ZOrderPosition;

                // Increase the Z-order by five to bring the shape forward in the layer stack
                shape.ZOrderPosition = currentZOrder + 5;

                // Output the new Z-order to verify the change
                Console.WriteLine($"Shape '{shape.Name}' Z-order changed from {currentZOrder} to {shape.ZOrderPosition}");
            }
            else
            {
                Console.WriteLine("No shapes found on the worksheet.");
            }

            // Save the workbook with the updated shape Z-order
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
