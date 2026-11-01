// Title: How to lower a worksheet shape's Z-order by three positions and verify it with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# reduce shape Z-order by a specific number | set shape ZOrderPosition to a lower value in a worksheet using Aspose.Cells | verify worksheet shape layering after adjusting Z-order in .NET | prevent negative Z-order when moving a shape backward with Aspose.Cells | sample code to move a shape behind others in an Excel file using Aspose.Cells
// Tags: shape ZOrderPosition adjustment Aspose.Cells | worksheet shape layering order Aspose.Cells | negative ZOrderPosition handling Aspose.Cells | C# example shape Z-order modification Aspose.Cells | save workbook after shape order change Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, sets its initial Z-order to 5, subtracts three from the current Z-order, prints the new Z-order, ensures the value does not become negative by resetting it to zero if needed, and saves the workbook as 'AdjustedShapeZOrder.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Set an initial Z-order value for demonstration purposes
            shape.ZOrderPosition = 5;

            // Adjust the shape's Z-order by subtracting three from its current value
            int currentZ = shape.ZOrderPosition;
            shape.ZOrderPosition = currentZ - 3;

            // Verify the new Z-order value
            Console.WriteLine("New Z-order: " + shape.ZOrderPosition);

            // Optional safety check: ensure Z-order does not become negative
            if (shape.ZOrderPosition < 0)
            {
                Console.WriteLine("Z-order became negative; resetting to 0.");
                shape.ZOrderPosition = 0;
            }

            // Save the workbook
            string outputPath = "AdjustedShapeZOrder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
