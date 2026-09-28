// Title: Subtract three from a shape's Z-order position and verify the new order using Aspose.Cells for .NET
// AI Prompts: Read a shape's current ZOrderPosition, decrement it by 3, and assign the result back using Aspose.Cells in C#. | Output the original and updated ZOrderPosition of a worksheet shape to the console. | Persist the workbook after changing the shape's Z-order to ensure the modification is saved.
// Common Searches: C# Aspose.Cells how to lower a shape's Z-order by a specific number | programmatically change shape layering order in Excel with Aspose.Cells .NET | retrieve and display shape ZOrderPosition after adjustment using Aspose.Cells
// Tags: adjust shape Z-order Aspose.Cells C# | decrement shape ZOrderPosition .NET | shape layering order Excel Aspose.Cells | read and set shape ZOrderPosition C# | save workbook after shape order change Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape, sets its ZOrderPosition, subtracts three from that value, prints the original and new Z-order to the console, and saves the file as ShapeZOrderAdjusted.xlsx.
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

            // Set an initial Z-order value for demonstration (optional)
            shape.ZOrderPosition = 5;

            // Store the original Z-order
            int originalZOrder = shape.ZOrderPosition;

            // Subtract three from the current Z-order
            shape.ZOrderPosition = originalZOrder - 3;

            // Verify the new Z-order
            int newZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Original Z-order: {originalZOrder}");
            Console.WriteLine($"New Z-order after subtracting 3: {newZOrder}");

            // Save the workbook
            workbook.Save("ShapeZOrderAdjusted.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
