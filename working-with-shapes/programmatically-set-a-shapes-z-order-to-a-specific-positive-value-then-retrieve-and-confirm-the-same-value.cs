// Title: Set and Verify Shape Z‑Order (ZOrderPosition) in Aspose.Cells for .NET
// Description: Shows how to add a rectangle shape to a worksheet, assign a positive ZOrderPosition, read the value back, and confirm the stacking order before saving the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | shape Z-order | ZOrderPosition | set shape stacking order | retrieve shape Z-order | Excel shape layering | programmatic shape order | Aspose.Cells shape ordering
// Common Searches: Aspose.Cells set shape ZOrderPosition | How to change shape stacking order in Aspose.Cells .NET | Retrieve Z-order of a shape after setting it | Verify shape ZOrderPosition in generated Excel file | C# Aspose.Cells shape layer control
// Developer Intent: Assign a specific positive ZOrderPosition to a shape, read the property, and ensure it matches the expected value.
// Use Cases: Place a particular shape above others by setting its ZOrderPosition when creating Excel reports. | Control the visual layering of charts, images, and drawings programmatically. | Validate that shape ordering persists after saving and reopening a workbook.
// AI Prompts: Generate C# code with Aspose.Cells that sets a shape's ZOrderPosition to 10 and confirms the value. | Explain how Aspose.Cells resolves conflicts when multiple shapes share the same ZOrderPosition. | Provide a C# snippet that iterates through all worksheet shapes and prints each shape's ZOrderPosition.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to a worksheet, assign a positive ZOrderPosition, read the value back, and confirm the stacking order before saving the workbook using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

        // Set the Z-order position to a specific positive value (e.g., 5)
        shape.ZOrderPosition = 5;

        // Retrieve the Z-order position to confirm it was set correctly
        int currentZOrder = shape.ZOrderPosition;
        Console.WriteLine("Current ZOrderPosition: " + currentZOrder);

        // Simple verification
        if (currentZOrder == 5)
        {
            Console.WriteLine("Z-order successfully set to the desired value.");
        }
        else
        {
            Console.WriteLine("Z-order value does not match the expected value.");
        }

        // Save the workbook to a file
        workbook.Save("ShapeZOrderDemo.xlsx");
    }
}
