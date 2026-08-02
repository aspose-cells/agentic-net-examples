// Title: Aspose.Cells for .NET: Set Shape Z‑Order to int.MaxValue to Bring It to the Front
// Description: Shows how to add a rectangle shape to a worksheet, assign the maximum Z‑order value (int.MaxValue) so it appears above every other object, optionally invoke ToFrontOrBack, verify the setting, and save the workbook.
// Keywords: Aspose.Cells | C# shape Z-order | int.MaxValue ZOrderPosition | bring shape to front | worksheet shape layering | Aspose.Cells shape ordering | ToFrontOrBack method | Excel shape hierarchy | Aspose.Cells example
// Common Searches: Aspose.Cells set shape ZOrderPosition | C# bring Excel shape to front | maximum Z-order value Aspose.Cells | how to use ToFrontOrBack Aspose.Cells | shape layering in Aspose.Cells .NET
// Developer Intent: Place a specific shape on top of all other worksheet objects by setting its Z‑order to the highest possible value.
// Use Cases: Add a watermark that must overlay all cells and charts. | Ensure a callout or annotation stays visible over dynamic content. | Create a clickable button that should not be hidden by later‑added shapes. | Maintain logo positioning above data in automated report generation.
// AI Prompts: Write C# code using Aspose.Cells to add an ellipse shape and set its ZOrderPosition to int.MaxValue. | Compare ZOrderPosition and ToFrontOrBack in Aspose.Cells: when should each method be used? | Provide a C# routine that scans all shapes on a worksheet and moves the shape named 'Header' to the front. | Explain how shape Z‑order influences rendering order in Excel files created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to a worksheet, assign the maximum Z‑order value (int.MaxValue) so it appears above every other object, optionally invoke ToFrontOrBack, verify the setting, and save the workbook.
class ShapeZOrderDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 100);

        // Set a large positive Z-order position to bring the shape to the front of all layers
        // Using int.MaxValue ensures it is above any other shape's Z-order
        shape.ZOrderPosition = int.MaxValue;

        // Verify that the Z-order position has been set
        Console.WriteLine("Shape ZOrderPosition set to: " + shape.ZOrderPosition);

        // Optionally, you can also bring the shape to front using ToFrontOrBack for extra assurance
        // shape.ToFrontOrBack(1);

        // Save the workbook to a file
        workbook.Save("ShapeZOrderDemo.xlsx");
    }
}
