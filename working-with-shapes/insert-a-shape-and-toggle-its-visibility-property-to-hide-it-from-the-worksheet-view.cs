// Title: Hide a Rectangle Shape in an Excel worksheet using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, insert a rectangle shape with Shapes.AddShape, toggle its IsHidden property to hide the shape, and save the file as HiddenShapeDemo.xlsx.
// Keywords: Aspose.Cells hide shape | C# add rectangle shape | IsHidden property Aspose.Cells | Excel shape visibility .NET | programmatically hide Excel shape
// Common Searches: Aspose.Cells hide shape C# | add rectangle shape and hide it in Excel | toggle shape visibility with Aspose.Cells | set IsHidden for worksheet shape | hide Excel shape using Aspose.Cells API
// Developer Intent: Add a shape to an Excel worksheet and then make it invisible by setting its IsHidden property.
// Use Cases: Create a hidden watermark that appears only when printing. | Insert placeholder graphics that can be revealed later. | Prepare a template with annotation shapes that are shown conditionally.
// AI Prompts: Generate C# code with Aspose.Cells to add a circle shape and hide it using IsHidden. | Show how to hide or show a shape based on a cell's value in an Excel file. | Provide an example that loops through all shapes in a worksheet and hides them with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new Workbook, insert a rectangle shape with Shapes.AddShape, toggle its IsHidden property to hide the shape, and save the file as HiddenShapeDemo.xlsx.
class HideShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet using ShapeCollection.AddShape
        // Parameters: type, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
        Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 2, 0, 100, 200);

        // Ensure the shape is initially visible
        shape.IsHidden = false;
        Console.WriteLine("Shape is visible: " + !shape.IsHidden);

        // Hide the shape from the worksheet view
        shape.IsHidden = true;
        Console.WriteLine("Shape is hidden: " + shape.IsHidden);

        // Save the workbook to a file
        workbook.Save("HiddenShapeDemo.xlsx");
    }
}
