// Title: Add and Hide a Rectangle Shape in Excel with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, inserts a rectangle shape on the first worksheet, sets its IsHidden property to true, and saves the file as HiddenShape.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | Insert shape | Rectangle shape | IsHidden property | Hide shape | Excel worksheet shape | shape visibility | Excel automation
// Common Searches: Aspose.Cells hide shape C# | how to make a shape invisible in Excel using Aspose.Cells | add rectangle shape programmatically Aspose.Cells | set IsHidden property for Excel shape | hide Excel shape with Aspose.Cells .NET
// Developer Intent: Add a shape to a worksheet and keep it hidden from the user view.
// Use Cases: Store a marker shape that can be detected by code while remaining invisible to end users. | Create a placeholder for conditional display based on runtime logic. | Embed a shape used for internal calculations or data validation without cluttering the UI.
// AI Prompts: Generate C# code with Aspose.Cells to add a circle shape and hide it using the IsHidden property. | Show how to toggle the IsHidden flag of an existing shape at runtime based on a condition. | Provide an example that hides all shapes in a worksheet by iterating with a loop in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, inserts a rectangle shape on the first worksheet, sets its IsHidden property to true, and saves the file as HiddenShape.xlsx using Aspose.Cells for .NET.
class InsertShapeAndHide
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a rectangle shape (top row, top offset, left column, left offset, height, width)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Hide the shape from the worksheet view
        shape.IsHidden = true;

        // Save the workbook with the hidden shape
        workbook.Save("HiddenShape.xlsx");
    }
}
