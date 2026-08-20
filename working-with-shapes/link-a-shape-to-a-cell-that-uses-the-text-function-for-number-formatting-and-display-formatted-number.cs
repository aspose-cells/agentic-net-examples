// Title: Link a Shape to a TEXT‑formatted Cell in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, store a raw number in A1, format it with the TEXT function in B1, add a rectangle shape, bind the shape to B1 using the LinkedCell property, refresh the displayed text, and save the file as ShapeLinkedToTextFunction.xlsx.
// Keywords: Aspose.Cells | C# shape LinkedCell | TEXT function formatting | Excel shape binding | number formatting in shape | currency format Aspose.Cells | dynamic shape text | link shape to formula cell | Aspose.Cells example | shape update value
// Common Searches: Aspose.Cells link shape to cell with TEXT formula | C# bind rectangle shape to formatted cell value | How to display TEXT‑formatted number in a shape using Aspose.Cells | LinkedCell property example with formula result | Refresh shape text after changing source cell in Aspose.Cells
// Developer Intent: Bind a worksheet shape to a cell that returns a TEXT‑formatted string so the shape shows the formatted number automatically.
// Use Cases: Financial dashboards where a shape displays a currency total calculated by a formula. | Invoice templates that show the amount with custom number formatting inside a shape. | Dynamic reports where shapes act as labels that update when the underlying formatted values change.
// AI Prompts: Show C# code to link a shape to a cell that uses the TEXT function and refresh the shape text in Aspose.Cells. | Explain how the LinkedCell property works with formula results and how to ensure the shape displays the formatted string. | Provide an example that updates a linked shape after modifying the source numeric value while keeping TEXT formatting.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // Demonstrates how to create a workbook, store a raw number in A1, format it with the TEXT function in B1, add a rectangle shape, bind the shape to B1 using the LinkedCell property, refresh the displayed text, and save the file as ShapeLinkedToTextFunction.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a raw numeric value in cell A1
            worksheet.Cells["A1"].PutValue(12345.6789);

            // Use the TEXT function to format the number in cell B1
            // The formula returns a formatted string, e.g., "$12,345.68"
            worksheet.Cells["B1"].Formula = @"=TEXT(A1, ""$#,##0.00"")";

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset, upper left offset,
            // height, width (all in points)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

            // Link the shape to the cell that contains the TEXT formula (B1)
            // Using the LinkedCell property (rule: Shape.LinkedCell)
            shape.LinkedCell = "$B$1";

            // Optionally, update the shape's displayed value immediately
            shape.UpdateSelectedValue();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ShapeLinkedToTextFunction.xlsx");
        }
    }
}
