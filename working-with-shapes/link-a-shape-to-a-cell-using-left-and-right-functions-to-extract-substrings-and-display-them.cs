// Title: Link a Rectangle Shape to a Cell with LEFT/RIGHT Formula using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, write a source string to A1, apply a LEFT‑RIGHT concatenation formula in B1, add a rectangle shape, and bind the shape to the formula cell via the LinkedCell property. The workbook is saved as LinkedShape.xlsx.
// Keywords: Aspose.Cells | C# | LinkedCell | rectangle shape | LEFT function | RIGHT function | Excel formula binding | shape to cell | dynamic label
// Common Searches: Aspose.Cells link shape to cell C# | how to bind rectangle to formula result in Aspose.Cells | LEFT RIGHT Excel formula with shape LinkedCell | C# example linking shape to cell using Aspose.Cells | dynamic shape text based on cell formula Aspose
// Developer Intent: Create a shape whose displayed text updates automatically from a cell that uses LEFT and RIGHT functions.
// Use Cases: Show a shortened version of a long string as a visual label. | Provide a dashboard element that reflects concatenated parts of a source cell. | Maintain a visual marker that updates when the underlying text changes.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape and link it to a cell containing a LEFT/RIGHT formula. | Explain how to resize and reposition a linked shape while keeping it synchronized with the formula cell. | Describe how to modify the LEFT/RIGHT formula at runtime and ensure the linked shape reflects the new result.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, write a source string to A1, apply a LEFT‑RIGHT concatenation formula in B1, add a rectangle shape, and bind the shape to the formula cell via the LinkedCell property. The workbook is saved as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a sample text into cell A1
        worksheet.Cells["A1"].PutValue("AsposeCells");

        // In cell B1 set a formula that extracts substrings using LEFT and RIGHT
        // Example: first 5 characters and last 3 characters concatenated
        worksheet.Cells["B1"].Formula = "=LEFT(A1,5)&RIGHT(A1,3)";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, height, width, placement type, shape index
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 200, 0, 0);

        // Link the shape to the cell containing the formula result (B1)
        shape.LinkedCell = "B1";

        // Save the workbook to a file
        workbook.Save("LinkedShape.xlsx");
    }
}
