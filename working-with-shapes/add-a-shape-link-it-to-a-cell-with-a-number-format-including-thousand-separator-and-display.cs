// Title: Link a Rectangle Shape to a Cell with Thousand‑Separator Number Format using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, write a numeric value to A1, apply the built‑in "#,##0" format (ID 3) for thousand separators, add a rectangle shape, link the shape to the formatted cell with SetLinkedCell, and save the file as ShapeLinkedWithNumberFormat.xlsx.
// Keywords: Aspose.Cells C# shape linking | rectangle shape SetLinkedCell | thousand separator number format Aspose.Cells | .NET workbook shape example | link shape to cell value
// Common Searches: Aspose.Cells link shape to cell C# | apply thousand separator format to cell Aspose.Cells | add rectangle shape and bind to cell value .NET | SetLinkedCell usage examples | format cell with #,##0 and link shape
// Developer Intent: Create a rectangle shape, bind it to a cell formatted with a thousand‑separator number format, and generate the workbook.
// Use Cases: Dynamic dashboards where a shape shows a formatted total that updates with the source cell. | Financial reports that display currency amounts in shapes with proper thousand separators. | Invoice templates where a shape reflects a formatted price pulled from a specific cell.
// AI Prompts: Generate C# code that adds a rectangle shape, formats a cell with "#,##0", links the shape to the cell, and saves the workbook using Aspose.Cells. | Explain how to refresh a linked shape after changing the cell value in an existing Aspose.Cells file. | Show different SetLinkedCell options for linking shapes to cells in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // Demonstrates how to create a workbook, write a numeric value to A1, apply the built‑in "#,##0" format (ID 3) for thousand separators, add a rectangle shape, link the shape to the formatted cell with SetLinkedCell, and save the file as ShapeLinkedWithNumberFormat.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value in cell A1
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue(1234567);

            // Apply a number format with thousand separator (format ID 3 => "#,##0")
            Style style = cell.GetStyle();
            style.Number = 3; // "#,##0"
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // top row index
                0,   // vertical offset (pixels) from the top row
                2,   // left column index
                0,   // horizontal offset (pixels) from the left column
                100, // height in pixels
                200  // width in pixels
            );

            // Link the shape's value to cell A1
            shape.SetLinkedCell("$A$1", false, true);

            // Save the workbook
            workbook.Save("ShapeLinkedWithNumberFormat.xlsx");
        }
    }
}
