// Title: Link a Rectangle Shape to a Cell with Thousand‑Separator Formatting using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, write a numeric value to A1, apply the "#,##0" thousand‑separator format, add a rectangle shape, link the shape to the formatted cell, and save the file as ShapeLinkedWithNumberFormat.xlsx.
// Keywords: Aspose.Cells shape linking | rectangle shape linked cell | thousand separator number format | #,##0 format Aspose.Cells | C# Aspose.Cells example | link shape to cell value | display formatted number in shape
// Common Searches: Aspose.Cells link shape to cell with number format | how to add rectangle shape linked to cell in .NET | apply thousand separator format to linked shape Aspose.Cells | C# code for shape linked cell Aspose.Cells | display formatted cell value in shape Aspose.Cells
// Developer Intent: Create a workbook, format a cell with "#,##0", add a rectangle shape, and bind the shape to that cell so it shows the formatted number.
// Use Cases: Dynamic dashboards where shapes reflect updated totals with thousand separators. | Financial reports that embed formatted monetary values inside linked shapes. | Interactive spreadsheets that automatically refresh shape text when the source cell changes.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape linked to cell A1 formatted as "#,##0" and ensure the shape displays the formatted value. | Explain how to change the linked cell's value at runtime and have the shape automatically update its displayed number. | Show how to adjust the text alignment and font of a linked shape while preserving the cell link and number format.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // Demonstrates how to create a workbook, write a numeric value to A1, apply the "#,##0" thousand‑separator format, add a rectangle shape, link the shape to the formatted cell, and save the file as ShapeLinkedWithNumberFormat.xlsx.
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

            // Apply thousand separator number format (#,##0) to the cell
            Style style = cell.GetStyle();
            style.Number = 3; // 3 corresponds to "#,##0"
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                topRow: 2,    // upper left row index
                top: 0,       // vertical offset in pixels
                leftColumn: 2,// upper left column index
                left: 0,      // horizontal offset in pixels
                height: 100,  // height in pixels
                width: 200    // width in pixels
            );

            // Link the shape's value to cell A1
            shape.SetLinkedCell("$A$1", isR1C1: false, isLocal: true);

            // Save the workbook
            workbook.Save("ShapeLinkedWithNumberFormat.xlsx");
        }
    }
}
