// Title: Aspose.Cells .NET – Add Rectangle Shape Linked to a Cell with Scientific Notation
// Description: Creates a workbook, writes a number to A1, applies a custom scientific notation format (0.00E+00), inserts a rectangle shape, links the shape to $A$1 via the LinkedCell property, prints the link and formatted value, and saves the file as ShapeLinkedCellScientific.xlsx.
// Keywords: Aspose.Cells | C# shape LinkedCell | rectangle shape Excel | custom number format scientific notation | 0.00E+00 format | link shape to cell | verify linked cell value
// Common Searches: Aspose.Cells link shape to cell | apply scientific notation format with Aspose.Cells .NET | how to use LinkedCell property in C# | add rectangle shape to worksheet Aspose | retrieve formatted cell value from linked shape
// Developer Intent: Add a rectangle shape, bind it to a cell formatted in scientific notation, and confirm the binding and displayed value.
// Use Cases: Show a large numeric value in scientific notation on a dashboard by linking a shape to the source cell. | Create interactive Excel reports where clicking a shape jumps to a cell containing a formatted scientific value. | Generate templates where shapes serve as visual anchors tied to cells with custom number formats for consistency.
// AI Prompts: Generate C# code using Aspose.Cells to insert a rectangle shape, set its LinkedCell to A1, and format A1 with 0.00E+00. | Demonstrate how to read and display the LinkedCell address and the cell's scientific notation string value. | Explain how to modify the scientific notation format of a linked cell after the shape is linked without breaking the connection.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes a number to A1, applies a custom scientific notation format (0.00E+00), inserts a rectangle shape, links the shape to $A$1 via the LinkedCell property, prints the link and formatted value, and saves the file as ShapeLinkedCellScientific.xlsx.
class ShapeLinkedCellScientificNotation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a numeric value in cell A1
        Cell targetCell = sheet.Cells["A1"];
        targetCell.PutValue(123456789.0);

        // Apply scientific notation number format to the cell (e.g., 0.00E+00)
        Style sciStyle = workbook.CreateStyle();
        sciStyle.Custom = "0.00E+00";
        targetCell.SetStyle(sciStyle);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        RectangleShape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 50);

        // Link the shape to cell A1 using the LinkedCell property
        rect.LinkedCell = "$A$1";

        // Verify the link and the formatted value
        Console.WriteLine("Shape linked to cell: " + rect.LinkedCell);
        Console.WriteLine("Cell A1 formatted value: " + targetCell.StringValue);

        // Save the workbook
        workbook.Save("ShapeLinkedCellScientific.xlsx");
    }
}
