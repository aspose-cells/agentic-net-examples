// Title: Create a rectangle shape linked to a cell formatted with thousand separators using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet and set its Text property to a formula that references a formatted cell. | Apply the custom number format '#,##0.00' to cell A1 and ensure a linked shape displays the formatted value. | Customize the shape’s fill and line colors, then save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells .NET how to bind shape text to a cell with custom number format | display formatted numeric value inside a rectangle shape using Aspose.Cells | add and style a shape that shows a thousand‑separator value from a worksheet cell | link Excel shape to cell value with '#,##0.00' format in C# Aspose.Cells | save workbook with shape linked to formatted cell Aspose.Cells example
// Tags: link shape text to cell value Aspose.Cells | custom number format '#,##0.00' Aspose.Cells | add rectangle shape worksheet .NET | shape fill line styling Aspose.Cells | save workbook as .xlsx Aspose.Cells | reference cell in shape formula Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This program creates a new workbook, writes a numeric value to A1, applies the '#,##0.00' thousand‑separator format, adds a rectangle shape, links the shape’s text to A1 so the formatted number appears, styles the shape’s fill and line, and saves the file as ShapeLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a numeric value in cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(1234567.89);

        // Apply a number format with thousand separator
        Style style = cell.GetStyle();
        style.Custom = "#,##0.00";
        cell.SetStyle(style);

        // Add a rectangle shape to the worksheet
        // Parameters: shape type, upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 120, 30);

        // Link the shape's text to cell A1 (the shape will display the formatted value)
        shape.Text = "=A1";

        // Optionally set the shape's fill and line style for better visibility
        shape.FillFormat.ForeColor = System.Drawing.Color.LightYellow;
        shape.LineFormat.Weight = 1.0;
        shape.LineFormat.ForeColor = System.Drawing.Color.DarkGray;

        // Save the workbook to a file
        workbook.Save("ShapeLinkedToCell.xlsx");
    }
}
