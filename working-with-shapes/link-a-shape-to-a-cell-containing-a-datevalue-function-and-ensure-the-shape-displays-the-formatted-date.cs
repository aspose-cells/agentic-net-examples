// Title: Link a TextBox shape to a DATEVALUE formula cell and display the formatted date with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link textbox shape to cell with DATEVALUE formula | display formatted date from linked shape in Excel workbook using Aspose.Cells | set shape text to cell A1 and keep date format when using Aspose.Cells .NET | move and size shape with cells after linking to date cell in Aspose.Cells | how to bind textbox shape to a date formula cell in Aspose.Cells for .NET
// Tags: textbox shape linked to DATEVALUE cell Aspose.Cells | formatted date display in shape Aspose.Cells | shape placement moveandsize Aspose.Cells | apply built‑in date style to cell Aspose.Cells | shape text formula binding Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a DATEVALUE formula in cell A1, applies a built‑in date number format, adds a TextBox shape, links the shape's text to the cell so the formatted date appears, sets the shape to move and size with cells, and saves the file as ShapeLinkedToDateValue.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a DATEVALUE formula in cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.Formula = @"=DATEVALUE(""2023-09-01"")";

            // Apply a date format to the cell so it displays as a formatted date
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format (e.g., mm/dd/yyyy)
            dateCell.SetStyle(dateStyle);

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            int upperLeftRow = 2;
            int upperLeftColumn = 1;
            int top = 5;
            int left = 5;
            int height = 50;
            int width = 200;

            Shape textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, height, width);

            // Link the shape's text to the cell containing the DATEVALUE formula
            // The shape will display the cell's formatted value and update automatically
            textBox.Text = "=A1";

            // Optional: set the shape to move and size with cells
            textBox.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("ShapeLinkedToDateValue.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
