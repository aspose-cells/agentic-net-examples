// Title: Create a rectangle shape anchored to a cell and apply a custom thousand‑separator number format with Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape to a specific cell using Aspose.Cells C# | apply custom number format "#,##0.00" to a cell with Aspose.Cells .NET | position a shape relative to a worksheet cell in Aspose.Cells for .NET | set text on a shape linked to a cell in Aspose.Cells C# example | save workbook with shape anchored to a cell using Aspose.Cells
// Tags: rectangle shape insertion Aspose.Cells .NET | custom number format #,##0.00 Aspose.Cells | shape positioning relative to cell Aspose.Cells | assign text to shape Aspose.Cells C# | export workbook with shapes Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to cell B2, applies the custom number format "#,##0.00" to show thousand separators, adds a rectangle shape positioned at B2 with the text "Linked to B2", and saves the file as ShapeLinkedToCell.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell B2
            Cell cell = sheet.Cells["B2"];
            cell.PutValue(1234567.89);

            // Apply a custom number format with thousand separator
            Style style = cell.GetStyle();
            style.Custom = "#,##0.00"; // thousand separator with two decimal places
            cell.SetStyle(style);

            // Add a rectangle shape anchored to cell B2 (row 1, column 1)
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 80, 200);
            shape.Text = "Linked to B2";

            // Note: Direct linking via SetLinkToCell is not available in this API version.
            // The shape is already positioned relative to cell B2 through the AddShape parameters.

            // Save the workbook
            string outputPath = "ShapeLinkedToCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
