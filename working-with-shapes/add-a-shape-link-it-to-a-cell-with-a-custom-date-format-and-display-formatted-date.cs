// Title: Link a Label Shape to a Formatted Date Cell using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, write the current date to cell B2, apply a custom "dd-mmm-yyyy" format, add a label shape at C5, link the shape to the formatted date cell with an A1 reference, center the text, and save the file as ShapeLinkedDate.xlsx.
// Keywords: Aspose.Cells | C# | label shape | linked cell | custom date format | Excel shape linking | SetLinkedCell | date formatting Aspose.Cells | Excel dashboard shape
// Common Searches: Aspose.Cells link shape to cell | C# add label shape linked to date cell | custom date format in Aspose.Cells | display formatted date in Excel shape | how to set linked cell for a shape in Aspose.Cells
// Developer Intent: Create a label shape that automatically shows a date from a worksheet cell formatted with a custom pattern.
// Use Cases: Add a dynamic generation date to a report header via a linked shape. | Show a deadline or invoice date on a dashboard without manual updates. | Design a template where a shape reflects a cell's formatted date for branding consistency.
// AI Prompts: Generate C# code with Aspose.Cells that adds a label shape linked to cell B2 formatted as "dd-mmm-yyyy". | Explain how to refresh the text of a linked label shape when the source date cell changes. | Provide an example that centers the linked text inside the shape and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, write the current date to cell B2, apply a custom "dd-mmm-yyyy" format, add a label shape at C5, link the shape to the formatted date cell with an A1 reference, center the text, and save the file as ShapeLinkedDate.xlsx.
class ShapeLinkedDateExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a date value in cell B2
        Cell dateCell = worksheet.Cells["B2"];
        dateCell.PutValue(DateTime.Now);

        // Apply a custom date format to the cell (e.g., "dd-mmm-yyyy")
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-mmm-yyyy";
        dateCell.SetStyle(dateStyle);

        // Add a label shape that will display the linked cell value
        // Parameters: upper left row, upper left column, top, left, height, width
        // Here we place the shape at row 4, column 2 (C5) with size 100x30 points
        Label label = (Label)worksheet.Shapes.AddLabel(4, 2, 0, 0, 100, 30);

        // Link the shape to the date cell (B2)
        // Using A1 style reference, not R1C1, and locale-aware formatting
        label.SetLinkedCell("$B$2", false, true);

        // Optionally, set the shape's text alignment
        label.TextHorizontalAlignment = TextAlignmentType.Center;
        label.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook to a file
        workbook.Save("ShapeLinkedDate.xlsx");
    }
}
