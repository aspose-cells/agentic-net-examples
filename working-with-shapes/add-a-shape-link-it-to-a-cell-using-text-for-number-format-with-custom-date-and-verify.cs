// Title: Link a Rectangle Shape to a TEXT‑formatted Date Cell with Aspose.Cells for .NET
// Description: Creates a workbook, stores a DateTime in A1, uses the TEXT function in B1 to produce a custom date string, adds a rectangle shape, links the shape to B1 via SetLinkedCell, prints the linked cell address and displayed value, and saves the file for visual verification.
// Keywords: Aspose.Cells | C# | SetLinkedCell | shape linking | rectangle shape | TEXT function | custom date format | linked cell verification | .NET spreadsheet | Excel automation
// Common Searches: Aspose.Cells link shape to cell with TEXT formula | SetLinkedCell custom date format C# | how to verify shape linked cell in Aspose.Cells | link rectangle to formatted date cell Aspose | retrieve linked cell address from shape .NET
// Developer Intent: Add a rectangle shape, link it to a cell that returns a formatted date via TEXT, and confirm the link programmatically.
// Use Cases: Dynamic reports where a shape displays a formatted date that updates with the source cell. | Dashboards that use shapes as labels bound to TEXT‑formatted values for real‑time data visualization. | Automated tests that validate shape‑to‑cell links by outputting the linked address and displayed string.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape, link it to a TEXT‑formatted date cell, and output the linked cell address and value. | Explain how SetLinkedCell works with A1‑style references and local formulas when linking a shape to a TEXT result in Aspose.Cells for .NET. | Provide step‑by‑step instructions to verify that a shape linked to a TEXT‑formatted date cell reflects changes when the original date cell is modified.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, stores a DateTime in A1, uses the TEXT function in B1 to produce a custom date string, adds a rectangle shape, links the shape to B1 via SetLinkedCell, prints the linked cell address and displayed value, and saves the file for visual verification.
class ShapeLinkedCellExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a date value into cell A1
        Cell dateCell = sheet.Cells["A1"];
        dateCell.PutValue(new DateTime(2023, 12, 25));

        // Apply a custom date format using TEXT function in cell B1
        // The TEXT function returns a string, which will be linked to the shape
        Cell linkedCell = sheet.Cells["B1"];
        linkedCell.Formula = @"TEXT(A1, ""dd-mmm-yyyy"")";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Link the shape to cell B1 (the TEXT result)
        // Using SetLinkedCell method; isR1C1 = false (A1 style), isLocal = true
        rect.SetLinkedCell("$B$1", false, true);

        // Verify: output the linked cell address and its current value
        Console.WriteLine("Shape's LinkedCell: " + rect.LinkedCell);
        Console.WriteLine("Linked cell value (as displayed): " + linkedCell.StringValue);

        // Save the workbook (optional, for visual verification)
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}
