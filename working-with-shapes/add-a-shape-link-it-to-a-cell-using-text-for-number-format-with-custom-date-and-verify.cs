// Title: Link a Rectangle Shape to a Custom‑Formatted Date Cell with Aspose.Cells for .NET
// Description: Demonstrates how to insert the current date into cell A1, apply a custom date format, add a rectangle shape, link the shape to the formatted cell using the LinkedCell property, retrieve the address with LinkedCell and GetLinkedCell, and save the workbook to verify persistence.
// Keywords: Aspose.Cells shape linking | C# rectangle shape LinkedCell | custom date format cell Aspose.Cells | GetLinkedCell method | save workbook with linked shape | Aspose.Cells for .NET examples
// Common Searches: Aspose.Cells link shape to cell | set custom date format and link shape C# | retrieve linked cell address from shape | how to save workbook with linked shape Aspose.Cells
// Developer Intent: Create a shape, bind it to a date‑formatted cell, and programmatically confirm the binding.
// Use Cases: Add a visual rectangle that points to a date value displayed in a specific format. | Validate that a shape’s LinkedCell matches the target cell using both properties and methods. | Generate Excel reports where shapes serve as anchors for formatted date cells.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape, link it to a cell formatted as dd-MMM-yyyy, and print the linked address. | Explain the purpose of the two boolean parameters in GetLinkedCell when obtaining a shape’s linked cell reference. | Provide error‑handling patterns for linking a shape to a cell that may be empty or contain an invalid value.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to insert the current date into cell A1, apply a custom date format, add a rectangle shape, link the shape to the formatted cell using the LinkedCell property, retrieve the address with LinkedCell and GetLinkedCell, and save the workbook to verify persistence.
class ShapeLinkedCellExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a date value into cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.PutValue(DateTime.Now);

            // Apply a custom date number format (e.g., "dd-MMM-yyyy")
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset (pixels),
            // upper left column offset (pixels), width (pixels), height (pixels)
            Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);

            // Link the shape to the cell containing the date
            rect.LinkedCell = "$A$1";

            // Verify the linked cell by reading the property
            Console.WriteLine("Shape's LinkedCell: " + rect.LinkedCell);

            // Optionally, retrieve the linked cell using GetLinkedCell (non‑R1C1, non‑local)
            string linkedFormula = rect.GetLinkedCell(false, false);
            Console.WriteLine("GetLinkedCell returned: " + linkedFormula);

            // Save the workbook to verify the shape persists
            workbook.Save("ShapeLinkedCellDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
