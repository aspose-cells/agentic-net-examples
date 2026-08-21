// Title: Aspose.Cells for .NET: Add a Rectangle Shape Linked to a Currency‑Formatted Cell (C#)
// Description: Creates a workbook, writes 1234.56 to C5, applies the custom format "$#,##0.00", adds a rectangle shape, links it to the cell with A1 notation, refreshes the shape to show the formatted value, and saves the file.
// Keywords: Aspose.Cells | C# shape linking | rectangle shape | custom currency format | SetLinkedCell | UpdateSelectedValue | linked cell display | Excel automation | financial dashboard
// Common Searches: Aspose.Cells link shape to cell C# | display custom number format in linked shape Aspose.Cells | SetLinkedCell A1 notation example | refresh shape text after linking Aspose.Cells | add rectangle shape programmatically .NET
// Developer Intent: Bind a rectangle shape to a cell that uses a custom currency format and ensure the shape displays the formatted value.
// Use Cases: Build a financial dashboard where shapes automatically reflect currency values from cells. | Create interactive reports with shapes that update their displayed text when the underlying formatted cell changes. | Automate workbook generation where visual elements are bound to formatted data for quick visual summaries.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape, formats a cell as currency, links the shape to the cell, and updates the shape's displayed text. | Explain the effect of the isR1C1 and isLocal parameters in SetLinkedCell and how they influence the shape's link to a formatted cell. | Provide troubleshooting steps when a linked shape does not show the custom currency format after calling UpdateSelectedValue.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes 1234.56 to C5, applies the custom format "$#,##0.00", adds a rectangle shape, links it to the cell with A1 notation, refreshes the shape to show the formatted value, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a numeric value into cell C5
        Cell targetCell = sheet.Cells["C5"];
        targetCell.PutValue(1234.56);

        // Apply a custom currency number format to the cell
        Style style = targetCell.GetStyle();
        style.Custom = "$#,##0.00";
        targetCell.SetStyle(style);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset (pixels), upper left offset (pixels), width (pixels), height (pixels)
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 40);

        // Link the shape to the cell C5 using A1 style (isR1C1 = false) and locale aware (isLocal = true)
        shape.SetLinkedCell("$C$5", false, true);

        // Refresh the shape so it displays the linked cell's formatted text
        shape.UpdateSelectedValue();

        // Verify the linked cell address
        Console.WriteLine("Linked cell: " + shape.LinkedCell);

        // Save the workbook
        workbook.Save("ShapeLinkedCellCurrency.xlsx");
    }
}
