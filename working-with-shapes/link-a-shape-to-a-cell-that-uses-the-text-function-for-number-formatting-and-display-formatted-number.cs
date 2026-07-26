// Title: C# – Link a Shape to a TEXT‑formatted Cell in Aspose.Cells (.NET)
// Description: Demonstrates how to place a raw number in A1, format it with the TEXT function in B1, add a rectangle shape, link the shape to the formatted cell, refresh the shape’s displayed value, and save the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape linking | C# shape linked cell | TEXT function number formatting | formatted value in shape | update shape from formula | Aspose.Cells rectangle shape
// Common Searches: Aspose.Cells link shape to cell with TEXT formula | display formatted number in shape C# | how to refresh linked shape after cell change Aspose.Cells | shape linked cell number formatting .NET | add rectangle shape linked to formatted cell Aspose
// Developer Intent: The developer needs a shape to show a number that is formatted by the TEXT function, updating automatically when the source cell changes.
// Use Cases: Dashboard label that shows a currency total with thousand separators. | Dynamic report header where a shape reflects a formula‑based, formatted value. | Printable invoice where shapes display formatted subtotals derived from raw data.
// AI Prompts: Write C# code with Aspose.Cells to link a rectangle shape to a cell that uses the TEXT function for number formatting and update the shape’s text. | Show how to refresh a shape after modifying the source cell that contains a TEXT formula in Aspose.Cells for .NET. | Explain step‑by‑step how to create a formatted number with TEXT, link a shape to that cell, and save the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to place a raw number in A1, format it with the TEXT function in B1, add a rectangle shape, link the shape to the formatted cell, refresh the shape’s displayed value, and save the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a raw numeric value into cell A1
        worksheet.Cells["A1"].PutValue(1234.567);

        // Use the TEXT function in cell B1 to format the number from A1
        // The format "#,##0.00" will display the number with thousand separators and two decimal places
        worksheet.Cells["B1"].Formula = "TEXT(A1, \"#,##0.00\")";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, width (pixels), height (pixels), offsetX, offsetY
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 150, 50, 0, 0);

        // Link the shape to the formatted cell B1
        shape.LinkedCell = "$B$1";

        // Update the shape so it reflects the current value of the linked cell
        shape.UpdateSelectedValue();

        // Save the workbook to a file
        workbook.Save("ShapeLinkedToFormattedCell.xlsx");
    }
}
