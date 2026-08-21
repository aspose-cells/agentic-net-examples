// Title: Aspose.Cells for .NET – Link a Shape to a Formula Cell for Live Value Display
// Description: Demonstrates how to add a rectangle shape to a worksheet, link its text to a formula cell using SetLinkedCell, refresh the displayed value with UpdateSelectedValue, and save the workbook. The shape automatically reflects any changes to the underlying formula.
// Keywords: Aspose.Cells C# shape linked cell | SetLinkedCell example | UpdateSelectedValue method | dynamic shape text Aspose.Cells | live formula result in shape | rectangle shape Excel automation | C# Excel shape binding | global developers Aspose.Cells
// Common Searches: how to bind a shape to a cell in Aspose.Cells | display formula result in a shape using Aspose.Cells for .NET | Aspose.Cells SetLinkedCell usage | refresh shape text after formula change Aspose | C# add rectangle shape linked to cell
// Developer Intent: Enable a shape’s displayed text to automatically show the current result of a formula cell, eliminating manual updates.
// Use Cases: Dashboard KPI: a shape shows a calculated metric that updates when source data changes. | Report label: a shape displays a total or tax amount derived from a formula without extra code. | Invoice template: a shape reflects a computed discount linked to a formula cell.
// AI Prompts: Generate C# code that adds multiple shapes, each linked to a different formula cell, and updates them after data modifications. | Explain the parameters of SetLinkedCell, including locale‑aware linking and absolute vs. relative references. | Show how to change the formula in a linked cell and immediately refresh the shape’s text using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a worksheet, link its text to a formula cell using SetLinkedCell, refresh the displayed value with UpdateSelectedValue, and save the workbook. The shape automatically reflects any changes to the underlying formula.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a formula in cell B2 (e.g., double the value of A1)
        worksheet.Cells["B2"].Formula = "=A1*2";

        // Add a rectangle shape that will display the linked cell value
        // Parameters: upper left row, upper left column, width, height, upper left offset X, offset Y
        Shape linkedShape = worksheet.Shapes.AddRectangle(2, 2, 120, 30, 0, 0);

        // Link the shape's displayed text to cell B2 (A1 style, locale‑aware)
        linkedShape.SetLinkedCell("$B$2", false, true);

        // Refresh the shape so it shows the current value of the linked cell
        linkedShape.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("LinkedShapeDemo.xlsx");
    }
}
