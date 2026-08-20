// Title: Aspose.Cells for .NET – Link a Rectangle Shape to a Custom‑Formatted Cell and Display the Formatted Value
// Description: Demonstrates how to create a workbook, apply a custom currency number format to cell B2, add a rectangle shape, set its LinkedCell to B2, invoke UpdateSelectedValue so the shape shows the formatted text, and save the result as ShapeLinkedCellFormatted.xlsx.
// Keywords: Aspose.Cells | C# | shape linked cell | custom number format | UpdateSelectedValue | rectangle shape | Excel automation | formatted value in shape | .NET example
// Common Searches: Aspose.Cells link shape to cell with custom format | C# rectangle shape shows formatted cell value | How to use UpdateSelectedValue in Aspose.Cells | Display currency format in linked shape Aspose | Aspose.Cells shape LinkedCell example
// Developer Intent: Show how to bind a shape to a cell that uses a custom number format and make the shape render the formatted text automatically.
// Use Cases: Financial dashboards where shapes reflect live currency values from the worksheet. | Printable reports that use shapes as dynamic placeholders for formatted dates or amounts. | Interactive spreadsheets where shape captions update instantly to match custom‑styled cell data.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape linked to a cell formatted as currency and verifies the shape displays the formatted amount. | Provide an example linking multiple shapes to cells with different custom formats (date, percentage, currency) and ensure each shape shows the correct formatted value. | Explain the role of UpdateSelectedValue when a shape is linked to a styled cell in Aspose.Cells, including when the cell format changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, apply a custom currency number format to cell B2, add a rectangle shape, set its LinkedCell to B2, invoke UpdateSelectedValue so the shape shows the formatted text, and save the result as ShapeLinkedCellFormatted.xlsx.
class ShapeLinkedCellDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Set up a cell with a custom number format
        // -------------------------------------------------
        // Target cell B2
        Cell targetCell = worksheet.Cells["B2"];
        // Put a numeric value
        targetCell.PutValue(1234.56);
        // Create a style with a custom format (e.g., currency)
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "$#,##0.00";
        // Apply the style to the cell
        targetCell.SetStyle(customStyle);

        // -------------------------------------------------
        // 2. Add a rectangle shape and link it to the cell
        // -------------------------------------------------
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftPixel, upperLeftPixel, width, height
        Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 50);
        // Link the shape to cell B2
        rectangle.LinkedCell = "B2";
        // Ensure the shape reflects the linked cell's current value/format
        rectangle.UpdateSelectedValue();

        // -------------------------------------------------
        // 3. Save the workbook
        // -------------------------------------------------
        workbook.Save("ShapeLinkedCellFormatted.xlsx");
    }
}
