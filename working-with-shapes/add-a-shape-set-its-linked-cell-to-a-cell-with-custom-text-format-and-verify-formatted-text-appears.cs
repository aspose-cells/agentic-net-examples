// Title: Link a Rectangle Shape to a Formatted Cell and Verify Text with Aspose.Cells for .NET
// Description: Creates a workbook, writes "Custom Formatted Text" to cell A1, applies a red bold 14‑pt style, adds a rectangle shape, sets its LinkedCell to A1, updates the shape value, prints shape.Text to confirm the formatted content, and saves the file as ShapeLinkedCellDemo.xlsx.
// Keywords: Aspose.Cells | C# | shape linked cell | rectangle shape | custom cell style | LinkedCell property | UpdateSelectedValue | verify shape text | Excel automation | styled cell value in shape
// Common Searches: Aspose.Cells link shape to cell with formatting | How to set LinkedCell for a rectangle in .NET | Retrieve shape.Text after linking to a styled cell | Update shape value after changing cell style Aspose.Cells | C# example linking shape to formatted Excel cell
// Developer Intent: Connect a shape to a styled cell and confirm the shape displays the same formatted text.
// Use Cases: Dynamic dashboards where shapes reflect styled cell values for visual emphasis. | Automated report templates that use linked shapes to show updated cell content. | Excel‑based UI components that sync shape captions with formatted data cells.
// AI Prompts: Generate C# code that links a rectangle shape to a formatted cell and verifies the displayed text using Aspose.Cells. | Show how to refresh a shape after changing the font color, weight, or size of its linked cell in Aspose.Cells for .NET. | Explain the steps to use LinkedCell and UpdateSelectedValue to keep shape text in sync with a styled Excel cell.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes "Custom Formatted Text" to cell A1, applies a red bold 14‑pt style, adds a rectangle shape, sets its LinkedCell to A1, updates the shape value, prints shape.Text to confirm the formatted content, and saves the file as ShapeLinkedCellDemo.xlsx.
class ShapeLinkedCellDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a custom formatted text in cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Custom Formatted Text");

        // Apply a custom style (e.g., red font, bold, 14pt)
        Style style = cell.GetStyle();
        style.Font.Color = Color.Red;
        style.Font.IsBold = true;
        style.Font.Size = 14;
        cell.SetStyle(style);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 50);

        // Link the shape to the formatted cell (A1)
        shape.LinkedCell = "A1";

        // Update the shape's displayed value based on the linked cell
        shape.UpdateSelectedValue();

        // Verify that the shape's text reflects the cell's value
        Console.WriteLine("Shape Text: " + shape.Text);

        // Save the workbook to a file
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}
