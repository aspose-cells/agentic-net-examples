// Title: Link a Label Shape to a Currency‑Formatted Cell and Verify Text with Aspose.Cells for .NET
// Description: This C# example creates a workbook, writes a numeric value to B2, applies a custom currency format ($#,##0.00), adds a label shape, links the shape to the formatted cell using SetLinkedCell, forces a refresh, reads the shape's Text property to confirm the formatted output, and saves the file.
// Keywords: Aspose.Cells | .NET | C# | label shape | SetLinkedCell | custom number format | currency format | shape text verification | Excel shape linking | Aspose.Cells example
// Common Searches: Aspose.Cells link shape to cell | SetLinkedCell currency format C# | verify shape text after linking Aspose.Cells | add label shape linked to cell .NET | custom number format for linked shape Aspose
// Developer Intent: Add a label shape, bind it to a currency‑formatted cell, refresh the shape, and confirm the displayed text matches the cell format.
// Use Cases: Create dashboards where shapes automatically show monetary values from cells. | Generate printable reports that keep shape captions in sync with formatted cell data. | Build interactive Excel workbooks where shape captions reflect locale‑aware number formats.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape linked to cell C5 using a date format and refresh its displayed value. | Explain how the three parameters of SetLinkedCell control locale awareness, absolute reference, and formula evaluation. | Create a unit test that asserts a linked label shape's Text property equals the custom formatted value of its source cell.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, writes a numeric value to B2, applies a custom currency format ($#,##0.00), adds a label shape, links the shape to the formatted cell using SetLinkedCell, forces a refresh, reads the shape's Text property to confirm the formatted output, and saves the file.
class ShapeLinkedCellExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a numeric value in cell B2
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue(1234.56);

            // Apply a custom number format for currency (e.g., $1,234.56)
            Style style = targetCell.GetStyle();
            style.Custom = "$#,##0.00";
            targetCell.SetStyle(style);

            // Add a label shape to the worksheet (positioned at row 4, column 2)
            // Parameters: upperRow, upperColumn, top, left, height, width
            Label label = (Label)sheet.Shapes.AddLabel(4, 2, 0, 0, 100, 30);

            // Link the label's displayed value to cell B2 (A1 style, locale‑aware)
            label.SetLinkedCell("$B$2", false, true);

            // Force the shape to refresh its displayed value from the linked cell
            label.UpdateSelectedValue();

            // Retrieve the displayed text from the label
            string shapeText = label.Text;
            Console.WriteLine("Shape displays: " + shapeText); // Expected: $1,234.56

            // Save the workbook to a file
            string outputPath = "ShapeLinkedCellCurrency.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
