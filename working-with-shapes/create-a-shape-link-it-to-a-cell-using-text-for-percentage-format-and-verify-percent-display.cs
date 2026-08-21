// Title: Link a Label Shape to a Percentage‑Formatted Cell and Verify Display with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, writes 0.25 to cell A1, applies the built‑in percentage format (style 10), confirms the cell is recognized as a percent, adds a label shape, links it to A1 with automatic updates, forces a refresh, reads the shape text (expected "25%"), and saves the file as ShapeLinkedPercent.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape linked cell | percentage format | label shape | SetLinkedCell | UpdateSelectedValue | verify shape text | dynamic dashboard label
// Common Searches: Aspose.Cells link shape to cell percentage | C# shape displays cell as percent | How to sync label shape with formatted cell in Aspose.Cells | Verify linked shape text shows 25% in .NET | Update shape text after cell format change Aspose
// Developer Intent: The developer needs to bind a worksheet shape to a cell that uses percentage formatting and confirm that the shape shows the correctly formatted percent value.
// Use Cases: Create live dashboard labels that automatically reflect percentage calculations. | Generate printable reports where shape captions mirror cell percentages without manual editing. | Synchronize annotations on charts or diagrams with underlying percentage data stored in worksheets.
// AI Prompts: Write C# code using Aspose.Cells to link a label shape to a percent‑formatted cell and validate the displayed text. | Explain the role of Shape.SetLinkedCell and Shape.UpdateSelectedValue for keeping shape text in sync with a percentage cell. | Provide error‑handling best practices for confirming that a linked shape displays the expected percent after cell updates.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // This example creates a workbook, writes 0.25 to cell A1, applies the built‑in percentage format (style 10), confirms the cell is recognized as a percent, adds a label shape, links it to A1 with automatic updates, forces a refresh, reads the shape text (expected "25%"), and saves the file as ShapeLinkedPercent.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a decimal value into cell A1
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue(0.25); // 25%

                // Apply percentage number format to the cell (built‑in format 10)
                Style style = cell.GetStyle();
                style.Number = 10;
                cell.SetStyle(style);

                // Verify that the cell style is recognized as a percent format
                Console.WriteLine("Cell A1 IsPercent: " + style.IsPercent); // should be True

                // Add a label shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, height (pixels), width (pixels)
                Shape shape = worksheet.Shapes.AddLabel(2, 2, 0, 0, 100, 30);

                // Link the shape to cell A1 (true for update on change, true for update on load)
                shape.SetLinkedCell("A1", true, true);

                // Force the shape to refresh its displayed value from the linked cell
                shape.UpdateSelectedValue();

                // Retrieve the text displayed in the shape
                string shapeText = shape.Text;

                Console.WriteLine("Shape linked text: " + shapeText); // should display "25%"

                // Save the workbook (ensure the directory exists)
                string outputPath = "ShapeLinkedPercent.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
