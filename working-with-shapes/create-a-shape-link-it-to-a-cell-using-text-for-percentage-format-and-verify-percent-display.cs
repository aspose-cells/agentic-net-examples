// Title: Aspose.Cells for .NET – Link a Label Shape to a Percentage‑Formatted Cell and Verify the Displayed Text (C#)
// Description: This example creates a workbook, writes 0.5 to cell A1, applies the built‑in percent format (index 10), adds a label shape, links the shape to A1, refreshes the shape value, reads the Text property, and confirms that the shape shows "50%". The workbook is saved for visual inspection.
// Keywords: Aspose.Cells | C# | label shape | linked cell | percentage format | SetLinkedCell | UpdateSelectedValue | Excel shape text | cell formatting | Aspose.Cells .NET example | automated Excel dashboard
// Common Searches: Aspose.Cells link shape to cell percentage | C# label shape display formatted cell value | SetLinkedCell percentage format Aspose.Cells | UpdateSelectedValue shape text verification | How to show 50% in a linked shape using Aspose.Cells
// Developer Intent: Link a label shape to a cell formatted as a percentage and confirm that the shape displays the correct "%" text.
// Use Cases: Create dynamic dashboards where shapes automatically reflect percentage values from worksheet cells. | Generate Excel reports that use linked shapes to present formatted numbers without manual updates. | Implement automated tests that validate shape text matches the cell's number format. | Build interactive Excel templates where shapes act as live data labels.
// AI Prompts: Generate C# code with Aspose.Cells that adds a label shape linked to cell B2, applies a currency format, and prints the shape's displayed text. | Explain the relationship between SetLinkedCell and UpdateSelectedValue for reflecting cell formatting in a linked shape. | Provide step‑by‑step instructions to verify that a shape linked to a percentage‑formatted cell shows the correct "%" string.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeLinkedCellPercentDemoApp
{
    // This example creates a workbook, writes 0.5 to cell A1, applies the built‑in percent format (index 10), adds a label shape, links the shape to A1, refreshes the shape value, reads the Text property, and confirms that the shape shows "50%". The workbook is saved for visual inspection.
    class ShapeLinkedCellPercentDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set a numeric value in cell A1
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue(0.5); // 50%

                // Apply percentage number format (built‑in format index 10)
                Style style = cell.GetStyle();
                style.Number = 10; // Percent format
                cell.SetStyle(style);

                // Add a label shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, height, width (pixels)
                Label label = worksheet.Shapes.AddLabel(2, 0, 0, 0, 100, 200);

                // Link the shape to cell A1 (the shape will display the cell's formatted text)
                label.SetLinkedCell("A1", false, false);
                // Refresh the shape's displayed value
                label.UpdateSelectedValue();

                // Retrieve the text displayed by the shape
                string shapeText = label.Text;

                // Output the shape's text to verify percentage display
                Console.WriteLine("Shape linked text: " + shapeText); // Expected output: "50%"

                // Save the workbook (optional, to view the result in Excel)
                string outputPath = "ShapeLinkedCellPercentDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
