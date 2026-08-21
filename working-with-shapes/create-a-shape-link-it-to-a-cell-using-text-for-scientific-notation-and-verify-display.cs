// Title: Link a TextBox shape to a cell formatted with TEXT() scientific notation in Aspose.Cells for .NET
// Description: Demonstrates how to place a number in A1, format it as scientific notation with the TEXT function in B1, recalculate the workbook, add a TextBox shape, link the shape to B1 using SetLinkedCell, update the shape's displayed text, and save the file.
// Keywords: Aspose.Cells | C# | .NET | TextBox shape | linked cell | TEXT function | scientific notation | SetLinkedCell | shape update | formula calculation
// Common Searches: Aspose.Cells link textbox to cell | display scientific notation in shape | SetLinkedCell parameters C# | update shape text after formula | link shape to TEXT formula Aspose
// Developer Intent: Create a textbox shape that automatically shows the scientific‑notation string produced by a TEXT formula.
// Use Cases: Financial or scientific reports where numbers need exponential display inside annotated shapes. | Dynamic dashboards that keep shape captions synchronized with formatted cell values. | Template designs that use linked shapes to reflect locale‑aware formatted data without manual updates.
// AI Prompts: Generate C# code with Aspose.Cells to add a TextBox, link it to a cell using TEXT() for scientific notation, recalculate, and refresh the shape text. | Explain the effect of the isR1C1 and isLocal flags in SetLinkedCell when linking a shape to an A1‑style address. | Outline steps to verify that a shape linked to a TEXT‑formatted cell displays the correct exponential value after workbook calculation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // Demonstrates how to place a number in A1, format it as scientific notation with the TEXT function in B1, recalculate the workbook, add a TextBox shape, link the shape to B1 using SetLinkedCell, update the shape's displayed text, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value in A1
            worksheet.Cells["A1"].PutValue(12345);

            // Use the TEXT function to format the number in scientific notation in B1
            worksheet.Cells["B1"].Formula = "TEXT(A1,\"0.00E+00\")";

            // Recalculate formulas so B1 contains the formatted text
            workbook.CalculateFormula();

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels), upper left offset (pixels), width, height
            TextBox shape = (TextBox)worksheet.Shapes.AddTextBox(2, 2, 0, 0, 200, 30);

            // Link the shape to cell B1 (the cell with scientific notation text)
            // isR1C1 = false (A1 style), isLocal = true (locale aware)
            shape.SetLinkedCell("$B$1", false, true);

            // Update the shape's displayed value from the linked cell
            shape.UpdateSelectedValue();

            // Verify the displayed text
            Console.WriteLine("Shape Text (should be scientific notation): " + shape.Text);

            // Save the workbook
            workbook.Save("ShapeLinkedCellScientificNotation.xlsx");
        }
    }
}
