// Title: C# – Add a Rectangle Shape Linked to a Cell with a Nested Formula using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill cells A1 and B1, set a nested formula in C1, insert a rectangle shape, link the shape to C1, recalculate formulas, retrieve the computed value, and save the file. The example verifies that the linked shape reflects the final formula result.
// Keywords: Aspose.Cells shape linking | C# rectangle shape linked cell | nested formula Aspose.Cells | verify linked shape value | Aspose.Cells for .NET example | calculate formulas programmatically | Excel shape to cell reference
// Common Searches: how to link a shape to a formula cell in Aspose.Cells | Aspose.Cells C# example linking rectangle to cell | retrieve calculated value from linked shape Aspose.Cells | link shape to cell with nested formula .NET | Aspose.Cells shape linked cell verification
// Developer Intent: Link a worksheet shape to a cell that contains a nested formula and confirm that the shape reflects the calculated result.
// Use Cases: Building interactive dashboards where shapes display live totals from complex formulas. | Automating report generation that adds visual markers tied to calculated cells for quick review. | Creating templates that require shapes to stay synchronized with underlying formula results after data updates.
// AI Prompts: Generate C# code with Aspose.Cells that adds a circle shape linked to cell D5 containing an IF‑nested formula and prints the linked value. | Show how to update a linked shape after changing source cells and re‑calculating formulas using Aspose.Cells for .NET. | Explain how to programmatically obtain the linked cell address and its evaluated value for any shape in a worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // Demonstrates how to create a workbook, fill cells A1 and B1, set a nested formula in C1, insert a rectangle shape, link the shape to C1, recalculate formulas, retrieve the computed value, and save the file. The example verifies that the linked shape reflects the final formula result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate data used in the nested formula
                cells["A1"].PutValue(5);
                cells["B1"].PutValue(10);

                // Set a nested formula in C1 with an initial displayed value
                cells["C1"].SetFormula("=SUM(A1,B1)*2", 30);

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape rect = sheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);

                // Link the shape to the cell containing the formula (C1)
                rect.LinkedCell = "$C$1";

                // Calculate all formulas so C1 gets its final result
                workbook.CalculateFormula();

                // Retrieve the linked cell address (remove $ to use with Cells collection)
                string linkedAddress = rect.LinkedCell.Replace("$", string.Empty); // "C1"

                // Get the calculated value from the linked cell
                Cell linkedCell = sheet.Cells[linkedAddress];
                object result = linkedCell.Value;

                // Output verification result
                Console.WriteLine($"Shape is linked to cell {rect.LinkedCell} with calculated value: {result}");

                // Save the workbook (lifecycle rule: save)
                string outputPath = "ShapeLinkedFormulaDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
