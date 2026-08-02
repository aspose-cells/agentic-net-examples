// Title: Link a Rectangle Shape to a Cell with a Nested Formula and Verify the Result – Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a workbook, fill cells A1‑B2, set a nested formula in C3 (SUM(A1:B1) * AVERAGE(A2:B2)), calculate the workbook, add a rectangle shape, link the shape to C3 using SetLinkedCell, update the shape's selected value, read the linked value, and confirm that the shape displays the same calculated result before saving the file.
// Keywords: Aspose.Cells C# shape linking | rectangle shape linked to cell | SetLinkedCell example | nested formula verification | UpdateSelectedValue Aspose.Cells | Excel shape cell synchronization | C# Aspose.Cells tutorial
// Common Searches: how to link a shape to a formula cell using Aspose.Cells .NET | verify shape displays calculated value in Aspose.Cells C# | Aspose.Cells SetLinkedCell rectangle example | update shape selected value after workbook.CalculateFormula | C# code to link Excel shape to cell result
// Developer Intent: Connect a rectangle shape to a cell that contains a nested formula and programmatically confirm that the shape shows the formula's final value.
// Use Cases: Build dynamic dashboards where shapes automatically reflect totals, averages, or other computed metrics. | Generate Excel reports with visual indicators (shapes) that stay in sync with underlying calculations. | Create automated tests to ensure shape‑cell links remain accurate after formula recalculation.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape, links it to a cell containing a nested formula, updates the shape, and validates the linked value matches the cell's result. | Explain the role of SetLinkedCell and UpdateSelectedValue when synchronizing a shape with a calculated cell in Aspose.Cells for .NET. | Provide a step‑by‑step guide to verify that a shape linked to a formula cell reflects the final computed value after calling workbook.CalculateFormula().

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // This C# example demonstrates how to create a workbook, fill cells A1‑B2, set a nested formula in C3 (SUM(A1:B1) * AVERAGE(A2:B2)), calculate the workbook, add a rectangle shape, link the shape to C3 using SetLinkedCell, update the shape's selected value, read the linked value, and confirm that the shape displays the same calculated result before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Create a new workbook ----------
                Workbook workbook = new Workbook();                     // create
                Worksheet sheet = workbook.Worksheets[0];              // first worksheet

                // ---------- 2. Prepare data for the nested formula ----------
                // A1 = 10, B1 = 20, A2 = 5, B2 = 15
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B1"].PutValue(20);
                sheet.Cells["A2"].PutValue(5);
                sheet.Cells["B2"].PutValue(15);

                // ---------- 3. Set a nested formula in cell C3 ----------
                // Formula: =SUM(A1:B1) * AVERAGE(A2:B2)
                sheet.Cells["C3"].SetFormula("=SUM(A1:B1)*AVERAGE(A2:B2)", null);

                // Calculate all formulas so C3 contains the final result.
                workbook.CalculateFormula();

                // Retrieve the calculated result for verification.
                double calculatedResult = sheet.Cells["C3"].DoubleValue;
                Console.WriteLine($"Calculated result in C3: {calculatedResult}");

                // ---------- 4. Add a rectangle shape ----------
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape rect = sheet.Shapes.AddRectangle(5, 2, 0, 0, 150, 50);

                // ---------- 5. Link the shape to the cell containing the formula ----------
                // Use SetLinkedCell(string formula, bool isR1C1, bool isLocal)
                rect.SetLinkedCell("$C$3", false, false);

                // Ensure the shape's selected value is updated from the linked cell.
                sheet.Shapes.UpdateSelectedValue();

                // ---------- 6. Verify that the shape reflects the cell's final result ----------
                // Get the linked cell address from the shape.
                string linkedAddress = rect.GetLinkedCell(false, false); // returns absolute address like "$C$3"

                // Remove the leading '$' characters to use with Cells collection.
                string cleanAddress = linkedAddress.Replace("$", "");

                // Read the value from the linked cell.
                double shapeLinkedValue = sheet.Cells[cleanAddress].DoubleValue;

                Console.WriteLine($"Shape linked cell address: {linkedAddress}");
                Console.WriteLine($"Value read via shape link: {shapeLinkedValue}");

                // Simple verification
                if (Math.Abs(calculatedResult - shapeLinkedValue) < 1e-9)
                    Console.WriteLine("Verification succeeded: Shape reflects the final result.");
                else
                    Console.WriteLine("Verification failed: Mismatch between shape and cell values.");

                // ---------- 7. Save the workbook ----------
                workbook.Save("ShapeLinkedFormulaDemo.xlsx"); // save
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
