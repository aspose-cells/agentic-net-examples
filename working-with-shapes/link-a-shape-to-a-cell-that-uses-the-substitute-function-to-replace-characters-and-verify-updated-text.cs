// Title: Aspose.Cells .NET: Link a TextBox Shape to a Cell with SUBSTITUTE Formula and Auto‑Update Text
// Description: Demonstrates how to add a TextBox shape, link it to a cell containing a SUBSTITUTE formula, recalculate the workbook, and verify that the shape’s displayed text updates when the source cell changes. Includes error handling and saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | SetLinkedCell | linked shape | textbox shape | SUBSTITUTE formula | dynamic shape text | workbook recalculation | cell formula evaluation | example code
// Common Searches: Aspose.Cells link shape to cell example | C# set linked textbox to formula cell | update shape text after cell value change Aspose.Cells | SUBSTITUTE function with linked shape .NET | how to refresh linked shape in Aspose.Cells
// Developer Intent: Show how to bind a TextBox shape to a cell that uses the SUBSTITUTE function and ensure the shape reflects formula results after any source‑cell modification.
// Use Cases: Display transformed data (e.g., character replacement) directly on a shape for visual reports. | Create dashboards where shapes automatically show the latest formula outcomes without manual refresh. | Generate workbooks with labels or annotations that stay synchronized with underlying cell values.
// AI Prompts: Write C# code with Aspose.Cells to link a TextBox shape to a cell that contains a SUBSTITUTE formula and verify the shape text after changing the source cell. | Explain the role of Shape.SetLinkedCell and Workbook.CalculateFormula in keeping linked shapes up‑to‑date. | Suggest best practices for error handling when linking shapes to formula cells in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkShapeExample
{
    // Demonstrates how to add a TextBox shape, link it to a cell containing a SUBSTITUTE formula, recalculate the workbook, and verify that the shape’s displayed text updates when the source cell changes. Includes error handling and saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, height, width
                Shape shape = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);

                // Link the shape to cell B1 (the shape will display the value of B1)
                shape.SetLinkedCell("B1", false, false);

                // Put initial value in A1
                sheet.Cells["A1"].PutValue("apple");

                // Set formula in B1 that substitutes character 'a' with 'b' in A1
                sheet.Cells["B1"].Formula = "=SUBSTITUTE(A1,\"a\",\"b\")";

                // Recalculate the workbook so that formulas are evaluated
                workbook.CalculateFormula();

                // Verify that the shape text reflects the substituted value
                Console.WriteLine("After first calculation:");
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
                Console.WriteLine($"Cell B1 (formula result): {sheet.Cells["B1"].StringValue}");
                Console.WriteLine($"Shape linked text: {shape.Text}");

                // Change the source value in A1
                sheet.Cells["A1"].PutValue("banana");

                // Recalculate again
                workbook.CalculateFormula();

                // Verify updated text in the shape
                Console.WriteLine("\nAfter updating A1:");
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
                Console.WriteLine($"Cell B1 (formula result): {sheet.Cells["B1"].StringValue}");
                Console.WriteLine($"Shape linked text: {shape.Text}");

                // Save the workbook
                workbook.Save("LinkedShapeWithSubstitute.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
