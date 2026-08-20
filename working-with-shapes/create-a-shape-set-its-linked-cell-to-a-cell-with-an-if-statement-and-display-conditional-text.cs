// Title: Aspose.Cells .NET: Link a Rectangle Shape to a Cell with IF Formula and Display Conditional Text
// Description: Demonstrates how to create a workbook, add a rectangle shape, link it to cell A1, assign an IF formula that returns "High" or "Low" based on B1, recalculate formulas, refresh the shape’s displayed value, and save the file as output.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# shape linked cell | rectangle shape IF formula | conditional text in shape | UpdateSelectedValue | CalculateFormula | .NET workbook automation | dynamic label Excel
// Common Searches: Aspose.Cells link shape to cell and show IF result | Refresh shape text after formula calculation .NET | How to display conditional text in a shape using Aspose.Cells | Set linked cell for rectangle shape in C# | Update shape value after workbook.CalculateFormula
// Developer Intent: Create a shape, bind it to a cell containing an IF statement, and have the shape automatically show the formula’s outcome.
// Use Cases: Status indicator that switches between "High" and "Low" when a metric exceeds a threshold. | Dashboard widgets that reflect real‑time calculations without manual updates. | Conditional annotations in reports where text changes based on underlying data.
// AI Prompts: Generate C# code that adds multiple linked shapes, each with its own IF formula, and updates all shapes after recalculating the workbook with Aspose.Cells. | Explain the steps to programmatically refresh a shape’s displayed text when the linked cell’s formula result changes in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a rectangle shape, link it to cell A1, assign an IF formula that returns "High" or "Low" based on B1, recalculate formulas, refresh the shape’s displayed value, and save the file as output.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape (acts like a label) at row 2, column 2 (C3)
        // Parameters: type, upper left row, upper left column, top, left, bottom, right
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 30);

        // Link the shape to cell A1
        shape.SetLinkedCell("A1", false, false);

        // Put an IF formula in A1 that shows "High" if B1>10, otherwise "Low"
        sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

        // Example value for B1 to trigger the condition
        sheet.Cells["B1"].PutValue(15);

        // Recalculate all formulas so the IF result is evaluated
        workbook.CalculateFormula();

        // Update the shape's displayed text based on the linked cell value
        shape.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
