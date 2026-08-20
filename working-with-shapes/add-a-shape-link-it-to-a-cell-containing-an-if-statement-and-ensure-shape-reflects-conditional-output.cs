// Title: Aspose.Cells for .NET – C# Example: Link a Rectangle Shape to an IF Formula Cell
// Description: Creates a workbook, writes a value to B1, adds an IF formula in A1, inserts a rectangle shape, links the shape to A1, and forces the shape to display the formula's result, updating automatically when the cell changes.
// Keywords: Aspose.Cells shape linking | C# rectangle shape Excel | linked cell IF formula | dynamic shape text Aspose | update shape value programmatically | Aspose.Cells .NET tutorial | Excel shape conditional display
// Common Searches: how to bind a shape to a cell with an IF statement using Aspose.Cells | Aspose.Cells C# update shape text when linked cell formula changes | link rectangle shape to Excel cell and show result | Aspose.Cells set linked cell for shape and refresh value | dynamic shape content based on worksheet formula Aspose
// Developer Intent: Create a shape that automatically shows the outcome of an IF formula linked to a worksheet cell.
// Use Cases: Dashboard indicator that switches between "High" and "Low" without manual edits. | Report generation where status shapes reflect pass/fail results derived from formulas. | Interactive Excel templates that display conditional messages inside shapes.
// AI Prompts: Generate C# code with Aspose.Cells to link a shape to a VLOOKUP cell and refresh its displayed value. | Show how to change a shape's fill color based on the IF result of its linked cell in Aspose.Cells. | Provide an example that links multiple shapes to different formula cells and updates them after workbook modifications.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes a value to B1, adds an IF formula in A1, inserts a rectangle shape, links the shape to A1, and forces the shape to display the formula's result, updating automatically when the cell changes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set up data for the IF condition
            // B1 will hold a numeric value that determines the IF result
            sheet.Cells["B1"].PutValue(15); // Change this value to test different outcomes

            // Place an IF formula in A1 that returns "High" if B1 > 10, otherwise "Low"
            sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape rectangle = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 30);

            // Link the shape to the cell containing the IF formula (A1)
            // The two boolean parameters indicate whether to update the shape when the linked cell changes
            rectangle.SetLinkedCell("A1", true, true);

            // Force the shape to read the current value of the linked cell
            rectangle.UpdateSelectedValue();

            // Optionally, set some visual properties for clarity
            rectangle.Placement = PlacementType.FreeFloating;
            rectangle.Text = "Result:"; // Prefix text; the linked value will appear after this text

            // Define output file path
            string outputPath = "ShapeLinkedToIf.xlsx";

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
