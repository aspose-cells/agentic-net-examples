// Title: Create a rectangle shape linked to an IF formula for conditional text using Aspose.Cells for .NET (C#)
// Description: Shows how to add a rectangle shape to a worksheet, link it to cell A1, assign an IF formula that returns "High" or "Low" based on the value in B1, set a sample value, refresh the shape with UpdateSelectedValue, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# shape linked cell | IF formula shape | conditional text shape | UpdateSelectedValue | rectangle shape Aspose | dynamic label Excel | Aspose.Cells .NET example | linked shape formula | Excel dashboard label
// Common Searches: Aspose.Cells link shape to cell | display IF result in shape Aspose.Cells | refresh shape after cell change Aspose | C# rectangle shape with conditional text | Aspose.Cells dynamic label example
// Developer Intent: Link a worksheet shape to a cell that contains an IF formula so the shape automatically shows the formula’s evaluated text.
// Use Cases: Show a status label that reads "High" or "Low" depending on a numeric cell value. | Create dashboard indicators that update their displayed text when source data changes. | Provide dynamic captions for charts or tables based on conditional logic. | Build interactive reports where shape text reflects real‑time calculations.
// AI Prompts: Generate Aspose.Cells C# code to create multiple shapes, each linked to a different IF formula for a status dashboard. | Explain how to programmatically refresh linked shapes after modifying source cells in a workbook. | Provide an example of linking a shape to a cell with a nested IF formula and customizing the shape’s color based on the result.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to a worksheet, link it to cell A1, assign an IF formula that returns "High" or "Low" based on the value in B1, set a sample value, refresh the shape with UpdateSelectedValue, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (acts as a label) to the worksheet
            // Parameters: drawing type, upper left row, upper left column,
            // top offset (pixels), left offset (pixels), height (pixels), width (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // top offset in pixels
                0,   // left offset in pixels
                100, // height in pixels
                200  // width in pixels
            );

            // Link the shape to cell A1 (the cell that will contain the IF formula)
            shape.SetLinkedCell("A1", true, true);

            // Define the IF formula in cell A1
            // Example: If B1 > 10 then display "High" else display "Low"
            sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

            // Optionally set a default value in B1 for demonstration
            sheet.Cells["B1"].PutValue(15); // Change this value to test the condition

            // Update the shape so it reflects the current value of the linked cell
            shape.UpdateSelectedValue();

            // Define output file path
            string outputPath = "ConditionalShapeOutput.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
