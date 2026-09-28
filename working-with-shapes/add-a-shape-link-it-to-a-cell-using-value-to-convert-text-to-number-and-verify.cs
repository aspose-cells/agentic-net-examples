// Title: Link a rectangle shape to a cell with the VALUE function to convert text to a number and verify the result using Aspose.Cells for .NET
// AI Prompts: Create a rectangular drawing object on the first worksheet, set its text to "=VALUE(B2)", run workbook.CalculateFormula, and output whether the shape now shows the numeric value 123. | Replace the source cell with C5 containing the string "456", adjust the shape's formula to reference C5, recalculate, and programmatically confirm the shape displays 456. | Once the shape's numeric display is confirmed, save the workbook both as an XLSX file and as a PDF using Aspose.Cells.
// Common Searches: how to bind a shape's text to a cell value using VALUE function in Aspose.Cells C# | convert text in a cell to number inside a shape formula with Aspose.Cells | verify shape displays calculated numeric result after workbook.CalculateFormula in .NET | create rectangular shape linked to cell B2 and read its evaluated text with Aspose.Cells | Aspose.Cells C# example linking shape to cell and checking numeric output
// Tags: insert rectangle shape Aspose.Cells C# | shape text formula VALUE Aspose.Cells | link shape to cell using VALUE function | recalculate workbook formulas Aspose.Cells | verify shape numeric display Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, writes the text "123" into cell B2, inserts a rectangle shape, assigns the formula "=VALUE(B2)" to the shape's text, recalculates all formulas, reads the evaluated shape text, checks that it equals the numeric value 123, and saves the workbook as ShapeLinkDemo.xlsx.
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

            // Put a text value "123" in cell B2 (as text)
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("123"); // text, not numeric

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                1,          // upper left row
                0,          // upper left column
                0,          // top offset (pixels)
                0,          // left offset (pixels)
                100,        // height (pixels)
                50);        // width (pixels)

            // Link the shape's text to the cell using VALUE to convert text to number
            shape.Text = "=VALUE(B2)";

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the evaluated text from the shape
            string shapeResult = shape.Text;

            // Verify that the shape now displays the numeric value 123
            if (double.TryParse(shapeResult, out double numericValue) && numericValue == 123)
            {
                Console.WriteLine("Verification succeeded: Shape displays numeric value 123.");
            }
            else
            {
                Console.WriteLine($"Verification failed: Shape displays '{shapeResult}'.");
            }

            // Save the workbook (optional)
            workbook.Save("ShapeLinkDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
