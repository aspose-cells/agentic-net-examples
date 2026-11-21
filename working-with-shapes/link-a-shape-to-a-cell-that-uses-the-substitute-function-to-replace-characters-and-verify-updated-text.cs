// Title: How to link a rectangle shape to a cell with a SUBSTITUTE formula and verify automatic text update using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape to cell formula and retrieve updated text | set rectangle shape AlternativeText to a cell reference for dynamic text | verify shape text matches cell value after CalculateFormula in .NET | dynamic shape text synchronization with cell formulas using Aspose.Cells drawing API | C# example linking shape to cell A1 to reflect SUBSTITUTE function result
// Tags: link shape to cell reference Aspose.Cells | rectangle shape AlternativeText property | SUBSTITUTE formula evaluation Aspose.Cells | dynamic shape text synchronization C# | workbook.CalculateFormula automatic update

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a SUBSTITUTE formula into cell A1, adds a rectangle shape, links the shape to the cell by setting its AlternativeText to "A1", clears the shape's text, recalculates formulas, and then confirms that the shape's displayed text matches the evaluated cell value before saving the file.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a SUBSTITUTE formula into cell A1
            // This will replace "World" with "Aspose" in the string "Hello World"
            cells["A1"].Formula = @"=SUBSTITUTE(""Hello World"",""World"",""Aspose"")";

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                5,   // upper left column
                5,   // top offset (pixels)
                5,   // left offset (pixels)
                50,  // height (pixels)
                200  // width (pixels)
            );

            // Link the shape to cell A1 by setting AlternativeText
            shape.AlternativeText = "A1";

            // Ensure the shape displays the linked cell's value as its text
            shape.Text = string.Empty; // Clear any existing text
            shape.IsLocked = false;    // Allow the shape to update automatically

            // Recalculate formulas so that the SUBSTITUTE function is evaluated
            workbook.CalculateFormula();

            // Retrieve the updated text from the linked cell
            string cellValue = cells["A1"].StringValue; // Expected: "Hello Aspose"

            // After linking, the shape's Text property reflects the cell's value
            string shapeText = shape.Text;

            // Output verification result
            Console.WriteLine("Cell A1 value: " + cellValue);
            Console.WriteLine("Shape linked text: " + shapeText);
            Console.WriteLine("Verification: " + (cellValue == shapeText ? "PASS" : "FAIL"));

            // Save the workbook (optional)
            string outputPath = "ShapeLinkExample.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
