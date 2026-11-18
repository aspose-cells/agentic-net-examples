// Title: Create a rectangle shape linked to a worksheet cell using the TEXT function for percentage formatting and verify the result with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to cell using TEXT function with percentage format | how to display cell value as percentage inside a shape in Aspose.Cells | verify shape text after workbook.CalculateFormula in Aspose.Cells .NET | add rectangle shape and bind its text to A1 with custom number format using Aspose.Cells | C# Aspose.Cells shape text evaluation example
// Tags: rectangle shape cell binding Aspose.Cells | percentage formatting with TEXT function in shape | shape text verification after formula recalculation | Aspose.Cells workbook.CalculateFormula usage | C# shape text evaluation from worksheet cell

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // The example creates a new workbook, writes a numeric value to cell A1, adds a rectangle shape, sets the shape's text to the result of the TEXT function with a custom percentage format referencing A1, recalculates all formulas, retrieves the evaluated shape text, compares it to the expected "12.34%" and outputs the verification result, then saves the workbook as ShapeLinkExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Put a numeric value in cell A1
                sheet.Cells["A1"].PutValue(0.1234);

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top, left, height, width
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

                // Link the shape's text to cell A1 using the TEXT function with a custom percentage format
                shape.Text = "=TEXT(A1,\"0.00%\")";

                // Recalculate all formulas so the shape's text is updated
                workbook.CalculateFormula();

                // Retrieve the evaluated text from the shape
                string shapeText = shape.Text;

                // Verify that the shape's text matches the expected formatted value
                string expectedText = "12.34%";
                if (shapeText == expectedText)
                {
                    Console.WriteLine("Verification succeeded: Shape text is correctly formatted as " + shapeText);
                }
                else
                {
                    Console.WriteLine($"Verification failed: Expected '{expectedText}' but got '{shapeText}'.");
                }

                // Save the workbook to a file (lifecycle rule: save)
                workbook.Save("ShapeLinkExample.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
