// Title: Link a textbox shape to a cell with a SUBSTITUTE formula and verify the displayed text using Aspose.Cells for .NET
// AI Prompts: Create C# code that adds a textbox shape to a worksheet, sets its Text property to a cell containing a SUBSTITUTE formula, runs workbook.CalculateFormula(), and asserts that the shape shows the substituted string. | Show how to bind a shape’s text to a cell that uses the SUBSTITUTE function, recalculate the workbook with Aspose.Cells, and programmatically confirm the shape displays the expected result.
// Common Searches: Aspose.Cells how to bind a textbox shape to a cell formula in C# | C# example linking shape text to SUBSTITUTE function result | verify that a shape updates after CalculateFormula in Aspose.Cells | display result of SUBSTITUTE formula in a textbox shape using Aspose.Cells for .NET
// Tags: textbox shape linked to cell formula Aspose.Cells | SUBSTITUTE formula binding to shape text .NET | validate shape text after CalculateFormula Aspose.Cells | link shape to calculated cell value C# | Aspose.Cells shape text update verification

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a SUBSTITUTE formula in B1, adds a textbox shape whose Text property is linked to B1, recalculates formulas, checks that the shape displays "Hello Aspose", and saves the file.
class ShapeLinkWithSubstitute
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put original text into cell A1
            sheet.Cells["A1"].PutValue("Hello World");

            // Set formula in B1 that substitutes "World" with "Aspose"
            sheet.Cells["B1"].Formula = "SUBSTITUTE(A1, \"World\", \"Aspose\")";

            // Add a textbox shape to the worksheet and obtain the TextBox object directly
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 5, 5, 100, 300);

            // Link the shape's text to cell B1 using a formula
            textBox.Text = "=B1";

            // Recalculate the workbook so the formula is evaluated
            workbook.CalculateFormula();

            // Retrieve the evaluated value from cell B1
            string cellResult = sheet.Cells["B1"].StringValue; // Expected: "Hello Aspose"

            // Verify the result
            string expected = "Hello Aspose";

            if (cellResult == expected)
            {
                Console.WriteLine("Success: Shape displays the updated text -> " + cellResult);
            }
            else
            {
                Console.WriteLine($"Failure: Expected '{expected}' but got '{cellResult}'");
            }

            // Save the workbook (optional, just to demonstrate saving)
            workbook.Save("ShapeLinkedWithSubstitute.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
