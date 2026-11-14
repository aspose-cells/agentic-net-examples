// Title: Link a textbox shape to worksheet cells with a CONCAT formula and verify the displayed text using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set textbox shape text to CONCAT(A1,B1) formula | verify that a shape shows concatenated cell values after workbook.CalculateFormula | link Excel shape to cells using CONCAT function with Aspose.Cells | read textbox shape text after formula evaluation in Aspose.Cells .NET | save workbook with a shape linked to cell values using Aspose.Cells
// Tags: Aspose.Cells textbox shape formula binding | C# shape text CONCAT function | shape displayed text verification after calculation | link shape to worksheet cells Aspose.Cells | save workbook with linked shape .xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes "Hello" and "World" to cells A1 and B1, adds a textbox shape, assigns its Text property to the formula "=CONCAT(A1,B1)", forces formula recalculation, reads the shape's displayed text, compares it with the expected concatenated string, outputs the verification result, and saves the file as ShapeLinkedToCell.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells A1 and B1 with sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Define shape bounds (rows and columns) and size (width, height in points)
            int upperLeftRow = 2;
            int upperLeftColumn = 1;
            int lowerRightRow = 5;
            int lowerRightColumn = 5;
            int width = 200;   // width in points
            int height = 100;  // height in points

            // Add a textbox shape to the worksheet
            Shape shape = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, width, height);

            // Link the shape's text to a CONCAT formula referencing cells A1 and B1
            shape.Text = "=CONCAT(A1,B1)";

            // Recalculate formulas so the shape displays the result
            workbook.CalculateFormula();

            // Retrieve the displayed text from the shape
            string displayedText = shape.Text;

            // Verify that the shape displays the concatenated string
            string expectedText = sheet.Cells["A1"].StringValue + sheet.Cells["B1"].StringValue;
            bool isCorrect = displayedText == expectedText;

            Console.WriteLine($"Shape displayed text: \"{displayedText}\"");
            Console.WriteLine($"Expected text: \"{expectedText}\"");
            Console.WriteLine($"Verification result: {(isCorrect ? "Success" : "Failure")}");

            // Save the workbook (optional, for visual inspection)
            string outputPath = "ShapeLinkedToCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
