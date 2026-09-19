// Title: Link a TextBox shape to a cell containing a SUBSTITUTE formula and verify the linked text with Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, assign the formula SUBSTITUTE("Hello World","World","Aspose") to cell A1, calculate the workbook, add a TextBox shape at row 5 column 5, set its Text property to "=A1", recalculate, then compare shape.Text with sheet.Cells["A1"].StringValue and print whether they match. | Adapt the code to link any shape (e.g., Rectangle, Oval) to a cell that holds a formula, ensure CalculateFormula is called after the link, and programmatically confirm that the shape's displayed text equals the evaluated cell value.
// Common Searches: asp.net how to bind a textbox shape to a cell formula using Aspose.Cells | example of linking a shape to a cell with SUBSTITUTE function in Aspose.Cells for .NET | verify that a linked shape updates after workbook.CalculateFormula in Aspose.Cells | Aspose.Cells .NET linking shape text to cell A1 and checking equality
// Tags: textbox shape cell linking Aspose.Cells | SUBSTITUTE function shape binding | shape text verification after CalculateFormula | Aspose.Cells workbook recalculation for linked shapes | link shape to formula result .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The sample creates a workbook, sets cell A1 with a SUBSTITUTE formula, calculates the formula, adds a TextBox shape, links the shape's text to the cell using "=A1", recalculates again, compares the shape's text with the cell's evaluated value, outputs the verification result, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1 that uses SUBSTITUTE to replace "World" with "Aspose"
            sheet.Cells["A1"].Formula = "SUBSTITUTE(\"Hello World\",\"World\",\"Aspose\")";

            // Calculate the formula so the cell contains the resulting value
            workbook.CalculateFormula();

            // Add a textbox shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
            TextBox shape = sheet.Shapes.AddTextBox(5, 5, 200, 50, 100, 200);

            // Link the shape's text to cell A1
            shape.Text = "=A1";

            // Recalculate to ensure the shape reflects the linked cell's value
            workbook.CalculateFormula();

            // Retrieve the cell's displayed value and the shape's text
            string cellValue = sheet.Cells["A1"].StringValue;
            string shapeText = shape.Text;

            // Output verification result
            Console.WriteLine($"Cell A1 value: {cellValue}");
            Console.WriteLine($"Shape text: {shapeText}");
            Console.WriteLine($"Verification: {(cellValue == shapeText ? "Passed" : "Failed")}");

            // Save the workbook (optional)
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
