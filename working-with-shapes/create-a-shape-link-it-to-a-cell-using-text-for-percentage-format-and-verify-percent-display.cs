// Title: Create a textbox shape linked to a cell via the TEXT function for percentage formatting and verify the percent sign with Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a TextBox shape into a worksheet, sets its Text property to =TEXT(A1,"0.00%"), runs workbook.CalculateFormula(), reads the shape's Text, and checks whether it contains a '%' character. | Demonstrate how to apply a percentage number format to a cell, link a shape’s text to that cell using the TEXT function, and programmatically confirm the formatted output in Aspose.Cells. | Adapt the sample to link the shape to another cell and use a custom percentage pattern such as "0.0%", then validate that the shape still displays the percent sign.
// Common Searches: aspnet cells how to attach a textbox shape to a worksheet cell with TEXT function for percent | c# aspose.cells show a cell's percentage value inside a shape's text | check that a shape linked to a cell displays a % symbol after workbook calculation | use custom percentage format in shape text linked to a different cell aspose.cells example
// Tags: textbox shape linked to cell Aspose.Cells | shape text TEXT function percentage format | calculate workbook formulas Aspose.Cells | validate percent sign in shape output C# | percentage number format cell Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, writes a numeric value to cell A1, applies a 0.00% number format, adds a TextBox shape whose Text is set to =TEXT(A1,"0.00%"), calculates all formulas so the shape displays the formatted value, reads the shape text to verify it contains a '%' character, and saves the file as ShapePercent.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(0.2567);

            // Apply percentage number format (0.00%)
            Style percentStyle = workbook.CreateStyle();
            percentStyle.Number = 10; // 0.00%
            cellA1.SetStyle(percentStyle); // apply style directly

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddTextBox(5, 0, 0, 200, 30, 200);

            // Link the shape's text to cell A1 using TEXT function for percentage format
            shape.Text = "=TEXT(A1,\"0.00%\")";

            // Calculate all formulas so the shape text is evaluated
            workbook.CalculateFormula();

            // Retrieve the displayed text from the shape
            string displayedText = shape.Text;

            // Output the result and verify that it contains a percent sign
            Console.WriteLine("Shape displayed text: " + displayedText);
            if (displayedText.Contains("%"))
                Console.WriteLine("Percent display verified.");
            else
                Console.WriteLine("Percent display NOT verified.");

            // Save the workbook
            workbook.Save("ShapePercent.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
