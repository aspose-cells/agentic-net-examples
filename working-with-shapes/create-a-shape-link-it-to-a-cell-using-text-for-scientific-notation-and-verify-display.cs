// Title: Add a textbox shape linked to a cell and format its value in scientific notation using the TEXT function with Aspose.Cells for .NET
// AI Prompts: Insert a TextBox shape at a given row and column, set its Text property to =TEXT(A1,"0.00E+00"), call Workbook.CalculateFormula(), and read back the displayed text from the shape. | Create a shape, bind it to cell A1 with a TEXT formula for scientific notation, recalculate formulas, and programmatically verify the shape shows the formatted value.
// Common Searches: how to link a textbox shape to a cell and show scientific notation with Aspose.Cells .NET | Aspose.Cells set shape text to =TEXT(A1,"0.00E+00") and recalculate | display numeric cell value in scientific notation inside a shape using Aspose.Cells for C# | Aspose.Cells shape formula binding example for scientific format
// Tags: textbox shape formula binding Aspose.Cells | scientific notation TEXT function shape | recalculate workbook for shape text Aspose.Cells | save workbook with linked shape .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a large number to cell A1, adds a textbox shape at row 2 column 1, assigns the formula =TEXT(A1,"0.00E+00") to the shape, recalculates formulas so the shape displays the number in scientific notation, prints the evaluated text, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Place a numeric value in cell A1
            worksheet.Cells["A1"].PutValue(123456789.0);

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            int upperRow = 2;      // Row index (0‑based)
            int upperColumn = 1;   // Column index (0‑based)
            int top = 5;           // Pixels from the top of the cell
            int left = 5;          // Pixels from the left of the cell
            int height = 50;       // Height in points
            int width = 200;       // Width in points

            // AddTextBox returns a Shape; we can work directly with the Shape object
            Shape shape = worksheet.Shapes.AddTextBox(upperRow, upperColumn, top, left, height, width);

            // Set the shape to contain a formula that formats the value of A1 in scientific notation
            // Note: In recent Aspose.Cells versions the IsFormula property is not required;
            // the shape automatically treats the text as a formula when it starts with '='.
            shape.Text = "=TEXT(A1,\"0.00E+00\")";

            // Recalculate formulas so the shape displays the formatted value
            workbook.CalculateFormula();

            // Output the evaluated text of the shape
            Console.WriteLine("Textbox displayed text: " + shape.Text); // Expected format like 1.23E+08

            // Save the workbook (optional, demonstrates lifecycle handling)
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
