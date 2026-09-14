// Title: Add a TextBox shape linked to a cell with a custom TEXT number format in Aspose.Cells for .NET
// AI Prompts: Create a TextBox shape on a worksheet and assign its Text property a TEXT formula that formats cell A1 as currency. | Invoke workbook.CalculateFormula() and capture the displayed text of the linked TextBox shape. | Save the workbook to confirm that the TextBox shows the formatted value from cell A1.
// Common Searches: aspocells .net link textbox shape to cell using TEXT function | how to apply custom number format to shape text in Aspose.Cells | retrieve linked shape text after CalculateFormula in C# | bind a shape to a cell value with TEXT formula Aspose.Cells example
// Tags: textbox shape linked via TEXT function Aspose.Cells | currency formatting for shape text .NET | recalculate workbook formulas for linked shapes | shape text verification after calculation C# | link shape to cell using custom number format Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Demonstrates creating a workbook, writing a numeric value to A1, adding a TextBox shape, linking its text to A1 with a TEXT formula that applies a custom currency format, recalculating formulas, outputting the shape's displayed text, and saving the file as LinkedShape.xlsx.
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
            sheet.Cells["A1"].PutValue(1234.567);

            // Add a TextBox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape textBox = sheet.Shapes.AddTextBox(2, 0, 0, 2, 200, 50);

            // Link the shape's text to cell A1 using the TEXT function with a custom number format
            // The formula will display the number as currency with two decimal places
            textBox.Text = "=TEXT(A1,\"$#,##0.00\")";

            // Recalculate all formulas so the shape's text is updated
            workbook.CalculateFormula();

            // Verify the shape's displayed text
            Console.WriteLine("Shape text after calculation: " + textBox.Text);

            // Save the workbook (optional verification step)
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
