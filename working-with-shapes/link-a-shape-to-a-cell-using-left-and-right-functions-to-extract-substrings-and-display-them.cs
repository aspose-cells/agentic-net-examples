// Title: Bind a TextBox shape to a worksheet cell using LEFT and RIGHT functions with Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that adds a TextBox shape to a worksheet and sets its Text property to an Excel formula that concatenates LEFT(cell, n) and RIGHT(cell, m) using Aspose.Cells. | Create code that links a shape’s displayed text to cell B2, showing the first 6 characters and the last 4 characters separated by a space, then saves the workbook. | Write a snippet demonstrating how to bind a shape to a cell formula for substring extraction in an Aspose.Cells workbook.
// Common Searches: aspnet aspose.cells bind textbox shape to cell formula left function | c# add textbox to excel worksheet and link text to cell using left and right | how to display substring of a cell in a shape with aspose.cells | asp.net aspose.cells shape text formula example
// Tags: Aspose.Cells bind shape text to cell formula | C# textbox shape substring extraction | LEFT and RIGHT functions in Aspose.Cells shape | link shape to worksheet cell Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, writes "AsposeCellsDemo" to cell B2, adds a TextBox shape, sets its Text property to the formula =LEFT(B2,6) & " " & RIGHT(B2,4) to display the first six and last four characters, and saves the file as LinkedShape.xlsx.
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

            // Put sample text in cell B2
            sheet.Cells["B2"].PutValue("AsposeCellsDemo");

            // Add a textbox shape to the worksheet
            // Parameters: upperRow, upperColumn, upperRowOffset, upperColumnOffset, height, width
            TextBox textBox = sheet.Shapes.AddTextBox(5, 5, 0, 0, 200, 50);

            // Link the shape's text to cell B2 using LEFT and RIGHT functions
            // This will display the first 6 characters and the last 4 characters separated by a space
            textBox.Text = "=LEFT(B2,6) & \" \" & RIGHT(B2,4)";

            // Save the workbook
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
