// Title: Link a textbox shape to a cell containing a REPT formula using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, write the formula =REPT("AB",5) into cell A1, add a textbox shape at C3, and set the shape's Text property to "=A1" with Aspose.Cells in C#. | Adjust the added textbox shape to hide its border while keeping the linked REPT result visible in the shape. | Produce an Excel file named LinkedShapeWithRept.xlsx where the textbox automatically updates to reflect any changes to the REPT formula in the linked cell.
// Common Searches: asp.net c# link textbox shape to cell formula using Aspose.Cells | display REPT function output in an Excel shape with Aspose.Cells .NET | how to bind shape text to a cell that contains a REPT formula in C# | Aspose.Cells hide shape border while linking shape text to a cell
// Tags: Aspose.Cells textbox shape linked to cell | C# insert REPT formula programmatically | Aspose.Cells hide shape border | Excel shape text reference cell .NET | link shape to REPT function output

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, sets cell A1 to a REPT formula (e.g., =REPT("AB",5)), adds a textbox shape at C3, links the shape's text to A1 so it displays the repeated pattern, hides the shape border, and saves the file as LinkedShapeWithRept.xlsx.
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

            // Set a formula in cell A1 that repeats a character pattern using REPT
            // Example: repeat the string "AB" 5 times -> "ABABABABAB"
            Cell targetCell = sheet.Cells["A1"];
            targetCell.Formula = @"=REPT(""AB"",5)";

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, height, width
            // Here we place it at row 2, column 2 (C3) with arbitrary size
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 0;
            int left = 0;
            int height = 100;
            int width = 200;
            Shape shape = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, height, width);

            // Link the shape's text to the cell A1 so it displays the repeated pattern
            shape.Text = "=A1";

            // Optionally, format the shape (e.g., remove border)
            shape.Line.Weight = 0; // Set line weight to 0 to hide border

            // Save the workbook to a file
            workbook.Save("LinkedShapeWithRept.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
