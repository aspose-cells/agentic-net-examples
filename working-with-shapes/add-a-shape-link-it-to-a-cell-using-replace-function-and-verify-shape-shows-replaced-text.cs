// Title: Create a textbox shape linked to a REPLACE formula and validate its displayed text with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add textbox shape and bind it to a REPLACE formula | How to show REPLACE function result inside a shape using Aspose.Cells .NET | Verify shape text after calling Workbook.CalculateFormula in C# | Link a shape's text to a cell formula and retrieve the evaluated value with Aspose.Cells | Save workbook with a textbox shape that displays a formula result in Aspose.Cells
// Tags: textbox shape formula binding Aspose.Cells | shape text REPLACE function .NET | Workbook.CalculateFormula for shapes | validate shape displayed text C# | persist shape formula result in XLSX

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a textbox shape on the first worksheet, assigns a REPLACE formula to the shape's Text property to modify the value in cell A1, recalculates formulas so the shape shows the evaluated result, prints the shape's text to the console for verification, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a sample value in cell A1
            worksheet.Cells["A1"].PutValue("Hello");

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            // In recent Aspose.Cells versions AddTextBox returns a TextBox object
            TextBox textBox = worksheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 50);
            Shape shape = textBox; // Cast to base Shape for further operations

            // Link the shape's text to a REPLACE formula that references cell A1
            // This will replace the first character of "Hello" with "J", resulting in "Jello"
            shape.Text = "=REPLACE(A1,1,1,\"J\")";

            // Recalculate formulas so the shape displays the evaluated result
            workbook.CalculateFormula();

            // Verify that the shape now shows the replaced text
            Console.WriteLine("Shape text after REPLACE: " + shape.Text);

            // Save the workbook (optional, demonstrates that the shape is persisted)
            string outputPath = "ShapeReplaceDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
