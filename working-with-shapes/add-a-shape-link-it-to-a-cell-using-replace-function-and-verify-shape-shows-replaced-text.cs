// Title: Create a rectangle shape, link its text to a REPLACE formula on cell A1, and read the calculated text with Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to the first worksheet, assign the Text property =REPLACE(A1,1,5,"Hi"), call workbook.CalculateFormula(), and output the shape's resulting text. | Swap the REPLACE formula with =CONCAT(A1,"_2023"), recalculate the workbook, and display the new shape text. | Persist the workbook after linking the shape to the formula, then reopen the file to verify the shape displays the expected text.
// Common Searches: Aspose.Cells bind shape text to REPLACE formula C# example | How to set a shape's Text property to an Excel formula using Aspose.Cells | Get shape's calculated text after workbook.CalculateFormula in .NET | Link rectangle shape to cell A1 formula with Aspose.Cells for .NET
// Tags: add rectangle shape with formula text Aspose.Cells | shape formula REPLACE C# | calculate shape formulas Aspose.Cells | retrieve shape displayed value .NET | link shape to cell A1 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, writes "HelloWorld" to cell A1, adds a rectangle shape, assigns a REPLACE formula (=REPLACE(A1,1,5,"Hi")) to the shape's Text property, forces formula calculation, and prints the resulting shape text ("HiWorld").
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put original text into cell A1
            Cell sourceCell = sheet.Cells["A1"];
            sourceCell.PutValue("HelloWorld");

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, height, width (points)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Link the shape's text to a REPLACE formula that operates on A1
            // This will replace the first 5 characters of A1 with "Hi"
            shape.Text = "=REPLACE(A1,1,5,\"Hi\")";

            // Force calculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the displayed text of the shape after calculation
            string displayedText = shape.Text;

            // Output the result to verify the replacement
            Console.WriteLine("Shape text after REPLACE formula: " + displayedText);
            // Expected output: "HiWorld"

            // (Optional) Save the workbook to inspect manually
            // workbook.Save("ShapeReplaceDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
