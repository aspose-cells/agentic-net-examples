// Title: Link a rectangle shape to a cell that uses the CHAR function to display a Unicode character with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape to cell with CHAR(9731) formula | display a Unicode snowflake in a rectangle shape linked to a cell using Aspose.Cells | verify linked shape text after calculating formulas in Aspose.Cells .NET | set shape.LinkedCell to a cell that contains a CHAR function in C# | example of linking an Excel shape to a cell with special characters using Aspose.Cells
// Tags: shape linked cell Aspose.Cells .NET | rectangle shape linked to cell value | CHAR function Unicode rendering Aspose.Cells | formula calculation before shape linking | verify shape text output C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, assigns the formula =CHAR(9731) (snowflake) to cell A1, calculates formulas, adds a rectangle shape, links the shape to A1 so it displays the evaluated Unicode character, prints the shape's linked text, and saves the workbook as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set cell A1 formula to use CHAR function (Unicode character 9731 – snowflake)
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=CHAR(9731)";

            // Calculate formulas so the cell value is evaluated
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Link the shape to cell A1; the shape will display the cell's value
            shape.LinkedCell = "A1";

            // Output the shape's linked text to the console
            Console.WriteLine("Shape linked text: " + shape.Text);

            // Save the workbook to a file
            string outputPath = "LinkedShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
