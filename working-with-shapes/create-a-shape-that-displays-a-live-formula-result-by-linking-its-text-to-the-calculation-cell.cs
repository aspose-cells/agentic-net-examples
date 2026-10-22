// Title: Add a rectangle shape that shows a live formula result by linking its text to a worksheet cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to insert a rectangle shape and assign its Text property the evaluated value of a specific cell after formula calculation. | Show how to calculate worksheet formulas first, then bind the resulting value to a shape's text before saving the workbook.
// Common Searches: Aspose.Cells C# bind shape text to cell value after formula calculation | display calculated cell result inside a rectangle shape in Excel using Aspose.Cells | set dynamic text for a shape based on worksheet formula result with Aspose.Cells for .NET
// Tags: Aspose.Cells set shape text from cell value | C# add rectangle shape with dynamic text | Aspose.Cells calculate formula before binding to shape | Excel shape linked to formula result using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, writes a SUM formula in cell B2, calculates the formula, adds a rectangle shape, sets the shape's text to the calculated value of B2, and saves the file as LiveFormulaShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Place a formula in cell B2 (example: sum of numbers)
            sheet.Cells["B2"].Formula = "=SUM(1,2,3)"; // Result will be 6

            // Calculate formulas so the value is available for the shape
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left row offset,
            // upper left column, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, 2, 0, 2, 0, 120, 30);

            // Set the shape's text to the calculated value of B2
            shape.Text = sheet.Cells["B2"].StringValue;

            // Save the workbook to a file
            string outputPath = "LiveFormulaShape.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
