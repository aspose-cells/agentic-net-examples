// Title: Create a rectangle shape linked to a SUM formula cell and verify its displayed value using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, sets its Text property to reference a cell containing a SUM formula, runs workbook.CalculateFormula, and asserts that the shape’s text matches the computed total. | Write a script that binds a shape’s text to cell B1 (=SUM(A1:A5)), triggers recalculation, parses the shape text, and confirms the value equals the expected sum.
// Common Searches: aspnet add rectangle shape linked to cell formula Aspose.Cells example | how to bind shape text to a SUM cell in C# using Aspose.Cells | verify that a shape displays the result of a formula in an Excel file with Aspose.Cells | C# Aspose.Cells shape text shows calculated total from range A1:A5
// Tags: link shape text to cell formula Aspose.Cells | add rectangle shape to worksheet C# | bind shape to SUM cell Aspose.Cells | verify shape displays calculated value .NET | shape text formula reference Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a workbook, fills A1:A5 with numbers, places a SUM formula in B1, adds a rectangle shape, links the shape's text to B1, recalculates the workbook, checks that the shape's displayed text equals the calculated sum, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data in A1:A5
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue(50);

            // Set SUM formula in B1 and calculate
            sheet.Cells["B1"].Formula = "=SUM(A1:A5)";
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 150, 30);

            // Link the shape's text to cell B1 (the shape will display the calculated sum)
            shape.Text = "=B1";

            // Recalculate to ensure the shape's displayed text is updated
            workbook.CalculateFormula();

            // Verify that the shape shows the total sum
            string shapeText = shape.Text;
            double expectedSum = sheet.Cells["B1"].DoubleValue;

            Console.WriteLine($"Shape text: {shapeText}");
            Console.WriteLine($"Expected sum: {expectedSum}");

            // Attempt to parse the shape text as a double for verification
            if (double.TryParse(shapeText, out double shapeValue) &&
                Math.Abs(shapeValue - expectedSum) < 0.0001)
            {
                Console.WriteLine("Verification passed: Shape displays the correct total sum.");
            }
            else
            {
                Console.WriteLine("Verification failed: Shape does not display the correct total sum.");
            }

            // Save the workbook (optional)
            string outputPath = "ShapeLinkedToSum.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
