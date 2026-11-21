// Title: Add a rectangle shape linked to a nested formula result and verify its text with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add rectangle shape and set its text to a formula result | how to bind a shape's text to a calculated cell value using Aspose.Cells | verify that a shape displays the same value as a cell after formula calculation in C# | link shape to cell C1 with nested SUM and multiplication formula in Aspose.Cells
// Tags: add rectangle shape to worksheet Aspose.Cells | set shape text from calculated cell value C# | nested formula evaluation with Aspose.Cells | shape text verification against cell value | save workbook with linked shape Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a workbook, fills cells A1, A2, and B1, assigns a nested formula (=SUM(A1:A2)*B1) to C1, calculates the formula, adds a rectangle shape, sets the shape's text to the computed value of C1, verifies the shape text matches the cell value, and saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells used in the nested formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(5);

            // Set a nested formula in C1 (e.g., =SUM(A1:A2)*B1)
            sheet.Cells["C1"].Formula = "=SUM(A1:A2)*B1";

            // Calculate all formulas so C1 contains the final result
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, topOffsetPixels, leftOffsetPixels, heightPixels, widthPixels
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,   // upper left row
                5,   // upper left column
                0,   // top offset in pixels
                0,   // left offset in pixels
                100, // height in pixels
                50   // width in pixels
            );

            // Link the shape's text to the calculated value of cell C1
            shape.Text = sheet.Cells["C1"].StringValue;

            // Verify that the shape reflects the final result
            Console.WriteLine("Cell C1 value: " + sheet.Cells["C1"].StringValue);
            Console.WriteLine("Shape text: " + shape.Text);
            if (shape.Text == sheet.Cells["C1"].StringValue)
            {
                Console.WriteLine("Verification passed: Shape reflects the final result.");
            }
            else
            {
                Console.WriteLine("Verification failed.");
            }

            // Save the workbook
            string outputPath = "ShapeLinkedToFormula.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
