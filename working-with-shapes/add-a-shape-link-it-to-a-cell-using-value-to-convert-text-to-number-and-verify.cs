// Title: Create a rectangle shape linked to a cell with the VALUE function and verify the numeric result using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape text to =VALUE(A1) and retrieve evaluated value | how to bind a rectangle shape to a worksheet cell value in Aspose.Cells | verify that a shape formula returns the same number as the source cell in Aspose.Cells | convert text in a cell to a number inside a shape using the VALUE function with Aspose.Cells | save an Excel file after linking a shape to a cell formula using Aspose.Cells .NET
// Tags: add rectangle shape linked to cell Aspose.Cells | shape text formula VALUE function | calculate workbook formulas Aspose.Cells | verify shape numeric value C# | bind shape to cell value Aspose.Cells | save workbook with linked shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes the string "456" to cell A1, adds a rectangle shape, sets its text to the formula "=VALUE(A1)", forces formula calculation, parses the evaluated shape text, compares it to the numeric value of the cell, outputs verification results, and saves the file as ShapeLinkValue.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Put a text that looks like a number into cell A1.
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("456");

            // Add a rectangle shape to the worksheet.
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height.
            // Use MsoDrawingType for compatibility with older Aspose.Cells versions.
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 30);

            // Link the shape's displayed text to the numeric conversion of A1 using the VALUE function.
            shape.Text = "=VALUE(A1)";

            // Force calculation of all formulas in the workbook.
            workbook.CalculateFormula();

            // Retrieve the evaluated text from the shape.
            string shapeText = shape.Text;

            // Retrieve the numeric value from the cell (as double).
            double cellNumber = Convert.ToDouble(cell.Value);

            // Verify that the shape's text matches the numeric value from the cell.
            if (double.TryParse(shapeText, out double shapeNumber))
            {
                Console.WriteLine($"Shape value: {shapeNumber}");
                Console.WriteLine($"Cell numeric value: {cellNumber}");
                Console.WriteLine(shapeNumber == cellNumber ? "Verification succeeded." : "Verification failed.");
            }
            else
            {
                Console.WriteLine("Shape text could not be parsed to a number.");
            }

            // Define output file path.
            string outputPath = "ShapeLinkValue.xlsx";

            // Ensure the directory exists before saving.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
