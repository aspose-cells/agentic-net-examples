// Title: Add a rectangle shape linked to a cell with the TEXT function using scientific notation in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link shape text to a cell using the TEXT function with scientific notation | Create a rectangle shape that displays a cell value in scientific format with Aspose.Cells | Calculate shape formulas after adding a shape in Aspose.Cells .NET | Set shape text to =TEXT(A1,"0.00E+00") using Aspose.Cells API | Save a workbook containing a shape linked to a cell in C# Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | link shape text to cell using TEXT function | apply scientific notation to shape text Aspose.Cells | calculate formulas for shapes Aspose.Cells | save workbook with linked shape .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, writes a numeric value to cell A1, adds a rectangle shape positioned at row 1 column 0, assigns the shape's text to the formula =TEXT(A1,"0.00E+00") to display the value in scientific notation, forces formula calculation, outputs the stored formula, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a numeric value in cell A1
                worksheet.Cells["A1"].PutValue(1234567.89);

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,    // upper left row
                    0,    // upper left column
                    0,    // top offset (pixels)
                    100,  // left offset (pixels)
                    30,   // height (pixels)
                    100   // width (pixels)
                );

                // Link the shape's text to cell A1 using the TEXT function with scientific notation
                shape.Text = "=TEXT(A1,\"0.00E+00\")";

                // Calculate all formulas so the shape displays the evaluated result
                workbook.CalculateFormula();

                // Output the formula stored in the shape (the displayed value is calculated by Excel)
                Console.WriteLine("Shape text (formula): " + shape.Text);

                // Save the workbook (optional verification file)
                string outputPath = "output.xlsx";

                // Ensure the directory exists before saving (handle cases where outputPath has no directory)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook safely
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Error saving workbook: " + saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
