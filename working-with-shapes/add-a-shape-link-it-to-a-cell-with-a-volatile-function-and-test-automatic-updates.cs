// Title: Create a rectangle shape linked to a volatile NOW() formula and verify automatic updates using Aspose.Cells for .NET
// Common Searches: how to bind a shape's text to a cell with a volatile NOW() formula in Aspose.Cells C# | update shape text after recalculating volatile functions with Aspose.Cells .NET | link rectangle shape to worksheet cell and refresh automatically in Aspose.Cells | use Aspose.Cells to recalculate NOW() and reflect changes in a linked shape | save Excel file with a shape linked to a volatile function using Aspose.Cells for .NET
// Tags: add rectangle shape linked to cell formula | recalculate volatile functions and update shape text | bind shape text to worksheet cell using formula | save workbook with linked shape Aspose.Cells | Aspose.Cells shape-to-cell binding example

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts the volatile NOW() formula into cell A1, adds a rectangle shape positioned at row 2 column A, links the shape's displayed text to cell A1 via a formula, performs an initial calculation, pauses for three seconds, recalculates to let the NOW() value change, and prints the shape text before and after the second calculation before saving the file as ShapeLinkedToVolatileFunction.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a volatile function (NOW) in cell A1
            sheet.Cells["A1"].Formula = "=NOW()";

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                0,   // upper left column
                0,   // top offset
                100, // left offset
                50,  // height
                100  // width
            );

            // Link the shape's displayed text to cell A1 using a formula
            shape.Text = "=A1";

            // Perform the first calculation
            workbook.CalculateFormula();

            // Show the shape text after the first calculation
            Console.WriteLine("Shape text after first calc: " + shape.Text);

            // Wait a few seconds so the volatile NOW() function can change
            Thread.Sleep(3000);

            // Recalculate to update the volatile function
            workbook.CalculateFormula();

            // Show the shape text after the second calculation
            Console.WriteLine("Shape text after second calc: " + shape.Text);

            // Define output file path
            string outputPath = "ShapeLinkedToVolatileFunction.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
