// Title: Add a rectangle shape linked to a cell with a TEXT formula for currency formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to a cell using TEXT function with custom currency format | display formatted currency from a worksheet cell inside a rectangle shape in Aspose.Cells | force formula calculation so shape text shows updated value in Aspose.Cells .NET | save Excel file with a shape that shows the result of a TEXT formula using Aspose.Cells | verify shape text after workbook.CalculateFormula in C#
// Tags: add rectangle shape Aspose.Cells .NET | link shape to cell using TEXT formula | custom currency format in shape text | force formula calculation for shape text | save workbook with linked shape Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, writes a numeric value to cell B2, adds a rectangle shape, sets the shape's Text property to a TEXT formula that formats the cell value as currency, forces formula calculation, reads and prints the evaluated shape text, and saves the workbook as ShapeLink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell B2 (the source for the TEXT formula)
            Cell sourceCell = sheet.Cells["B2"];
            sourceCell.PutValue(1234.56);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column,
            //             row offset, column offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,          // upper left row
                5,          // upper left column
                0,          // row offset (pixels)
                0,          // column offset (pixels)
                30,         // height (pixels)
                150);       // width (pixels)

            // Link the shape to the cell using a TEXT formula with a custom currency format
            shape.Text = $"=TEXT({sourceCell.Name},\"$#,##0.00\")";

            // Force calculation of all formulas so the shape text is updated
            workbook.CalculateFormula();

            // Verify: read the evaluated text from the shape and output it
            string displayedText = shape.Text;
            Console.WriteLine($"Shape displays: {displayedText}");

            // Save the workbook (optional, demonstrates that the shape is persisted)
            string outputPath = "ShapeLink.xlsx";
            try
            {
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"File I/O error: {ioEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
