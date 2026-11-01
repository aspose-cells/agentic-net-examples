// Title: Create a rectangle shape linked to a custom‑formatted cell and validate the displayed text with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, formats cell A1 with a custom date pattern, links the shape to that cell via the LinkedCell property, runs CalculateFormula, and asserts that shape.Text matches the cell's formatted string. | Write an Aspose.Cells example that applies a custom number format to a cell, attaches a shape through LinkedCell, recalculates the workbook, and prints whether the shape shows the same formatted value as the cell.
// Common Searches: Aspose.Cells C# link rectangle shape to a cell with custom date format | how to verify shape text matches formatted cell value in Aspose.Cells | using LinkedCell property to display formatted cell content in a shape Aspose.Cells .NET | C# example for linking shape to cell and checking formatted output with Aspose.Cells
// Tags: Aspose.Cells LinkedCell with custom styled cell | rectangle shape bound to cell value | apply custom date format to worksheet cell C# | compare shape text to cell formatted string | recalculate workbook before reading shape text

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, applying a custom date format to cell A1, adding a rectangle shape linked to that cell via LinkedCell, recalculating formulas, and confirming that the shape's displayed text equals the cell's formatted value before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a custom text format for cell A1 (e.g., date format)
            Cell linkedCell = worksheet.Cells["A1"];
            linkedCell.PutValue(DateTime.Now);
            Style cellStyle = linkedCell.GetStyle();
            cellStyle.Custom = "dd-mmm-yyyy";
            linkedCell.SetStyle(cellStyle);

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,   // upper left row
                5,   // upper left column
                5,   // top offset (pixels)
                5,   // left offset (pixels)
                30,  // height (pixels)
                100  // width (pixels)
            );

            // Link the shape to cell A1 so it displays the cell's formatted value
            shape.LinkedCell = "A1";

            // Recalculate formulas to ensure the shape text is updated
            workbook.CalculateFormula();

            // Retrieve the formatted text from the cell and the shape
            string formattedCellText = linkedCell.StringValue; // formatted according to custom style
            string shapeText = shape.Text;

            // Verify that the shape displays the same formatted text as the cell
            if (shapeText == formattedCellText)
            {
                Console.WriteLine("Formatted text appears correctly: " + shapeText);
            }
            else
            {
                Console.WriteLine("Mismatch detected.");
                Console.WriteLine("Cell formatted text: " + formattedCellText);
                Console.WriteLine("Shape text: " + shapeText);
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "ShapeLinkedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
