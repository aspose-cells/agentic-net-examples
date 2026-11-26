// Title: How to link a textbox shape to a cell containing a REPT formula using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link textbox shape text to cell formula result | display REPT function output in a shape with Aspose.Cells .NET | set shape.Text = "=A1" to show repeated characters in Excel workbook | remove border from textbox shape after linking to a cell in Aspose.Cells | save workbook with shape linked to REPT formula using Aspose.Cells C#
// Tags: textbox shape bound to cell value Aspose.Cells | show REPT formula result in shape text C# | remove border from Aspose.Cells shape | add textbox shape on worksheet Aspose.Cells .NET | shape text formula reference Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // The example creates a new workbook, writes a REPT formula in cell A1 to repeat a character, adds a textbox shape, binds the shape's text to the cell by setting Text = "=A1", optionally removes the shape border, and saves the file as ShapeLinkedToReptCell.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set a formula in cell A1 that repeats a character pattern using REPT
                // Example: repeat the character "*" 10 times
                Cell targetCell = sheet.Cells["A1"];
                targetCell.Formula = @"=REPT(""*"",10)";

                // Add a textbox shape to the worksheet
                // Parameters: upper left row, upper left column, upper left row offset (pixels),
                // upper left column offset (pixels), height (pixels), width (pixels)
                // Here we place it at row 2, column 0 with a size of 200x50 pixels
                Shape shape = sheet.Shapes.AddTextBox(2, 0, 0, 0, 200, 50);

                // Link the shape's displayed text to the cell containing the REPT formula
                // By setting the shape's text to a formula reference, the shape will show the evaluated result
                shape.Text = "=A1";

                // Optional: format the shape (e.g., remove border)
                shape.Line.Weight = 0.0; // No border

                // Save the workbook to a file
                string outputPath = "ShapeLinkedToReptCell.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
