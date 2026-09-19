// Title: Create a rectangle shape linked to a text‑formatted cell and validate its displayed text using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, applies the text number format (code 49) to cell A1, sets the shape's LinkedCell to A1, and copies the cell's string value into the shape's Text property. | Write a C# snippet that calls Workbook.CalculateFormula, asserts that shape.Text equals the linked cell's string value, and then saves the workbook to an XLSX file. | Demonstrate how to format a cell as text (Number format "@") in Aspose.Cells before linking it to a shape and retrieving the formatted value.
// Common Searches: Aspose.Cells how to link a shape to a cell with text format in C# | verify shape text equals linked cell value Aspose.Cells .NET example | apply number format 49 to a cell before linking shape Aspose.Cells | C# code to add rectangle shape and bind it to a formatted cell using Aspose.Cells
// Tags: shape.LinkedCell with text formatted cell Aspose.Cells | add rectangle shape to worksheet Aspose.Cells .NET | apply text number format 49 cell Aspose.Cells | verify shape.Text matches linked cell C# | workbook.Save after shape linking Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, formats cell A1 as text using number format code 49, adds a rectangle shape linked to A1, copies the cell's string value into the shape's Text property, asserts that the displayed text matches the cell value after recalculating formulas, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set cell A1 with text format and a string value
            Cell linkedCell = worksheet.Cells["A1"];
            // Apply text (string) format: Number format code 49 corresponds to "@"
            Style textStyle = linkedCell.GetStyle();
            textStyle.Number = 49;
            linkedCell.SetStyle(textStyle);
            linkedCell.PutValue("Hello Aspose!");

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 30);

            // Link the shape to the formatted cell (A1)
            shape.LinkedCell = "A1";

            // Ensure the shape displays the linked cell's value
            shape.Text = linkedCell.StringValue;

            // Recalculate formulas (not strictly required here but kept for completeness)
            workbook.CalculateFormula();

            // Verify that the shape's displayed text matches the linked cell's value
            string shapeText = shape.Text; // Should be "Hello Aspose!"
            Debug.Assert(shapeText == linkedCell.StringValue, "Shape text does not match linked cell value.");

            // Output verification result
            Console.WriteLine($"Shape text: {shapeText}");

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
