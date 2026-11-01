// Title: Add a rectangle shape linked to a cell that displays a TEXT‑formatted number using Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape on the first worksheet, set its LinkedCell to "B1", and confirm the shape shows the TEXT‑formatted value from cell A1. | Apply a TEXT formula with a custom number format to cell B1, run workbook.CalculateFormula(), and read the shape's Text property to verify the output. | Save the workbook as an .xlsx file after linking the shape and validating that the displayed text matches the expected formatted string.
// Common Searches: how to link a shape to a cell using TEXT formula in Aspose.Cells .NET | display custom formatted numeric value in a rectangle shape with Aspose.Cells | Aspose.Cells shape LinkedCell property example with TEXT function | calculate formulas before reading linked shape text in Aspose.Cells | save workbook with shape linked to cell using Aspose.Cells for .NET
// Tags: add rectangle shape linkedcell Aspose.Cells | TEXT formula custom number format Aspose.Cells | calculate workbook formulas before reading shape text | save workbook with linked shape Aspose.Cells | shape text verification Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to A1, sets B1 to a TEXT formula with a custom format, adds a rectangle shape, links it to B1, calculates formulas, prints the shape's displayed text (expected "Value: 1234.57"), and saves the file as ShapeLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value in cell A1
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(1234.567);

            // Set cell B1 to a TEXT formula with custom format
            Cell cellB1 = sheet.Cells["B1"];
            cellB1.Formula = "=TEXT(A1, \"Value: 0.00\")";

            // Add a rectangle shape to the worksheet and link it to cell B1
            // AddShape returns a Shape object in recent Aspose.Cells versions
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                200  // width (pixels)
            );

            // Link the shape to cell B1 so it displays the cell's value
            shape.LinkedCell = "B1";

            // Calculate formulas so the TEXT result is evaluated
            workbook.CalculateFormula();

            // Verify the shape's displayed text
            Console.WriteLine("Shape text: " + shape.Text); // Expected output: Value: 1234.57

            // Save the workbook (lifecycle rule)
            workbook.Save("ShapeLinkedToCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
