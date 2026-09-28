// Title: Link a rectangle shape to a CONCATENATE formula cell and display the combined text using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, write a CONCATENATE formula in a cell, calculate the workbook, add a rectangle shape, and assign the shape's Text property to the evaluated cell value with Aspose.Cells in C#. | Generate a .xlsx file where a rectangle shape automatically reflects the result of a CONCATENATE formula after the workbook is calculated, using the Aspose.Cells .NET API. | Programmatically bind a shape's displayed text to the value of a formula cell, ensuring the formula is evaluated first, via Aspose.Cells C# methods.
// Common Searches: Aspose.Cells C# link shape text to cell formula result | set rectangle shape text from CONCATENATE formula in Excel using Aspose.Cells | display calculated CONCATENATE value inside a shape with Aspose.Cells .NET
// Tags: Aspose.Cells bind shape to calculated cell | C# rectangle shape displays formula result | evaluate workbook before shape assignment Aspose.Cells | Excel shape text from CONCATENATE cell using .NET | link shape to cell value Aspose.Cells API

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a CONCATENATE formula in B2, calculates the formula, adds a rectangle shape, sets the shape's Text property to the evaluated string from B2, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a CONCATENATE formula in cell B2
            sheet.Cells["B2"].Formula = "=CONCATENATE(\"Hello \",\"World\")";

            // Evaluate the formula so the cell contains the combined text
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                0,    // upper left column
                0,    // top offset (pixels)
                100,  // left offset (pixels)
                30,   // height (pixels)
                200   // width (pixels)
            );

            // Link the shape's displayed text to the value of cell B2
            shape.Text = sheet.Cells["B2"].StringValue;

            // Save the workbook
            workbook.Save("LinkedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
