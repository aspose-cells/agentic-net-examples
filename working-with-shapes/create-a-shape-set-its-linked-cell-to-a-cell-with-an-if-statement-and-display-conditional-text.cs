// Title: Add a rectangle shape linked to an IF formula cell and display conditional text in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a rectangle shape, sets its LinkedCell to a cell containing an IF formula, and displays the formula result as the shape's text. | Show how to center the text inside a shape that is linked to a formula cell by configuring TextHorizontalAlignment and TextVerticalAlignment in Aspose.Cells. | Demonstrate adjusting a shape's position and dimensions while preserving its link to a conditional IF formula result in an Excel file using Aspose.Cells.
// Common Searches: aspnet c# aspose.cells link rectangle shape to cell with if formula | how to display result of an IF formula inside a shape using Aspose.Cells | set linkedcell property for shape to formula cell in Aspose.Cells .NET | center text in a shape linked to a cell in Aspose.Cells workbook | change shape size and keep linked cell after adding IF formula in Aspose.Cells
// Tags: add rectangle shape linkedcell aspose.cells | if formula result in shape text aspose.cells | centered shape text aspose.cells c# | adjust shape size linkedcell aspose.cells | excel workbook conditional shape text aspose.cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, writes a numeric value to A1, defines an IF formula in B2, adds a rectangle shape, links the shape to B2 so it shows the formula's result, centers the text within the shape, optionally adjusts the shape's size and position, and saves the file as ShapeLinkedCell.xlsx.
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

            // Put a numeric value in A1 that will be used in the IF condition
            sheet.Cells["A1"].PutValue(15);

            // Set an IF formula in B2 that returns text based on the value in A1
            sheet.Cells["B2"].Formula = "=IF(A1>10,\"High\",\"Low\")";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column,
            //            upper left row offset (pixels), upper left column offset (pixels),
            //            height (pixels), width (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5,          // upper left row
                5,          // upper left column
                0,          // row offset
                0,          // column offset
                100,        // height
                200);       // width

            // Link the shape to cell B2 so the shape displays the result of the IF formula
            shape.LinkedCell = "B2";

            // Optional: center the text inside the shape
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;

            // Define output file path
            string outputPath = "ShapeLinkedCell.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
