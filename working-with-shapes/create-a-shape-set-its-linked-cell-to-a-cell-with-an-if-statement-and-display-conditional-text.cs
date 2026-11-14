// Title: Create a rectangle shape linked to an IF formula cell to display conditional text with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bind rectangle shape to cell with IF formula | show Pass or Fail text in a shape using linked cell in Aspose.Cells | link shape to cell that contains conditional formula in .NET | example of shape line styling after linking to formula cell Aspose.Cells | create shape that updates text based on another cell value using Aspose.Cells
// Tags: rectangle shape linked cell Aspose.Cells | IF formula driven shape text Aspose.Cells | link shape to worksheet cell C# | shape line styling Aspose.Cells | conditional shape content based on cell value

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;
using System.IO;

// The example creates a new workbook, writes an IF formula in cell A1 that returns "Pass" or "Fail" based on B1, adds a rectangle shape, links the shape to the formula cell, applies basic line styling, ensures the output folder exists, and saves the file as ShapeLinkedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the linked cell (A1) with an IF formula.
            // The formula displays "Pass" if B1 >= 50, otherwise "Fail".
            Cell linkedCell = sheet.Cells["A1"];
            linkedCell.Formula = "=IF(B1>=50,\"Pass\",\"Fail\")";

            // Set a sample value in B1 to demonstrate the conditional text.
            sheet.Cells["B1"].PutValue(75); // Change this value to test different outcomes

            // Add a rectangle shape to the worksheet.
            // Parameters: shape type, upper-left row, upper-left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper-left row
                0,   // upper-left column
                0,   // top offset
                30,  // left offset
                30,  // height
                100  // width
            );

            // Link the shape to the cell containing the IF formula.
            shape.LinkedCell = "A1";

            // Optional: format the shape (border, fill, etc.).
            shape.Line.Weight = 1;
            shape.Line.DashStyle = MsoLineDashStyle.Solid;
            // The following properties may not be available in older Aspose.Cells versions.
            // Uncomment if your version supports them.
            // shape.Line.Color = Color.Black;
            // shape.Fill.ForeColor = Color.LightYellow;

            // Save the workbook to a file.
            string outputPath = "ShapeLinkedCell.xlsx";

            // Ensure the directory exists before saving.
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
