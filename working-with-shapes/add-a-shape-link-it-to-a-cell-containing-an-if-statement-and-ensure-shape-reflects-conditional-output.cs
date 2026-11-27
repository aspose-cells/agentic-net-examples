// Title: Create a rectangle shape linked to an IF formula cell in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link shape text to a cell formula | How to display IF function result in a rectangle shape with Aspose.Cells | Set shape text to reference a worksheet cell in Aspose.Cells .NET | Add rectangle shape and bind it to A1 formula using Aspose.Cells C# | Conditional shape content based on cell value in Aspose.Cells workbook
// Tags: rectangle shape bound to cell formula Aspose.Cells | shape text referencing IF result C# | Aspose.Cells conditional shape content | C# add shape linked to worksheet cell | Excel shape displaying formula output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes an IF formula in A1 that returns "High" or "Low" based on B1, adds a rectangle shape at row 2 column 0, sets the shape's text to "=A1" so it reflects the formula result, optionally formats the shape, and saves the file as ShapeLinkedToIf.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a value in B1 that the IF formula will evaluate
            sheet.Cells["B1"].PutValue(15); // Change this value to test different outcomes

            // Place an IF formula in A1 that depends on B1
            sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

            // Add a rectangle shape that will display the result of the IF formula
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                0,    // upper left column
                0,    // top offset (in pixels)
                0,    // left offset (in pixels)
                200,  // width (in pixels)
                50    // height (in pixels)
            );

            // Link the shape's text to cell A1 by using a formula string
            shape.Text = "=A1";

            // Optional: format the shape's appearance
            shape.Line.Weight = 1.0;
            shape.Line.DashStyle = MsoLineDashStyle.Solid;
            // Note: Color properties may not be available in some Aspose.Cells versions; they are omitted for compatibility.

            // Optional: set fill color if supported
            // shape.Fill.ForeColor = System.Drawing.Color.LightYellow; // Uncomment if the API provides this property

            // Determine output path and ensure the directory exists
            string outputPath = "ShapeLinkedToIf.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
