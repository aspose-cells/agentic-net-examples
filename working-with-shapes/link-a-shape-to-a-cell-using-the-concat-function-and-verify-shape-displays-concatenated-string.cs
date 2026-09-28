// Title: How to link a rectangle shape to a cell containing a CONCAT formula and display the result using Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet and set its Text property to reference cell C1 that holds a CONCAT formula. | Execute Workbook.CalculateFormula() and then read Shape.Text to confirm the concatenated string appears inside the shape. | Save the workbook as an .xlsx file after linking the shape to ensure the linked text is persisted.
// Common Searches: Aspose.Cells .NET link shape text to cell formula example | display CONCAT result inside a shape using Aspose.Cells | bind rectangle shape to cell C1 and get updated text after calculation | how to use shape.Text = "=C1" with CONCAT in Aspose.Cells | verify shape displays concatenated values from A1 and B1 in Aspose.Cells
// Tags: shape-to-cell formula binding Aspose.Cells | rectangle shape displays CONCAT output | calculate workbook formulas before shape read | Aspose.Cells linked shape persistence | C# linking shape text to cell reference

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a workbook, writes values to A1 and B1, sets a CONCAT formula in C1, adds a rectangle shape, links the shape's Text property to C1, calculates formulas, prints the shape's displayed text (expected "Hello World!"), and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate cells A1 and B1 with sample text
            ws.Cells["A1"].PutValue("Hello ");
            ws.Cells["B1"].PutValue("World!");

            // Set formula in C1 that concatenates A1 and B1 using CONCAT
            ws.Cells["C1"].Formula = "CONCAT(A1,B1)";

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            Shape shape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 120, 30);

            // Link the shape's text to cell C1
            shape.Text = "=C1";

            // Calculate all formulas in the workbook
            wb.CalculateFormula();

            // Verify that the shape now displays the concatenated string
            Console.WriteLine("Shape text: " + shape.Text); // Expected output: "Hello World!"

            // Save the workbook (optional)
            string outputPath = "LinkedShape.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
