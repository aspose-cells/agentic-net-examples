// Title: How to link a rectangle shape’s text to a cell using LEFT and RIGHT functions in Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet and assign its Text property a formula that concatenates LEFT and RIGHT functions referencing a specific cell. | Modify the cell reference and the character counts in the LEFT/RIGHT functions of the shape’s formula to show different parts of the source text. | Save the workbook, open the generated Excel file, and confirm that the shape displays the combined substrings from the linked cell.
// Common Searches: Aspose.Cells C# set shape text to Excel formula with LEFT and RIGHT functions | link shape text to cell value using formula in Aspose.Cells for .NET | display part of a cell’s content in a rectangle shape with Aspose.Cells | how to bind a shape to a cell using substring functions in Aspose.Cells | C# Aspose.Cells example linking shape to cell using LEFT and RIGHT
// Tags: shape text formula binding Aspose.Cells | rectangle shape with LEFT function C# | RIGHT function substring for shape display | link shape to cell value .NET | Aspose.Cells shape text Excel formula example

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample creates a new workbook, writes "AsposeCellsDemo" to cell A1, adds a rectangle shape, sets its Text property to the formula "=LEFT(A1,6) & \" \" & RIGHT(A1,4)" to show substrings of A1, and saves the file as LinkedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put sample text into cell A1
            worksheet.Cells["A1"].PutValue("AsposeCellsDemo");

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top (points), left (points), height (points), width (points)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                0,    // upper left column
                0,    // top (in points)
                150,  // left (in points)
                30,   // height (in points)
                150   // width (in points)
            );

            // Link the shape's text to cell A1 using LEFT and RIGHT functions
            shape.Text = "=LEFT(A1,6) & \" \" & RIGHT(A1,4)";

            // Define output file path
            string outputPath = "LinkedShape.xlsx";

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
