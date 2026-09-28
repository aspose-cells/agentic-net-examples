// Title: Add a rectangle shape to a worksheet cell that contains a column header and verify its alignment using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a rectangle shape at the cell of a specific header (e.g., C1) by assigning UpperLeftRow and UpperLeftColumn with Aspose.Cells. | Compare the shape's UpperLeftRow and UpperLeftColumn to the header cell coordinates and output whether they match. | Save the workbook after positioning the shape and display the full file path of the saved Excel file.
// Common Searches: Aspose.Cells C# place shape on specific cell coordinates | how to position a rectangle shape over a header row in Excel using Aspose.Cells | verify shape alignment with cell in Aspose.Cells .NET example | set shape UpperLeftRow UpperLeftColumn to match header cell Aspose.Cells | C# Aspose.Cells add shape and check alignment with column header
// Tags: add rectangle shape to worksheet cell Aspose.Cells | set shape UpperLeftRow UpperLeftColumn C# | align shape with header cell Aspose.Cells | validate shape position against cell coordinates | save workbook after shape alignment Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Demonstrates creating a workbook, inserting a header value, adding a rectangle shape at the same cell, checking that the shape's UpperLeftRow and UpperLeftColumn match the header cell, and saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the header text and its location (e.g., column C, row 1)
            int headerRow = 0;          // zero‑based index for the first row
            int headerColumn = 2;       // zero‑based index for column C
            sheet.Cells[headerRow, headerColumn].PutValue("TargetHeader");

            // Add a rectangle shape positioned at the same cell as the header
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, headerRow, headerColumn, 0, 0, 50, 100);
            shape.Name = "AlignedShape";

            // Validate that the shape's upper‑left cell matches the header cell
            bool isAligned = shape.UpperLeftRow == headerRow && shape.UpperLeftColumn == headerColumn;

            Console.WriteLine(isAligned ? "Shape aligned correctly." : "Shape misaligned.");

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "AlignedShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
