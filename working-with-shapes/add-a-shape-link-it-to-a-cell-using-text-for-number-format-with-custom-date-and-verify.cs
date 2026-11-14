// Title: Create a rectangle shape linked to a cell with a custom date format using the TEXT function in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to a cell using TEXT formula with custom date format | display formatted date from cell A1 inside a rectangle shape in Aspose.Cells | force formula calculation for shape text after adding a shape in Aspose.Cells .NET | apply built‑in date number format and bind it to a shape in Aspose.Cells | save workbook with shape linked to cell using TEXT function in C#
// Tags: rectangle shape linked to worksheet cell Aspose.Cells | TEXT function custom date format in shape text | calculate shape formulas Aspose.Cells .NET | built‑in date number format ID 14 Aspose.Cells | bind shape text to cell C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes the current date to cell A1, applies the built‑in date format (ID 14), adds a rectangle shape at row 2 column 2, sets the shape's text to a TEXT formula with a custom "yyyy-MM-dd" format referencing A1, forces formula calculation so the shape displays the formatted date, reads the displayed text, and saves the file as ShapeLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Put a date value into cell A1
            ws.Cells["A1"].PutValue(DateTime.Now);

            // Set a built‑in date format for the cell (Number format ID 14)
            Style dateStyle = ws.Cells["A1"].GetStyle();
            dateStyle.Number = 14;
            ws.Cells["A1"].SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 30, 100);

            // Link the shape's text to cell A1 using TEXT function with a custom date format
            shape.Text = "=TEXT(A1,\"yyyy-MM-dd\")";

            // Calculate formulas so the shape displays the result
            wb.CalculateFormula();

            // Verify: read the displayed text from the shape
            string displayedText = shape.Text;
            Console.WriteLine("Shape text: " + displayedText);

            // Define output file path
            string outputPath = "ShapeLinkedToCell.xlsx";

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
