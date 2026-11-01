// Title: Create a rectangle shape linked to a cell formatted with thousand separators using Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape that displays a cell value with thousand separators in Aspose.Cells C# | Aspose.Cells C# linking shape text to a worksheet cell using a formula | apply custom number format '#,##0.00' to a cell and show it inside a shape with Aspose.Cells | center text inside a rectangle shape on an Excel worksheet with Aspose.Cells for .NET | save a workbook containing a linked shape as an .xlsx file using Aspose.Cells
// Tags: add rectangle shape to worksheet Aspose.Cells | link shape text to cell formula Aspose.Cells | custom numeric format with thousand separator Aspose.Cells | center shape text horizontally vertically Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new workbook, writes a numeric value to cell A1, applies the custom number format "#,##0.00" (thousand separator), adds a rectangle shape at row 2 column 2, links the shape's text to cell A1 with a formula, centers the text inside the shape, and saves the file as Output.xlsx.
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

            // Set a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234567.89);

            // Apply custom number format with thousand separator
            Style style = cell.GetStyle();
            style.Custom = "#,##0.00";
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Link the shape's text to cell A1 using a formula
            shape.Text = "=A1";

            // Center the text inside the shape
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
