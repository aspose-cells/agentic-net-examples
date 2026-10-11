// Title: How to link a rectangle shape to a cell that uses the TEXT function for currency formatting in Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, write a raw number to A1, apply the TEXT formula in B1 for currency formatting, add a rectangle shape, and set its Text property to "=B1" using Aspose.Cells in C#. | Generate a .xlsx file where a rectangle shape displays the result of a TEXT function applied to a numeric cell by linking the shape's text to the formula cell with the Aspose.Cells API. | Programmatically bind a shape's displayed text to a cell containing a TEXT formula that formats a number as currency, using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# link shape text to cell with TEXT formula for currency | How to display formatted number from TEXT function inside an Excel shape using Aspose.Cells | Set rectangle shape text to reference a formula cell in Aspose.Cells .NET | C# Aspose.Cells example linking shape to cell that uses TEXT function | Display TEXT function result in shape text property Aspose.Cells
// Tags: link shape text to formula cell Aspose.Cells | rectangle shape TEXT function binding C# | currency formatting with TEXT function Aspose.Cells | Aspose.Cells shape text reference .xlsx | C# Aspose.Cells bind shape to cell

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a workbook, stores a raw number in A1, formats it as currency in B1 using the TEXT function, adds a rectangle shape, links the shape's text to B1 so it shows the formatted value, and saves the file as an .xlsx workbook.
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

            // Put a raw number into cell A1
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(12345.6789);

            // In cell B1 use the TEXT function to format the number as currency
            Cell cellB1 = sheet.Cells["B1"];
            cellB1.Formula = @"=TEXT(A1,""$#,##0.00"")";

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 0, 60, 200);

            // Link the shape's text to cell B1 so it displays the formatted number
            shape.Text = "=B1";

            // Optional: set shape appearance
            shape.Line.Weight = 0.75;
            shape.Line.DashStyle = MsoLineDashStyle.Solid;
            // Line color and fill color properties may not be available in all versions; they are omitted for compatibility.

            // Save the workbook
            workbook.Save("LinkedShapeWithTextFunction.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
