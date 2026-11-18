// Title: How to bind a rectangle shape to a VLOOKUP result cell in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link rectangle shape text to a cell formula result | Update shape text automatically when VLOOKUP value changes with Aspose.Cells | Set shape text to =D2 in Aspose.Cells workbook C# example | Bind Excel shape to lookup cell using Aspose.Cells .NET API | Dynamic shape content based on VLOOKUP in Aspose.Cells C#
// Tags: bind shape text to worksheet cell Aspose.Cells | rectangle shape linked to VLOOKUP output C# | dynamic shape content from cell reference .NET | Aspose.Cells set shape text via formula result | Excel shape auto-updates with lookup value using Aspose

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new workbook, builds a lookup table, places a VLOOKUP formula in D2, adds a rectangle shape, assigns the shape's text to "=D2" so it reflects the lookup result, applies basic line formatting, and saves the file as ShapeLinkedToVlookup.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Build lookup table (A5:B9) -----
            sheet.Cells["A5"].PutValue("Apple");
            sheet.Cells["B5"].PutValue(10);
            sheet.Cells["A6"].PutValue("Banana");
            sheet.Cells["B6"].PutValue(20);
            sheet.Cells["A7"].PutValue("Cherry");
            sheet.Cells["B7"].PutValue(30);
            sheet.Cells["A8"].PutValue("Date");
            sheet.Cells["B8"].PutValue(40);
            sheet.Cells["A9"].PutValue("Elderberry");
            sheet.Cells["B9"].PutValue(50);

            // ----- Value to look up (A2) -----
            sheet.Cells["A2"].PutValue("Cherry");

            // ----- VLOOKUP formula result in D2 -----
            sheet.Cells["D2"].Formula = "VLOOKUP(A2,$A$5:$B$9,2,FALSE)";

            // ----- Add a rectangle shape -----
            // Parameters: type, upper left row, upper left row offset, upper left column, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 200, 50);

            // Link shape's text to the VLOOKUP result cell D2
            shape.Text = "=D2";

            // Optional formatting for visibility
            shape.Line.Weight = 1;
            shape.Line.DashStyle = MsoLineDashStyle.Solid;

            // Save the workbook
            workbook.Save("ShapeLinkedToVlookup.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
