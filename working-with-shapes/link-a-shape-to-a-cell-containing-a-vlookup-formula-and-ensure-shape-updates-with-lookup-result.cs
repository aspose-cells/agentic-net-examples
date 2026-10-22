// Title: How to link a rectangle shape to a VLOOKUP result cell so the shape text updates automatically using Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape on a worksheet and set its Text property to "=D2" to display the result of a VLOOKUP formula with Aspose.Cells. | Write C# code that builds a lookup table, adds a VLOOKUP formula in a cell, and binds a shape’s displayed text to that formula cell using Aspose.Cells. | Show how to make a shape’s text refresh automatically when the lookup key changes by linking the shape to the VLOOKUP result cell.
// Common Searches: asp.net c# link shape text to vlookup cell aspose.cells | how to bind an Excel shape to a formula result using Aspose.Cells | dynamic shape content based on VLOOKUP in Aspose.Cells C# example | update rectangle shape automatically when VLOOKUP key changes Aspose.Cells
// Tags: link shape text to formula cell Aspose.Cells | rectangle shape bound to VLOOKUP result .NET | dynamic shape content from worksheet calculation C# | Aspose.Cells bind shape to cell value | Excel shape updates with lookup value using Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, defines a lookup table, inserts a VLOOKUP formula, adds a rectangle shape, links the shape's Text property to the formula cell so the displayed text reflects the lookup result, applies basic formatting, and saves the file.
class ShapeLinkVlookupExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate lookup table (A2:B5)
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("Date");
            sheet.Cells["B5"].PutValue(40);

            // Cell that will contain the lookup key (C2)
            sheet.Cells["C2"].PutValue("Cherry");

            // Cell with VLOOKUP formula that returns the price (D2)
            // =VLOOKUP(C2, A2:B5, 2, FALSE)
            sheet.Cells["D2"].Formula = "=VLOOKUP(C2, A2:B5, 2, FALSE)";

            // Add a rectangle shape and obtain the Shape object directly
            Shape rect = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                0,   // upper left column
                5,   // top offset (pixels)
                0,   // left offset (pixels)
                200, // width (pixels)
                50   // height (pixels)
            );

            // Link the shape's text to the VLOOKUP result cell (D2)
            rect.Text = "=D2";

            // Optional: format the shape (fill color, line, etc.)
            // FillColor may not be available in older library versions; comment if not supported
            // rect.FillColor = Color.LightYellow;
            rect.Line.Weight = 1.0;
            rect.Line.DashStyle = MsoLineDashStyle.Solid;

            // Save the workbook
            string outputPath = "ShapeLinkedToVlookup.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
