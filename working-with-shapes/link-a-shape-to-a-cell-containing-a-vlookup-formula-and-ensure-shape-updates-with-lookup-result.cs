// Title: How to bind a rectangle shape’s text to a VLOOKUP formula cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook in C# that creates a VLOOKUP table, writes a VLOOKUP formula in a cell, inserts a rectangle shape, and sets the shape's displayed text to reference that formula cell so the shape shows the lookup result. | Write C# code with Aspose.Cells to add a rectangle drawing to a worksheet, link its visible text to cell C2 containing =VLOOKUP(...), apply line formatting, and save the workbook as an .xlsx file. | Create a program that formats a rectangle shape (line weight, dash style) after binding its text to a VLOOKUP result cell, then exports the file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# bind shape text to cell with formula | display VLOOKUP result inside a shape using Aspose.Cells for .NET | set shape text to =C2 in generated Excel file with Aspose.Cells | make rectangle shape update automatically from a VLOOKUP cell in Aspose.Cells
// Tags: Aspose.Cells bind shape text to formula cell | C# rectangle shape displaying VLOOKUP output | link shape displayed text to cell C2 with Aspose.Cells | dynamic shape content driven by VLOOKUP in Excel | shape line formatting after linking in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, building a VLOOKUP table, inserting a rectangle shape, linking the shape's displayed text to the VLOOKUP result cell, applying line formatting, and saving the file as an .xlsx using Aspose.Cells for .NET.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Setup data for VLOOKUP ----------
            // Lookup table (E2:F5)
            sheet.Cells["E2"].PutValue("Key");
            sheet.Cells["F2"].PutValue("Value");
            sheet.Cells["E3"].PutValue("A");
            sheet.Cells["F3"].PutValue(100);
            sheet.Cells["E4"].PutValue("B");
            sheet.Cells["F4"].PutValue(200);
            sheet.Cells["E5"].PutValue("C");
            sheet.Cells["F5"].PutValue(300);

            // Input key for lookup (A2)
            sheet.Cells["A2"].PutValue("B");

            // Cell with VLOOKUP formula (C2)
            // =VLOOKUP(A2, $E$2:$F$5, 2, FALSE)
            sheet.Cells["C2"].Formula = "=VLOOKUP(A2, $E$2:$F$5, 2, FALSE)";

            // ---------- Add a shape and link it to the VLOOKUP result ----------
            // Insert a rectangle shape
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                0,   // upper left column
                5,   // top offset (pixels)
                0,   // left offset (pixels)
                200, // height (pixels)
                50   // width (pixels)
            );

            // Link shape's text to the cell containing the VLOOKUP formula.
            // In Excel, setting the text to a formula (e.g., "=C2") makes the shape display the cell's value.
            shape.Text = "=C2";

            // Optional: format the shape's line
            shape.Line.Weight = 1.0;
            shape.Line.DashStyle = MsoLineDashStyle.Solid;

            // Save the workbook
            string outputPath = "ShapeLinkedToVLookup.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
