// Title: Link a TEXTJOIN formula result to a rectangle shape in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that fills cells A1:A5, inserts a TEXTJOIN formula in C1, adds a rectangle shape, and sets the shape's text to =C1. | Show how to format a rectangle shape (fill color, line weight) and save the workbook as an .xlsx file after linking its text to a cell formula using Aspose.Cells. | Create a C# example that links multiple shapes to different TEXTJOIN results across several worksheets with Aspose.Cells.
// Common Searches: how to bind a shape's text to a cell formula with Aspose.Cells C# | using TEXTJOIN in Aspose.Cells and displaying result in a shape | Aspose.Cells C# rectangle shape linked to cell value example | set shape text to =C1 in Excel workbook via Aspose.Cells .NET | concatenate range values with TEXTJOIN and show in shape using Aspose.Cells
// Tags: Aspose.Cells link shape to cell formula | C# TEXTJOIN formula Aspose.Cells | rectangle shape text binding Aspose.Cells | format shape fill color Aspose.Cells | save workbook with linked shape .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Demonstrates creating a workbook, populating A1:A5, applying TEXTJOIN in C1, adding a rectangle shape, linking its text to the formula result, styling the shape, and saving as ShapeLinkedWithTextJoin.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells A1:A5 with sample text values
            sheet.Cells["A1"].PutValue("Apple");
            sheet.Cells["A2"].PutValue("Banana");
            sheet.Cells["A3"].PutValue("Cherry");
            sheet.Cells["A4"].PutValue("Date");
            sheet.Cells["A5"].PutValue("Elderberry");

            // In cell C1, set a TEXTJOIN formula that concatenates the values in A1:A5
            // The result will be: Apple, Banana, Cherry, Date, Elderberry
            sheet.Cells["C1"].Formula = @"TEXTJOIN("","", TRUE, A1:A5)";

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 1, 0, 0, 100, 200);

            // Link the shape's displayed text to the cell containing the TEXTJOIN result (C1)
            shape.Text = "=C1";

            // Optionally format the shape (e.g., fill color, line)
            shape.FillFormat.ForeColor = System.Drawing.Color.LightYellow;
            shape.Line.Weight = 1.0;

            // Define output file path
            string outputPath = "ShapeLinkedWithTextJoin.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
