// Title: Link a textbox shape to a DATEVALUE formula cell and display the formatted date using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link textbox shape to a cell with DATEVALUE formula | display formatted date in a linked shape textbox in Excel using Aspose.Cells | center text of a shape that references a date formula cell with Aspose.Cells | set built‑in date number format for a cell and bind it to a shape in C# | save workbook with shape text bound to DATEVALUE cell using Aspose.Cells
// Tags: textbox shape cell binding Aspose.Cells C# | DATEVALUE formula cell formatting Aspose.Cells | centered shape text reference date cell | apply built‑in date number format Aspose.Cells | save workbook with linked shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a DATEVALUE formula to cell B1, applies a built‑in date format, adds a textbox shape, links the shape's text to the formatted date cell, centers the text, and saves the file as LinkedShapeDate.xlsx.
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

            // Set a DATEVALUE formula in cell B1
            Cell dateCell = sheet.Cells["B1"];
            dateCell.Formula = "DATEVALUE(\"2023-09-01\")";

            // Apply a date format to the cell so it displays as a formatted date
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format (mm-dd-yy)
            dateCell.SetStyle(dateStyle);

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            TextBox shape = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);

            // Link the shape's text to the cell containing the DATEVALUE formula
            shape.Text = "=B1";

            // Align the shape's text to the center
            shape.TextHorizontalAlignment = TextAlignmentType.Center;

            // Save the workbook
            string outputPath = "LinkedShapeDate.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
