// Title: Create a rectangle shape linked to a worksheet cell using the TEXT function with a custom date format in Aspose.Cells for .NET
// Common Searches: how to bind a shape's text to a cell value using TEXT function in Aspose.Cells C# | Aspose.Cells set shape text to formatted date from cell A1 | recalculate formulas after linking shape to cell in Aspose.Cells .NET | add rectangle shape to worksheet and display date with custom format using Aspose.Cells | retrieve evaluated shape text after formula calculation in Aspose.Cells C#
// Tags: rectangle shape addition Aspose.Cells | shape text formula linking Aspose.Cells | custom date formatting in shape text | recalculate worksheet formulas Aspose.Cells | export workbook with linked shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes the current date into cell A1, applies a built‑in date style, adds a rectangle shape, sets the shape's Text property to the formula =TEXT(A1,"yyyy-MM-dd"), forces formula calculation, reads the evaluated text from the shape, outputs it, and saves the workbook as ShapeLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put the current date into cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.PutValue(DateTime.Now);

            // Apply a built‑in date format (optional)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,                        // upper left row (zero‑based)
                1,                        // upper left column (zero‑based)
                5,                        // top offset (pixels)
                5,                        // left offset (pixels)
                30,                       // height (pixels)
                200);                     // width (pixels)

            // Link the shape's text to cell A1 using the TEXT function with a custom date format
            shape.Text = "=TEXT(A1,\"yyyy-MM-dd\")";

            // Recalculate all formulas so the shape displays the evaluated value
            workbook.CalculateFormula();

            // Retrieve and display the evaluated text from the shape
            string displayedText = shape.Text;
            Console.WriteLine("Shape displays: " + displayedText);

            // Save the workbook (optional)
            string outputPath = "ShapeLinkedToCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
