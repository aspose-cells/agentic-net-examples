using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure each worksheet has a summary cell (A1) with some text
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            ws.Name = $"Sheet{i + 1}";
            ws.Cells["A1"].PutValue($"Summary for {ws.Name}");
        }

        // Iterate through all worksheets, add a shape, link it to the summary cell, and log the action
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Add a rectangle shape (row 2, column 2, height 80, width 200)
            Shape shape = ws.Shapes.AddRectangle(2, 0, 2, 0, 80, 200);
            shape.Name = $"LinkShape_{ws.Name}";

            // Link the shape to cell A1 of the current worksheet
            shape.SetLinkedCell("A1", false, false);

            // Retrieve the linked cell address for logging
            string linkedCell = shape.GetLinkedCell(false, false);

            // Output the result
            Console.WriteLine($"Added shape '{shape.Name}' to worksheet '{ws.Name}' linked to cell {linkedCell}");
        }

        // Save the workbook
        workbook.Save("ShapesLinked.xlsx");
    }
}