using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPerformanceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a column with sample data (A1:A1000)
            for (int i = 0; i < 1000; i++)
            {
                sheet.Cells[i, 0].Value = i + 1; // Values 1..1000
            }

            // Add a large number of ListBox shapes, each linked to a different cell
            const int shapeCount = 5000; // Adjust to simulate a "large" workbook
            for (int i = 0; i < shapeCount; i++)
            {
                // Position each shape in a grid to avoid overlap
                int row = i / 10;
                int col = i % 10;
                int upperLeftRow = row * 2;
                int upperLeftColumn = col * 2;

                // Add ListBox shape
                Shape shape = sheet.Shapes.AddListBox(upperLeftRow, upperLeftColumn, 0, 0, 120, 120);
                // Set the input range (same for all shapes)
                shape.SetInputRange("$A$1:$A$1000", false, false);
                // Link each shape to a unique cell in column B
                string linkedCell = $"$B${i + 1}";
                shape.SetLinkedCell(linkedCell, false, true);
                // Initialize linked cell value
                sheet.Cells[i, 1].Value = (i % 1000) + 1; // Value within input range
            }

            // Ensure formulas are calculated (if any)
            workbook.CalculateFormula();

            // Measure performance of updating selected values for all shapes
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Update selected values for all shapes in the worksheet
            sheet.Shapes.UpdateSelectedValue();

            sw.Stop();

            Console.WriteLine($"Time taken to update selected values for {shapeCount} shapes: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional, just to verify that changes are persisted)
            workbook.Save("PerformanceDemo.xlsx");
        }
    }
}