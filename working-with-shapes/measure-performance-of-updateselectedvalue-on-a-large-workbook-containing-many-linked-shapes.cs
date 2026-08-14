// Title: Aspose.Cells for .NET – Benchmark UpdateSelectedValue with Thousands of Linked ListBox Shapes
// Description: C# sample that creates a workbook, fills column A with values 1‑1000, adds 1,000 ListBox shapes each linked to a distinct cell in column B, and measures the execution time of sheet.Shapes.UpdateSelectedValue using Stopwatch before saving the file.
// Keywords: Aspose.Cells | UpdateSelectedValue | performance benchmark | .NET | C# example | ListBox shape | linked cell | large workbook | shape processing speed | stopwatch timing | GitHub
// Common Searches: Aspose.Cells UpdateSelectedValue performance test | measure time for UpdateSelectedValue with many shapes | benchmark ListBox shape linked cells .NET | how fast is Shapes.UpdateSelectedValue on large workbook | C# code to time UpdateSelectedValue for thousands of shapes
// Developer Intent: Evaluate how quickly Shapes.UpdateSelectedValue processes a worksheet that contains a high volume of linked ListBox controls.
// Use Cases: Confirm that bulk updating of ListBox selections meets latency requirements for reporting dashboards. | Compare runtimes before and after changing input ranges, linked cells, or disabling events. | Validate scalability of interactive Excel reports generated with Aspose.Cells in high‑volume scenarios.
// AI Prompts: Generate C# code that adds 5,000 ListBox shapes, links each to a unique cell, and records the duration of sheet.Shapes.UpdateSelectedValue with high‑precision timing. | Suggest ways to accelerate UpdateSelectedValue for worksheets containing many linked shapes, such as turning off calculation or event handling. | Provide a snippet that logs the elapsed time of UpdateSelectedValue in seconds to a JSON file for later analysis.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPerformanceDemo
{
    // C# sample that creates a workbook, fills column A with values 1‑1000, adds 1,000 ListBox shapes each linked to a distinct cell in column B, and measures the execution time of sheet.Shapes.UpdateSelectedValue using Stopwatch before saving the file.
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
                sheet.Cells[i, 0].Value = i + 1;
            }

            // Add a large number of ListBox shapes, each linked to a different cell
            int shapeCount = 1000; // adjust as needed for performance testing
            for (int i = 0; i < shapeCount; i++)
            {
                // Position each shape in a separate row to avoid overlap
                int row = i;
                int col = 2; // column C
                int upperLeftRow = row;
                int upperLeftColumn = col;
                int top = 5;
                int left = 5;
                int width = 100;
                int height = 20;

                // Add ListBox shape
                Shape shape = sheet.Shapes.AddListBox(upperLeftRow, upperLeftColumn, top, left, width, height);
                // Set the input range (same for all shapes in this example)
                shape.SetInputRange("$A$1:$A$10", false, false);
                // Link each shape to a unique cell in column B (e.g., B1, B2, ...)
                string linkedCell = $"$B${i + 1}";
                shape.SetLinkedCell(linkedCell, false, true);
                // Initialize linked cell with a value that matches one of the input items
                sheet.Cells[i, 1].Value = (i % 10) + 1; // values 1..10
            }

            // Measure performance of updating selected values for all shapes
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Update selected values for all shapes in the worksheet
            sheet.Shapes.UpdateSelectedValue();

            sw.Stop();
            Console.WriteLine($"Time taken to update selected values for {shapeCount} shapes: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (using the standard save method)
            workbook.Save("PerformanceUpdateSelectedValue.xlsx");
        }
    }
}
