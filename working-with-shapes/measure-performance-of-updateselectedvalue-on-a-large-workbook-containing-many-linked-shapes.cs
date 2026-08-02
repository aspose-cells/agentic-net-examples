// Title: Benchmark Shape.UpdateSelectedValue with 1,000 Linked ListBox Shapes in Aspose.Cells for .NET
// Description: C# sample that creates a workbook, fills column A with sequential values, adds 1,000 ListBox shapes each linked to its row cell, then measures the time required for Shapes.UpdateSelectedValue to synchronize all linked shapes before saving the file.
// Keywords: Aspose.Cells | .NET | C# | Shape.UpdateSelectedValue | performance benchmark | linked shapes | ListBox shape | large workbook | Excel automation | code example | GitHub sample
// Common Searches: Aspose.Cells Shape.UpdateSelectedValue performance | benchmark linked ListBox shapes in .NET | how long does UpdateSelectedValue take with many shapes | measure shape update speed Aspose.Cells | C# example for updating linked shape values
// Developer Intent: Evaluate the execution time of Shapes.UpdateSelectedValue when thousands of ListBox shapes are linked to worksheet cells.
// Use Cases: Determine if shape‑linked updates meet latency requirements for real‑time dashboards. | Compare performance impact of different shape counts during workbook generation. | Identify bottlenecks before optimizing large Excel reports that use many linked controls.
// AI Prompts: Generate a C# script that logs Shape.UpdateSelectedValue execution time for 100, 500, and 2000 linked ListBox shapes and plots the results. | Suggest optimization techniques to reduce the runtime of UpdateSelectedValue when handling thousands of linked shapes in Aspose.Cells. | Create a unit test that verifies UpdateSelectedValue completes within 150 ms for 500 linked ListBox shapes.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPerformanceDemo
{
    // C# sample that creates a workbook, fills column A with sequential values, adds 1,000 ListBox shapes each linked to its row cell, then measures the time required for Shapes.UpdateSelectedValue to synchronize all linked shapes before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Number of shapes to create for the performance test
            const int shapeCount = 1000;

            // Populate cells that will be linked to the shapes
            for (int i = 0; i < shapeCount; i++)
            {
                // Fill column A with sequential numbers (these will be the linked values)
                sheet.Cells[i, 0].Value = i + 1;
            }

            // Add ListBox shapes and link each one to a corresponding cell in column A
            for (int i = 0; i < shapeCount; i++)
            {
                // Position each shape in its own row to avoid overlap
                int row = i;
                int col = 2; // start from column C to leave space for data

                // Add a ListBox shape (height and width are arbitrary)
                Shape shape = sheet.Shapes.AddListBox(row, col, 0, 0, 120, 20);

                // Set the input range (the list of items) – using the same range for simplicity
                shape.SetInputRange("$A$1:$A$10", false, false);

                // Link the shape's selected value to the cell in column A of the same row
                string linkedCellAddress = $"$A${row + 1}";
                shape.SetLinkedCell(linkedCellAddress, false, true);
            }

            // Measure the time taken to update selected values for all shapes
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // This updates the selected value of each shape based on its linked cell
            sheet.Shapes.UpdateSelectedValue();

            sw.Stop();
            Console.WriteLine($"Time taken to update {shapeCount} linked shapes: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional, just to verify that everything works)
            workbook.Save("PerformanceTestResult.xlsx");
        }
    }
}
