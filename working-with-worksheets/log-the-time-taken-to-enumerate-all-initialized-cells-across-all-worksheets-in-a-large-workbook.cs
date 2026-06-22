using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the large workbook (replace with the actual file path)
        Workbook workbook = new Workbook("LargeWorkbook.xlsx");

        // Start timing the enumeration process
        Stopwatch stopwatch = Stopwatch.StartNew();

        long initializedCellCount = 0;

        // Enumerate initialized cells in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Obtain the cells enumerator for the current worksheet
            IEnumerator enumerator = sheet.Cells.GetEnumerator();

            // Iterate through all instantiated cells
            while (enumerator.MoveNext())
            {
                // Cast the current object to Cell (required by the enumerator)
                Cell cell = (Cell)enumerator.Current;

                // Example processing: simply count the cell
                initializedCellCount++;
            }
        }

        // Stop timing
        stopwatch.Stop();

        // Log the results
        Console.WriteLine($"Total initialized cells: {initializedCellCount}");
        Console.WriteLine($"Time taken (ms): {stopwatch.ElapsedMilliseconds}");

        // Save the workbook if needed (demonstrates the save rule)
        workbook.Save("ProcessedWorkbook.xlsx");
    }
}