// Title: Benchmark Find operation in a large named range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills a 10,000‑row × 10‑column area, defines a named range over that block, configures FindOptions with a matching CellArea, searches for the value in the last cell, and uses Stopwatch to report the elapsed milliseconds. The workbook can then be saved.
// Keywords: Aspose.Cells | C# | .NET | FindOptions | named range | performance benchmark | search time | cell lookup | large worksheet | stopwatch timing | Excel processing speed
// Common Searches: measure find performance Aspose.Cells C# | time search in named range Aspose.Cells | benchmark cell lookup .NET | how long does Find take in large Excel sheet | Aspose.Cells search duration example
// Developer Intent: Determine the execution time of a Find call when it is limited to a large named range.
// Use Cases: Compare different FindOptions settings to identify the fastest configuration for massive worksheets. | Validate that a critical value can be located within acceptable latency in reporting pipelines. | Log search duration as a health metric to detect performance regressions in automated Excel workflows.
// AI Prompts: Generate C# code that uses Aspose.Cells FindOptions with SetRange to search a named range and measures the elapsed time. | Suggest optimizations for speeding up Find operations on large ranges in Aspose.Cells, such as changing search order or using alternative APIs. | Create a loop that repeats the timed search multiple times and returns average, min, and max durations for reliable benchmarking.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook, fills a 10,000‑row × 10‑column area, defines a named range over that block, configures FindOptions with a matching CellArea, searches for the value in the last cell, and uses Stopwatch to report the elapsed milliseconds. The workbook can then be saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a large range (e.g., 10,000 rows × 10 columns)
            int totalRows = 10000;
            int totalCols = 10;
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalCols; j++)
                {
                    cells[i, j].PutValue($"R{i}C{j}");
                }
            }

            // Create a named range that covers the populated area
            Aspose.Cells.Range range = cells.CreateRange(0, 0, totalRows, totalCols);
            range.Name = "LargeRange";

            // Configure FindOptions to limit the search to the named range
            FindOptions findOptions = new FindOptions();
            CellArea searchArea = new CellArea
            {
                StartRow = range.FirstRow,
                StartColumn = range.FirstColumn,
                EndRow = range.FirstRow + range.RowCount - 1,
                EndColumn = range.FirstColumn + range.ColumnCount - 1
            };
            findOptions.SetRange(searchArea);

            // Value to search for (placed at the very end to force full scan)
            string targetValue = $"R{totalRows - 1}C{totalCols - 1}";

            // Measure the time taken to perform the search
            Stopwatch stopwatch = Stopwatch.StartNew();
            Cell foundCell = cells.Find(targetValue, null, findOptions);
            stopwatch.Stop();

            Console.WriteLine($"Search completed in {stopwatch.ElapsedMilliseconds} ms.");
            Console.WriteLine(foundCell != null ? $"Found at {foundCell.Name}" : "Value not found");

            // Save the workbook (optional)
            workbook.Save("SearchTiming.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
