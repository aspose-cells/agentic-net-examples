// Title: Benchmark Find Performance in a Large Named Range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills a 10,000‑row × 10‑column area, defines it as a named range, configures FindOptions to limit the search, uses Stopwatch to time cells.Find for a value near the end, prints elapsed milliseconds and the cell address, then saves the file.
// Keywords: Aspose.Cells | C# | .NET | Find performance | named range search | benchmark cell lookup | Stopwatch timing | FindOptions | CellArea | large worksheet | search duration | performance measurement
// Common Searches: Aspose.Cells measure Find execution time | How to time a Find operation in a named range using C# | Benchmark cell search performance Aspose.Cells .NET | Search duration for large worksheet Aspose.Cells | FindOptions SetRange performance test
// Developer Intent: Determine how long a Find call takes when it is restricted to a large named range in an Aspose.Cells workbook.
// Use Cases: Profile and optimize search speed for massive data grids. | Validate that a specific value exists within a defined range while capturing latency. | Compare the runtime of full‑sheet searches versus named‑range‑limited searches. | Log search metrics for performance monitoring in data‑processing pipelines. | Create unit tests that assert acceptable lookup times for critical worksheets.
// AI Prompts: Generate C# code that times a cells.Find call inside a named range using Aspose.Cells and outputs the elapsed milliseconds. | Show how to modify the example to write the search duration and result to a log file instead of the console. | Provide a version that repeats the Find operation for multiple values and aggregates average search time. | Explain how to integrate this benchmark into an automated performance test suite for Aspose.Cells. | Suggest ways to visualize the timing results using a chart library after the search completes.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, fills a 10,000‑row × 10‑column area, defines it as a named range, configures FindOptions to limit the search, uses Stopwatch to time cells.Find for a value near the end, prints elapsed milliseconds and the cell address, then saves the file.
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
            int totalColumns = 10;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    cells[row, col].PutValue(row * totalColumns + col);
                }
            }

            // Create a named range that covers the populated area
            AsposeRange largeRange = cells.CreateRange(0, 0, totalRows, totalColumns);
            largeRange.Name = "LargeRange";

            // Configure FindOptions to limit the search to the named range
            FindOptions findOptions = new FindOptions();
            CellArea searchArea = new CellArea
            {
                StartRow = largeRange.FirstRow,
                StartColumn = largeRange.FirstColumn,
                EndRow = largeRange.FirstRow + largeRange.RowCount - 1,
                EndColumn = largeRange.FirstColumn + largeRange.ColumnCount - 1
            };
            findOptions.SetRange(searchArea);
            findOptions.LookInType = LookInType.Values;
            findOptions.LookAtType = LookAtType.EntireContent;

            // Measure the time taken to find a value near the end of the range
            Stopwatch stopwatch = Stopwatch.StartNew();
            Cell foundCell = cells.Find(totalRows * totalColumns - 1, null, findOptions);
            stopwatch.Stop();

            // Output the duration and result
            Console.WriteLine($"Search duration: {stopwatch.ElapsedMilliseconds} ms");
            if (foundCell != null)
            {
                Console.WriteLine($"Value found at cell: {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("Value not found.");
            }

            // Save the workbook (optional, demonstrates the save rule)
            string outputPath = "LargeRangeSearch.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
