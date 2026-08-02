// Title: Aspose.Cells .NET Benchmark: MaxDisplayRange vs Full Sheet Enumeration Speed
// Description: Creates a 5,000 × 100 workbook, fills each cell with a numeric value, then measures and displays the elapsed time for iterating cells using the worksheet's MaxDisplayRange enumerator versus the full Cells enumerator. The workbook is saved after the test.
// Keywords: Aspose.Cells | .NET | C# | cell enumeration | MaxDisplayRange | performance benchmark | worksheet traversal | speed test | large spreadsheet | enumeration speed
// Common Searches: Aspose.Cells MaxDisplayRange performance | benchmark cell iteration Aspose.Cells | enumerate used cells vs all cells Aspose.Cells | C# Aspose.Cells enumeration speed | measure worksheet traversal time Aspose.Cells
// Developer Intent: Compare the execution time of iterating cells through MaxDisplayRange with iterating the entire worksheet to identify the faster enumeration method.
// Use Cases: Optimize data‑processing loops for large Excel files | Choose the most efficient iteration strategy in memory‑constrained applications | Profile spreadsheet handling before exporting or performing calculations | Integrate enumeration speed tests into CI pipelines for performance regression detection
// AI Prompts: Create a benchmark that runs each enumeration method multiple times and reports average, min, and max durations. | Generate a function that selects the faster enumeration technique based on worksheet dimensions and density. | Suggest code changes to skip empty rows/columns while still leveraging MaxDisplayRange for maximum speed.

using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // Creates a 5,000 × 100 workbook, fills each cell with a numeric value, then measures and displays the elapsed time for iterating cells using the worksheet's MaxDisplayRange enumerator versus the full Cells enumerator. The workbook is saved after the test.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a large area with data to have a meaningful display range
                int totalRows = 5000;
                int totalCols = 100;
                for (int r = 0; r < totalRows; r++)
                {
                    for (int c = 0; c < totalCols; c++)
                    {
                        cells[r, c].PutValue(r * totalCols + c);
                    }
                }

                // Benchmark: enumerate cells using MaxDisplayRange (display range)
                Aspose.Cells.Range displayRange = cells.MaxDisplayRange; // Includes data, merged cells and shapes
                if (displayRange == null)
                {
                    Console.WriteLine("Worksheet is empty, cannot benchmark display range.");
                    return;
                }

                Stopwatch swDisplay = Stopwatch.StartNew();
                IEnumerator displayEnum = displayRange.GetEnumerator();
                while (displayEnum.MoveNext())
                {
                    Cell cell = (Cell)displayEnum.Current;
                    var val = cell.Value; // Simulate work
                }
                swDisplay.Stop();

                // Benchmark: enumerate all cells in the worksheet (full traversal)
                Stopwatch swFull = Stopwatch.StartNew();
                IEnumerator fullEnum = cells.GetEnumerator();
                while (fullEnum.MoveNext())
                {
                    Cell cell = (Cell)fullEnum.Current;
                    var val = cell.Value; // Simulate work
                }
                swFull.Stop();

                // Output the results
                Console.WriteLine($"Enumeration using MaxDisplayRange: {swDisplay.ElapsedMilliseconds} ms");
                Console.WriteLine($"Enumeration using full sheet traversal: {swFull.ElapsedMilliseconds} ms");

                // Save the workbook (optional, just to keep the data)
                string outputPath = "BenchmarkDisplayRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
