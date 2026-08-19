// Title: Aspose.Cells .NET: Benchmark Row‑Major vs Column‑Major Worksheet Population
// Description: Creates a 2000 × 2000 double array, writes it to an Aspose.Cells worksheet using row‑major and column‑major loops, measures each elapsed time, clears the sheet, and saves the workbook with the last method's data.
// Keywords: Aspose.Cells | C# | .NET | worksheet population performance | row-major loop | column-major loop | Excel write speed | large dataset export | benchmarking | performance testing
// Common Searches: Aspose.Cells row major vs column major speed | benchmark Excel write performance C# | optimal loop order for filling large worksheets | measure Aspose.Cells cell insertion time | how to speed up Excel export with Aspose.Cells
// Developer Intent: Identify which iteration order—row‑major or column‑major—yields faster worksheet creation with Aspose.Cells.
// Use Cases: Compare loop strategies before implementing high‑volume Excel exports. | Select the fastest data‑population method for reporting tools that generate massive spreadsheets. | Validate performance impact of different cell‑writing orders in performance‑critical applications.
// AI Prompts: Show how to parallelize column‑major population with Parallel.For to improve Aspose.Cells write speed. | Explain the benefits of Aspose.Cells ImportData or FastDataProvider over manual cell loops for large datasets.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // Creates a 2000 × 2000 double array, writes it to an Aspose.Cells worksheet using row‑major and column‑major loops, measures each elapsed time, clears the sheet, and saves the workbook with the last method's data.
    class Program
    {
        static void Main()
        {
            // Parameters for the test data
            const int rows = 2000;      // number of rows
            const int cols = 2000;      // number of columns
            const string outputFile = "PerformanceResult.xlsx";

            // Prepare a 2‑dimensional array with sample data
            double[,] data = new double[rows, cols];
            Random rnd = new Random();
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    data[r, c] = rnd.NextDouble();

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------- Row‑major population --------------------
            Stopwatch swRow = Stopwatch.StartNew();

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(data[r, c]);
                }
            }

            swRow.Stop();
            Console.WriteLine($"Row‑major population time: {swRow.ElapsedMilliseconds} ms");

            // Clear the sheet before the next test
            cells.Clear();

            // -------------------- Column‑major population --------------------
            Stopwatch swCol = Stopwatch.StartNew();

            for (int c = 0; c < cols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    cells[r, c].PutValue(data[r, c]);
                }
            }

            swCol.Stop();
            Console.WriteLine($"Column‑major population time: {swCol.ElapsedMilliseconds} ms");

            // Save the workbook (contains the data populated by the last method)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
    }
}
