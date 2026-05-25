using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
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

                // Populate a large area with sample data (e.g., 2000 rows x 2000 columns)
                const int totalRows = 2000;
                const int totalCols = 2000;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        cells[row, col].PutValue(row * col);
                    }
                }

                // -----------------------------------------------------------------
                // Benchmark: enumeration using MaxDisplayRange (DisplayRange)
                // -----------------------------------------------------------------
                // Resolve ambiguity between Aspose.Cells.Range and System.Range
                Aspose.Cells.Range displayRange = cells.MaxDisplayRange; // Gets the range that includes data, merged cells and shapes
                Stopwatch sw = new Stopwatch();
                sw.Start();

                IEnumerator displayEnum = displayRange.GetEnumerator();
                while (displayEnum.MoveNext())
                {
                    // Access each cell to simulate realistic work
                    Cell cell = (Cell)displayEnum.Current;
                    var _ = cell.Value;
                }

                sw.Stop();
                Console.WriteLine($"DisplayRange enumeration time: {sw.ElapsedMilliseconds} ms");

                // -----------------------------------------------------------------
                // Benchmark: enumeration over the entire Cells collection (full sheet)
                // -----------------------------------------------------------------
                sw.Reset();
                sw.Start();

                IEnumerator cellsEnum = cells.GetEnumerator();
                while (cellsEnum.MoveNext())
                {
                    Cell cell = (Cell)cellsEnum.Current;
                    var _ = cell.Value;
                }

                sw.Stop();
                Console.WriteLine($"Full Cells enumeration time: {sw.ElapsedMilliseconds} ms");

                // Save the workbook (optional, just to keep the file)
                string outputPath = "BenchmarkResult.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during execution: {ex.Message}");
            }
        }
    }
}