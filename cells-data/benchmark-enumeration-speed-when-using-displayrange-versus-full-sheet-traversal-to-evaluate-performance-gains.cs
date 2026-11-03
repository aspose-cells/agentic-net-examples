// Title: Benchmarking cell enumeration speed: DisplayRange (MinRow/MaxDataRow) vs full sheet traversal in Aspose.Cells for .NET C#
// Common Searches: C# Aspose.Cells benchmark iteration over visible range versus full worksheet cells | measure performance of DisplayRange-like traversal in Aspose.Cells .NET | how fast is cell enumeration using MinRow and MaxDataRow compared to looping all cells in Aspose.Cells | Aspose.Cells performance test for summing values in a large worksheet C# | compare execution time of sheet.Cells loop vs calculated data range in Aspose.Cells
// Tags: Aspose.Cells cell enumeration performance | DisplayRange traversal benchmark .NET | full worksheet iteration timing Aspose.Cells | MinRow MaxDataRow range iteration C# | Excel workbook performance testing Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example creates a workbook with 10,000 rows and 100 columns of numeric data, then uses Stopwatch to compare the elapsed time for summing values by iterating over the entire Cells collection versus iterating only within the calculated data range (using MinRow, MinColumn, MaxDataRow, MaxDataColumn as a DisplayRange substitute). It outputs both timings and sums, and saves the workbook as an XLSX file.
class BenchmarkDisplayRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with sample data (10,000 rows x 100 columns)
            const int totalRows = 10000;
            const int totalCols = 100;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    sheet.Cells[row, col].PutValue(row * totalCols + col);
                }
            }

            // Benchmark full sheet traversal (using Cells collection)
            Stopwatch swFull = Stopwatch.StartNew();
            double sumFull = 0;
            for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
                {
                    object val = sheet.Cells[row, col].Value;
                    if (val is double d)
                        sumFull += d;
                }
            }
            swFull.Stop();

            // Determine the range of cells that contain data (used as a substitute for DisplayRange)
            int firstRow = sheet.Cells.MinRow;
            int firstCol = sheet.Cells.MinColumn;
            int lastRow = sheet.Cells.MaxDataRow;
            int lastCol = sheet.Cells.MaxDataColumn;

            // Benchmark traversal using the determined range (visible cells only)
            Stopwatch swDisplay = Stopwatch.StartNew();
            double sumDisplay = 0;
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    object val = sheet.Cells[row, col].Value;
                    if (val is double d)
                        sumDisplay += d;
                }
            }
            swDisplay.Stop();

            // Output the benchmark results
            Console.WriteLine($"Full sheet traversal:   {swFull.ElapsedMilliseconds} ms, Sum = {sumFull}");
            Console.WriteLine($"DisplayRange traversal: {swDisplay.ElapsedMilliseconds} ms, Sum = {sumDisplay}");

            // Save the workbook
            string outputPath = "BenchmarkResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
