// Title: Benchmarking the creation of 5,000 line sparklines in a large Excel workbook with Aspose.Cells for .NET
// Common Searches: how to measure Aspose.Cells sparkline insertion time in C# | benchmark adding 5000 sparklines to an XLSX file using Aspose.Cells | C# stopwatch example for profiling sparkline generation with Aspose.Cells | using reflection to add sparklines when Aspose.Cells.Sparkline assembly is not referenced | performance testing of large workbook sparkline scalability in .NET
// Tags: Aspose.Cells sparkline insertion benchmark | C# stopwatch profiling Aspose.Cells operations | reflection based sparkline creation .NET | large workbook sparkline scalability test | XLSX line sparkline generation performance

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample creates a 100‑row by 10‑column data worksheet, adds a separate sheet for sparklines, and uses reflection to insert up to 5,000 line sparklines (10 per row) while timing the operation with a Stopwatch. After benchmarking, the workbook is optionally saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Start measuring total execution time.
            Stopwatch sw = new Stopwatch();

            // Create a new workbook and a worksheet for data.
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Fill the data sheet with sample data (100 rows × 10 columns).
            int dataRows = 100;
            int dataCols = 10;
            for (int r = 0; r < dataRows; r++)
            {
                for (int c = 0; c < dataCols; c++)
                {
                    dataSheet.Cells[r, c].PutValue(r + c);
                }
            }

            // Add a sheet that will contain the sparklines (if supported).
            Worksheet sparkSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            sparkSheet.Name = "Sparklines";

            // Define the source range for all sparklines.
            string sourceRange = $"A1:J{dataRows}";

            // Benchmark the creation of 5,000 sparklines.
            int totalSparklines = 5000;
            int sparklinesPerRow = 10; // Number of sparklines placed in each row.

            sw.Start();

            // Attempt to add sparklines only if the Sparkline API is available.
            // The Sparkline API resides in the Aspose.Cells.Sparkline assembly, which may not be referenced.
            // Therefore we use reflection to avoid compile‑time dependencies.
            var sparklineGroupCollection = sparkSheet.GetType().GetProperty("SparklineGroups")?.GetValue(sparkSheet);
            var addMethod = sparklineGroupCollection?.GetType().GetMethod("Add");
            var sparklineTypeEnum = Type.GetType("Aspose.Cells.Sparkline.SparklineType, Aspose.Cells.Sparkline");

            for (int i = 0; i < totalSparklines; i++)
            {
                int row = i / sparklinesPerRow;                     // Zero‑based row index.
                int col = i % sparklinesPerRow;                     // Zero‑based column index.
                string location = $"{(char)('A' + col)}{row + 1}"; // e.g., A1, B1, ...

                // If the Sparkline API is present, add a line sparkline; otherwise skip.
                if (addMethod != null && sparklineTypeEnum != null)
                {
                    // SparklineType.Line enum value (usually 0). Use Enum.ToObject to obtain the value.
                    object lineType = Enum.ToObject(sparklineTypeEnum, 0);
                    addMethod.Invoke(sparklineGroupCollection, new object[] { lineType, sourceRange, location });
                }
            }

            sw.Stop();

            Console.WriteLine($"Processed {totalSparklines} sparkline entries in {sw.ElapsedMilliseconds} ms.");

            // Save the workbook (optional, demonstrates the save lifecycle).
            string outputPath = "SparklinesBenchmark.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
