// Title: Aspose.Cells C# Benchmark: Save Speed and File Size with OoxmlCompressionLevel1 vs Level9
// Description: C# sample that loads an Excel workbook with Aspose.Cells, saves it twice using OoxmlSaveOptions (CompressionType Level1 and Level9), measures each operation with Stopwatch, and reports the elapsed time and resulting file sizes.
// Keywords: Aspose.Cells C# compression benchmark | OoxmlCompressionType Level1 | OoxmlCompressionType Level9 | Excel save performance .NET | file size reduction Aspose.Cells | OoxmlSaveOptions example | measure workbook save time
// Common Searches: Aspose.Cells compare Level1 and Level9 compression speed | how to benchmark Excel save time with Aspose.Cells | C# code to measure file size after Ooxml compression | fast vs maximum compression in Aspose.Cells | save workbook with OoxmlSaveOptions compression type
// Developer Intent: Determine which OoxmlCompressionType (Level1 or Level9) offers the best trade‑off between save speed and output file size for a given workbook.
// Use Cases: Select an optimal compression level for large‑scale report generation. | Create minimal‑size Excel files for distribution by using Level9 compression. | Implement a rapid‑save mode for temporary files where speed outweighs size, using Level1 compression.
// AI Prompts: Generate C# code that logs both elapsed time and memory usage when saving a workbook with OoxmlCompressionType.Level1 and Level9. | Explain how to analyze the benchmark results to choose the appropriate compression level for production workloads. | Provide a script that runs the compression benchmark on multiple workbooks in parallel and aggregates timing and size statistics.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace CompressionPerformanceDemo
{
    // C# sample that loads an Excel workbook with Aspose.Cells, saves it twice using OoxmlSaveOptions (CompressionType Level1 and Level9), measures each operation with Stopwatch, and reports the elapsed time and resulting file sizes.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with an actual file path)
            string sourcePath = "input.xlsx";

            // Load the workbook using the string constructor (rule: Workbook(string))
            Workbook workbook = new Workbook(sourcePath);

            // Measure saving with fast compression (Level1)
            OoxmlSaveOptions fastOptions = new OoxmlSaveOptions(); // rule: OoxmlSaveOptions()
            fastOptions.CompressionType = OoxmlCompressionType.Level1; // rule: set CompressionType

            Stopwatch swFast = Stopwatch.StartNew();
            // Save using Save(string, SaveOptions) (rule)
            workbook.Save("output_Level1.xlsx", fastOptions);
            swFast.Stop();

            Console.WriteLine($"Saving with OoxmlCompressionType.Level1 took {swFast.ElapsedMilliseconds} ms.");

            // Measure saving with maximum compression (Level9)
            OoxmlSaveOptions maxOptions = new OoxmlSaveOptions();
            maxOptions.CompressionType = OoxmlCompressionType.Level9;

            Stopwatch swMax = Stopwatch.StartNew();
            workbook.Save("output_Level9.xlsx", maxOptions);
            swMax.Stop();

            Console.WriteLine($"Saving with OoxmlCompressionType.Level9 took {swMax.ElapsedMilliseconds} ms.");

            // Optional: compare file sizes
            long sizeLevel1 = new System.IO.FileInfo("output_Level1.xlsx").Length;
            long sizeLevel9 = new System.IO.FileInfo("output_Level9.xlsx").Length;

            Console.WriteLine($"File size with Level1 compression: {sizeLevel1} bytes.");
            Console.WriteLine($"File size with Level9 compression: {sizeLevel9} bytes.");
        }
    }
}
