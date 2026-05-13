using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCompressionComparison
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Prepare save options for Level1 compression (fastest)
            OoxmlSaveOptions level1Options = new OoxmlSaveOptions();
            level1Options.CompressionType = OoxmlCompressionType.Level1;

            // Prepare save options for Level9 compression (best compression, slowest)
            OoxmlSaveOptions level9Options = new OoxmlSaveOptions();
            level9Options.CompressionType = OoxmlCompressionType.Level9;

            // Measure saving time with Level1 compression
            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save("output_Level1.xlsx", level1Options);
            sw.Stop();
            Console.WriteLine($"Level1 compression saved in {sw.ElapsedMilliseconds} ms.");

            // Measure saving time with Level9 compression
            sw.Restart();
            workbook.Save("output_Level9.xlsx", level9Options);
            sw.Stop();
            Console.WriteLine($"Level9 compression saved in {sw.ElapsedMilliseconds} ms.");
        }
    }
}