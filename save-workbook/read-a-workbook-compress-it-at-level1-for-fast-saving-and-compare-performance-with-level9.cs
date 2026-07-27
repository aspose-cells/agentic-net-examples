using System;
using System.Diagnostics;
using Aspose.Cells;

namespace CompressionPerformanceDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // -------------------- Save with Level1 (fastest) --------------------
            // Create OoxmlSaveOptions and set compression to Level1
            OoxmlSaveOptions level1Options = new OoxmlSaveOptions();
            level1Options.CompressionType = OoxmlCompressionType.Level1;

            // Measure the time taken to save with Level1 compression
            Stopwatch stopwatch = Stopwatch.StartNew();
            workbook.Save("output_Level1.xlsx", level1Options);
            stopwatch.Stop();
            Console.WriteLine($"Level1 compression save time: {stopwatch.ElapsedMilliseconds} ms");

            // -------------------- Save with Level9 (best compression) --------------------
            // Create OoxmlSaveOptions and set compression to Level9
            OoxmlSaveOptions level9Options = new OoxmlSaveOptions();
            level9Options.CompressionType = OoxmlCompressionType.Level9;

            // Measure the time taken to save with Level9 compression
            stopwatch.Restart();
            workbook.Save("output_Level9.xlsx", level9Options);
            stopwatch.Stop();
            Console.WriteLine($"Level9 compression save time: {stopwatch.ElapsedMilliseconds} ms");
        }
    }
}