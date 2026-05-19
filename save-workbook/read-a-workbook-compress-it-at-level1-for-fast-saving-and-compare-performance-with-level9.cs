using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Prepare OoxmlSaveOptions for Level1 compression (fastest)
            OoxmlSaveOptions optionsLevel1 = new OoxmlSaveOptions();
            optionsLevel1.CompressionType = OoxmlCompressionType.Level1;

            // Measure time for saving with Level1
            Stopwatch sw = Stopwatch.StartNew();
            workbook.Save("output_level1.xlsx", optionsLevel1);
            sw.Stop();
            Console.WriteLine($"Saving with OoxmlCompressionType.Level1 took {sw.ElapsedMilliseconds} ms.");

            // Prepare OoxmlSaveOptions for Level9 compression (best compression, slower)
            OoxmlSaveOptions optionsLevel9 = new OoxmlSaveOptions();
            optionsLevel9.CompressionType = OoxmlCompressionType.Level9;

            // Measure time for saving with Level9
            sw.Restart();
            workbook.Save("output_level9.xlsx", optionsLevel9);
            sw.Stop();
            Console.WriteLine($"Saving with OoxmlCompressionType.Level9 took {sw.ElapsedMilliseconds} ms.");
        }
    }
}