// Title: C# Example: Compare Aspose.Cells Default (Level2) and Level5 Ooxml Compression File Sizes
// Description: Shows how to set OoxmlCompressionType.Level5 with OoxmlSaveOptions, save a workbook, read the resulting file sizes, and display the byte reduction versus the default Level2 compression. Ideal for measuring storage savings and performance impact in .NET applications.
// Keywords: Aspose.Cells C# | OoxmlCompressionType Level5 | Level2 compression | OoxmlSaveOptions | Excel file size comparison | Aspose.Cells compression benchmark | C# workbook save options | file size reduction | performance impact | GitHub sample
// Common Searches: Aspose.Cells set OoxmlCompressionType Level5 C# | compare Excel file size default vs Level5 compression Aspose | how to measure Aspose.Cells workbook size after compression | C# code example OoxmlSaveOptions compression levels | Aspose.Cells compression performance benchmark
// Developer Intent: The developer wants to apply Level5 Ooxml compression when saving an Aspose.Cells workbook and compare the resulting file size to the default Level2 compression.
// Use Cases: Quantify storage savings for large generated reports. | Select the optimal compression level for batch export services that balance size and speed. | Validate that compressed workbooks meet email attachment or upload size limits. | Create automated tests that verify compression settings produce expected size reductions.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook with OoxmlCompressionType.Level5 and prints the size difference compared to the default compression. | Explain the trade‑offs between Level2 and Level5 Ooxml compression in Aspose.Cells regarding save time and file size. | Write a method that returns the percentage reduction when using Level5 compression versus the default level. | Provide a GitHub‑style README snippet describing how to run the compression comparison example.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionComparison
{
    // Shows how to set OoxmlCompressionType.Level5 with OoxmlSaveOptions, save a workbook, read the resulting file sizes, and display the byte reduction versus the default Level2 compression. Ideal for measuring storage savings and performance impact in .NET applications.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Save with default compression (Level2)
            string defaultPath = "default_compression.xlsx";
            workbook.Save(defaultPath); // uses default OoxmlSaveOptions

            // Save with Level5 compression
            string level5Path = "level5_compression.xlsx";
            OoxmlSaveOptions level5Options = new OoxmlSaveOptions();
            level5Options.CompressionType = OoxmlCompressionType.Level5;
            workbook.Save(level5Path, level5Options);

            // Get file sizes
            long defaultSize = new FileInfo(defaultPath).Length;
            long level5Size = new FileInfo(level5Path).Length;

            // Output comparison
            Console.WriteLine($"Default compression file size (Level2): {defaultSize} bytes");
            Console.WriteLine($"Level5 compression file size: {level5Size} bytes");
            Console.WriteLine($"Size reduction: {defaultSize - level5Size} bytes");
        }
    }
}
