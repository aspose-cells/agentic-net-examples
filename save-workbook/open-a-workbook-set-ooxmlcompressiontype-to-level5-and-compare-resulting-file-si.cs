using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with enough data to see a size difference
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // -----------------------------------------------------------------
            // Save with default compression (default is Level2)
            // -----------------------------------------------------------------
            string defaultPath = "DefaultCompression.xlsx";
            workbook.Save(defaultPath); // uses default OoxmlSaveOptions

            // Get file size for default compression
            long defaultSize = new FileInfo(defaultPath).Length;

            // -----------------------------------------------------------------
            // Save with OoxmlCompressionType.Level5
            // -----------------------------------------------------------------
            string level5Path = "Level5Compression.xlsx";

            // Create OoxmlSaveOptions and set the compression type to Level5
            OoxmlSaveOptions level5Options = new OoxmlSaveOptions();
            level5Options.CompressionType = OoxmlCompressionType.Level5;

            // Save the workbook using the specified compression options
            workbook.Save(level5Path, level5Options);

            // Get file size for Level5 compression
            long level5Size = new FileInfo(level5Path).Length;

            // -----------------------------------------------------------------
            // Compare and output the results
            // -----------------------------------------------------------------
            Console.WriteLine($"Default compression (Level2) file size: {defaultSize} bytes");
            Console.WriteLine($"Level5 compression file size: {level5Size} bytes");

            if (level5Size < defaultSize)
            {
                Console.WriteLine("Level5 compression produced a smaller file.");
            }
            else if (level5Size > defaultSize)
            {
                Console.WriteLine("Level5 compression produced a larger file (unlikely for this data).");
            }
            else
            {
                Console.WriteLine("Both files have the same size.");
            }
        }
    }
}