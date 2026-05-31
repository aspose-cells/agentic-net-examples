using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Save with default compression (Level2 is default)
            string defaultPath = "DefaultCompression.xlsx";
            workbook.Save(defaultPath);

            // Save with OoxmlCompressionType.Level5
            string level5Path = "Level5Compression.xlsx";
            OoxmlSaveOptions level5Options = new OoxmlSaveOptions();
            level5Options.CompressionType = OoxmlCompressionType.Level5;
            workbook.Save(level5Path, level5Options);

            // Get file sizes
            long defaultSize = new FileInfo(defaultPath).Length;
            long level5Size = new FileInfo(level5Path).Length;

            // Output comparison
            Console.WriteLine($"Default compression file size: {defaultSize} bytes");
            Console.WriteLine($"Level5 compression file size: {level5Size} bytes");
            if (level5Size < defaultSize)
            {
                Console.WriteLine("Level5 compression produced a smaller file.");
            }
            else if (level5Size > defaultSize)
            {
                Console.WriteLine("Level5 compression produced a larger file.");
            }
            else
            {
                Console.WriteLine("Both files have the same size.");
            }
        }
    }
}