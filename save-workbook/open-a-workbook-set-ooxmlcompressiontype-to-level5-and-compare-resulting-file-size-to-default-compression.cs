using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionComparison
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with an actual file if needed)
            string sourcePath = "source.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // -------------------------------------------------
            // Save with default compression (no OoxmlSaveOptions)
            // -------------------------------------------------
            string defaultCompressedPath = "default_compressed.xlsx";
            workbook.Save(defaultCompressedPath); // uses default OoxmlCompressionType.Level2

            // Get file size of the default compressed file
            long defaultSize = new FileInfo(defaultCompressedPath).Length;

            // -------------------------------------------------
            // Save with OoxmlCompressionType.Level5
            // -------------------------------------------------
            string level5CompressedPath = "level5_compressed.xlsx";

            // Create OoxmlSaveOptions and set the compression type to Level5
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.CompressionType = OoxmlCompressionType.Level5;

            // Save the workbook using the specified options
            workbook.Save(level5CompressedPath, saveOptions);

            // Get file size of the Level5 compressed file
            long level5Size = new FileInfo(level5CompressedPath).Length;

            // -------------------------------------------------
            // Compare the file sizes and output the result
            // -------------------------------------------------
            Console.WriteLine($"Default compression (Level2) file size: {defaultSize} bytes");
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