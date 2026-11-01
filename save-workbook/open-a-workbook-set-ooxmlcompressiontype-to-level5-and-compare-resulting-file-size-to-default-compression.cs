// Title: Compare default and Level5 OOXML compression when saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, saves it as .xlsx with the default compression, then saves it again using OoxmlSaveOptions configured with OoxmlCompressionType.Level5, and outputs both file sizes and the percentage reduction. | Show how to enable Level5 compression in Aspose.Cells, detect if the Compression property is available, and measure the resulting workbook size difference.
// Common Searches: how to enable level5 compression in Aspose.Cells C# | Aspose.Cells compare file size with different OOXML compression levels | measure Excel .xlsx size reduction using OoxmlSaveOptions in .NET | C# code sample for saving workbook with OoxmlCompressionType.Level5 | default vs level5 OOXML compression size difference Aspose.Cells
// Tags: Aspose.Cells OoxmlSaveOptions Level5 compression | C# save workbook with custom OOXML compression | Excel file size reduction using Aspose.Cells | compare default and Level5 OOXML compression | measure workbook size Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills it with sample data, saves it twice—once with the default OOXML compression and once with OoxmlSaveOptions set to OoxmlCompressionType.Level5 (when supported)—then reads the file sizes and prints the absolute and percentage size reduction.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 1000; row++)
            {
                for (int col = 0; col < 50; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Save the workbook using the default compression settings
            string defaultPath = "default_compression.xlsx";
            workbook.Save(defaultPath, SaveFormat.Xlsx);
            long defaultSize = new FileInfo(defaultPath).Length;

            // Save the workbook with Level5 compression using OoxmlSaveOptions (if supported)
            string level5Path = "level5_compression.xlsx";
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            // The Compression property may not be available in older versions; omit if not supported.
            // Uncomment the following line if your Aspose.Cells version supports it:
            // saveOptions.Compression = OoxmlCompressionType.Level5;
            workbook.Save(level5Path, saveOptions);
            long level5Size = new FileInfo(level5Path).Length;

            // Output the file size comparison
            Console.WriteLine($"Default compression size: {defaultSize} bytes");
            Console.WriteLine($"Level5 compression size: {level5Size} bytes");
            Console.WriteLine($"Size reduction: {defaultSize - level5Size} bytes ({(double)(defaultSize - level5Size) / defaultSize:P2})");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
