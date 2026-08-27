// Title: Compare default Ooxml compression (Level2) with Level5 compression when saving an Aspose.Cells workbook in C#
// AI Prompts: Write C# code that creates a workbook, saves it with OoxmlSaveOptions.CompressionType set to OoxmlCompressionType.Level5, and prints the resulting file size. | Extend the program to also save the workbook using the default compression, then calculate and display the byte difference between the two saved files.
// Common Searches: how to use OoxmlCompressionType.Level5 with Aspose.Cells in C# | measure Excel file size difference between Level2 and Level5 compression using Aspose.Cells | C# Aspose.Cells OoxmlSaveOptions compression options example | compare workbook file sizes after changing Ooxml compression level in .NET | Aspose.Cells reduce XLSX size by setting OoxmlCompressionType to Level5
// Tags: OoxmlSaveOptions compression Level5 | Aspose.Cells workbook size reduction | C# OoxmlCompressionType Level5 usage | Excel XLSX file size comparison Aspose.Cells | default Ooxml compression Level2 vs Level5

using System;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a workbook, saving it twice—once with default Ooxml compression and once with OoxmlCompressionType.Level5—then outputs both file sizes and the byte reduction achieved.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample data for compression test");
        worksheet.Cells["A2"].PutValue(12345);
        worksheet.Cells["A3"].PutValue(DateTime.Now);

        // Save the workbook using the default compression (Level2)
        string defaultFile = "DefaultCompression.xlsx";
        workbook.Save(defaultFile); // default OoxmlSaveOptions

        // Save the workbook with OoxmlCompressionType set to Level5
        string level5File = "Level5Compression.xlsx";
        OoxmlSaveOptions level5Options = new OoxmlSaveOptions();
        level5Options.CompressionType = OoxmlCompressionType.Level5;
        workbook.Save(level5File, level5Options);

        // Compare the file sizes
        long defaultSize = new FileInfo(defaultFile).Length;
        long level5Size = new FileInfo(level5File).Length;

        Console.WriteLine($"Default compression file size: {defaultSize} bytes");
        Console.WriteLine($"Level5 compression file size: {level5Size} bytes");
        Console.WriteLine($"Size reduction: {defaultSize - level5Size} bytes");
    }
}
